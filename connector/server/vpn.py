"""Self-hosted AmneziaWG VPN provisioning for the connector.

Gives remote devices (whose ISP blocks SMB port 445) access to the firm SMB share
over a split-tunnel AmneziaWG VPN. The connector server runs as SYSTEM, so it can
register peers on the live AmneziaWG tunnel service (`awg set`) and persist them to
the server .conf — the same privileged-subprocess model already used for SMB users
and firewall rules.

Design notes:
  * Entirely ADDITIVE and behind cfg["vpn"]["enabled"] (default off). If anything
    here fails, the bootstrap handler swallows it so SMB/heartbeat are unaffected.
  * Server-generated keypair per device (trusted firm env, delivered over HTTPS +
    device token). Stored in device_vpn so re-bootstrap returns the same config.
  * Peer is added LIVE via `awg set` (does not disconnect other peers) AND appended
    to the server .conf so it survives a tunnel-service/host restart.
  * Split-tunnel: client AllowedIPs always contains the VPN subnet and may contain
    explicitly configured server routes. This lets existing UNC paths keep using
    the canonical server IP without routing unrelated internet traffic into VPN.
"""

import ipaddress
import subprocess
import threading
from datetime import datetime, timezone

# Serializes the allocate -> register -> persist -> save sequence (and conf rewrites) so two
# concurrent first-time bootstraps cannot grab the same VPN IP or interleave conf appends.
_vpn_lock = threading.Lock()

# AWG 1.5 obfuscation keys that the deployed server tunnel uses; all must be present and match.
_REQUIRED_OBF = ("Jc", "Jmin", "Jmax", "S1", "S2", "H1", "H2", "H3", "H4")


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def get_vpn_cfg(cfg: dict) -> dict:
    v = cfg.get("vpn")
    return v if isinstance(v, dict) else {}


def is_enabled(cfg: dict) -> bool:
    return bool(get_vpn_cfg(cfg).get("enabled", False))


def _awg_exe(vcfg: dict) -> str:
    return str(vcfg.get("awg_exe", r"C:\Program Files\AmneziaWG\awg.exe"))


def _run(args: list[str], stdin: str | None = None) -> str:
    run = subprocess.run(args, input=stdin, capture_output=True, text=True)
    if run.returncode != 0:
        raise RuntimeError((run.stderr or run.stdout or "command failed").strip())
    return run.stdout.strip()


def _gen_keypair(vcfg: dict) -> tuple[str, str]:
    awg = _awg_exe(vcfg)
    priv = _run([awg, "genkey"])
    pub = _run([awg, "pubkey"], stdin=priv)
    return priv, pub


def _obfuscation_lines(vcfg: dict) -> str:
    obf = vcfg.get("obfuscation", {})
    order = ["Jc", "Jmin", "Jmax", "S1", "S2", "S3", "S4", "H1", "H2", "H3", "H4", "I1", "Itime"]
    lines = []
    for k in order:
        if k in obf and obf[k] not in (None, ""):
            lines.append(f"{k} = {obf[k]}")
    return "\n".join(lines)


# ---- DB layer (mirrors device_access) -------------------------------------------------

def get_device_vpn(device_id: str, conn_factory) -> dict | None:
    with conn_factory() as conn:
        row = conn.execute(
            "SELECT public_key, private_key, vpn_address, config FROM device_vpn WHERE device_id = ?",
            (device_id,),
        ).fetchone()
    if not row:
        return None
    return {"public_key": row[0], "private_key": row[1], "vpn_address": row[2], "config": row[3]}


def save_device_vpn(device_id: str, rec: dict, conn_factory) -> None:
    with conn_factory() as conn:
        conn.execute(
            """
            INSERT INTO device_vpn(device_id, public_key, private_key, vpn_address, config, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(device_id) DO UPDATE SET
                public_key=excluded.public_key,
                private_key=excluded.private_key,
                vpn_address=excluded.vpn_address,
                config=excluded.config,
                updated_at=excluded.updated_at
            """,
            (
                device_id,
                str(rec["public_key"]),
                str(rec["private_key"]),
                str(rec["vpn_address"]),
                str(rec["config"]),
                _utc_now(),
            ),
        )


def _used_addresses(conn_factory) -> set[str]:
    with conn_factory() as conn:
        rows = conn.execute("SELECT vpn_address FROM device_vpn").fetchall()
    return {str(r[0]) for r in rows if r and r[0]}


def _allocate_ip(vcfg: dict, conn_factory) -> str:
    """Next free host address in the VPN subnet, skipping the server's own address."""
    subnet = ipaddress.ip_network(str(vcfg.get("subnet", "10.77.123.0/24")), strict=False)
    server_ip = str(vcfg.get("server_address", "10.77.123.1"))
    used = _used_addresses(conn_factory)
    used.add(server_ip)
    for host in subnet.hosts():
        cand = str(host)
        if cand not in used:
            return cand
    raise RuntimeError("VPN subnet exhausted")


# ---- peer registration (privileged; server runs as SYSTEM) ---------------------------

def _register_peer_live(vcfg: dict, public_key: str, address: str) -> None:
    awg = _awg_exe(vcfg)
    tunnel = str(vcfg.get("tunnel_name", "awgserver"))
    _run([awg, "set", tunnel, "peer", public_key, "allowed-ips", f"{address}/32"])


def _persist_peer_to_conf(vcfg: dict, device_id: str, public_key: str, address: str) -> None:
    conf_path = str(vcfg.get("server_conf", r"C:\awg\awgserver.conf"))
    try:
        with open(conf_path, "r", encoding="utf-8") as f:
            content = f.read()
    except OSError:
        return  # live `awg set` already applied; conf persistence is best-effort
    if public_key in content:
        return
    block = f"\n[Peer]\n# device {device_id}\nPublicKey = {public_key}\nAllowedIPs = {address}/32\n"
    with open(conf_path, "a", encoding="utf-8") as f:
        f.write(block)


def _remove_peer_live(vcfg: dict, public_key: str) -> None:
    awg = _awg_exe(vcfg)
    tunnel = str(vcfg.get("tunnel_name", "awgserver"))
    _run([awg, "set", tunnel, "peer", public_key, "remove"])


def _strip_peer_from_conf(vcfg: dict, public_key: str) -> None:
    """Rewrite the server .conf with the [Peer] block whose PublicKey matches removed."""
    conf_path = str(vcfg.get("server_conf", r"C:\awg\awgserver.conf"))
    try:
        with open(conf_path, "r", encoding="utf-8") as f:
            text = f.read()
    except OSError:
        return
    if public_key not in text:
        return
    # split into the [Interface] head + [Peer] blocks; keep all but the matching peer
    chunks = text.split("\n[Peer]")
    head = chunks[0]
    kept = [c for c in chunks[1:] if public_key not in c]
    rebuilt = head + "".join("\n[Peer]" + c for c in kept)
    if not rebuilt.endswith("\n"):
        rebuilt += "\n"
    with open(conf_path, "w", encoding="utf-8") as f:
        f.write(rebuilt)


def _validate_vpn_cfg(vcfg: dict) -> None:
    """Fail loudly server-side rather than shipping a silent non-handshaking client .conf."""
    missing = []
    if not str(vcfg.get("server_public_key", "")).strip():
        missing.append("server_public_key")
    if not str(vcfg.get("endpoint_host", "")).strip():
        missing.append("endpoint_host")
    obf = vcfg.get("obfuscation", {})
    for k in _REQUIRED_OBF:
        if k not in obf or obf[k] in (None, ""):
            missing.append(f"obfuscation.{k}")
    if missing:
        raise RuntimeError("vpn config incomplete: missing " + ", ".join(missing))


# ---- client config assembly -----------------------------------------------------------

def _routed_ips_for_device(vcfg: dict, device_id: str) -> list[str]:
    raw_allowlist = vcfg.get("routed_ips_device_allowlist")
    if isinstance(raw_allowlist, list) and raw_allowlist:
        allowed_devices = {str(item).strip() for item in raw_allowlist if str(item).strip()}
        if device_id not in allowed_devices:
            return []

    raw_routes = vcfg.get("routed_ips", [])
    if not isinstance(raw_routes, list):
        raise RuntimeError("vpn.routed_ips must be a list")

    routes: list[str] = []
    for raw in raw_routes:
        value = str(raw).strip()
        if not value:
            continue
        try:
            normalized = str(ipaddress.ip_network(value, strict=False))
        except ValueError as exc:
            raise RuntimeError(f"invalid vpn.routed_ips entry: {value}") from exc
        if normalized not in routes:
            routes.append(normalized)
    return routes


def _client_allowed_ips(vcfg: dict, device_id: str) -> list[str]:
    subnet = str(ipaddress.ip_network(str(vcfg.get("subnet", "10.77.123.0/24")), strict=False))
    return [subnet, *[route for route in _routed_ips_for_device(vcfg, device_id) if route != subnet]]


def _build_client_conf(vcfg: dict, client_priv: str, address: str, device_id: str) -> str:
    endpoint_host = str(vcfg.get("endpoint_host", ""))
    listen_port = int(vcfg.get("listen_port", 9994))
    server_pub = str(vcfg.get("server_public_key", ""))
    obf = _obfuscation_lines(vcfg)
    parts = [
        "[Interface]",
        f"PrivateKey = {client_priv}",
        f"Address = {address}/32",
    ]
    if obf:
        parts.append(obf)
    parts += [
        "",
        "[Peer]",
        f"PublicKey = {server_pub}",
        f"Endpoint = {endpoint_host}:{listen_port}",
        f"AllowedIPs = {', '.join(_client_allowed_ips(vcfg, device_id))}",
        "PersistentKeepalive = 25",
        "",
    ]
    return "\n".join(parts)


# ---- public API -----------------------------------------------------------------------

def provision_vpn(device_id: str, cfg: dict, conn_factory) -> dict:
    """Create a peer for the device: keypair + IP + live registration + persisted conf.

    Serialized by _vpn_lock so concurrent first-time bootstraps cannot collide on an IP.
    """
    vcfg = get_vpn_cfg(cfg)
    _validate_vpn_cfg(vcfg)
    priv, pub = _gen_keypair(vcfg)
    with _vpn_lock:
        address = _allocate_ip(vcfg, conn_factory)
        # persist the DB row FIRST so a later step failing can't leave a live peer with no record
        rec = {"public_key": pub, "private_key": priv, "vpn_address": address,
               "config": _build_client_conf(vcfg, priv, address, device_id)}
        save_device_vpn(device_id, rec, conn_factory)
        _register_peer_live(vcfg, pub, address)
        _persist_peer_to_conf(vcfg, device_id, pub, address)
    return rec


def get_or_create_device_vpn(device_id: str, cfg: dict, conn_factory, force_rotate: bool = False) -> dict:
    if not force_rotate:
        existing = get_device_vpn(device_id, conn_factory)
        if existing and existing.get("config"):
            vcfg = get_vpn_cfg(cfg)
            refreshed_config = _build_client_conf(
                vcfg,
                existing["private_key"],
                existing["vpn_address"],
                device_id,
            )
            if refreshed_config != existing["config"]:
                existing = {**existing, "config": refreshed_config}
                save_device_vpn(device_id, existing, conn_factory)
            # re-assert the live peer AND the conf entry (idempotent) in case the tunnel/conf was rebuilt
            try:
                _register_peer_live(vcfg, existing["public_key"], existing["vpn_address"])
                _persist_peer_to_conf(vcfg, device_id, existing["public_key"], existing["vpn_address"])
            except Exception:
                pass
            return existing
    return provision_vpn(device_id, cfg, conn_factory)


def deprovision_device_vpn(device_id: str, cfg: dict, conn_factory) -> bool:
    """Remove a device's VPN peer (live + conf) and DB row. Best-effort; never raises.

    Call from token revoke/delete so a revoked device cannot keep dialing the tunnel.
    """
    try:
        rec = get_device_vpn(device_id, conn_factory)
        if not rec:
            return False
        vcfg = get_vpn_cfg(cfg)
        with _vpn_lock:
            try:
                _remove_peer_live(vcfg, rec["public_key"])
            except Exception:
                pass
            try:
                _strip_peer_from_conf(vcfg, rec["public_key"])
            except Exception:
                pass
            with conn_factory() as conn:
                conn.execute("DELETE FROM device_vpn WHERE device_id = ?", (device_id,))
        return True
    except Exception:
        return False


def is_enabled_for_device(cfg: dict, device_id: str) -> bool:
    """Global enable AND (optional) per-device allowlist for staged rollout.

    If vpn.device_allowlist is a non-empty list, only those device_ids get VPN (canary);
    remove/empty the list to enable for everyone.
    """
    if not is_enabled(cfg):
        return False
    allow = get_vpn_cfg(cfg).get("device_allowlist")
    if isinstance(allow, list) and len(allow) > 0:
        return device_id in allow
    return True


def vpn_bundle_for_bootstrap(device_id: str, cfg: dict, conn_factory) -> dict:
    """Shape returned in the bootstrap response. Never raises into the caller's happy path."""
    if not is_enabled_for_device(cfg, device_id):
        return {"enabled": False}
    vcfg = get_vpn_cfg(cfg)
    rec = get_or_create_device_vpn(device_id, cfg, conn_factory)
    routed_ips = _routed_ips_for_device(vcfg, device_id)
    routed_smb_unc = str(vcfg.get("routed_smb_unc", "")).strip()
    smb_unc = routed_smb_unc if routed_ips and routed_smb_unc else str(vcfg.get("smb_unc", ""))
    return {
        "enabled": True,
        "tunnel_name": str(vcfg.get("tunnel_name", "awgserver")),
        "address": rec["vpn_address"],
        "config": rec["config"],
        "smb_unc": smb_unc,
        "server_vpn_ip": str(vcfg.get("server_address", "10.77.123.1")),
    }
