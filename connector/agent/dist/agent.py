import json
import subprocess
import socket
import time
import urllib.error
import urllib.request
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
CFG_PATH = BASE_DIR / "agent.json"
VERSION = "0.1.0"


def load_cfg() -> dict:
    if not CFG_PATH.exists():
        raise RuntimeError("Missing agent.json (copy agent.example.json)")
    return json.loads(CFG_PATH.read_text(encoding="utf-8-sig"))


def decrypt_token(path: Path) -> str:
    if not path.exists():
        raise RuntimeError(f"Encrypted token file not found: {path}")

    ps = (
        "$s = Get-Content -Path '"
        + str(path).replace("'", "''")
        + "' | ConvertTo-SecureString;"
        "$b=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($s);"
        "[Runtime.InteropServices.Marshal]::PtrToStringBSTR($b)"
    )
    run = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps],
        capture_output=True,
        text=True,
        timeout=15,
    )
    if run.returncode != 0:
        raise RuntimeError(f"Token decrypt failed: {run.stderr.strip()}")
    token = run.stdout.strip()
    if not token:
        raise RuntimeError("Token decrypt returned empty value")
    return token


def get_token(cfg: dict) -> str:
    inline = cfg.get("device_token", "").strip()
    if inline:
        return inline

    enc_path = cfg.get("token_encrypted_path", "").strip()
    if not enc_path:
        raise RuntimeError("Missing token: set device_token or token_encrypted_path")

    return decrypt_token(Path(enc_path))


def public_ip() -> str:
    with urllib.request.urlopen("https://api.ipify.org", timeout=10) as r:
        return r.read().decode("utf-8").strip()


def heartbeat(cfg: dict) -> None:
    payload = {
        "device_id": cfg["device_id"],
        "public_ip": public_ip(),
        "hostname": socket.gethostname(),
        "agent_version": VERSION,
    }
    headers = {"X-Device-Token": get_token(cfg), "Content-Type": "application/json"}
    url = cfg["server_url"].rstrip("/") + "/heartbeat"
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=15) as r:
        if r.status >= 400:
            raise RuntimeError(f"HTTP {r.status}")


def main() -> None:
    cfg = load_cfg()
    interval = int(cfg.get("heartbeat_seconds", 60))
    while True:
        try:
            heartbeat(cfg)
            print("heartbeat ok")
        except (urllib.error.URLError, RuntimeError, OSError, ValueError) as e:
            print(f"heartbeat error: {e}")
        time.sleep(interval)


if __name__ == "__main__":
    main()
