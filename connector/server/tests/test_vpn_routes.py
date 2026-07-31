import vpn


def _config(**overrides):
    value = {
        "endpoint_host": "62.113.36.107",
        "listen_port": 9994,
        "subnet": "10.77.123.0/24",
        "server_public_key": "server-public-key",
        "routed_ips": ["62.113.36.107/32"],
        "routed_ips_device_allowlist": ["canary-device"],
        "obfuscation": {},
    }
    value.update(overrides)
    return value


def test_client_config_adds_canonical_server_route_for_canary():
    config = vpn._build_client_conf(_config(), "private-key", "10.77.123.2", "canary-device")

    assert "AllowedIPs = 10.77.123.0/24, 62.113.36.107/32" in config


def test_client_config_keeps_vpn_subnet_only_for_other_devices():
    config = vpn._build_client_conf(_config(), "private-key", "10.77.123.3", "other-device")

    assert "AllowedIPs = 10.77.123.0/24" in config
    assert "62.113.36.107/32" not in config


def test_routed_ips_are_normalized_and_deduplicated():
    config = _config(
        routed_ips=["62.113.36.107", "62.113.36.107/32"],
        routed_ips_device_allowlist=[],
    )

    assert vpn._client_allowed_ips(config, "any-device") == [
        "10.77.123.0/24",
        "62.113.36.107/32",
    ]
