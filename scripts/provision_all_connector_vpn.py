"""Ensure every active Connector token has a complete VPN peer/config.

Run on the Connector server with CONNECTOR_CONFIG_PATH, CONNECTOR_DB_PATH, and
CONNECTOR_DB_URL loaded from the production runtime environment. The server
module performs the actual key/IP allocation, live peer registration, and
persistent awgserver.conf update.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


SERVER_SOURCE = Path(r"C:\Connector\src\connector\server")
if str(SERVER_SOURCE) not in sys.path:
    sys.path.insert(0, str(SERVER_SOURCE))

import app  # noqa: E402
import vpn  # noqa: E402


def main() -> None:
    config = app.load_config()
    if not vpn.is_enabled(config):
        raise SystemExit("VPN is disabled in Connector server config")

    with app.db_connect() as connection:
        rows = connection.execute(
            """
            SELECT device_id
            FROM device_tokens
            WHERE revoked_at IS NULL
            ORDER BY device_id
            """
        ).fetchall()

    provisioned = []
    failed = []
    for (device_id,) in rows:
        try:
            record = vpn.get_or_create_device_vpn(
                device_id,
                config,
                app.db_connect,
            )
            provisioned.append(
                {
                    "device_id": device_id,
                    "vpn_address": record["vpn_address"],
                }
            )
        except Exception as exc:  # report all devices in one run
            failed.append({"device_id": device_id, "error": str(exc)})

    print(
        json.dumps(
            {
                "active_count": len(rows),
                "provisioned": provisioned,
                "failed": failed,
            },
            ensure_ascii=False,
        )
    )
    if failed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
