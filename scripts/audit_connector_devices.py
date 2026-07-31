"""Print a secret-free readiness report for active Connector devices.

The script expects CONNECTOR_DB_URL in the environment. It deliberately emits
only presence flags and non-secret routing/version metadata so the report can
be attached to diagnostics without exposing device, SMB, web, or VPN secrets.
"""

from __future__ import annotations

import json
import os

import psycopg2


QUERY = """
SELECT t.device_id AS device_id,
       COALESCE(t.issued_to, '') AS issued_to,
       t.created_at AS token_created_at,
       t.last_used_at AS token_last_used_at,
       CASE WHEN COALESCE(t.token_value, '') <> '' THEN 1 ELSE 0 END AS has_token,
       CASE WHEN a.device_id IS NOT NULL THEN 1 ELSE 0 END AS has_smb,
       COALESCE(a.smb_share_unc, '') AS smb_share_unc,
       CASE
           WHEN COALESCE(a.smb_login, '') <> ''
            AND COALESCE(a.smb_password, '') <> ''
           THEN 1 ELSE 0
       END AS smb_complete,
       CASE WHEN v.device_id IS NOT NULL THEN 1 ELSE 0 END AS has_vpn,
       COALESCE(v.vpn_address, '') AS vpn_address,
       CASE
           WHEN COALESCE(v.config, '') <> ''
            AND COALESCE(v.public_key, '') <> ''
            AND COALESCE(v.private_key, '') <> ''
           THEN 1 ELSE 0
       END AS vpn_complete,
       CASE
           WHEN COALESCE(v.config, '') LIKE '%62.113.36.107/32%'
           THEN 1 ELSE 0
       END AS canonical_server_route,
       COALESCE(d.hostname, '') AS hostname,
       COALESCE(d.agent_version, '') AS agent_version,
       COALESCE(d.public_ip, '') AS last_seen_ip,
       d.updated_at AS device_updated_at,
       CASE WHEN COALESCE(w.speckle_url, '') <> '' THEN 1 ELSE 0 END AS has_speckle_url,
       CASE WHEN COALESCE(w.nextcloud_url, '') <> '' THEN 1 ELSE 0 END AS has_nextcloud_url,
       CASE
           WHEN COALESCE(w.speckle_login, '') <> ''
            AND COALESCE(w.speckle_password, '') <> ''
           THEN 1 ELSE 0
       END AS speckle_credentials_complete,
       CASE
           WHEN COALESCE(w.nextcloud_login, '') <> ''
            AND COALESCE(w.nextcloud_password, '') <> ''
           THEN 1 ELSE 0
       END AS nextcloud_credentials_complete
FROM device_tokens t
LEFT JOIN device_access a ON a.device_id = t.device_id
LEFT JOIN device_vpn v ON v.device_id = t.device_id
LEFT JOIN devices d ON d.device_id = t.device_id
LEFT JOIN device_web_access w ON w.device_id = t.device_id
WHERE t.revoked_at IS NULL
ORDER BY t.device_id
"""


def main() -> None:
    db_url = os.environ.get("CONNECTOR_DB_URL", "").strip()
    if not db_url:
        raise SystemExit("CONNECTOR_DB_URL is not set")

    with psycopg2.connect(db_url) as connection:
        with connection.cursor() as cursor:
            cursor.execute(QUERY)
            columns = [description[0] for description in cursor.description]
            devices = [
                dict(zip(columns, row, strict=True))
                for row in cursor.fetchall()
            ]

            cursor.execute(
                "SELECT COUNT(*) FROM device_tokens WHERE revoked_at IS NOT NULL"
            )
            revoked_count = cursor.fetchone()[0]

    print(
        json.dumps(
            {
                "active_devices": devices,
                "active_count": len(devices),
                "revoked_count": revoked_count,
            },
            ensure_ascii=False,
            default=str,
        )
    )


if __name__ == "__main__":
    main()
