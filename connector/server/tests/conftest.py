"""Shared pytest fixtures.

Изолируем app под временные DB и config через CONNECTOR_DB_PATH /
CONNECTOR_CONFIG_PATH (они же — путь к проду в нормальном запуске).
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pytest


SERVER_DIR = Path(__file__).resolve().parent.parent

if str(SERVER_DIR) not in sys.path:
    sys.path.insert(0, str(SERVER_DIR))


def _build_test_config(token_export_dir: Path) -> dict:
    return {
        "api_bind": "127.0.0.1",
        "api_port": 8080,
        "admin_api_key": "test-admin-key",
        "admin_username": "test_admin",
        "admin_password": "test-password",
        "managed_ports": [],
        "allowed_stale_minutes": 15,
        "default_heartbeat_seconds": 60,
        "smb_server_host": "127.0.0.1",
        "smb_share_name": "BIM_Models",
        "smb_share_path": str(token_export_dir),
        "smb_user_prefix": "bim_",
        "token_export_dir": str(token_export_dir),
        "devices": {},
    }


@pytest.fixture(scope="session")
def runtime_paths(tmp_path_factory):
    runtime = tmp_path_factory.mktemp("runtime")
    db_path = runtime / "connector.db"
    config_path = runtime / "config.json"
    token_export = runtime / "tokens"
    token_export.mkdir(parents=True, exist_ok=True)

    config_path.write_text(
        json.dumps(_build_test_config(token_export), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    os.environ["CONNECTOR_CONFIG_PATH"] = str(config_path)
    os.environ["CONNECTOR_DB_PATH"] = str(db_path)

    yield {
        "runtime": runtime,
        "db": db_path,
        "config": config_path,
    }

    for var in ("CONNECTOR_CONFIG_PATH", "CONNECTOR_DB_PATH"):
        os.environ.pop(var, None)


@pytest.fixture(scope="session")
def app_module(runtime_paths):
    """Импортирует app один раз после установки env vars."""
    import importlib

    if "app" in sys.modules:
        del sys.modules["app"]
    return importlib.import_module("app")


@pytest.fixture()
def client(app_module):
    from fastapi.testclient import TestClient

    # Enter the context so FastAPI lifespan runs exactly as it does in
    # production (database migrations and config initialization).
    with TestClient(app_module.app) as test_client:
        yield test_client
