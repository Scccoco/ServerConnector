"""Authorization and checksum regression tests."""


def test_devices_requires_admin_auth(client):
    response = client.get("/devices")
    assert response.status_code == 401


def test_devices_with_admin_key(client):
    response = client.get("/devices", headers={"X-Admin-Key": "test-admin-key"})
    assert response.status_code == 200
    assert "items" in response.json()


def test_normalize_sha256_digest(app_module):
    digest = "A" * 64
    assert app_module.normalize_sha256_digest(digest) == digest.lower()
    assert app_module.normalize_sha256_digest(f"sha256:{digest}") == digest.lower()
    assert app_module.normalize_sha256_digest("not-a-digest") == ""
