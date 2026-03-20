from __future__ import annotations

import types


def test_secrets_service_no_keyring(monkeypatch):
    import app.services.secrets_service as secrets_mod
    from app.services.secrets_service import SecretsService

    monkeypatch.setattr(secrets_mod, "keyring", None, raising=True)
    svc = SecretsService()
    assert svc.available is False
    # Should not raise
    ref = svc.vpn_password_ref("VPN Demo")
    svc.set_secret(ref, "pwd")
    assert svc.get_secret(ref) is None
    svc.delete_secret(ref)


def test_secrets_service_with_keyring_stub(monkeypatch):
    import app.services.secrets_service as secrets_mod
    from app.services.secrets_service import SecretsService

    store = {}

    def set_password(service, key, value):
        store[(service, key)] = value

    def get_password(service, key):
        return store.get((service, key))

    def delete_password(service, key):
        store.pop((service, key), None)

    stub = types.SimpleNamespace(
        set_password=set_password, get_password=get_password, delete_password=delete_password
    )

    monkeypatch.setattr(secrets_mod, "keyring", stub, raising=True)
    svc = SecretsService(namespace="TestNS")
    assert svc.available is True

    ref = svc.vpn_password_ref("VPN Demo")
    svc.set_secret(ref, "pwd1")
    assert svc.get_secret(ref) == "pwd1"
    svc.delete_secret(ref)
    assert svc.get_secret(ref) is None

