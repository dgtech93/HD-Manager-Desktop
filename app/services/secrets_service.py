from __future__ import annotations

from dataclasses import dataclass


try:
    import keyring  # type: ignore
except Exception:  # pragma: no cover
    keyring = None


@dataclass(frozen=True, slots=True)
class SecretRef:
    service: str
    key: str

    def serialize(self) -> str:
        return f"{self.service}:{self.key}"


class SecretsService:
    """Stores secrets in OS credential manager via `keyring`.

    Falls back gracefully if keyring isn't available.
    """

    def __init__(self, namespace: str = "HDManagerDesktop") -> None:
        self.namespace = namespace

    @property
    def available(self) -> bool:
        return keyring is not None

    def set_secret(self, ref: SecretRef, value: str) -> None:
        if not self.available:
            return
        keyring.set_password(ref.service, ref.key, value)

    def get_secret(self, ref: SecretRef) -> str | None:
        if not self.available:
            return None
        return keyring.get_password(ref.service, ref.key)

    def delete_secret(self, ref: SecretRef) -> None:
        if not self.available:
            return
        try:
            keyring.delete_password(ref.service, ref.key)
        except Exception:
            # If it's missing, ignore.
            return

    # Helpers to build stable refs
    def vpn_password_ref(self, connection_name: str) -> SecretRef:
        return SecretRef(self.namespace, f"vpn:{connection_name.strip()}")

    def credential_password_ref(self, client_id: int, product_id: int, credential_name: str) -> SecretRef:
        return SecretRef(self.namespace, f"cred:{int(client_id)}:{int(product_id)}:{credential_name.strip()}")

