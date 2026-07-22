"""Self-signed TLS material for the loopback job-intake listener.

Office requires add-in content to be served over HTTPS even on localhost, and
it only accepts a self-signed certificate whose *root CA* is installed in the
Windows "Trusted Root Certification Authorities" store. So this generates a
local root CA once, then a leaf certificate signed by it.

The leaf deliberately carries **both** `localhost` (DNS) and `127.0.0.1` (IP)
in its SAN: Office accepts either hostname in a manifest's SourceLocation, and
a cert issued for only one of them fails when the other spelling is used. This
is a known, easy-to-miss footgun (OfficeDev/Office-Addin-Scripts#514).

This module is self-contained on purpose - master_app has a similar `web_tls`
for its LAN web companion, but master_app now embeds *this* repo, so reaching
back into it would be a dependency cycle.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
import hashlib
from ipaddress import ip_address
from pathlib import Path

from paths import APP_DIR


CERT_DIR = APP_DIR / "_runtime" / "certs"
CA_CERT_PATH = CERT_DIR / "odd_job_intake_root_ca.crt.pem"
CA_KEY_PATH = CERT_DIR / "odd_job_intake_root_ca.key.pem"
CERT_PATH = CERT_DIR / "odd_job_intake.crt.pem"
KEY_PATH = CERT_DIR / "odd_job_intake.key.pem"

# Loopback only - this listener is never bound to a routable interface.
LOOPBACK_HOSTS = ("localhost", "127.0.0.1")

CA_COMMON_NAME = "Odd Job Intake Local Root CA"
LEAF_COMMON_NAME = "Odd Job Intake Loopback"
ORGANIZATION_NAME = "Battleshield Internal"


class CertificateError(RuntimeError):
    pass


def _crypto():
    try:
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID
    except ImportError as exc:  # pragma: no cover - dependency is declared
        raise CertificateError(
            "HTTPS needs the cryptography package. Run: "
            r"C:\Tools\.venv\Scripts\python.exe -m pip install cryptography"
        ) from exc
    return x509, hashes, serialization, rsa, ExtendedKeyUsageOID, NameOID


def _is_ip(value: str) -> bool:
    try:
        ip_address(value)
    except ValueError:
        return False
    return True


def _expires_within(cert, days: int) -> bool:
    expires_at = getattr(cert, "not_valid_after_utc", None)
    if expires_at is None:  # older cryptography returns a naive UTC datetime
        expires_at = cert.not_valid_after.replace(tzinfo=timezone.utc)
    return expires_at <= datetime.now(timezone.utc) + timedelta(days=days)


def _write_key(path: Path, key, serialization) -> None:
    path.write_bytes(
        key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )


def _ensure_ca():
    x509, hashes, serialization, rsa, _eku, NameOID = _crypto()
    CERT_DIR.mkdir(parents=True, exist_ok=True)

    if CA_CERT_PATH.exists() and CA_KEY_PATH.exists():
        try:
            cert = x509.load_pem_x509_certificate(CA_CERT_PATH.read_bytes())
            key = serialization.load_pem_private_key(CA_KEY_PATH.read_bytes(), password=None)
            constraints = cert.extensions.get_extension_for_class(x509.BasicConstraints).value
            if bool(constraints.ca) and not _expires_within(cert, 30):
                return cert, key
        except Exception:
            # Unreadable/expired CA: fall through and mint a fresh one. The old
            # one stays trusted-but-unused in the store until manually removed.
            pass

    key = rsa.generate_private_key(public_exponent=65537, key_size=3072)
    name = x509.Name(
        [
            x509.NameAttribute(NameOID.COMMON_NAME, CA_COMMON_NAME),
            x509.NameAttribute(NameOID.ORGANIZATION_NAME, ORGANIZATION_NAME),
        ]
    )
    now = datetime.now(timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(days=1))
        .not_valid_after(now + timedelta(days=3650))
        .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=True,
                crl_sign=True,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .sign(key, hashes.SHA256())
    )
    CA_CERT_PATH.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    _write_key(CA_KEY_PATH, key, serialization)
    return cert, key


def _leaf_is_usable(ca_cert, hosts: tuple[str, ...]) -> bool:
    if not CERT_PATH.exists() or not KEY_PATH.exists():
        return False
    try:
        from cryptography import x509

        cert = x509.load_pem_x509_certificate(CERT_PATH.read_bytes())
        if cert.issuer != ca_cert.subject or _expires_within(cert, 14):
            return False
        san = cert.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
        dns_names = {value.casefold() for value in san.get_values_for_type(x509.DNSName)}
        ip_names = {str(value) for value in san.get_values_for_type(x509.IPAddress)}
    except Exception:
        return False

    for host in hosts:
        if _is_ip(host):
            if host not in ip_names:
                return False
        elif host.casefold() not in dns_names:
            return False
    return True


def ensure_loopback_certificate(
    hosts: tuple[str, ...] = LOOPBACK_HOSTS,
) -> tuple[Path, Path, Path]:
    """Return (cert_path, key_path, ca_cert_path), generating them if needed.

    Cheap and idempotent: existing material that still covers every host and
    isn't near expiry is reused, so this is safe to call on every startup.
    """
    x509, hashes, serialization, rsa, ExtendedKeyUsageOID, NameOID = _crypto()
    CERT_DIR.mkdir(parents=True, exist_ok=True)
    ca_cert, ca_key = _ensure_ca()
    if _leaf_is_usable(ca_cert, hosts):
        return CERT_PATH, KEY_PATH, CA_CERT_PATH

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    alt_names = [
        x509.IPAddress(ip_address(host)) if _is_ip(host) else x509.DNSName(host)
        for host in hosts
    ]
    now = datetime.now(timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(
            x509.Name(
                [
                    x509.NameAttribute(NameOID.COMMON_NAME, LEAF_COMMON_NAME),
                    x509.NameAttribute(NameOID.ORGANIZATION_NAME, ORGANIZATION_NAME),
                ]
            )
        )
        .issuer_name(ca_cert.subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(days=1))
        # 825 days is the maximum lifetime browsers/webviews accept for a
        # server certificate; longer ones are rejected outright.
        .not_valid_after(now + timedelta(days=825))
        .add_extension(x509.SubjectAlternativeName(alt_names), critical=False)
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
        .sign(ca_key, hashes.SHA256())
    )
    CERT_PATH.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    _write_key(KEY_PATH, key, serialization)
    return CERT_PATH, KEY_PATH, CA_CERT_PATH


def certificate_sha256_fingerprint(path: Path) -> str:
    digest = hashlib.sha256(path.read_bytes()).hexdigest().upper()
    return ":".join(digest[index : index + 2] for index in range(0, len(digest), 2))
