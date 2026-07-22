"""Tests for the loopback certificate.

The properties asserted here are the ones Office is strict about; getting any
of them wrong shows up as an opaque "can't open this add-in" in Outlook rather
than a useful error, so they're pinned down.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from ipaddress import ip_address
from pathlib import Path

import pytest
from cryptography import x509
from cryptography.x509.oid import ExtendedKeyUsageOID

import job_intake_tls


@pytest.fixture(autouse=True)
def temp_cert_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Never touch the real _runtime/certs - a test must not replace the CA
    the user has already trusted in Windows."""
    cert_dir = tmp_path / "certs"
    monkeypatch.setattr(job_intake_tls, "CERT_DIR", cert_dir)
    monkeypatch.setattr(job_intake_tls, "CA_CERT_PATH", cert_dir / "ca.crt.pem")
    monkeypatch.setattr(job_intake_tls, "CA_KEY_PATH", cert_dir / "ca.key.pem")
    monkeypatch.setattr(job_intake_tls, "CERT_PATH", cert_dir / "leaf.crt.pem")
    monkeypatch.setattr(job_intake_tls, "KEY_PATH", cert_dir / "leaf.key.pem")
    return cert_dir


def _load(path: Path):
    return x509.load_pem_x509_certificate(path.read_bytes())


def test_certificate_covers_both_localhost_and_the_loopback_ip() -> None:
    """Office accepts either spelling in a manifest's SourceLocation, and a
    cert issued for only one of them fails on the other."""
    cert_path, _key_path, _ca_path = job_intake_tls.ensure_loopback_certificate()
    san = _load(cert_path).extensions.get_extension_for_class(x509.SubjectAlternativeName).value

    assert "localhost" in {name.casefold() for name in san.get_values_for_type(x509.DNSName)}
    assert ip_address("127.0.0.1") in san.get_values_for_type(x509.IPAddress)


def test_leaf_is_signed_by_the_generated_root_ca() -> None:
    cert_path, _key_path, ca_path = job_intake_tls.ensure_loopback_certificate()
    leaf, ca = _load(cert_path), _load(ca_path)

    assert leaf.issuer == ca.subject
    assert ca.extensions.get_extension_for_class(x509.BasicConstraints).value.ca is True
    # The leaf must not itself be a CA, and must be usable as a server.
    assert leaf.extensions.get_extension_for_class(x509.BasicConstraints).value.ca is False
    eku = leaf.extensions.get_extension_for_class(x509.ExtendedKeyUsage).value
    assert ExtendedKeyUsageOID.SERVER_AUTH in eku


def test_leaf_lifetime_stays_within_the_825_day_limit() -> None:
    """Browsers and webviews reject server certs valid for longer."""
    cert_path, _key, _ca = job_intake_tls.ensure_loopback_certificate()
    leaf = _load(cert_path)
    not_before = getattr(leaf, "not_valid_before_utc", None) or leaf.not_valid_before.replace(
        tzinfo=timezone.utc
    )
    not_after = getattr(leaf, "not_valid_after_utc", None) or leaf.not_valid_after.replace(
        tzinfo=timezone.utc
    )
    assert not_after - not_before <= timedelta(days=826)
    assert not_after > datetime.now(timezone.utc)


def test_calling_twice_reuses_the_same_material(temp_cert_dir: Path) -> None:
    """Startup calls this every time; regenerating would silently invalidate
    the CA the user installed into the Windows trust store."""
    cert_path, key_path, ca_path = job_intake_tls.ensure_loopback_certificate()
    first = (cert_path.read_bytes(), key_path.read_bytes(), ca_path.read_bytes())

    job_intake_tls.ensure_loopback_certificate()

    assert (cert_path.read_bytes(), key_path.read_bytes(), ca_path.read_bytes()) == first


def test_a_leaf_missing_a_required_host_is_reissued() -> None:
    """Guards the upgrade path: an older cert covering only localhost must be
    replaced once 127.0.0.1 is also required."""
    localhost_only, _key, ca_path = job_intake_tls.ensure_loopback_certificate(("localhost",))
    before = localhost_only.read_bytes()
    ca_before = ca_path.read_bytes()

    reissued, _key2, ca_after_path = job_intake_tls.ensure_loopback_certificate()

    assert reissued.read_bytes() != before
    # Only the leaf is replaced; the trusted root CA is left alone.
    assert ca_after_path.read_bytes() == ca_before
    san = _load(reissued).extensions.get_extension_for_class(x509.SubjectAlternativeName).value
    assert ip_address("127.0.0.1") in san.get_values_for_type(x509.IPAddress)


def test_fingerprint_is_stable_and_readable() -> None:
    _cert, _key, ca_path = job_intake_tls.ensure_loopback_certificate()
    fingerprint = job_intake_tls.certificate_sha256_fingerprint(ca_path)

    assert fingerprint == job_intake_tls.certificate_sha256_fingerprint(ca_path)
    assert len(fingerprint.split(":")) == 32
