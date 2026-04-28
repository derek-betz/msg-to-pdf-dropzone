from __future__ import annotations

from pathlib import Path

import pytest

from msg_to_pdf_dropzone.web_server import ServerSettings, load_server_settings, validate_startup_contract


def test_load_server_settings_reads_tls_values(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    certfile = tmp_path / "server.crt"
    keyfile = tmp_path / "server.key"
    monkeypatch.setenv("MSG_TO_PDF_HOST", "10.1.13.203")
    monkeypatch.setenv("MSG_TO_PDF_PORT", "443")
    monkeypatch.setenv("MSG_TO_PDF_TLS_CERTFILE", str(certfile))
    monkeypatch.setenv("MSG_TO_PDF_TLS_KEYFILE", str(keyfile))
    monkeypatch.setenv("MSG_TO_PDF_SERVER_NAMES", "emailpdf.hanson-inc.com")

    settings = load_server_settings()

    assert settings.host == "10.1.13.203"
    assert settings.port == 443
    assert settings.tls_certfile == certfile
    assert settings.tls_keyfile == keyfile
    assert settings.tls_server_names == ("emailpdf.hanson-inc.com",)


def test_production_non_loopback_requires_tls(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    settings = ServerSettings(
        host="10.1.13.203",
        port=443,
        tls_certfile=None,
        tls_keyfile=None,
        tls_server_names=(),
    )

    with pytest.raises(SystemExit, match="loopback binding unless direct TLS is configured"):
        validate_startup_contract(settings)


def test_tls_requires_existing_files(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    settings = ServerSettings(
        host="10.1.13.203",
        port=443,
        tls_certfile=tmp_path / "missing.crt",
        tls_keyfile=tmp_path / "missing.key",
        tls_server_names=("emailpdf.hanson-inc.com",),
    )

    with pytest.raises(SystemExit, match="TLS cert file was not found"):
        validate_startup_contract(settings)


def test_tls_allows_non_loopback_when_files_exist(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    certfile = tmp_path / "server.crt"
    keyfile = tmp_path / "server.key"
    certfile.write_text("cert", encoding="utf-8")
    keyfile.write_text("key", encoding="utf-8")
    settings = ServerSettings(
        host="10.1.13.203",
        port=443,
        tls_certfile=certfile,
        tls_keyfile=keyfile,
        tls_server_names=("emailpdf.hanson-inc.com",),
    )

    validate_startup_contract(settings)
