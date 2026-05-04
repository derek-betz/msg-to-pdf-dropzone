from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

import pytest

import msg_to_pdf_dropzone.msg_parser as msg_parser


class _FakeAttachment:
    def __init__(
        self,
        *,
        long_filename: str,
        cid: str | None = None,
        content_id: str | None = None,
        mimetype: str | None = None,
        data: bytes | None = None,
    ) -> None:
        self.longFilename = long_filename
        self.filename = long_filename
        self.cid = cid
        self.contentId = content_id
        self.mimetype = mimetype
        self.data = data or b""


class _FakeMessage:
    def __init__(self) -> None:
        self.subject = "Sample Subject"
        self.date = "Wed, 04 Mar 2026 14:50:00 -0700"
        self.sender = "sender@example.com"
        self.to = "to@example.com"
        self.cc = "cc@example.com"
        self.body = "Body text"
        self.htmlBody = "<html><body><p>Body</p></body></html>"
        self.attachments = [
            _FakeAttachment(
                long_filename="image001.png",
                cid="image001@abc",
                content_id="image001@abc",
                mimetype="image/png",
                data=b"png-data",
            ),
            _FakeAttachment(
                long_filename="contract.pdf",
                cid="pdf-inline",
                mimetype="application/pdf",
                data=b"pdf-data",
            ),
            _FakeAttachment(
                long_filename="notes.txt",
                cid=None,
                mimetype="text/plain",
                data=b"text-data",
            ),
        ]

    def close(self) -> None:
        return


def test_parse_msg_file_extracts_inline_cid_images(monkeypatch, tmp_path: Path) -> None:
    msg_path = tmp_path / "sample.msg"
    msg_path.write_text("dummy", encoding="utf-8")

    fake_message = _FakeMessage()
    monkeypatch.setattr(msg_parser.extract_msg, "Message", lambda _path: fake_message)

    record = msg_parser.parse_msg_file(msg_path)

    assert record.subject == "Sample Subject"
    assert record.attachment_names == ["image001.png", "contract.pdf", "notes.txt"]
    assert len(record.inline_images) == 1
    inline_asset = record.inline_images[0]
    assert inline_asset.cid == "image001@abc"
    assert inline_asset.mime_type == "image/png"
    assert inline_asset.filename == "image001.png"
    assert inline_asset.data == b"png-data"
    assert inline_asset.size_bytes == len(b"png-data")


def test_parse_msg_file_raises_for_non_msg_extension(tmp_path: Path) -> None:
    txt_path = tmp_path / "file.txt"
    txt_path.write_text("dummy", encoding="utf-8")
    with pytest.raises(ValueError, match="Unsupported file type"):
        msg_parser.parse_msg_file(txt_path)


def test_parse_sent_date_returns_none_for_none() -> None:
    assert msg_parser._parse_sent_date(None) is None


def test_parse_sent_date_returns_none_for_empty_string() -> None:
    assert msg_parser._parse_sent_date("") is None


def test_parse_sent_date_passes_through_datetime_object() -> None:
    dt = datetime(2026, 1, 15, 12, 0, tzinfo=timezone.utc)
    assert msg_parser._parse_sent_date(dt) == dt


def test_parse_sent_date_parses_rfc2822_string() -> None:
    result = msg_parser._parse_sent_date("Wed, 04 Mar 2026 14:50:00 -0700")
    assert result is not None
    assert result.year == 2026
    assert result.month == 3
    assert result.day == 4


def test_parse_sent_date_returns_none_for_invalid_string() -> None:
    assert msg_parser._parse_sent_date("not-a-date") is None


def test_as_html_decodes_utf8_bytes() -> None:
    assert msg_parser._as_html(b"<p>hello</p>") == "<p>hello</p>"


def test_as_html_returns_empty_string_for_none() -> None:
    assert msg_parser._as_html(None) == ""


def test_as_html_strips_string_input() -> None:
    assert msg_parser._as_html("  <p>content</p>  ") == "<p>content</p>"


def test_as_bytes_handles_bytearray() -> None:
    result = msg_parser._as_bytes(bytearray(b"hello"))
    assert result == b"hello"


def test_as_bytes_handles_memoryview() -> None:
    result = msg_parser._as_bytes(memoryview(b"hello"))
    assert result == b"hello"


def test_as_bytes_returns_empty_for_none() -> None:
    assert msg_parser._as_bytes(None) == b""


def test_normalize_cid_strips_angle_brackets() -> None:
    assert msg_parser._normalize_cid("<image001@abc>") == "image001@abc"


def test_normalize_cid_strips_cid_prefix() -> None:
    assert msg_parser._normalize_cid("cid:image001@abc") == "image001@abc"


def test_normalize_cid_strips_cid_prefix_with_angle_brackets() -> None:
    assert msg_parser._normalize_cid("<cid:image001@abc>") == "image001@abc"


def test_normalize_cid_returns_empty_for_empty_input() -> None:
    assert msg_parser._normalize_cid("") == ""


def test_guess_image_mime_returns_image_type_for_png() -> None:
    result = msg_parser._guess_image_mime("photo.png")
    assert result.startswith("image/")


def test_guess_image_mime_returns_empty_for_non_image() -> None:
    assert msg_parser._guess_image_mime("file.pdf") == ""


def test_guess_image_mime_returns_empty_for_empty_filename() -> None:
    assert msg_parser._guess_image_mime("") == ""


def test_parse_msg_file_deduplicates_cids_in_inline_images(monkeypatch, tmp_path: Path) -> None:
    msg_path = tmp_path / "sample.msg"
    msg_path.write_text("dummy", encoding="utf-8")

    class _FakeDupAttachment:
        def __init__(self, cid: str) -> None:
            self.longFilename = "img.png"
            self.filename = "img.png"
            self.cid = cid
            self.contentId = cid
            self.mimetype = "image/png"
            self.data = b"png-data"

    class _FakeDupMessage:
        subject = "Dup CID"
        date = "Wed, 04 Mar 2026 14:50:00 +0000"
        sender = "a@example.com"
        to = "b@example.com"
        cc = ""
        body = "body"
        htmlBody = None
        attachments = [_FakeDupAttachment("dup@abc"), _FakeDupAttachment("dup@abc")]

        def close(self) -> None:
            return

    monkeypatch.setattr(msg_parser.extract_msg, "Message", lambda _path: _FakeDupMessage())
    record = msg_parser.parse_msg_file(msg_path)
    assert len(record.inline_images) == 1
