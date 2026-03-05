from __future__ import annotations

from pathlib import Path

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
