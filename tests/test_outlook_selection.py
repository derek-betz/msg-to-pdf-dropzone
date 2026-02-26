from __future__ import annotations

from msg_to_pdf_dropzone.outlook_selection import is_likely_outlook_drop


def test_is_likely_outlook_drop_checks_source_markers() -> None:
    assert is_likely_outlook_drop("", ("FileGroupDescriptorW",))
    assert is_likely_outlook_drop("", ("RenPrivateMessages",))
    assert not is_likely_outlook_drop("", ("DND_Files",))


def test_is_likely_outlook_drop_checks_raw_data_text() -> None:
    assert is_likely_outlook_drop("Outlook Drag Payload", ())
    assert not is_likely_outlook_drop("regular file drop", ())
