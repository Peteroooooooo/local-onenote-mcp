import base64

from local_onenote_mcp.backup import BackupStore

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ONE = f"{{{NS}}}"
IMAGE = base64.b64encode(b"\x89PNG" + b"x" * 5000).decode()


def page(stamp: str, text: str = "hello", image: str = IMAGE) -> str:
    return (
        f'<one:Page xmlns:one="{NS}" ID="P" lastModifiedTime="{stamp}">'
        f"<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[{text}]]></one:T></one:OE></one:OEChildren></one:Outline>"
        f'<one:Image format="png"><one:Data>{image}</one:Data></one:Image>'
        '<one:InsertedFile preferredName="doc.pdf" pathCache="C:/missing.bin"/>'
        "</one:Page>"
    )


def store(tmp_path, **limits) -> BackupStore:
    options = {"max_bytes": 10 * 2**20, "versions": 10, "days": 30} | limits
    return BackupStore(tmp_path / "backups.sqlite3", **options)


def count(backups: BackupStore, table: str) -> int:
    return backups._conn().execute(f"select count(*) from {table}").fetchone()[0]


def test_images_are_stored_once_across_snapshots(tmp_path):
    backups = store(tmp_path)
    backups.save(page("t1", "one"), page_id="P", reason="edit")
    backups.save(page("t2", "two"), page_id="P", reason="edit")

    assert count(backups, "snapshots") == 2
    assert count(backups, "blobs") == 1


def test_unchanged_page_is_not_saved_again(tmp_path):
    backups = store(tmp_path)
    first = backups.save(page("t1"), page_id="P", reason="edit")
    again = backups.save(page("t1"), page_id="P", reason="edit")

    assert first == again and count(backups, "snapshots") == 1


def test_version_cap_keeps_newest_and_drops_orphan_blobs(tmp_path):
    backups = store(tmp_path, versions=2)
    for n in range(4):
        image = base64.b64encode(f"image {n}".encode()).decode()
        backups.save(page(f"t{n}", image=image), page_id="P", reason="edit")

    assert [b["backup_id"] for b in backups.list("P")] == [4, 3]
    assert count(backups, "blobs") == 2


def test_size_cap_drops_oldest_but_keeps_latest(tmp_path):
    backups = store(tmp_path, max_bytes=1)
    for n in range(3):
        backups.save(page(f"t{n}"), page_id=f"P{n}", reason="edit")

    assert [b["page_id"] for b in backups.list()] == ["P2"]


def test_age_cap_drops_old_snapshots(tmp_path):
    backups = store(tmp_path, days=1)
    backups.save(page("t1"), page_id="P", reason="edit")
    backups._conn().execute("update snapshots set created = created - 3 * 86400")
    backups.save(page("t2"), page_id="Q", reason="edit")

    assert [b["page_id"] for b in backups.list()] == ["Q"]


def test_load_restores_images_and_attachments(tmp_path):
    backups = store(tmp_path)
    snapshot_id = backups.save(page("t1"), page_id="P", path="NB/S/Page", title="Page", reason="edit", read_file=lambda node: b"%PDF")

    meta, root, files = backups.load(snapshot_id)

    assert meta["path"] == "NB/S/Page"
    assert root.find(f".//{ONE}Data").text == IMAGE
    attached = root.find(f".//{ONE}InsertedFile")
    assert files[attached.attrib["mcpBlob"]] == b"%PDF"


def test_disabled_store_does_nothing():
    backups = BackupStore(None, max_bytes=1, versions=1, days=1)

    assert backups.save(page("t1"), page_id="P", reason="edit") is None
    assert backups.list() == [] and backups.stats()["enabled"] is False
