from pathlib import Path

from app.database import Database
from app.sources.base import SpreadsheetInfo, VersionInfo
from app.version_catalog import VersionCatalog


SHEET = SpreadsheetInfo("site", "drive", "book", "CQLPA123.xlsx", "/book.xlsx")


def versions(last: int):
    return tuple(
        VersionInfo(str(i), f"1.{i}", is_current=i == last)
        for i in range(1, last + 1)
    )


class Source:
    def __init__(self, remote):
        self.remote = remote
        self.full_calls = 0
        self.current_calls = 0
        self.delta_calls = []

    def list_versions(self, _sheet, progress_callback=None):
        self.full_calls += 1
        return self.remote

    def get_current_version(self, _sheet):
        self.current_calls += 1
        return self.remote[-1]

    def list_version_delta(self, _sheet, anchor_id, anchor_label, progress_callback=None):
        self.delta_calls.append((anchor_id, anchor_label))
        start = next(i for i, version in enumerate(self.remote) if version.id == anchor_id)
        return self.remote[start:]


def database(tmp_path: Path):
    db = Database(tmp_path / "audit.db")
    db.initialize()
    return db


def test_second_open_uses_current_metadata_and_zero_historical_pages(tmp_path):
    db = database(tmp_path)
    source = Source(versions(17_575))
    first = VersionCatalog(db, source).synchronize(SHEET)
    second = VersionCatalog(db, source).synchronize(SHEET)

    assert first.source == "full_rebuild"
    assert len(first.versions) == 17_575
    assert second.source == "local_only"
    assert source.full_calls == 1
    assert source.current_calls == 1
    assert source.delta_calls == []
    db.close()


def test_new_current_fetches_only_catalog_tail_and_promotes_cached_current(tmp_path):
    db = database(tmp_path)
    source = Source(versions(3))
    VersionCatalog(db, source).synchronize(SHEET)
    source.remote = versions(5)

    result = VersionCatalog(db, source).synchronize(SHEET)

    assert result.source == "delta"
    assert source.full_calls == 1
    assert source.delta_calls == [("3", "1.3")]
    assert [(v.id, v.is_current) for v in result.versions[-3:]] == [
        ("3", False), ("4", False), ("5", True)
    ]
    db.close()


def test_checkpoint_is_not_a_catalog_frontier(tmp_path):
    db = database(tmp_path)
    source = Source(versions(20))
    VersionCatalog(db, source).synchronize(SHEET)
    db.connection.execute(
        "INSERT INTO planilha (drive_item_id,nome_atual,site_id,drive_id) VALUES ('book','x','site','drive')"
    )
    planilha_id = db.connection.execute("SELECT id FROM planilha").fetchone()[0]
    db.connection.execute(
        "INSERT INTO checkpoint (planilha_id,versao_id,versao_numero) VALUES (?,?,?)",
        (planilha_id, "2", "1.2"),
    )
    db.connection.commit()

    result = VersionCatalog(db, source).synchronize(SHEET)

    assert result.source == "local_only"
    assert source.delta_calls == []
    assert source.full_calls == 1
    db.close()


def test_conflicting_anchor_is_the_only_path_to_full_rebuild(tmp_path):
    db = database(tmp_path)
    source = Source(versions(3))
    VersionCatalog(db, source).synchronize(SHEET)
    source.remote = versions(4)
    source.remote = (
        VersionInfo("1", "1.1"), VersionInfo("2", "1.2"),
        VersionInfo("3", "CONFLICT"), VersionInfo("4", "1.4", is_current=True),
    )

    result = VersionCatalog(db, source).synchronize(SHEET)

    assert result.source == "full_rebuild"
    assert source.full_calls == 2
    db.close()
