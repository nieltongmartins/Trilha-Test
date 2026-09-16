import hashlib
from pathlib import Path

from app.integrity import sha256_file


def test_sha256_file_is_deterministic_and_uses_binary_content(tmp_path: Path) -> None:
    first = tmp_path / "first.xlsx"
    copy = tmp_path / "copy.xlsx"
    different = tmp_path / "different.xlsx"
    content = (b"conteudo-binario\x00" * 100_000) + b"fim"
    first.write_bytes(content)
    copy.write_bytes(content)
    different.write_bytes(content + b"diferente")

    expected = hashlib.sha256(content).hexdigest()
    assert sha256_file(first, chunk_size=31) == expected
    assert sha256_file(copy) == expected
    assert sha256_file(different) != expected
