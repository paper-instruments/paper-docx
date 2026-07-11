"""Out-of-band installation diagnostics for the frozen ``docx`` import."""

from __future__ import annotations

import base64
import hashlib
from types import SimpleNamespace

import pytest

import paper_docx_doctor as doctor


class _Distribution:
    def __init__(self, root, record: str, version: str = "0.1.1"):
        self._root = root
        self._record = record
        self.version = version

    def read_text(self, name: str):
        return self._record if name == "RECORD" else None

    def locate_file(self, path):
        return self._root.joinpath(*path.parts)


def _hash(data: bytes) -> str:
    digest = hashlib.sha256(data).digest()
    encoded = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
    return f"sha256={encoded}"


def _paper_distribution(tmp_path, data: bytes = b"paper bytes") -> _Distribution:
    package = tmp_path / "docx"
    package.mkdir()
    module = package / "__init__.py"
    module.write_bytes(data)
    record = (
        f"docx/__init__.py,{_hash(data)},{len(data)}\n"
        "docx/__pycache__/__init__.pyc,,0\n"
        "../../../bin/paper-docx-doctor,,0\n"
    )
    return _Distribution(tmp_path, record)


def _install_lookup(monkeypatch, paper=None, upstream=None) -> None:
    def lookup(name: str):
        found = {"paper-docx": paper, "python-docx": upstream}[name]
        if found is None:
            raise doctor.PackageNotFoundError(name)
        return found

    monkeypatch.setattr(doctor, "distribution", lookup)


class DescribePaperDocxDoctor:
    def it_accepts_a_clean_install(self, monkeypatch, tmp_path):
        paper = _paper_distribution(tmp_path)
        _install_lookup(monkeypatch, paper=paper)
        monkeypatch.setattr(
            doctor.importlib,
            "import_module",
            lambda _name: SimpleNamespace(
                __file__=str(tmp_path / "docx" / "__init__.py"),
                __paper_version__="0.1.1",
            ),
        )

        assert doctor.verify_install() == "0.1.1"

    def it_refuses_missing_paper_distribution_metadata(self, monkeypatch):
        _install_lookup(monkeypatch)

        with pytest.raises(doctor.DoctorError, match="metadata is missing"):
            doctor.verify_install()

    def it_refuses_when_both_distributions_are_installed(
        self, monkeypatch, tmp_path
    ):
        paper = _paper_distribution(tmp_path)
        _install_lookup(monkeypatch, paper=paper, upstream=object())

        with pytest.raises(doctor.DoctorError, match="both installed"):
            doctor.verify_install()

    def it_refuses_a_missing_recorded_docx_file(self, monkeypatch, tmp_path):
        paper = _paper_distribution(tmp_path)
        (tmp_path / "docx" / "__init__.py").unlink()
        _install_lookup(monkeypatch, paper=paper)

        with pytest.raises(doctor.DoctorError, match="file is missing"):
            doctor.verify_install()

    def it_refuses_a_hash_mismatch(self, monkeypatch, tmp_path):
        paper = _paper_distribution(tmp_path)
        (tmp_path / "docx" / "__init__.py").write_bytes(b"upstream bytes")
        _install_lookup(monkeypatch, paper=paper)

        with pytest.raises(doctor.DoctorError, match="hash mismatch"):
            doctor.verify_install()

    def it_refuses_a_matching_sentinel_outside_the_owned_docx_root(
        self, monkeypatch, tmp_path
    ):
        paper = _paper_distribution(tmp_path)
        foreign_module = tmp_path / "shadow" / "docx" / "__init__.py"
        foreign_module.parent.mkdir(parents=True)
        foreign_module.write_bytes(b"paper bytes")
        _install_lookup(monkeypatch, paper=paper)
        monkeypatch.setattr(
            doctor.importlib,
            "import_module",
            lambda _name: SimpleNamespace(
                __file__=str(foreign_module), __paper_version__="0.1.1"
            ),
        )

        with pytest.raises(doctor.DoctorError, match="outside.*owned docx root"):
            doctor.verify_install()

    def it_resolves_the_import_path_before_checking_ownership(
        self, monkeypatch, tmp_path
    ):
        paper = _paper_distribution(tmp_path)
        foreign_module = tmp_path / "shadow" / "__init__.py"
        foreign_module.parent.mkdir()
        foreign_module.write_bytes(b"paper bytes")
        linked_module = tmp_path / "docx" / "shadow.py"
        linked_module.symlink_to(foreign_module)
        _install_lookup(monkeypatch, paper=paper)
        monkeypatch.setattr(
            doctor.importlib,
            "import_module",
            lambda _name: SimpleNamespace(
                __file__=str(linked_module), __paper_version__="0.1.1"
            ),
        )

        with pytest.raises(doctor.DoctorError, match="outside.*owned docx root"):
            doctor.verify_install()

    @pytest.mark.parametrize(
        ("sentinel", "message"),
        [(None, "is missing"), ("9.9.9", "does not match")],
    )
    def it_refuses_a_missing_or_wrong_sentinel(
        self, monkeypatch, tmp_path, sentinel, message
    ):
        paper = _paper_distribution(tmp_path)
        _install_lookup(monkeypatch, paper=paper)
        imported = SimpleNamespace(
            __file__=str(tmp_path / "docx" / "__init__.py")
        )
        if sentinel is not None:
            imported.__paper_version__ = sentinel
        monkeypatch.setattr(
            doctor.importlib, "import_module", lambda _name: imported
        )

        with pytest.raises(doctor.DoctorError, match=message):
            doctor.verify_install()

    def it_prints_a_clean_reinstall_remedy(self, monkeypatch, capsys):
        monkeypatch.setattr(
            doctor,
            "verify_install",
            lambda: (_ for _ in ()).throw(doctor.DoctorError("broken")),
        )

        assert doctor.main() == 1
        stderr = capsys.readouterr().err
        assert "uninstall -y python-docx paper-docx" in stderr
        assert "--force-reinstall paper-docx" in stderr
