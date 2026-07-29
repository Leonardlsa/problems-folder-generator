from __future__ import annotations

import hashlib
import struct
import zipfile
import zlib
from pathlib import Path

import openpyxl
import pytest
from docx import Document

import core


class LogSink:
    def __init__(self, confirm_missing: bool = False) -> None:
        self.messages: list[str] = []
        self.textBrowser = self
        self.confirm_missing = confirm_missing
        self.confirmations: list[dict[str, list[int]]] = []

    def append(self, message: str) -> None:
        self.messages.append(message)

    def confirm_missing_images(
        self, missing_by_student: dict[str, list[int]]
    ) -> bool:
        self.confirmations.append(missing_by_student)
        return self.confirm_missing


def _png_chunk(kind: bytes, data: bytes) -> bytes:
    checksum = zlib.crc32(kind + data) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", checksum)


def write_png(
    path: Path,
    color: tuple[int, int, int],
    width: int = 120,
    height: int = 60,
) -> None:
    signature = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    row = b"\x00" + bytes(color) * width
    image = zlib.compress(row * height)
    path.write_bytes(
        signature
        + _png_chunk(b"IHDR", ihdr)
        + _png_chunk(b"IDAT", image)
        + _png_chunk(b"IEND", b"")
    )


def write_feedback(path: Path, rows: list[tuple[str, object]]) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["姓名", "错题号"])
    for row in rows:
        sheet.append(list(row))
    workbook.save(path)


def document_text(path: Path) -> str:
    return "\n".join(paragraph.text for paragraph in Document(path).paragraphs)


def media_hashes(path: Path) -> set[str]:
    hashes: set[str] = set()
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("word/media/"):
                hashes.add(hashlib.sha256(archive.read(name)).hexdigest())
    return hashes


def file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def prepare_two_editions(root: Path) -> tuple[Path, Path]:
    student_images = root / "student-images"
    teacher_images = root / "teacher-images"
    student_images.mkdir()
    teacher_images.mkdir()
    student_colors = ((255, 0, 0), (0, 255, 0), (0, 0, 255))
    teacher_colors = ((255, 255, 0), (255, 0, 255), (0, 255, 255))
    for number, (student_color, teacher_color) in enumerate(
        zip(student_colors, teacher_colors), start=1
    ):
        write_png(student_images / f"{number:02d}.png", student_color)
        write_png(teacher_images / f"{number:02d}.png", teacher_color)
    return student_images, teacher_images


def test_generate_coupled_student_and_teacher_versions(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    student_images, teacher_images = prepare_two_editions(tmp_path)
    feedback = tmp_path / "feedback.xlsx"
    write_feedback(feedback, [("张三", "1,3"), ("李四", "2")])
    monkeypatch.chdir(tmp_path)

    result = core.generate(
        str(student_images),
        str(teacher_images),
        str(feedback),
        "高一1班",
        "物理",
        "03",
        "0729",
        LogSink(),
    )

    assert len(result.student_version_paths) == 2
    assert len(result.teacher_version_paths) == 2
    zhang_student = tmp_path / "学生/张三/物理学生版/第03周错题集.docx"
    zhang_teacher = tmp_path / "学生/张三/物理教师版/第03周错题集.docx"
    li_student = tmp_path / "学生/李四/物理学生版/第03周错题集.docx"
    li_teacher = tmp_path / "学生/李四/物理教师版/第03周错题集.docx"
    for path in (zhang_student, zhang_teacher, li_student, li_teacher):
        assert path.is_file()

    assert media_hashes(zhang_student) == {
        file_hash(student_images / "01.png"),
        file_hash(student_images / "03.png"),
    }
    assert media_hashes(zhang_teacher) == {
        file_hash(teacher_images / "01.png"),
        file_hash(teacher_images / "03.png"),
    }
    assert media_hashes(li_student) == {file_hash(student_images / "02.png")}
    assert media_hashes(li_teacher) == {file_hash(teacher_images / "02.png")}

    assert "物理学生版" in Document(zhang_student).sections[0].header.paragraphs[0].text
    assert "物理教师版" in Document(zhang_teacher).sections[0].header.paragraphs[0].text
    assert not (tmp_path / "教师").exists()


def test_special_states_are_mirrored_to_both_editions(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    student_images, teacher_images = prepare_two_editions(tmp_path)
    feedback = tmp_path / "feedback.xlsx"
    write_feedback(feedback, [("甲", "全对"), ("乙", "未交")])
    monkeypatch.chdir(tmp_path)

    core.generate(
        str(student_images),
        str(teacher_images),
        str(feedback),
        "一班",
        "数学",
        "1",
        "0729",
        LogSink(),
    )

    for edition in ("学生版", "教师版"):
        all_correct = tmp_path / f"学生/甲/数学{edition}/第1周错题集.docx"
        unsubmitted = tmp_path / f"学生/乙/数学{edition}/第1周错题集.docx"
        assert "全对" in document_text(all_correct)
        assert len(media_hashes(unsubmitted)) == 3


def test_image_sets_must_have_matching_problem_numbers(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    student_images, teacher_images = prepare_two_editions(tmp_path)
    (teacher_images / "03.png").unlink()
    feedback = tmp_path / "feedback.xlsx"
    write_feedback(feedback, [("甲", "1")])
    monkeypatch.chdir(tmp_path)

    with pytest.raises(core.GenerateError, match="图片题号不对应"):
        core.generate(
            str(student_images),
            str(teacher_images),
            str(feedback),
            "一班",
            "物理",
            "1",
            "0729",
            LogSink(),
        )

    assert not (tmp_path / "学生").exists()


def test_duplicate_problem_images_are_rejected(tmp_path: Path) -> None:
    images = tmp_path / "images"
    images.mkdir()
    write_png(images / "1.png", (255, 0, 0))
    write_png(images / "01.png", (0, 255, 0))

    with pytest.raises(core.GenerateError, match="重复题号: 1"):
        core._discover_images(str(images), "学生版")


def test_missing_requested_problem_cancel_writes_nothing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    student_images, teacher_images = prepare_two_editions(tmp_path)
    feedback = tmp_path / "feedback.xlsx"
    write_feedback(feedback, [("甲", "4")])
    monkeypatch.chdir(tmp_path)
    sink = LogSink(confirm_missing=False)

    with pytest.raises(core.GenerateError, match="用户取消"):
        core.generate(
            str(student_images),
            str(teacher_images),
            str(feedback),
            "一班",
            "物理",
            "1",
            "0729",
            sink,
        )

    assert sink.confirmations == [{"甲": [4]}]
    assert not (tmp_path / "学生").exists()
    assert not (tmp_path / "记录.docx").exists()


def test_second_date_appends_and_backup_restores_both_editions(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    student_images, teacher_images = prepare_two_editions(tmp_path)
    feedback = tmp_path / "feedback.xlsx"
    write_feedback(feedback, [("甲", "全对")])
    monkeypatch.chdir(tmp_path)
    args = (
        str(student_images),
        str(teacher_images),
        str(feedback),
        "一班",
        "物理",
        "1",
    )

    core.generate(*args, "0728", LogSink())
    core.generate(*args, "0729", LogSink())

    student_path = tmp_path / "学生/甲/物理学生版/第1周错题集.docx"
    teacher_path = tmp_path / "学生/甲/物理教师版/第1周错题集.docx"
    assert "0728" in document_text(student_path)
    assert "0729" in document_text(student_path)
    assert "0728" in document_text(teacher_path)
    assert "0729" in document_text(teacher_path)
    assert (tmp_path / "backup/学生/甲/物理学生版/第1周错题集.docx").is_file()
    assert (tmp_path / "backup/学生/甲/物理教师版/第1周错题集.docx").is_file()

    restored = core.restore_outputs()
    assert restored == ("学生（含学生版和教师版）",)
    assert "0728" in document_text(student_path)
    assert "0729" not in document_text(student_path)
    assert "0728" in document_text(teacher_path)
    assert "0729" not in document_text(teacher_path)


def test_restore_accepts_v3_legacy_student_backup(tmp_path: Path) -> None:
    backup = tmp_path / "backup" / "张三" / "物理"
    backup.mkdir(parents=True)
    (backup / "legacy.txt").write_text("legacy", encoding="utf-8")
    current = tmp_path / "学生" / "old"
    current.mkdir(parents=True)

    restored = core.restore_outputs(tmp_path, tmp_path / "backup")

    assert restored == ("学生（旧版备份）",)
    assert (
        tmp_path / "学生" / "张三" / "物理" / "legacy.txt"
    ).read_text(encoding="utf-8") == "legacy"


def test_problem_number_parser_accepts_mixed_separators_and_deduplicates() -> None:
    assert core._parse_problem_numbers("1，2, 3  2 4.0") == [1, 2, 3, 4]


def test_duplicate_student_names_are_rejected(tmp_path: Path) -> None:
    feedback = tmp_path / "feedback.xlsx"
    write_feedback(feedback, [("张三", "1"), ("张三", "2")])

    with pytest.raises(core.GenerateError, match="重复姓名: 张三"):
        core._read_feedback(str(feedback), (1, 2))
