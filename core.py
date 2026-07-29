"""Core generation logic for 错题集生成器 v3.1."""

from __future__ import annotations

import os
import re
import shutil
import tempfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

import openpyxl
from docx import Document
from docx.document import Document as DocumentType
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Cm
from openpyxl.utils.exceptions import InvalidFileException

from version import APP_VERSION

SPLIT_PATTERN = r"[,，\s]+"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}


class GenerateError(Exception):
    """Raised when generation cannot continue safely."""


class GenerateWidget(Protocol):
    """Minimal UI surface required by the core generator."""

    textBrowser: object

    def confirm_missing_images(
        self, missing_by_student: dict[str, list[int]]
    ) -> bool: ...


@dataclass(frozen=True)
class StudentRecord:
    """Normalized feedback for one student."""

    name: str
    raw_feedback: str
    problem_numbers: tuple[int, ...]
    status: str


@dataclass(frozen=True)
class GenerationResult:
    """Paths written by one successful coupled generation."""

    student_version_paths: tuple[str, ...]
    teacher_version_paths: tuple[str, ...]
    record_path: str


def _validate_excel(file_path: str) -> None:
    path = Path(file_path)
    if not path.is_file():
        raise GenerateError(f"错题反馈表不存在: {file_path}")
    if path.suffix.lower() != ".xlsx":
        raise GenerateError(f"不支持的文件格式: {path.suffix}，请使用 .xlsx 文件")


def _discover_images(folder_path: str, label: str) -> dict[int, str]:
    """Return a strict problem-number-to-image mapping for one edition."""
    folder = Path(folder_path)
    if not folder.is_dir():
        raise GenerateError(f"{label}图片文件夹不存在: {folder_path}")

    images: dict[int, str] = {}
    duplicates: list[int] = []
    for entry in sorted(folder.iterdir(), key=lambda item: item.name.lower()):
        if not entry.is_file() or entry.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        if not entry.stem.isdigit():
            continue

        number = int(entry.stem)
        if number <= 0:
            continue
        if number in images:
            duplicates.append(number)
            continue
        images[number] = str(entry)

    if duplicates:
        duplicate_text = "、".join(map(str, sorted(set(duplicates))))
        raise GenerateError(
            f"{label}图片文件夹存在重复题号: {duplicate_text}。"
            "请确保每道题只有一张图片。"
        )
    if not images:
        raise GenerateError(
            f"{label}图片文件夹中没有有效题目图片。"
            "请使用 01.png、02.jpg 等数字文件名。"
        )
    return dict(sorted(images.items()))


def _validate_corresponding_images(
    student_images: dict[int, str], teacher_images: dict[int, str]
) -> None:
    student_numbers = set(student_images)
    teacher_numbers = set(teacher_images)
    if student_numbers == teacher_numbers:
        return

    details: list[str] = []
    student_only = sorted(student_numbers - teacher_numbers)
    teacher_only = sorted(teacher_numbers - student_numbers)
    if student_only:
        details.append("仅学生版有: " + "、".join(map(str, student_only)))
    if teacher_only:
        details.append("仅教师版有: " + "、".join(map(str, teacher_only)))
    raise GenerateError("学生版与教师版图片题号不对应；" + "；".join(details))


def _parse_problem_numbers(value: str) -> list[int]:
    """Parse a feedback cell into unique positive problem numbers."""
    if not value or value.strip() in ("", "None"):
        return []

    numbers: list[int] = []
    seen: set[int] = set()
    for part in re.split(SPLIT_PATTERN, value.strip()):
        if not part:
            continue
        try:
            number = int(float(part))
        except ValueError:
            continue
        if number > 0 and number not in seen:
            numbers.append(number)
            seen.add(number)
    return numbers


def _read_feedback(
    file_path: str, all_problem_numbers: tuple[int, ...]
) -> list[StudentRecord]:
    """Read and normalize the first worksheet of the feedback workbook."""
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    except InvalidFileException as exc:
        raise GenerateError(f"无法打开 Excel 文件: {exc}") from exc
    except Exception as exc:
        raise GenerateError(f"读取 Excel 失败: {exc}") from exc

    try:
        rows = list(workbook.active.iter_rows(min_row=1, values_only=True))
    finally:
        workbook.close()

    if len(rows) < 2:
        raise GenerateError("错题反馈表为空或只有表头")

    records: list[StudentRecord] = []
    seen_names: set[str] = set()
    duplicate_names: list[str] = []

    for row in rows[1:]:
        if not row or row[0] is None or not str(row[0]).strip():
            continue

        name = str(row[0]).strip()
        if name in seen_names:
            duplicate_names.append(name)
            continue
        seen_names.add(name)

        raw = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        if raw in ("未带", "未交"):
            status = "unsubmitted"
            numbers = all_problem_numbers
        elif raw == "全对":
            status = "all_correct"
            numbers = ()
        elif raw in ("", "None", "无"):
            status = "empty"
            numbers = ()
        else:
            parsed = tuple(_parse_problem_numbers(raw))
            status = "errors" if parsed else "invalid"
            numbers = parsed

        records.append(StudentRecord(name, raw, tuple(numbers), status))

    if duplicate_names:
        names = "、".join(dict.fromkeys(duplicate_names))
        raise GenerateError(f"错题反馈表中存在重复姓名: {names}")
    if not records:
        raise GenerateError("错题反馈表中没有找到有效的学生数据")
    return records


def _new_document(header_text: str) -> DocumentType:
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    header = section.header.paragraphs[0]
    header.text = header_text
    header.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    return document


def _load_document(
    path: Path,
    classes: str,
    name: str,
    edition_subject: str,
    weeknumber: str,
) -> DocumentType:
    if path.exists():
        return Document(path)
    return _new_document(
        f"{classes}    {name}    {edition_subject}    第{weeknumber}周"
    )


def _format_missing_message(missing_by_student: dict[str, list[int]]) -> str:
    lines = ["以下错题在学生版和教师版图片中均不存在："]
    for name, numbers in list(missing_by_student.items())[:10]:
        lines.append(f"{name}: {', '.join(map(str, numbers))}")
    if len(missing_by_student) > 10:
        lines.append(f"……另有 {len(missing_by_student) - 10} 名学生")
    lines.append("\n是否忽略缺失图片并继续生成两版文档？")
    return "\n".join(lines)


def _confirm_missing_images(
    widget: GenerateWidget, missing_by_student: dict[str, list[int]]
) -> None:
    if not missing_by_student:
        return
    callback = getattr(widget, "confirm_missing_images", None)
    if not callable(callback) or not callback(missing_by_student):
        raise GenerateError("用户取消：反馈表中包含缺少图片的题号")


def _append_student_content(
    document: DocumentType,
    record: StudentRecord,
    image_paths: dict[int, str],
    date: str,
    edition_subject: str,
    log: Callable[[str], None],
) -> None:
    document.add_heading(date, level=1)

    if record.status == "all_correct":
        document.add_paragraph("全对")
        log(f"{record.name} {date}日 {edition_subject} 全对")
        return

    if record.status == "empty":
        document.add_paragraph("（无错题信息）")
        log(f"<font color='orange'>警告：</font>{record.name} 的错题信息为空")
        return

    if record.status == "invalid":
        document.add_paragraph(f"（无法解析: {record.raw_feedback}）")
        log(
            f"<font color='orange'>警告：</font>{record.name} "
            f"的错题号无法解析: {record.raw_feedback}"
        )
        return

    for number in record.problem_numbers:
        image_path = image_paths.get(number)
        if image_path:
            document.add_picture(image_path, width=Cm(18))

    missing = [number for number in record.problem_numbers if number not in image_paths]
    if missing:
        document.add_paragraph(f"（缺少题目图片：{', '.join(map(str, missing))}）")

    if record.status == "unsubmitted":
        log(f"{record.name} {date}日 {edition_subject} {record.raw_feedback}")
    else:
        log(f"{record.name} {date}日 {edition_subject} 错题集已完成")


def _build_record_document(
    path: Path,
    student_images_dir: str,
    teacher_images_dir: str,
    feedback_file: str,
    classes: str,
    subject: str,
    weeknumber: str,
    date: str,
) -> DocumentType:
    record = Document(path) if path.exists() else Document()
    record.add_heading(date, level=1)
    record.add_paragraph(f"学生版图片文件夹: {student_images_dir}")
    record.add_paragraph(f"教师版图片文件夹: {teacher_images_dir}")
    record.add_paragraph(f"反馈表: {feedback_file}")
    record.add_paragraph(f"科目: {subject}（学生版 + 教师版）")
    record.add_paragraph(f"班级: {classes}")
    record.add_paragraph(f"周数: {weeknumber}")
    record.add_paragraph("")
    return record


def _save_documents_atomically(
    documents: list[tuple[Path, DocumentType]]
) -> None:
    """Write every document to a temp file before replacing any target."""
    temporary_files: list[tuple[Path, Path]] = []
    try:
        for final_path, document in documents:
            final_path.parent.mkdir(parents=True, exist_ok=True)
            handle, temporary_name = tempfile.mkstemp(
                prefix=f".{final_path.name}.",
                suffix=".tmp",
                dir=final_path.parent,
            )
            os.close(handle)
            temporary_path = Path(temporary_name)
            document.save(temporary_path)
            temporary_files.append((temporary_path, final_path))

        for temporary_path, final_path in temporary_files:
            os.replace(temporary_path, final_path)
    except Exception:
        for temporary_path, _ in temporary_files:
            temporary_path.unlink(missing_ok=True)
        raise


def _backup_outputs(
    student_root: Path, backup_dir: Path, log: Callable[[str], None]
) -> None:
    """Back up the complete per-student delivery tree as one snapshot."""
    if backup_dir.exists():
        shutil.rmtree(backup_dir)
    if not student_root.exists():
        log("学生目录不存在，跳过备份")
        return

    backup_dir.mkdir(parents=True)
    shutil.copytree(student_root, backup_dir / "学生")
    log("已备份完整学生目录（含学生版和教师版）")


def restore_outputs(
    base_dir: str | Path = ".", backup_dir: str | Path = "backup"
) -> tuple[str, ...]:
    """Restore the complete per-student tree, including v3 legacy backups."""
    base_path = Path(base_dir)
    backup_path = Path(backup_dir)
    if not backup_path.exists():
        raise GenerateError("备份不存在")

    student_path = base_path / "学生"
    if student_path.exists():
        shutil.rmtree(student_path)

    structured_source = backup_path / "学生"
    if structured_source.exists():
        shutil.copytree(structured_source, student_path)
        restored = ("学生（含学生版和教师版）",)
    else:
        # v3.0 stored the contents of 学生 directly in backup/.
        shutil.copytree(backup_path, student_path)
        restored = ("学生（旧版备份）",)

    shutil.rmtree(backup_path)
    return restored


def generate(
    student_images_dir: str,
    teacher_images_dir: str,
    feedback_file: str,
    classes: str,
    subject: str,
    weeknumber: str,
    date: str,
    widget: GenerateWidget,
) -> GenerationResult:
    """Generate two image-backed editions for every student in one operation."""
    log = widget.textBrowser.append

    log("正在验证输入文件...")
    _validate_excel(feedback_file)
    student_images = _discover_images(student_images_dir, "学生版")
    teacher_images = _discover_images(teacher_images_dir, "教师版")
    _validate_corresponding_images(student_images, teacher_images)
    all_problem_numbers = tuple(student_images)
    log(f"找到 {len(student_images)} 组一一对应的题目图片")

    log("正在读取错题反馈表...")
    records = _read_feedback(feedback_file, all_problem_numbers)
    log(f"共 {len(records)} 名学生；每人生成学生版和教师版各一份")

    missing_by_student = {
        record.name: [
            number
            for number in record.problem_numbers
            if number not in student_images
        ]
        for record in records
        if any(number not in student_images for number in record.problem_numbers)
    }
    _confirm_missing_images(widget, missing_by_student)

    base_dir = Path(".")
    student_root = base_dir / "学生"
    record_path = base_dir / "记录.docx"
    edition_inputs = (
        ("学生版", f"{subject}学生版", student_images),
        ("教师版", f"{subject}教师版", teacher_images),
    )

    documents: list[tuple[Path, DocumentType]] = []
    student_version_paths: list[Path] = []
    teacher_version_paths: list[Path] = []

    for record in records:
        for edition_label, edition_subject, image_paths in edition_inputs:
            output_path = (
                student_root
                / record.name
                / edition_subject
                / f"第{weeknumber}周错题集.docx"
            )
            document = _load_document(
                output_path,
                classes,
                record.name,
                edition_subject,
                weeknumber,
            )
            _append_student_content(
                document,
                record,
                image_paths,
                date,
                edition_subject,
                log,
            )
            documents.append((output_path, document))
            if edition_label == "学生版":
                student_version_paths.append(output_path)
            else:
                teacher_version_paths.append(output_path)

    record_document = _build_record_document(
        record_path,
        student_images_dir,
        teacher_images_dir,
        feedback_file,
        classes,
        subject,
        weeknumber,
        date,
    )
    documents.append((record_path, record_document))

    _backup_outputs(student_root, base_dir / "backup", log)

    log("<font color='blue'>开始同时保存学生版和教师版...</font>")
    try:
        _save_documents_atomically(documents)
    except Exception as exc:
        raise GenerateError(f"保存文档失败: {exc}") from exc

    for record, student_path, teacher_path in zip(
        records, student_version_paths, teacher_version_paths
    ):
        log(f"{record.name} 的学生版已保存: {student_path}")
        log(f"{record.name} 的教师版已保存: {teacher_path}")

    log(
        f"<font color='blue'>本次处理学生人数：{len(records)}；"
        f"生成文档数：{len(student_version_paths) + len(teacher_version_paths)}</font>"
    )
    log("<font color='green'>学生版和教师版生成完成！</font>")

    return GenerationResult(
        tuple(str(path) for path in student_version_paths),
        tuple(str(path) for path in teacher_version_paths),
        str(record_path),
    )
