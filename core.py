"""核心生成逻辑 - 错题集生成器 v3.1。"""

from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Protocol

import openpyxl
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Cm
from openpyxl.utils.exceptions import InvalidFileException

from version import __version__

APP_VERSION = __version__
SPLIT_PATTERN = r"[,，\s]+"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
VERSION_SUFFIXES = (("学生版", "student"), ("教师版", "teacher"))


class GenerateError(Exception):
    """生成过程中的可预期错误。"""


class LogTarget(Protocol):
    textBrowser: object


@dataclass(frozen=True)
class Student:
    name: str
    problem_text: str


def _log_func(widget: LogTarget) -> Callable[[str], None]:
    try:
        return widget.textBrowser.append
    except AttributeError as exc:
        raise GenerateError("内部错误：日志组件不可用") from exc


def _validate_excel(file_path: str) -> None:
    path = Path(file_path)
    if not path.is_file():
        raise GenerateError(f"错题反馈表不存在: {file_path}")
    if path.suffix.lower() != ".xlsx":
        raise GenerateError(f"不支持的文件格式: {path.suffix}，请使用 .xlsx 文件")


def _image_number(file_name: str) -> int | None:
    stem = Path(file_name).stem
    if not stem.isdigit():
        return None
    number = int(stem)
    return number if number > 0 else None


def _validate_folder(folder_path: str, label: str) -> dict[int, str]:
    """验证截图目录并返回题号到图片路径的映射。"""
    folder = Path(folder_path)
    if not folder.is_dir():
        raise GenerateError(f"{label}截图文件夹不存在: {folder_path}")

    images: dict[int, str] = {}
    duplicates: list[int] = []
    for entry in folder.iterdir():
        if not entry.is_file() or entry.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        number = _image_number(entry.name)
        if number is None:
            continue
        if number in images:
            duplicates.append(number)
            continue
        images[number] = str(entry)

    if duplicates:
        duplicate_text = ", ".join(str(number) for number in sorted(set(duplicates)))
        raise GenerateError(f"{label}截图文件夹存在重复题号: {duplicate_text}")
    if not images:
        raise GenerateError(
            f"{label}截图文件夹中没有找到有效题目图片"
            "（请使用 01.png、02.jpg 等数字文件名）"
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
    only_student = sorted(student_numbers - teacher_numbers)
    only_teacher = sorted(teacher_numbers - student_numbers)
    if only_student:
        details.append("仅学生版有: " + ", ".join(map(str, only_student)))
    if only_teacher:
        details.append("仅教师版有: " + ", ".join(map(str, only_teacher)))
    raise GenerateError("学生版与教师版图片题号不对应；" + "；".join(details))


def _parse_problem_numbers(problem_text: str) -> list[int]:
    if not problem_text or problem_text.strip() in ("", "None"):
        return []

    numbers: list[int] = []
    for part in re.split(SPLIT_PATTERN, problem_text.strip()):
        if not part:
            continue
        try:
            number = int(float(part))
        except ValueError:
            continue
        if number > 0 and number not in numbers:
            numbers.append(number)
    return numbers


def _read_students(feedback_file: str) -> list[Student]:
    try:
        workbook = openpyxl.load_workbook(feedback_file, read_only=True, data_only=True)
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

    students: list[Student] = []
    for row in rows[1:]:
        if not row or row[0] is None or not str(row[0]).strip():
            continue
        name = str(row[0]).strip()
        problem_text = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        students.append(Student(name=name, problem_text=problem_text))

    if not students:
        raise GenerateError("错题反馈表中没有找到有效的学生数据")
    return students


def _backup_students(student_dir: Path, backup_dir: Path, log: Callable[[str], None]) -> None:
    if backup_dir.exists():
        shutil.rmtree(backup_dir)
    if student_dir.exists():
        shutil.copytree(student_dir, backup_dir)
        log("已创建学生目录备份")
    else:
        log("学生文件夹不存在，跳过备份")


def _new_student_doc(classes: str, name: str, version_subject: str, week: str) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    header = section.header.paragraphs[0]
    header.text = f"{classes}    {name}    {version_subject}    第{week}周"
    header.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    return doc


def _load_or_create_doc(
    path: Path, classes: str, name: str, version_subject: str, week: str
) -> Document:
    if path.exists():
        return Document(str(path))
    return _new_student_doc(classes, name, version_subject, week)


def _confirm_missing(widget: LogTarget, name: str, version_subject: str, missing: list[int]) -> bool:
    """询问是否忽略缺图；无 UI 确认能力时默认终止，避免静默漏题。"""
    callback = getattr(widget, "confirm_missing_images", None)
    if callback is None:
        return False
    return bool(callback(name, version_subject, missing))


def _append_student_content(
    doc: Document,
    student: Student,
    date: str,
    version_subject: str,
    images: dict[int, str],
    widget: LogTarget,
    log: Callable[[str], None],
    missing_confirmed: bool = False,
) -> None:
    doc.add_heading(date, level=1)
    problem_text = student.problem_text.strip()

    if problem_text in ("未带", "未交"):
        log(f"{student.name} {date}日 {version_subject} {problem_text}")
        for image_path in images.values():
            doc.add_picture(image_path, width=Cm(18))
        return

    if problem_text == "全对":
        log(f"{student.name} {date}日 {version_subject} 全对")
        doc.add_paragraph("全对")
        return

    if problem_text in ("", "None", "无"):
        log(f"<font color='orange'>警告：</font>{student.name} 的错题信息为空，跳过")
        doc.add_paragraph("（无错题信息）")
        return

    problem_numbers = _parse_problem_numbers(problem_text)
    if not problem_numbers:
        log(f"<font color='orange'>警告：</font>{student.name} 的错题号无法解析: {problem_text}")
        doc.add_paragraph(f"（无法解析: {problem_text}）")
        return

    missing = [number for number in problem_numbers if number not in images]
    if missing:
        missing_text = ", ".join(map(str, missing))
        log(
            f"<font color='red'>警告：</font>{student.name} 的 {version_subject} "
            f"缺少题目图片: {missing_text}"
        )
        if not missing_confirmed and not _confirm_missing(widget, student.name, version_subject, missing):
            raise GenerateError(f"已取消：{student.name} 的 {version_subject} 缺少图片 {missing_text}")

    for number in problem_numbers:
        image_path = images.get(number)
        if image_path:
            doc.add_picture(image_path, width=Cm(18))

    log(f"{student.name} {date}日 {version_subject} 错题集已完成")


def _save_record(
    date: str,
    classes: str,
    subject: str,
    week: str,
    feedback_file: str,
    student_images_dir: str,
    teacher_images_dir: str,
) -> None:
    record_path = Path("记录.docx")
    record = Document(str(record_path)) if record_path.exists() else Document()
    record.add_heading(date, level=1)
    record.add_paragraph(f"学生版截图文件夹: {student_images_dir}")
    record.add_paragraph(f"教师版截图文件夹: {teacher_images_dir}")
    record.add_paragraph(f"反馈表: {feedback_file}")
    record.add_paragraph(f"科目: {subject}（学生版 + 教师版）")
    record.add_paragraph(f"班级: {classes}")
    record.add_paragraph(f"周数: {week}")
    record.add_paragraph("")
    record.save(str(record_path))


def generate(
    student_images_dir: str,
    teacher_images_dir: str,
    feedback_file: str,
    classes: str,
    subject: str,
    weeknumber: str,
    date: str,
    widget: LogTarget,
) -> None:
    """一次生成每名学生的学生版和教师版错题集。"""
    log = _log_func(widget)
    log("正在验证输入文件...")
    _validate_excel(feedback_file)
    student_images = _validate_folder(student_images_dir, "学生版")
    teacher_images = _validate_folder(teacher_images_dir, "教师版")
    _validate_corresponding_images(student_images, teacher_images)
    log(f"找到 {len(student_images)} 组一一对应的题目图片")

    log("正在读取错题反馈表...")
    students = _read_students(feedback_file)
    log(f"共 {len(students)} 名学生；每人生成学生版和教师版各一份")

    student_root = Path("学生")
    backup_root = Path("backup")
    _backup_students(student_root, backup_root, log)

    version_inputs = (
        ("学生版", f"{subject}学生版", student_images),
        ("教师版", f"{subject}教师版", teacher_images),
    )
    docs_to_save: list[tuple[Path, Document, str, str]] = []

    for student in students:
        problem_text = student.problem_text.strip()
        missing_confirmed = False
        if problem_text not in ("", "None", "无", "全对", "未带", "未交"):
            requested_numbers = _parse_problem_numbers(problem_text)
            missing_numbers = [number for number in requested_numbers if number not in student_images]
            if missing_numbers:
                missing_text = ", ".join(map(str, missing_numbers))
                log(
                    f"<font color='red'>警告：</font>{student.name} 的学生版和教师版 "
                    f"均缺少题目图片: {missing_text}"
                )
                combined_subject = f"{subject}学生版和教师版"
                if not _confirm_missing(widget, student.name, combined_subject, missing_numbers):
                    raise GenerateError(
                        f"已取消：{student.name} 的学生版和教师版缺少图片 {missing_text}"
                    )
                missing_confirmed = True

        for version_label, version_subject, images in version_inputs:
            output_dir = student_root / student.name / version_subject
            output_path = output_dir / f"第{weeknumber}周错题集.docx"
            output_dir.mkdir(parents=True, exist_ok=True)
            doc = _load_or_create_doc(
                output_path, classes, student.name, version_subject, weeknumber
            )
            _append_student_content(
                doc,
                student,
                date,
                version_subject,
                images,
                widget,
                log,
                missing_confirmed=missing_confirmed,
            )
            docs_to_save.append((output_path, doc, student.name, version_label))

    log("<font color='blue'>开始保存学生版和教师版错题集...</font>")
    for output_path, doc, student_name, version_label in docs_to_save:
        doc.save(str(output_path))
        log(f"{student_name} 的{version_label}错题集已保存")

    _save_record(
        date,
        classes,
        subject,
        weeknumber,
        feedback_file,
        student_images_dir,
        teacher_images_dir,
    )
    log(f"<font color='blue'>本次处理学生人数：{len(students)}，生成文档数：{len(docs_to_save)}</font>")
    log("<font color='green'>学生版和教师版均已生成完成！</font>")
