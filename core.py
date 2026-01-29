"""核心生成逻辑 - 错题集生成器 v3.0"""

import os
import re
import shutil
from typing import Optional

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from PyQt5.QtWidgets import QMessageBox
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Cm

print("错题集生成器 v3.0")
print("作者：Leonard")

# 错题号分隔符：中英文逗号、空格
SPLIT_PATTERN = r'[,，\s]+'


class GenerateError(Exception):
    """生成过程中的错误"""
    pass


def _validate_excel(file_path: str) -> None:
    """验证Excel文件有效性"""
    if not os.path.isfile(file_path):
        raise GenerateError(f"错题反馈表不存在: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ('.xlsx', '.xls'):
        raise GenerateError(f"不支持的文件格式: {ext}，请使用.xlsx或.xls文件")


def _validate_folder(folder_path: str) -> list[str]:
    """验证截图文件夹并返回图片列表"""
    if not os.path.isdir(folder_path):
        raise GenerateError(f"截图文件夹不存在: {folder_path}")

    images = []
    for f in os.listdir(folder_path):
        if f.lower().endswith(('.png', '.jpg', '.jpeg')):
            name = os.path.splitext(f)[0]
            if name.isdigit():
                images.append(f)

    if not images:
        raise GenerateError("截图文件夹中没有找到有效的题目图片（需要数字命名的png/jpg文件）")

    return sorted(images)


def _parse_problem_numbers(pn_str: str) -> list[int]:
    """解析错题号字符串，返回题号列表"""
    if not pn_str or pn_str.strip() in ('', 'None'):
        return []

    parts = re.split(SPLIT_PATTERN, pn_str.strip())
    numbers = []

    for part in parts:
        part = part.strip()
        if not part:
            continue
        try:
            # 支持浮点数格式（如Excel中的1.0）
            num = int(float(part))
            if num > 0:
                numbers.append(num)
        except ValueError:
            # 忽略无法解析的部分
            continue

    return numbers


def _get_image_path(folder: str, number: int) -> Optional[str]:
    """根据题号获取图片路径"""
    # 尝试不同格式：01.png, 1.png, 01.jpg, 1.jpg
    candidates = [
        f"{number:02d}.png",
        f"{number}.png",
        f"{number:02d}.jpg",
        f"{number}.jpg",
        f"{number:02d}.jpeg",
        f"{number}.jpeg",
    ]

    for name in candidates:
        path = os.path.join(folder, name)
        if os.path.isfile(path):
            return path

    return None


def _backup_students(student_dir: str, backup_dir: str, log_func) -> None:
    """备份学生文件夹"""
    if os.path.exists(backup_dir):
        shutil.rmtree(backup_dir)

    if os.path.exists(student_dir):
        shutil.copytree(student_dir, backup_dir)
        log_func("已创建备份")
    else:
        log_func("学生文件夹不存在，跳过备份")


def _create_student_doc(stupath: str, classes: str, name: str, subject: str, weeknumber: str) -> Document:
    """创建新的学生错题集文档"""
    doc = Document()

    # 设置页边距
    section = doc.sections[0]
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    # 设置页眉
    header = section.header
    header.paragraphs[0].text = f"{classes}    {name}    {subject}    第{weeknumber}周"
    header.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    doc.save(stupath)
    return doc


def generate(jyfoldername: str, cuotifkbname: str, classes: str,
             subject: str, weeknumber: str, date: str, widget) -> None:
    """
    生成错题集主函数

    Args:
        jyfoldername: 截图文件夹路径
        cuotifkbname: 错题反馈表路径
        classes: 班级名称
        subject: 科目
        weeknumber: 周数
        date: 日期 (MMDD格式)
        widget: UI适配器，需要有textBrowser.append方法
    """
    log = widget.textBrowser.append

    # 验证输入
    log("正在验证输入文件...")
    _validate_excel(cuotifkbname)
    available_images = _validate_folder(jyfoldername)
    log(f"找到 {len(available_images)} 张题目图片")

    # 打开Excel
    log("正在读取错题反馈表...")
    try:
        wb = openpyxl.load_workbook(cuotifkbname, read_only=True, data_only=True)
    except InvalidFileException as e:
        raise GenerateError(f"无法打开Excel文件: {e}")
    except Exception as e:
        raise GenerateError(f"读取Excel失败: {e}")

    sheet = wb.active
    rows = list(sheet.iter_rows(min_row=1, values_only=True))
    wb.close()

    if len(rows) < 2:
        raise GenerateError("错题反馈表为空或只有表头")

    # 提取姓名列（跳过表头）
    students = []
    for i, row in enumerate(rows[1:], start=1):
        if row[0] is None or str(row[0]).strip() == '':
            continue
        name = str(row[0]).strip()
        pn_raw = str(row[1]) if len(row) > 1 and row[1] is not None else ''
        students.append((i, name, pn_raw))

    if not students:
        raise GenerateError("错题反馈表中没有找到有效的学生数据")

    log(f"共 {len(students)} 名学生")

    # 备份
    student_dir = os.path.join(".", "学生")
    backup_dir = os.path.join(".", "backup")
    _backup_students(student_dir, backup_dir, log)

    # 处理每个学生
    docs_to_save = []  # (path, doc) 延迟保存

    for idx, name, pn_raw in students:
        studir = os.path.join(student_dir, name, subject)
        stupath = os.path.join(studir, f"第{weeknumber}周错题集.docx")

        # 确保目录存在
        os.makedirs(studir, exist_ok=True)

        # 加载或创建文档
        if os.path.exists(stupath):
            doc = Document(stupath)
        else:
            doc = _create_student_doc(stupath, classes, name, subject, weeknumber)

        # 添加日期标题
        doc.add_heading(date, level=1)

        # 处理错题
        pn_clean = pn_raw.strip()

        if pn_clean in ('未带', '未交'):
            # 未交作业：添加所有题目
            log(f"{name} {date}日 {subject} {pn_clean}")
            for img in available_images:
                img_path = os.path.join(jyfoldername, img)
                doc.add_picture(img_path, width=Cm(18))

        elif pn_clean == '全对':
            log(f"{name} {date}日 {subject} 全对")
            doc.add_paragraph("全对")

        elif pn_clean in ('', 'None', '无'):
            log(f"<font color='orange'>警告：</font>{name} 的错题信息为空，跳过")
            doc.add_paragraph("（无错题信息）")

        else:
            # 解析错题号
            problem_numbers = _parse_problem_numbers(pn_clean)

            if not problem_numbers:
                log(f"<font color='orange'>警告：</font>{name} 的错题号无法解析: {pn_clean}")
                doc.add_paragraph(f"（无法解析: {pn_clean}）")
            else:
                missing = []
                for num in problem_numbers:
                    img_path = _get_image_path(jyfoldername, num)
                    if img_path:
                        doc.add_picture(img_path, width=Cm(18))
                    else:
                        missing.append(num)

                if missing:
                    missing_str = ', '.join(str(n) for n in missing)
                    log(f"<font color='red'>警告：</font>{name} 缺少题目图片: {missing_str}")

                    # 询问是否继续
                    result = QMessageBox.warning(
                        widget,
                        '缺少图片',
                        f'{name} 的错题 {missing_str} 没有找到对应图片。\n是否忽略并继续？',
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes
                    )
                    if result != QMessageBox.Yes:
                        raise GenerateError(f"用户取消：缺少题目图片 {missing_str}")

                log(f"{name} {date}日错题集已完成")

        docs_to_save.append((stupath, doc))

    # 批量保存
    log("<font color='blue'>开始保存错题集...</font>")
    for stupath, doc in docs_to_save:
        doc.save(stupath)
        name = os.path.basename(os.path.dirname(os.path.dirname(stupath)))
        log(f"{name} 的 {subject} 错题集已保存")

    log(f"<font color='blue'>本次处理学生人数：{len(students)}</font>")

    # 保存记录
    log("正在保存记录...")
    record_path = '记录.docx'
    if os.path.exists(record_path):
        record = Document(record_path)
    else:
        record = Document()

    record.add_heading(date, level=1)
    record.add_paragraph(f"截图文件夹: {jyfoldername}")
    record.add_paragraph(f"反馈表: {cuotifkbname}")
    record.add_paragraph(f"科目: {subject}")
    record.add_paragraph(f"班级: {classes}")
    record.add_paragraph(f"周数: {weeknumber}")
    record.add_paragraph("")
    record.save(record_path)

    log("<font color='green'>生成完成！</font>")
