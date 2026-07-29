# 错题集生成器 v3.1

面向教育培训场景的桌面工具。根据一份 Excel 错题反馈表和两套一一对应的题目图片，一次为每名学生生成学生版与教师版错题集。

## v3.1 的核心变化

旧版需要分别运行两次，并把科目填写为“物理学生版”和“物理教师版”。v3.1 将这一流程整合为一次操作：

- 选择学生版图片文件夹；
- 选择教师版图片文件夹；
- 读取一次错题反馈表；
- 每位学生同时得到两份结构相同、图片来源不同的文档。

输出结构：

```text
学生/
└── 张三/
    ├── 物理学生版/
    │   └── 第3周错题集.docx
    └── 物理教师版/
        └── 第3周错题集.docx
```

因此可以直接压缩整个学生姓名目录发给家长。

## 功能

- 两套图片按题号严格对应，生成前自动校验；
- 支持 PNG、JPG、JPEG，文件名使用 `1.png`、`01.png` 等数字格式；
- 同一目录中重复题号会被拒绝，避免误用图片；
- 支持普通错题号、`全对`、`未交`、`未带`；
- 同一周后续日期会追加到已有文档；
- 每次生成前备份完整 `学生/` 目录，可同时恢复两版；
- 自动迁移 v3.0 的单图片目录配置和旧备份；
- 深色/浅色主题、配置记忆和生成日志。

## 输入格式

### 两套图片

两套目录必须具有相同题号，但图片内容可以不同：

```text
学生版图片/                 教师版图片/
├── 01.png                  ├── 01.png
├── 02.png                  ├── 02.jpg
└── 03.png                  └── 03.png
```

### 错题反馈表

使用 `.xlsx` 文件，第一列为姓名，第二列为错题信息：

| 姓名 | 错题号 |
|---|---|
| 张三 | 1,3,5 |
| 李四 | 全对 |
| 王五 | 未交 |

错题号支持中英文逗号或空格分隔。`未交`、`未带`会在两版中插入各自目录的全部题目图片。

## 安装与运行

要求 Python 3.10+。

```bash
git clone https://github.com/Leonardlsa/problems-folder-generator.git
cd problems-folder-generator
python -m venv .venv

# Linux / macOS
source .venv/bin/activate

# Windows PowerShell
# .venv\Scripts\Activate.ps1

pip install -r requirements.txt
python main.py
```

## 测试

```bash
pip install -e ".[dev]"
pytest -q
```

测试覆盖两套图片隔离、题号对应校验、特殊状态、缺图取消、周内追加、双版本备份恢复和 v3.0 旧备份兼容。

## CI 与发布

- 每次 push 到 `master` 或创建 PR 时，CI 会在 Python 3.10、3.11、3.12 上运行编译检查和测试。
- 推送 `v*` 标签时，GitHub Actions 会先测试，再构建：
  - Windows x86_64 单文件 EXE；
  - Linux x86_64 可执行文件压缩包；
  - Python wheel 和 source distribution；
  - SHA-256 校验文件；
  - GitHub Release。

发布 v3.1：

```bash
git tag -a v3.1.0 -m "Problems Folder Generator v3.1.0"
git push origin v3.1.0
```

## 项目结构

```text
PFG/
├── main.py
├── core.py
├── version.py
├── ui/
├── tests/
├── .github/workflows/
├── requirements.txt
├── requirements-dev.txt
└── pyproject.toml
```

## 许可证

[Apache License 2.0](LICENSE)

## 致谢

项目初版由 **Leonard** 于 2022 年在 **武汉天任教育** 开发，**Lrj** 参与命令行初版开发。v3.1 恢复并整合了旧版通过“科目学生版/科目教师版”区分两套材料的实际工作流。
