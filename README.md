# 错题集生成器 v3.1

一款面向教育培训场景的桌面工具，可根据错题反馈表，为每位学生同时生成**学生版**和**教师版**两份错题集。

## v3.1 核心变化

v3.1 恢复并整合了旧版的实际工作流：过去需要将科目分别填写为“物理学生版”和“物理教师版”，切换两套图片并运行两次；现在只需选择两套一一对应的图片文件夹，点击一次即可完成两版生成。

每名学生的输出结构如下：

```text
学生/
└── 张三/
    ├── 物理学生版/
    │   └── 第3周错题集.docx
    └── 物理教师版/
        └── 第3周错题集.docx
```

因此可以直接打包学生姓名目录发给家长。

## 功能特性

- **双版本同步生成**：同一份反馈表一次生成学生版和教师版
- **两套图片严格对应**：启动生成前校验两边题号集合是否一致
- **同结构不同图片**：两版使用相同学生、错题号、日期和排版，只替换图片来源
- **多科目支持**：数学、物理、化学、生物
- **特殊状态同步**：`全对`、`未交`、`未带`同时应用于两版
- **备份恢复**：每次操作前备份完整 `学生` 目录，可同时恢复两版文件
- **配置记忆**：保存两套图片目录、反馈表、班级、日期和周数
- **操作日志**：记录两套输入目录和每份文档的生成结果

## 快速开始

### 环境要求

- Python 3.10+
- Windows / Linux / macOS

### 安装与运行

```bash
git clone https://github.com/Leonardlsa/problems-folder-generator.git
cd problems-folder-generator
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows
pip install -r requirements.txt
python main.py
```

## 使用说明

### 1. 准备两套题目图片

学生版与教师版图片必须使用相同题号，扩展名可以不同：

```text
学生版图片/                 教师版图片/
├── 01.png                  ├── 01.png
├── 02.png                  ├── 02.jpg
└── 03.png                  └── 03.png
```

文件名支持 `1.png`、`01.png`、`01.jpg`、`01.jpeg` 等数字格式。若同一目录内同时出现 `1.png` 和 `01.jpg`，会被视为重复题号并停止生成。

### 2. 准备错题反馈表

使用 `.xlsx` 文件，第一列为姓名，第二列为错题号：

| 姓名 | 错题号 |
|------|--------|
| 张三 | 1,3,5 |
| 李四 | 全对 |
| 王五 | 未交 |

错题号支持：

- 数字列表：`1,3,5`、`1，3，5` 或空格分隔
- `全对`
- `未交` / `未带`：两版均插入各自目录中的全部题目图片

### 3. 一次生成两版

1. 选择科目并填写班级
2. 输入日期（MMDD）和周数
3. 分别选择学生版图片目录、教师版图片目录
4. 选择错题反馈表
5. 点击 **GO!**

程序将输出：

```text
学生/[姓名]/[科目]学生版/第X周错题集.docx
学生/[姓名]/[科目]教师版/第X周错题集.docx
```

若同一周文档已经存在，新日期内容会继续追加到原文档。

## 测试

```bash
pip install -e '.[dev]'
pytest -q
```

测试覆盖两版图片隔离、特殊状态、题号对应校验、缺图取消、历史追加和完整目录备份。

## 构建可执行文件

```bash
pip install pyinstaller

# Windows
pyinstaller --onefile --windowed --name="错题集生成器" --add-data="ui/icons;ui/icons" main.py

# Linux
pyinstaller --onefile --windowed --name="cuotiji-generator" --add-data="ui/icons:ui/icons" main.py
```

## 项目结构

```text
PFG/
├── main.py
├── core.py
├── version.py
├── ui/
├── tests/
├── requirements.txt
├── pyproject.toml
└── LICENSE
```

## 许可证

[Apache License 2.0](LICENSE)

## 致谢

项目初版由 **Leonard** 于 2022 年在 **武汉天任教育** 开发，**Lrj** 参与命令行初版开发。v3.1 将旧版通过“科目学生版/教师版”区分两套材料的方式整合为一次生成流程。
