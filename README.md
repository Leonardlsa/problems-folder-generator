# 错题集生成器 v3.0

一款自动化生成学生错题集的桌面工具，专为教育培训场景设计。

<p align="center">
  <img src="assets/screenshot_dark.png" alt="深色主题" width="400"/>
  <img src="assets/screenshot_light.png" alt="浅色主题" width="400"/>
</p>

## 功能特性

- **自动生成错题集**：根据错题反馈表，自动为每位学生生成个性化的错题集文档
- **多科目支持**：数学、物理、化学、生物
- **深色/浅色主题**：护眼暗色主题，支持一键切换
- **备份恢复**：每次操作前自动备份，可一键恢复
- **配置记忆**：自动保存上次输入的班级、日期、周数等信息
- **操作日志**：自动保存每次生成的详细日志

## 快速开始

### 环境要求

- Python 3.10+
- Windows / Linux / macOS

### 安装

```bash
# 克隆仓库
git clone https://github.com/Leonardlsa/problems-folder-generator.git
cd problems-folder-generator

# 创建虚拟环境并安装依赖
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# 或 .venv\Scripts\activate  # Windows

pip install -r requirements.txt
```

### 运行

```bash
python main.py
```

## 使用说明

### 1. 准备截图文件夹

将讲义题目截图保存为 PNG/JPG 格式，以题号命名：

```
截图文件夹/
├── 01.png
├── 02.png
├── 03.png
└── ...
```

### 2. 准备错题反馈表

创建 Excel 文件（.xlsx），格式如下：

| 姓名 | 错题号 |
|------|--------|
| 张三 | 1,3,5 |
| 李四 | 全对 |
| 王五 | 未交 |

**错题号格式：**
- 数字列表：`1,3,5` 或 `1，3，5`（支持中英文逗号）
- 全对：`全对`
- 未交/未带：`未交` 或 `未带`（将包含所有题目）

### 3. 填写信息并生成

1. 选择科目、填写班级
2. 输入日期（MMDD）和周数
3. 选择截图文件夹和错题反馈表
4. 点击 **GO!** 开始生成

生成的错题集保存在 `学生/[姓名]/[科目]/第X周错题集.docx`

## 项目结构

```
PFG/
├── main.py              # 程序入口
├── core.py              # 核心生成逻辑
├── ui/                  # 用户界面模块
│   ├── __init__.py
│   ├── main_window.py   # 主窗口
│   ├── styles.py        # 主题样式
│   ├── widgets.py       # 自定义控件
│   └── icons/           # 图标资源
├── requirements.txt     # Python 依赖
├── pyproject.toml       # 项目配置
└── LICENSE              # Apache 2.0 许可证
```

## 构建可执行文件

### 使用 PyInstaller

```bash
pip install pyinstaller

# Windows
pyinstaller --onefile --windowed --name="错题集生成器" --add-data="ui/icons;ui/icons" main.py

# Linux
pyinstaller --onefile --windowed --name="cuotiji-generator" --add-data="ui/icons:ui/icons" main.py
```

生成的可执行文件位于 `dist/` 目录。

## 许可证

[Apache License 2.0](LICENSE)

## 致谢

本项目初版由 **Leonard** 于 2022 年在 **武汉天任教育** 开发，**Lrj** 参与了命令行初版的开发工作。

**v3.0** 是 2026 年 1 月由 Leonard 翻新的第三个版本，也是第二个图形化版本。

感谢孙总、硕姐和万幸哥对我的肯定和对本项目研发的支持。

希望这个工具能为后续的朋友们减轻繁复的劳动，祝大家越过越开心！

## 作者

**Leonard** - [GitHub](https://github.com/Leonardlsa)
