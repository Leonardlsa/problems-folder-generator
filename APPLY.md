# 应用 v3.1 修改

本压缩包只包含新增或需要替换的文件，不包含仓库原有图标、截图和许可证等未改动资源。

在仓库根目录执行：

```bash
# 建议先创建分支
git switch -c feat/v3.1-coupled-student-teacher-output

# 将压缩包解压后，覆盖到仓库根目录
# Windows 可直接在文件管理器中合并覆盖；Linux/macOS 示例：
unzip -o problems-folder-generator-v3.1-changes.zip -d .

python -m compileall -q core.py main.py version.py ui
pip install pytest pillow openpyxl python-docx
pytest -q

git add core.py main.py version.py ui/main_window.py README.md CHANGELOG.md \
  pyproject.toml tests .github/workflows

git commit -m "feat: generate coupled student and teacher versions"
git push -u origin feat/v3.1-coupled-student-teacher-output
```

输出结构：

```text
学生/<姓名>/<科目>学生版/第X周错题集.docx
学生/<姓名>/<科目>教师版/第X周错题集.docx
```
