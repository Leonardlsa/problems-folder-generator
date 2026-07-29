# Changelog

## 3.1.1 - 2026-07-29

### Added

- macOS Apple Silicon 与 Intel 的 `.app` / `.dmg` 自动构建。
- README 增加真实主界面截图、平台下载表、输入格式与恢复说明。

### Fixed

- Release 附件改用跨平台稳定的 ASCII 文件名，避免中文文件名前缀在 GitHub Actions 中丢失。
- Windows、Linux、macOS 的 `.sha256` 文件统一为可跨平台校验的格式。
- 发布源码 tag、版本号与所有二进制产物严格对应同一提交。

## 3.1.0 - 2026-07-29

### Added

- 学生版和教师版两套图片目录输入。
- 一次读取反馈表，同时为每名学生生成两版错题集。
- 输出目录采用 `[科目]学生版` 与 `[科目]教师版`，便于按学生整体打包。
- 两套图片题号集合一致性和重复题号校验。
- push/PR 多版本测试 CI。
- Windows、Linux、macOS（Apple Silicon / Intel）、wheel、sdist 和 GitHub Release 自动构建。
- 双图片隔离、特殊状态、缺图、追加和备份恢复回归测试。

### Changed

- `全对`、`未交`、`未带`同步应用到两版。
- 备份与恢复覆盖完整 `学生/` 家长交付目录。
- v3.0 单目录配置自动迁移为学生版目录。
- 文档保存先写临时文件，再替换正式文件。
- 版本号统一由 `version.py` 与 `pyproject.toml` 管理。
- 发布附件改用跨平台稳定的 ASCII 文件名，避免中文文件名在 GitHub Actions 中丢失。
