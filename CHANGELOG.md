# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [1.0.1] - 2026-04-02

### Added
- 新增扫描目录功能，支持递归扫描 xlsx/xlsm 文件
- 新增子表名称搜索和文件名搜索
- 新增"打开文件"功能，使用系统默认程序打开
- 新增"定位文件"功能，在文件管理器中定位文件
- 新增"复制路径"功能
- 新增扫描耗时显示（支持毫秒/秒/分钟显示）
- 新增增量扫描，只更新有变化的文件

### Improved
- 使用 PyQt5 重构 GUI
- 优化扫描性能：
  - 使用 zipfile 直接读取 xlsx 内部结构，速度提升 5-10 倍
  - 支持多线程并发扫描（默认 8 线程）
- 优化跨平台兼容性

### Fixed
- 修复 Windows 端定位文件功能失效的问题
- 修复数据库写入 bug（INSERT OR REPLACE 后 lastrowid 错误）