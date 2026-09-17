# Windows 打包说明

如果只是使用拾文，直接从 GitHub Releases 下载安装包，不用执行下面的步骤。

## 1. 打包独立程序

准备 Windows 10 / 11 和 64 位 Python 3.12。在仓库根目录运行：

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.venv\Scripts\python.exe scripts/build_windows.py
```

结果在 `build/bundle/Shiwen`。运行其中的 `Shiwen.exe` 即可使用；分发这个版本时，需要发送整个文件夹。

脚本只收集 `desktop.py`、它导入的程序代码、页面文件和必要依赖，不会复制 `data`、`runtime.json` 或开发环境。

## 2. 制作安装包

准备 [Inno Setup 6](https://jrsoftware.org/isdl.php)。1.0.0 安装包使用 6.7.3 构建；请根据自己的用途遵守其许可条件。

将中文安装语言文件下载到生成目录：

```powershell
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/kira-96/Inno-Setup-Chinese-Simplified-Translation/main/ChineseSimplified.isl' -OutFile 'build/ChineseSimplified.isl'
```

语言文件来自 [Inno Setup 简体中文翻译项目](https://github.com/kira-96/Inno-Setup-Chinese-Simplified-Translation)。

用 Inno Setup 打开 `packaging/installer.iss`，点击 Build → Compile。或者在 `ISCC.exe` 已加入 PATH 时运行：

```powershell
ISCC.exe packaging/installer.iss
```

安装包输出到 `release/拾文-Setup-1.0.0-Windows-x64.exe`。将它和 `packaging/使用说明.txt` 一起压缩，即可发给朋友。

安装版按当前 Windows 用户安装，不需要管理员权限。数据保存在 `%LOCALAPPDATA%\Shiwen\data`，更新和卸载保留数据。

## 3. 发布前检查

- 在空数据目录验证首次打开没有预装文章和公众号。
- 验证导入、日期与关键词筛选、18 列 Excel 导出，以及退出后重新打开。
- 验证更新与卸载不会删除已有数据库。
- 安装包中不应出现个人数据库、日志、本机环境配置或任何凭据。

测试可用 `COLLECTOR_DATA_DIR` 指向临时数据目录，`COLLECTOR_PORT=0` 使用空闲端口。`Shiwen.exe --no-browser` 启动时不自动打开浏览器，`Shiwen.exe --shutdown` 退出同一数据目录对应的后台程序。

首次发布已在构建电脑上实测独立程序运行、空库、导入去重、筛选导出、重启保留、重复启动、端口冲突回退、安装、覆盖更新和卸载。没有在所有 Windows 版本和安全软件组合上验证。
