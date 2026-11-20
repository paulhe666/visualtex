# VisualTeX 0.1.0 — Windows / Linux 构建

项目已配置两个 GitHub Actions 工作流。

## 1. 手动测试构建

工作流：`.github/workflows/build-windows-linux.yml`

触发方式：

1. 将当前源码推送到 GitHub 仓库的 `main` 分支。
2. 打开仓库的 **Actions** 页面。
3. 选择 **Build Windows and Linux**。
4. 点击 **Run workflow**。
5. 等待 Windows 和 Linux 两个任务完成。
6. 在该次运行页面底部的 **Artifacts** 下载：
   - `VisualTeX-0.1.0-Windows-x64`
   - `VisualTeX-0.1.0-Linux-x64`

Windows artifact 包含：

- NSIS `.exe` 安装程序
- `.msi` 安装程序

Linux artifact 包含：

- `.AppImage`
- Debian/Ubuntu `.deb`

## 2. 创建 GitHub Release

工作流：`.github/workflows/release-v0.1.0.yml`

有两种触发方式：

### 方法 A：手动触发

在 Actions 页面选择 **Release VisualTeX 0.1.0**，点击 **Run workflow**。

### 方法 B：推送版本标签

```bash
git tag v0.1.0
git push origin v0.1.0
```

完成后 GitHub 会建立一个 **Draft Release**，并附上 Windows 和 Linux 安装包。确认内容无误后，在 Releases 页面手动点击发布。

## 3. GitHub 权限

Release 工作流需要仓库允许 Actions 写入 Release：

1. 打开仓库 **Settings**。
2. 进入 **Actions → General**。
3. 找到 **Workflow permissions**。
4. 选择 **Read and write permissions**。
5. 保存。

## 4. OCR 运行环境

安装包本身不内置 Python、PaddlePaddle 和模型权重。用户首次使用 OCR 时，VisualTeX 会在用户数据目录创建独立虚拟环境并安装依赖。

Windows OCR 前提：

- 64 位 Windows
- 64 位 Python 3.9–3.13
- 可通过 `python`、`python.exe` 或 `py` 启动

Linux OCR 前提：

- x86_64 Linux
- Python 3.9–3.13
- 系统能够访问 Python/PaddleOCR 下载源

不使用 OCR 时，公式编辑、工具栏、源码编辑、历史记录和导入导出不依赖 Python。

## 5. 签名说明

当前 0.1.0 构建未配置 Windows 代码签名，因此 Windows SmartScreen 可能在首次安装时显示未知发布者警告。Linux AppImage 和 `.deb` 也未做发行版仓库签名。
