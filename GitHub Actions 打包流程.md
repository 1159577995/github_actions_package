# GitHub Actions 一键打包流程（Windows / macOS / Linux / 信创）

> 本文件夹包含打包所需的全部文件与 GitHub Actions 工作流。
> **上传到 GitHub 时，把本文件夹内的所有内容放到仓库根目录**（含 `.github` 隐藏目录）。

---

## 一、文件夹结构（= GitHub 仓库根目录）

```
<你的仓库根目录>/
├── .github/
│   └── workflows/
│       └── build.yml            ← GitHub Actions 工作流（自动打包核心）
├── launcher.py                  ← 程序启动入口（双击后从这里开始记录日志）
├── excel_processor.py           ← 主程序
├── excel_processor.spec         ← PyInstaller 配置（Windows/Linux 用）
├── excel_processor_macos.spec   ← PyInstaller 配置（macOS .app 用）
├── requirements.txt             ← Python 依赖
├── build_all.sh                 ← 一键打包主入口
├── build_macos.sh               ← macOS 打包脚本
├── build_linux.sh               ← Linux/信创打包脚本
├── build_windows.bat            ← Windows 打包脚本
└── GitHub Actions 打包流程.md    ← 本文档
```

> 注意：此文件夹是**打包快照**。若以后修改了主程序（`excel_processor.py` 等），
> 请把更新后的文件同步回本文件夹（或直接在仓库中修改）。

---

## 二、上传到 GitHub

1. 在 GitHub 新建一个仓库（Public / Private 均可）；
2. 把本文件夹**内所有内容**上传到仓库根目录，两种方式任选：
   - **网页上传**：进入仓库页面 → `Add file` → `Upload files`，把本文件夹里的
     `launcher.py`、`excel_processor.py`、`excel_processor.spec`、`excel_processor_macos.spec`、
     `requirements.txt`、`build_all.sh`、`build_macos.sh`、`build_linux.sh`、
     `build_windows.bat`、`GitHub Actions 打包流程.md` 以及 `.github/workflows/build.yml`
     全部拖进去（注意 `.github` 是隐藏文件夹，拖入整个文件夹即可）；
   - **命令行上传**：
     ```bash
     cd github_actions_package
     git init
     git add .
     git commit -m "init"
     git remote add origin https://github.com/<你的用户名>/<仓库名>.git
     git push -u origin main
     ```
3. 确认上传后仓库根目录存在 `.github/workflows/build.yml`（Actions 只识别这个固定路径）。

---

## 三、触发打包

| 方式 | 操作 |
|---|---|
| **手动（推荐）** | 仓库页面 → `Actions` → 左侧选择 `打包发布` → `Run workflow` → 绿色按钮 |
| 打标签 | 本地 `git tag v1.0 && git push origin v1.0`，自动触发 |

触发后会自动并行运行 4 个任务：
`Windows 打包`、`macOS 打包`、`Linux x86_64 打包`、`Linux ARM64 打包`。

---

## 四、下载产物

1. 仓库页面 → `Actions` → 点开最新一次运行的记录；
2. 页面底部 **Artifacts** 区域有 4 个压缩包：

| 产物名 | 对应系统 | 解压后内容 |
|---|---|---|
| `excel-processor-windows` | Windows | `Excel处理器-Dify答案匹配工具-windows.exe` |
| `excel-processor-macos` | macOS（Apple Silicon） | `Excel处理器-Dify答案匹配工具.app` |
| `excel-processor-linux-x64` | Linux / 信创（x86_64，海光/兆芯等） | `...-linux-x86_64` |
| `excel-processor-linux-arm64` | 信创（aarch64，鲲鹏/飞腾等） | `...-linux-aarch64` |

3. 点击下载 zip，解压后即可分发/使用（macOS 需先 `chmod +x` 内层二进制或用 `xattr -dr com.apple.quarantine` 解除隔离，详见"五"）。

---

## 五、产物使用与分发注意

- **macOS**：双击 `.app` 直接出界面（无命令行、含启动画面）；拷贝给他人如提示
  "无法打开，因为来自身份不明的开发者"，右键 → 打开，或执行
  `xattr -dr com.apple.quarantine "Excel处理器-Dify答案匹配工具.app"`。
  > 默认产物为 **arm64（Apple Silicon）**；如需 Intel 版，见文末附录。
- **Windows**：双击 `.exe`；首次启动较慢（单文件自解压）；被杀毒软件误报时添加信任。
- **Linux / 信创**：`chmod +x` 后双击运行；建议在真实信创机器（麒麟/统信）上做一次冒烟验证，
  选择与机器 CPU 架构匹配的产物（x86_64 或 aarch64）。
- 文件处理 Agent 的沙箱工作区 `agent_workspace/` 优先生成在可执行文件同目录；若同级目录不可写，则自动回退到平台用户数据目录：Windows 为 `%APPDATA%\Excel处理器-Dify答案匹配工具\agent_workspace\`，macOS 为 `~/Library/Application Support/Excel处理器-Dify答案匹配工具/agent_workspace/`，Linux 为 `~/.Excel处理器-Dify答案匹配工具/agent_workspace/`。
- **运行日志**：双击运行后优先在可执行文件同目录生成 `运行日志.log`（启动/处理/异常全程记录）；若同级目录不可写，则回退到平台用户数据目录。启动失败还会生成 `启动错误.log`（完整堆栈）。用户反馈"双击无反应"时，先看这两个文件。

---

## 六、常见问题（Actions）

| 现象 | 原因与解决 |
|---|---|
| Actions 页没有 `打包发布` 工作流 | `.github/workflows/build.yml` 未放在仓库根目录；或上传后需等待几秒刷新 |
| `Run workflow` 按钮灰色 | 仓库刚创建还没默认分支，先 push 一次代码 |
| 某任务失败（红叉） | 点击失败步骤查看日志：多为 pip 网络波动，直接 `Run workflow` 重跑即可 |
| macOS 产物提示开发者无法验证 | 见"五"中的 Gatekeeper 处理 |
| 免费额度 | GitHub 免费账号：Public 仓库无限分钟；Private 仓库每月 2000 分钟（4 个任务一次约 10~20 分钟） |

---

## 附录：增加 macOS Intel（x86_64）产物

在 `build.yml` 的 `build-macos` 任务后追加一个任务即可：

```yaml
  build-macos-intel:
    name: macOS Intel 打包
    runs-on: macos-13        # macos-13 为 Intel runner
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
      - name: 执行 macOS 打包
        run: |
          chmod +x build_all.sh build_macos.sh
          ./build_macos.sh
      - name: 上传 macOS Intel 产物
        uses: actions/upload-artifact@v4
        with:
          name: excel-processor-macos-intel
          path: dist/*.app
```

（`macos-13` 为 Intel 架构，产物为 x86_64；`macos-latest` 当前为 Apple Silicon arm64。）
