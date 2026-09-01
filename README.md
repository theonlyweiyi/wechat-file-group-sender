# 微信文件群发工具 · WeChat File Group Sender

> 按「文件名关键词 → 微信备注名/昵称」的映射，批量把某个文件夹里的文件自动发给对应的微信联系人。
> 运行在 **个人微信 (Windows 客户端)**，界面为原生窗口程序（无控制台、不弹浏览器）。

[![Release](https://img.shields.io/github/v/release/theonlyweiyi/wechat-file-group-sender?label=Release)](https://github.com/theonlyweiyi/wechat-file-group-sender/releases)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](#)

---

## ✨ 功能特性

- **文件夹选择**：点击按钮 → 系统原生文件夹对话框，无需手填路径。
- **两种映射录入方式**：
  - 上传 Excel（关键词列 + 微信备注名/昵称列）；
  - 直接在文本框里按 `关键词,微信备注名/昵称` 逐行输入。
- **同人合并去重**：多个关键词映射到同一个联系人时，自动合并为**一次发送**、文件去重，绝不重复发送。
- **模板内嵌**：无需附带任何文件，界面点「下载模板」即可获取映射表模板（已编译进 exe 资源）。
- **双视图实时预览**：
  - 「按联系人」= 实际发送计划（一人一行，合并后的结果）；
  - 「按关键词」= 匹配明细（便于核对每个关键词命中了哪些文件）。
- **可选附言**：发送文件前可填一条文字消息。
- **进度可视化**：前端实时显示进度与每步状态，支持中途「停止」。
- **自动窗口管理**：点击「开始发送」后自动切到微信、恢复最小化窗口并置顶，发送结束把焦点拉回本程序。

---

## 🚀 快速开始

### 方式一：直接下载 exe（推荐普通用户）

1. 打开 [Releases 页面](https://github.com/theonlyweiyi/wechat-file-group-sender/releases)，下载 `WeChatFileSender.exe`。
2. 确认 **微信 PC 版已登录**（前台或最小化窗口均可，程序会自动处理）。
3. 双击 `WeChatFileSender.exe`，直接弹出原生窗口（无黑框、不弹浏览器）。
4. 选文件夹 → 填/上传映射（也可点「下载模板」）→ 预览匹配 → 点「开始发送」。
5. 发送期间**请勿操作鼠标键盘**，直到状态栏汇总完成。

> **环境依赖**：Windows 10/11 (x64)，且已安装 **Microsoft Edge WebView2 Runtime**（Win10/11 一般自带；若窗口打不开，去微软官网安装 WebView2 Runtime）。

### 方式二：从源码运行（开发者）

```bash
# 需要 Python 3.11+
pip install flask pywebview openpyxl pyautogui pywin32

# 启动（会同时起 Flask 线程并弹出原生窗口）
python launch.py
```

如需自行打包成单文件 exe：

```bash
pip install pyinstaller
pyinstaller "微信文件群发工具.spec"
# 产物在 dist/微信文件群发工具.exe（模板已通过 --add-data 编译进资源）
```

---

## 📖 使用流程

1. 启动程序 → 弹出原生窗口。
2. 点击「选择文件夹」选定待发送文件所在目录。
3. 录入映射（上传 Excel 或直接输入文本，一行一条）。
4. 点「预览匹配」→ 确认每个文件对应的收件人。
5. 点「开始发送」→ 程序逐一对每个联系人：
   - 在微信搜索框输入备注名/昵称、回车打开聊天；
   - 把该联系人的文件写入剪贴板（CF_HDROP）并粘贴发送；
   - 若配置了附言，先粘贴发送文字；
   - 切换下一个联系人，循环直至完成。
6. 发送结束，状态栏汇总成功 / 失败。

---

## 🧩 工作原理

微信 4.1+ 使用 Qt Quick 渲染，不暴露 UI Automation 控件，因此 `wxauto` 等自动化库失效。
本项目采用 **键盘 + 剪贴板模拟** 方案完成发送：

```
┌─────────────────────────────────────────────┐
│            微信文件群发工具.exe               │
│  (PyInstaller --onefile --noconsole 打包)     │
│                                               │
│  ┌────────────────┐      ┌─────────────────┐ │
│  │  pywebview 窗口 │◄────►│  Flask 本地服务  │ │
│  │ (Edge WebView2) │ JS   │  127.0.0.1:5890 │ │
│  │   原生 HTML 界面 │ API  │  (守护线程运行) │ │
│  └────────────────┘      └────────┬────────┘ │
│                                    │ 调用       │
│                          ┌─────────▼─────────┐ │
│                          │  wechat_sender.py  │ │
│                          │  键鼠/剪贴板模拟   │ │
│                          └─────────┬─────────┘ │
│                                    │ 操控       │
│                          ┌─────────▼─────────┐ │
│                          │    Windows 微信客户端 │ │
│                          └───────────────────┘ │
└─────────────────────────────────────────────┘
```

| 组件 | 作用 | 说明 |
|------|------|------|
| **Flask** | 后端 API + 页面托管 | 轻量，界面与逻辑解耦 |
| **pywebview** | 把 HTML 渲染进原生窗口 | 底层用系统 Edge WebView2，比弹浏览器更像「普通程序」 |
| **PyInstaller --noconsole** | 打包成单文件 exe | 满足「无控制台、不弹浏览器」的最终要求 |
| **键鼠 + 剪贴板模拟** | 实际发送动作 | 微信 4.1+ 不暴露 UIA 控件，只能走输入模拟 |
| **win32api / pyautogui** | 窗口查找、置顶、粘贴 | 绕过 Windows 前台锁、处理中文输入 |

---

## 📁 目录结构

```
项目根目录/
├── launch.py                 # pywebview 入口（Flask 线程 + 原生窗口）
├── web_app.py                # Flask 后端 + 内嵌 HTML/CSS/JS 前端（主程序）
├── wechat_sender.py          # 独立发送模块（键鼠/剪贴板模拟，核心逻辑）
├── wechat_file_sender.py     # 早期探索版本（已弃用，保留参考）
├── create_template.py        # 生成「发送映射模板.xlsx」
├── 发送映射模板.xlsx          # 关键词 / 微信备注名-昵称 模板（已编译进 exe，可单独下载）
├── 微信文件群发工具.spec      # PyInstaller 打包规格
├── test_match_logic.py        # 关键词匹配逻辑单元测试
├── test_web_api.py            # Web API 测试脚本
└── .gitignore                 # 排除 .workbuddy/、dist/、build/ 等
```

> 说明：`dist/`、`build/`、`.workbuddy/` 已在 `.gitignore` 中排除；
> 编译好的 exe 通过 [GitHub Release](https://github.com/theonlyweiyi/wechat-file-group-sender/releases) 分发，保持仓库源码整洁。

---

## ⚠️ 已知限制

1. **依赖输入模拟**：发送时占用键鼠，期间请避免手动操作；发送速度受模拟节奏影响（已加适当延时保证稳定）。
2. **联系人靠微信搜索定位**：程序把第二列的字符串原样粘进微信搜索框并回车取第一个结果，因此**备注名 / 昵称 / 微信号都能填**。推荐优先级：**微信号（唯一） > 备注名 > 原始昵称**（昵称易重名、对方改名即失效）。若出现重名，可能开错会话。
3. **无自动重试**：单个文件发送失败会记录但跳过，未做失败重试队列。
4. **仅 Windows**：方案深度绑定 Win32 API，无法跨平台。

---

## 📜 免责声明

- 本项目仅供个人批量发送文件所用，**请遵守微信使用规范，勿用于 spam / 骚扰**。
- 微信产品行为可能随版本变更，如发送失效需等待适配。
- 本项目与腾讯公司无官方关联。

---

## 📝 License

MIT © theonlyweiyi
