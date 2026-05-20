# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev       # 启动开发模式（Vite 开发服务器 + Electron，支持 HMR）
npm run build     # 生产构建：tsc 编译 → Vite 打包 → Electron Builder 生成可执行文件
npm test          # 运行 Vitest 单元测试（pretest 自动以 test 模式重新构建）
```

输出目录：
- `dist/` — 渲染进程产物（Vite 打包）
- `dist-electron/` — 主进程和预加载脚本产物
- `release/` — 最终可执行安装包（按版本号分目录）

## Architecture

项目是一个 **Electron + Vite + React + TypeScript** 桌面应用，包含两个独立的进程：

### 主进程（`electron/`）

- `electron/main/index.ts` — 创建 `BrowserWindow`、管理应用生命周期，开发模式加载 `http://127.0.0.1:7777/`，生产模式加载 `dist/index.html`
- `electron/main/update.ts` — 基于 `electron-updater` 的自动更新逻辑，通过 IPC 向渲染进程推送更新事件
- `electron/preload/index.ts` — 预加载脚本，通过 `contextBridge` 将 `ipcRenderer` 挂载到 `window.ipcRenderer`

### 渲染进程（`src/`）

标准 React 18 应用，入口为 `src/main.tsx`。路径别名 `@/` 映射到 `src/`。

- `src/App.tsx` — 根组件
- `src/components/update/` — 自动更新 UI 组件（Modal + Progress）
- `src/demos/ipc.ts` — IPC 通信用法示例

### IPC 通信约定

渲染进程通过 `window.ipcRenderer.invoke(channel, ...)` 调用主进程，主进程通过 `event.sender.send(channel, ...)` 推送消息：

| Channel | 方向 | 说明 |
|---|---|---|
| `check-update` | invoke | 检查是否有新版本 |
| `start-download` | invoke | 开始下载更新 |
| `cancel-download` | invoke | 取消下载 |
| `quit-and-install` | invoke | 安装并重启 |
| `update-can-available` | main→renderer | 更新可用通知 |
| `download-progress` | main→renderer | 下载进度 |
| `update-downloaded` | main→renderer | 下载完成通知 |
| `update-error` | main→renderer | 更新错误通知 |

### 构建工具配置

- `vite.config.ts` — 使用 `vite-plugin-electron/simple`，主进程和预加载脚本分别编译到 `dist-electron/main` 和 `dist-electron/preload`；开发端口为 **7777**
- `electron-builder.json` — Windows 平台生成 NSIS 安装程序（x64），macOS 生成 DMG；自动更新提供商为 generic
- `tailwind.config.js` — 已禁用 Tailwind preflight（全局样式重置）
