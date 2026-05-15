# Quick Document

Quick Document 是一个 macOS / Windows 兼容的 AI 文档处理助手。核心模式是：选择本地文档目录，输入一句修改要求，AI 根据文件路径调用内置 Office Skills 处理 Word、Excel、PPT。

AI快速处理文档软件，高效自带SKILL。

## 当前能力

- 桌面驻留：关闭窗口后可隐藏到托盘继续运行。
- 路径驱动：扫描本地目录中的 `.docx`、`.xlsx`、`.ppt`、`.pptx` 文件。
- 轻量指令：例如“把 xxx.docx 第二段润色一下”或“把表格里第 3 行金额改成 1200”。
- 多模态上下文：可粘贴截图、拖入图片/音频/视频/文件作为文档处理参考，但核心仍限定在 Word、Excel、PPT 工作流。
- Docling 解析增强：如果本机已安装 Docling，应用会自动把候选 PDF、Word、Excel、PPT、HTML、Markdown、CSV、图片等解析成 Markdown 上下文交给 AI；没有安装时自动回退到内置轻量预览。
- Skill 路由：
  - `.docx` → `resources/skills/documents`
  - `.xlsx` → `resources/skills/spreadsheets`
  - `.ppt` / `.pptx` → `resources/skills/presentations`
- 打包发布：`resources/skills` 会随应用一起进入 DMG / EXE。
- AI 配置：支持 OpenAI-compatible、OpenAI Responses 和 Anthropic-compatible。默认自动跟随本机 cc-switch 当前 Codex/OpenAI 配置，并可回退到 Claude 或手动配置。

## 开发运行

```bash
npm install
npm run dev
```

## 打包

```bash
npm run build
npm run dist:mac
npm run dist:win
```

`dist:mac` 输出 `.dmg`，`dist:win` 输出 NSIS `.exe` 安装包。建议在 macOS 构建 DMG，在 Windows 或 Windows CI Runner 构建 EXE。

项目已包含 `.github/workflows/build-desktop.yml`，推送到 GitHub 后可自动分别构建 macOS DMG 和 Windows EXE。

macOS 正式分发签名说明见 [docs/macos-signing.md](docs/macos-signing.md)。

## 可选：启用 Docling

Docling 用于增强“读懂文档”的能力，不负责原样修改 Office 文件。修改和复制仍由 Quick Document 的本地执行器和内置 Office Skills 完成。

安装任一方式即可被自动检测：

```bash
pipx install docling
# 或
python3 -m pip install docling
```

如果希望运行时允许 `uvx --from docling docling` 自动拉起 Docling，可设置：

```bash
export QUICK_DOCUMENT_DOCLING_AUTO_UVX=1
```
