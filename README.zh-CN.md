# office-agent

Office Agent 是一个本地优先的工具，可以把自然语言需求转换为 PowerPoint `.pptx`
文件。项目包含：

- 用于生成与渲染的 Python CLI 和 HTTP API
- 用于流程可视化、主题切换和双语界面的 React + TypeScript 前端

## 功能特性

- 支持 OpenAI-compatible 与 Ollama 两类 provider
- 两阶段生成流程：prompt -> 结构化 `PresentationSpec` -> `.pptx`
- 当前支持 10 类 slide：
  - `title`
  - `section`
  - `bullets`
  - `two_column`
  - `image`
  - `timeline`
  - `quote`
  - `comparison`
  - `summary`
  - `table`
- 支持预设主题切换与自定义主题覆盖
- 前端支持主题 JSON 导出与导入
- 页面展示输入、LLM 解析、渲染、Office 收尾四阶段可视化流程
- 前端 API Key 只在当前会话中使用，不落盘
- 网页支持英文与简体中文切换

## 主题结构

共享主题 schema 如下：

```json
{
  "preset": "default | executive | editorial",
  "custom": {
    "primary_color": "#RRGGBB",
    "accent_color": "#RRGGBB",
    "background_color": "#RRGGBB",
    "heading_font": "字体名",
    "body_font": "字体名",
    "cover_style": "band | centered | minimal"
  }
}
```

可以直接使用预设主题，也可以叠加自定义字段。

## 后端启动

使用 `uv` 安装 Python 依赖：

```bash
uv sync
```

运行 CLI：

```bash
uv run office-agent generate --prompt "生成一份项目汇报" --out output.pptx --theme executive
```

运行 API 服务：

```bash
uv run office-agent-api
```

可用接口：

- `GET /api/health`
- `POST /api/specs`
- `POST /api/presentations`

## 前端启动

安装前端依赖：

```bash
cd frontend
npm install
```

启动 Vite 开发服务：

```bash
npm run dev -- --port 5174
```

前端会把 `/api/*` 代理到 `http://127.0.0.1:8000`，开发时需要同时启动 Python API。

## 环境变量

- `OFFICE_AGENT_PROVIDER`
- `OFFICE_AGENT_MODEL`
- `OPENAI_API_KEY`
- `OPENAI_BASE_URL`
- `OLLAMA_BASE_URL`

前端不会持久化保存 API Key。用户在页面中输入的 key 只会随本次请求发送，不会写入本地文件。

## 文档

- 英文版：`README.md`
- 中文版：`README.zh-CN.md`

## 测试

运行后端测试：

```bash
uv run python -m unittest discover -s tests
```

运行前端测试：

```bash
cd frontend
npm test
```
