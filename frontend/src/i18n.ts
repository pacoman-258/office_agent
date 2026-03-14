import type { Language, WorkflowStageKey, WorkflowStageStatus } from "./types";

type LocaleText = {
  appTitle: string;
  heroTitle: string;
  heroDescription: string;
  providersMetric: string;
  flowMetric: string;
  outputMetric: string;
  configuration: string;
  runtimeControls: string;
  provider: string;
  model: string;
  apiKey: string;
  openaiBaseUrl: string;
  ollamaBaseUrl: string;
  finalizeTitle: string;
  finalizeEnabled: string;
  finalizeModel: string;
  finalizeMaxRounds: string;
  finalizeHint: string;
  finalizeSummary: string;
  finalizeIssues: string;
  finalizeOperations: string;
  finalizeRounds: string;
  finalizeStatus: string;
  themeStudio: string;
  themeControls: string;
  themePreset: string;
  primaryColor: string;
  accentColor: string;
  backgroundColor: string;
  headingFont: string;
  bodyFont: string;
  coverStyle: string;
  themeJson: string;
  exportThemeJson: string;
  importThemeJson: string;
  templateTitle: string;
  templateFile: string;
  templateHint: string;
  templateSelected: string;
  clearTemplate: string;
  templateMode: string;
  templateModeManual: string;
  templateModePreview: string;
  templateManualMapping: string;
  templateManualHint: string;
  templatePageNumber: string;
  templatePageNumberHint: string;
  templatePreview: string;
  templatePreviewAction: string;
  templatePreviewLoading: string;
  templatePreviewEmpty: string;
  templatePreviewError: string;
  templateCleanupNotice: string;
  templateMappings: string;
  templatePartOpening: string;
  templatePartAgenda: string;
  templatePartContent: string;
  templatePartClosing: string;
  placeholderRoles: string;
  slideIndex: string;
  themeSummary: string;
  input: string;
  describeDeck: string;
  prompt: string;
  promptPlaceholder: string;
  filename: string;
  generateSpec: string;
  generatingSpec: string;
  confirmGenerate: string;
  renderingPpt: string;
  copyJson: string;
  workflow: string;
  pipelineStatus: string;
  structuredSpec: string;
  validatedJson: string;
  noSpecTitle: string;
  noSpecDescription: string;
  slidePreview: string;
  presentationOutline: string;
  noSlidesTitle: string;
  noSlidesDescription: string;
  officePlaceholder: string;
  officeRunning: string;
  keyPointsLabel: string;
  nextStepsLabel: string;
  language: string;
  english: string;
  chinese: string;
  partPrefix: string;
  templateSelectionRequired: string;
};

export const locales: Record<Language, LocaleText> = {
  en: {
    appTitle: "Office Agent Studio",
    heroTitle: "Generate, review, and auto-fix the deck in one flow.",
    heroDescription:
      "Build the presentation spec, render the PPTX, then optionally let a visual review model inspect exported slide screenshots and apply safe PowerPoint edits automatically.",
    providersMetric: "OpenAI / Ollama",
    flowMetric: "Spec + render + review",
    outputMetric: "Template-aware PPTX",
    configuration: "Configuration",
    runtimeControls: "Runtime controls",
    provider: "Provider",
    model: "Model",
    apiKey: "API Key",
    openaiBaseUrl: "OpenAI Base URL",
    ollamaBaseUrl: "Ollama Base URL",
    finalizeTitle: "Office Visual Review",
    finalizeEnabled: "Enable Visual Review",
    finalizeModel: "Review Model",
    finalizeMaxRounds: "Max Review Rounds",
    finalizeHint: "Uses an OpenAI-compatible vision model to inspect slide screenshots and apply safe PowerPoint edits after rendering.",
    finalizeSummary: "Visual Review Summary",
    finalizeIssues: "Issues",
    finalizeOperations: "Operations",
    finalizeRounds: "Rounds",
    finalizeStatus: "Status",
    themeStudio: "Theme & Template",
    themeControls: "Theme, template mapping, and optional preview",
    themePreset: "Theme Preset",
    primaryColor: "Primary Color",
    accentColor: "Accent Color",
    backgroundColor: "Background Color",
    headingFont: "Heading Font",
    bodyFont: "Body Font",
    coverStyle: "Cover Style",
    themeJson: "Theme JSON",
    exportThemeJson: "Export Theme JSON",
    importThemeJson: "Import Theme JSON",
    templateTitle: "PPTX Template",
    templateFile: "Import PPTX Template",
    templateHint: "Default mode does not require template preview. Upload the file and assign page numbers directly, or switch to preview mode if you want thumbnails.",
    templateSelected: "Selected template",
    clearTemplate: "Clear Template",
    templateMode: "Template Mode",
    templateModeManual: "Manual Mapping",
    templateModePreview: "Preview-assisted Mapping",
    templateManualMapping: "Manual Template Mapping",
    templateManualHint: "Enter 1-based page numbers from the uploaded template. These mappings are used even if template preview is unavailable.",
    templatePageNumber: "Template page number",
    templatePageNumberHint: "PowerPoint page numbers start at 1.",
    templatePreview: "Template Preview",
    templatePreviewAction: "Load Template Preview",
    templatePreviewLoading: "Generating slide thumbnails...",
    templatePreviewEmpty: "Switch to preview mode and load thumbnails if you want a visual reference. Manual mapping already works without this step.",
    templatePreviewError: "Template preview error",
    templateCleanupNotice:
      "Template rendering preserves branding elements such as logos and footer marks while clearing existing body content, pictures, tables, charts, and embedded objects.",
    templateMappings: "Preview-based Mapping",
    templatePartOpening: "Opening",
    templatePartAgenda: "Agenda",
    templatePartContent: "Content",
    templatePartClosing: "Closing",
    placeholderRoles: "Placeholder roles",
    slideIndex: "Slide",
    themeSummary: "Theme Summary",
    input: "Input",
    describeDeck: "Describe the deck",
    prompt: "Prompt",
    promptPlaceholder: "Describe the presentation you want to generate.",
    filename: "Filename",
    generateSpec: "Generate Spec",
    generatingSpec: "Generating Spec...",
    confirmGenerate: "Confirm and Generate PPT",
    renderingPpt: "Rendering PPT...",
    copyJson: "Copy JSON",
    workflow: "Workflow",
    pipelineStatus: "Pipeline status",
    structuredSpec: "Structured Spec",
    validatedJson: "Validated JSON",
    noSpecTitle: "No spec generated yet",
    noSpecDescription: "Generate a spec first to inspect the validated JSON before rendering.",
    slidePreview: "Slide Preview",
    presentationOutline: "Presentation outline",
    noSlidesTitle: "Waiting for a slide plan",
    noSlidesDescription: "Once the LLM returns a valid spec, slide cards will appear here.",
    officePlaceholder: "Visual review is optional. If enabled, it will export screenshots, review layout, and apply safe fixes.",
    officeRunning: "Exporting screenshots, reviewing layout, and applying PowerPoint edits...",
    keyPointsLabel: "Key points",
    nextStepsLabel: "Next steps",
    language: "Language",
    english: "English",
    chinese: "中文",
    partPrefix: "Part",
    templateSelectionRequired: "Complete all four template part mappings before generating a spec.",
  },
  "zh-CN": {
    appTitle: "Office Agent Studio",
    heroTitle: "一条链路完成生成、审查与自动修正。",
    heroDescription:
      "先生成结构化方案并渲染 PPT，再按需启用视觉审查模型查看每页截图，通过 PowerPoint 安全回写修正排版问题。",
    providersMetric: "OpenAI / Ollama",
    flowMetric: "方案 + 渲染 + 审查",
    outputMetric: "支持模板的 PPTX",
    configuration: "配置",
    runtimeControls: "运行控制",
    provider: "Provider",
    model: "模型",
    apiKey: "API Key",
    openaiBaseUrl: "OpenAI Base URL",
    ollamaBaseUrl: "Ollama Base URL",
    finalizeTitle: "Office 视觉审查",
    finalizeEnabled: "启用视觉修正",
    finalizeModel: "审查模型",
    finalizeMaxRounds: "最大审查轮数",
    finalizeHint: "使用 OpenAI-compatible 视觉模型查看幻灯片截图，并在渲染后通过 PowerPoint 自动应用安全修正。",
    finalizeSummary: "视觉审查摘要",
    finalizeIssues: "问题数",
    finalizeOperations: "修改数",
    finalizeRounds: "轮次",
    finalizeStatus: "状态",
    themeStudio: "主题与模板",
    themeControls: "主题、模板映射与可选预览",
    themePreset: "预设主题",
    primaryColor: "主色",
    accentColor: "强调色",
    backgroundColor: "背景色",
    headingFont: "标题字体",
    bodyFont: "正文字体",
    coverStyle: "封面样式",
    themeJson: "主题 JSON",
    exportThemeJson: "导出主题 JSON",
    importThemeJson: "导入主题 JSON",
    templateTitle: "PPTX 模板",
    templateFile: "导入 PPTX 模板",
    templateHint: "默认模式不依赖模板预览。上传文件后可直接填写页码映射；如果你想看缩略图，再切换到预览模式。",
    templateSelected: "当前模板",
    clearTemplate: "清除模板",
    templateMode: "模板模式",
    templateModeManual: "手动映射",
    templateModePreview: "预览辅助映射",
    templateManualMapping: "手动模板映射",
    templateManualHint: "请输入上传模板中的页码，页码从 1 开始。即使模板预览不可用，这组映射也能直接用于生成。",
    templatePageNumber: "模板页码",
    templatePageNumberHint: "PowerPoint 页码从第 1 页开始。",
    templatePreview: "模板预览",
    templatePreviewAction: "加载模板预览",
    templatePreviewLoading: "正在生成模板缩略图...",
    templatePreviewEmpty: "如果你需要可视化参考，可切换到预览模式并加载缩略图。手动映射已经可以直接使用。",
    templatePreviewError: "模板预览错误",
    templateCleanupNotice: "套用模板时会保留 logo、页脚角标等品牌元素，同时清除模板里原有的正文、内容图片、表格、图表和嵌入对象，避免干扰生成结果。",
    templateMappings: "基于预览的映射",
    templatePartOpening: "开题",
    templatePartAgenda: "目录",
    templatePartContent: "正文",
    templatePartClosing: "结束",
    placeholderRoles: "占位符角色",
    slideIndex: "第",
    themeSummary: "当前主题摘要",
    input: "输入",
    describeDeck: "描述演示稿",
    prompt: "提示词",
    promptPlaceholder: "描述你想生成的演示稿内容。",
    filename: "文件名",
    generateSpec: "生成方案",
    generatingSpec: "正在生成方案...",
    confirmGenerate: "确认并生成 PPT",
    renderingPpt: "正在渲染 PPT...",
    copyJson: "复制 JSON",
    workflow: "流程",
    pipelineStatus: "工作流状态",
    structuredSpec: "结构化方案",
    validatedJson: "已校验 JSON",
    noSpecTitle: "还没有生成方案",
    noSpecDescription: "先生成方案，再检查已校验 JSON 是否符合预期。",
    slidePreview: "页面预览",
    presentationOutline: "演示结构摘要",
    noSlidesTitle: "等待页面结构",
    noSlidesDescription: "当 LLM 返回合法方案后，这里会展示每页内容摘要。",
    officePlaceholder: "视觉审查是可选步骤。开启后会导出截图、检查排版并自动应用安全修正。",
    officeRunning: "正在导出截图、审查版式并应用 PowerPoint 修正...",
    keyPointsLabel: "关键结论",
    nextStepsLabel: "后续动作",
    language: "语言",
    english: "English",
    chinese: "中文",
    partPrefix: "部分",
    templateSelectionRequired: "请先完成四个模板部分的映射，再生成方案。",
  },
};

export const workflowLabels: Record<Language, Record<WorkflowStageKey, { title: string; subtitle: string }>> = {
  en: {
    input: { title: "1. Capture Intent", subtitle: "Collect prompt, provider, theme, template mapping, and review settings." },
    llm: { title: "2. Parse to Spec", subtitle: "Call the model and generate a section-tagged JSON spec." },
    render: { title: "3. Render PPT", subtitle: "Generate the PPTX with the selected template pages and theme." },
    office: { title: "4. Visual Review", subtitle: "Export screenshots, review layout, and apply safe PowerPoint edits." },
  },
  "zh-CN": {
    input: { title: "1. 收集需求", subtitle: "整理提示词、模型配置、主题、模板映射和审查设置。" },
    llm: { title: "2. 生成结构化方案", subtitle: "调用模型，输出带部分标记的 JSON 方案。" },
    render: { title: "3. 渲染 PPT", subtitle: "按所选主题与模板页生成最终 PPTX。" },
    office: { title: "4. 视觉审查", subtitle: "导出截图、审查版式，并应用安全的 PowerPoint 修正。" },
  },
};

export const statusLabels: Record<Language, Record<WorkflowStageStatus, string>> = {
  en: {
    idle: "idle",
    running: "running",
    success: "success",
    error: "error",
    skipped: "skipped",
  },
  "zh-CN": {
    idle: "未开始",
    running: "进行中",
    success: "成功",
    error: "失败",
    skipped: "跳过",
  },
};
