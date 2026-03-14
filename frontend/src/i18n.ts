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
  templatePreview: string;
  templatePreviewLoading: string;
  templatePreviewEmpty: string;
  templatePreviewError: string;
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
    heroTitle: "Preview the template, map each section, then render the deck.",
    heroDescription:
      "Upload a PPTX template, inspect each template slide, map it to opening, agenda, content, and closing, then generate a JSON-backed deck from the same page.",
    providersMetric: "OpenAI / Ollama",
    flowMetric: "Template-aware spec",
    outputMetric: "Preview + mapped PPTX",
    configuration: "Configuration",
    runtimeControls: "Runtime controls",
    provider: "Provider",
    model: "Model",
    apiKey: "API Key",
    openaiBaseUrl: "OpenAI Base URL",
    ollamaBaseUrl: "Ollama Base URL",
    themeStudio: "Theme & Template",
    themeControls: "Theme, template preview, and mappings",
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
    templateHint: "Upload a PPTX template to preview each slide and map it to the four presentation parts.",
    templateSelected: "Selected template",
    clearTemplate: "Clear Template",
    templatePreview: "Template Preview",
    templatePreviewLoading: "Generating slide thumbnails...",
    templatePreviewEmpty: "Upload a PPTX template to preview available layouts.",
    templatePreviewError: "Template preview error",
    templateMappings: "Template Part Mapping",
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
    noSpecDescription: "Generate a spec first to inspect the mapped JSON before rendering.",
    slidePreview: "Slide Preview",
    presentationOutline: "Presentation outline",
    noSlidesTitle: "Waiting for a slide plan",
    noSlidesDescription: "Once the LLM returns a valid spec, slide cards will appear here.",
    officePlaceholder: "Placeholder only in this version. No Office automation yet.",
    keyPointsLabel: "Key points",
    nextStepsLabel: "Next steps",
    language: "Language",
    english: "English",
    chinese: "中文",
    partPrefix: "Part",
    templateSelectionRequired: "Complete all four template part selections before generating a spec.",
  },
  "zh-CN": {
    appTitle: "Office Agent Studio",
    heroTitle: "先预览模板并映射四个部分，再生成演示文稿。",
    heroDescription:
      "上传 PPTX 模板后，先查看每一页模板样式，再把它们映射到开题、目录、正文和结束，最后生成使用该模板的 JSON 方案与 PPT。",
    providersMetric: "OpenAI / Ollama",
    flowMetric: "模板感知方案",
    outputMetric: "预览 + 映射 PPTX",
    configuration: "运行配置",
    runtimeControls: "模型与接口控制",
    provider: "Provider",
    model: "模型",
    apiKey: "API Key",
    openaiBaseUrl: "OpenAI Base URL",
    ollamaBaseUrl: "Ollama Base URL",
    themeStudio: "主题与模板",
    themeControls: "主题、模板预览与版式映射",
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
    templateHint: "上传 PPTX 模板后，可以预览模板每一页，并分别映射到开题、目录、正文和结束。",
    templateSelected: "当前模板",
    clearTemplate: "清除模板",
    templatePreview: "模板预览",
    templatePreviewLoading: "正在生成模板缩略图...",
    templatePreviewEmpty: "上传 PPTX 模板后，这里会显示每一页的预览。",
    templatePreviewError: "模板预览错误",
    templateMappings: "模板部分映射",
    templatePartOpening: "开题",
    templatePartAgenda: "目录",
    templatePartContent: "正文",
    templatePartClosing: "结束",
    placeholderRoles: "占位符角色",
    slideIndex: "第",
    themeSummary: "当前主题摘要",
    input: "输入",
    describeDeck: "描述演示稿需求",
    prompt: "提示词",
    promptPlaceholder: "描述你想生成的演示文稿内容。",
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
    noSpecDescription: "先生成方案，再检查带模板映射的 JSON 是否符合预期。",
    slidePreview: "页面预览",
    presentationOutline: "演示结构摘要",
    noSlidesTitle: "等待页面结构",
    noSlidesDescription: "当 LLM 返回合法方案后，这里会展示每页内容摘要。",
    officePlaceholder: "当前版本仅保留占位，不执行 Office 自动化。",
    keyPointsLabel: "关键结论",
    nextStepsLabel: "后续动作",
    language: "语言",
    english: "English",
    chinese: "中文",
    partPrefix: "部分",
    templateSelectionRequired: "请先完成四个模板部分的选择，再生成方案。",
  },
};

export const workflowLabels: Record<Language, Record<WorkflowStageKey, { title: string; subtitle: string }>> = {
  en: {
    input: { title: "1. Capture Intent", subtitle: "Collect prompt, provider, theme, and template mapping." },
    llm: { title: "2. Parse to Spec", subtitle: "Call the model and generate a section-tagged JSON spec." },
    render: { title: "3. Render PPT", subtitle: "Use the selected template slides while generating the final PPTX." },
    office: { title: "4. Final Office Step", subtitle: "Reserved for future animation and review." },
  },
  "zh-CN": {
    input: { title: "1. 收集需求", subtitle: "整理提示词、模型配置、主题与模板映射。" },
    llm: { title: "2. 生成结构化方案", subtitle: "调用模型，输出带部分标记的 JSON 方案。" },
    render: { title: "3. 渲染 PPT", subtitle: "按已选模板页生成最终 PPTX。" },
    office: { title: "4. Office 收尾", subtitle: "预留给动画和最终检查。" },
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
