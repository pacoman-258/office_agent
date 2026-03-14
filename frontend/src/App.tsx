import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";

import { downloadBlob, generateSpec, previewTemplate, renderPresentation } from "./api";
import { locales, statusLabels, workflowLabels } from "./i18n";
import type {
  CustomThemeConfig,
  GenerateSpecRequest,
  Language,
  PresentationSpec,
  Provider,
  SlideSpec,
  TemplatePart,
  TemplatePreviewSlide,
  TemplateSelectionSpec,
  ThemePreset,
  ThemeSpec,
  WorkflowStageKey,
  WorkflowStageStatus,
} from "./types";

type StageState = Record<WorkflowStageKey, WorkflowStageStatus>;
type TemplateMappingDraft = Partial<Record<TemplatePart, number>>;

interface FormState {
  provider: Provider;
  apiKey: string;
  model: string;
  theme: ThemeSpec;
  openaiBaseUrl: string;
  ollamaBaseUrl: string;
  prompt: string;
  filename: string;
}

const defaultModels: Record<Provider, string> = {
  openai: "gpt-4o-mini",
  ollama: "qwen3:8b",
};

const headingFontOptions = ["Microsoft YaHei", "Georgia", "Segoe UI", "Times New Roman"];
const bodyFontOptions = ["Microsoft YaHei", "Segoe UI", "Arial", "Georgia"];

const presetThemeDefaults: Record<ThemePreset, Required<CustomThemeConfig>> = {
  default: {
    primary_color: "#16324F",
    accent_color: "#2A6F97",
    background_color: "#F5F7FA",
    heading_font: "Microsoft YaHei",
    body_font: "Microsoft YaHei",
    cover_style: "band",
  },
  executive: {
    primary_color: "#0E2A47",
    accent_color: "#C28B2C",
    background_color: "#F7F3EC",
    heading_font: "Georgia",
    body_font: "Microsoft YaHei",
    cover_style: "centered",
  },
  editorial: {
    primary_color: "#3A243B",
    accent_color: "#D95D39",
    background_color: "#FFF9F1",
    heading_font: "Georgia",
    body_font: "Microsoft YaHei",
    cover_style: "minimal",
  },
};

const templatePartOrder: TemplatePart[] = ["opening", "agenda", "content", "closing"];

const initialStages = (): StageState => ({
  input: "idle",
  llm: "idle",
  render: "idle",
  office: "idle",
});

const initialTheme = (): ThemeSpec => ({
  preset: "default",
  custom: presetThemeDefaults.default,
});

const initialFormState = (): FormState => ({
  provider: "openai",
  apiKey: "",
  model: defaultModels.openai,
  theme: initialTheme(),
  openaiBaseUrl: "https://api.openai.com/v1",
  ollamaBaseUrl: "http://localhost:11434",
  prompt: "",
  filename: "office-agent-deck",
});

function getInitialLanguage(): Language {
  return navigator.language.toLowerCase().startsWith("zh") ? "zh-CN" : "en";
}

function normalizeStageError(message: string): string {
  return message.replace(/secret|bearer\s+[a-z0-9._-]+/gi, "****");
}

function safeThemeCustom(custom?: CustomThemeConfig | null, preset: ThemePreset = "default"): Required<CustomThemeConfig> {
  return {
    primary_color: custom?.primary_color ?? presetThemeDefaults[preset].primary_color,
    accent_color: custom?.accent_color ?? presetThemeDefaults[preset].accent_color,
    background_color: custom?.background_color ?? presetThemeDefaults[preset].background_color,
    heading_font: custom?.heading_font ?? presetThemeDefaults[preset].heading_font,
    body_font: custom?.body_font ?? presetThemeDefaults[preset].body_font,
    cover_style: custom?.cover_style ?? presetThemeDefaults[preset].cover_style,
  };
}

function parseImportedTheme(rawValue: string): ThemeSpec {
  const parsed = JSON.parse(rawValue) as Partial<ThemeSpec>;
  if (!parsed || typeof parsed !== "object") {
    throw new Error("Invalid theme JSON.");
  }
  if (!parsed.preset || !["default", "executive", "editorial"].includes(parsed.preset)) {
    throw new Error("Theme preset must be default, executive, or editorial.");
  }
  const custom = safeThemeCustom(parsed.custom, parsed.preset as ThemePreset);
  for (const color of [custom.primary_color, custom.accent_color, custom.background_color]) {
    if (!/^#[0-9A-Fa-f]{6}$/.test(color ?? "")) {
      throw new Error("Theme colors must use #RRGGBB format.");
    }
  }
  if (!["band", "centered", "minimal"].includes(custom.cover_style ?? "")) {
    throw new Error("Cover style must be band, centered, or minimal.");
  }
  return { preset: parsed.preset as ThemePreset, custom };
}

function buildTemplateMapping(mapping: TemplateMappingDraft): TemplateSelectionSpec | undefined {
  if (templatePartOrder.every((part) => typeof mapping[part] === "number")) {
    return {
      opening: mapping.opening as number,
      agenda: mapping.agenda as number,
      content: mapping.content as number,
      closing: mapping.closing as number,
    };
  }
  return undefined;
}

function buildGenerateSpecRequest(form: FormState, mapping: TemplateSelectionSpec | undefined): GenerateSpecRequest {
  return {
    prompt: form.prompt,
    provider: form.provider,
    model: form.model,
    theme: form.theme,
    templateMapping: mapping,
    apiKey: form.provider === "openai" ? form.apiKey : undefined,
    openaiBaseUrl: form.openaiBaseUrl,
    ollamaBaseUrl: form.ollamaBaseUrl,
  };
}

function summarizeSlide(slide: SlideSpec, labels: { keyPointsLabel: string; nextStepsLabel: string }): string[] {
  switch (slide.type) {
    case "title":
    case "section":
      return slide.subtitle ? [slide.subtitle] : [];
    case "bullets":
      return slide.bullets;
    case "two_column":
      return [...slide.left_bullets, ...slide.right_bullets];
    case "image":
      return slide.bullets.length > 0 ? slide.bullets : [slide.caption ?? slide.image];
    case "timeline":
      return slide.events.map((event) => `${event.label}: ${event.title}`);
    case "quote":
      return [slide.quote, [slide.attribution, slide.source].filter(Boolean).join(" / ")].filter(Boolean);
    case "comparison":
      return [slide.left.title, ...slide.left.bullets, slide.right.title, ...slide.right.bullets];
    case "summary":
      return [labels.keyPointsLabel, ...slide.key_points, labels.nextStepsLabel, ...slide.next_steps];
    case "table":
      return [slide.headers.join(" | "), ...slide.rows.map((row) => row.join(" | "))];
  }
}

export default function App() {
  const [language, setLanguage] = useState<Language>(getInitialLanguage);
  const [form, setForm] = useState<FormState>(initialFormState);
  const [stages, setStages] = useState<StageState>(initialStages);
  const [spec, setSpec] = useState<PresentationSpec | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [themeJsonDraft, setThemeJsonDraft] = useState(JSON.stringify(initialTheme(), null, 2));
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [templateSlides, setTemplateSlides] = useState<TemplatePreviewSlide[]>([]);
  const [templateMappingDraft, setTemplateMappingDraft] = useState<TemplateMappingDraft>({});
  const [templatePreviewError, setTemplatePreviewError] = useState<string | null>(null);
  const [isTemplatePreviewLoading, setIsTemplatePreviewLoading] = useState(false);
  const [isGeneratingSpec, setIsGeneratingSpec] = useState(false);
  const [isGeneratingPpt, setIsGeneratingPpt] = useState(false);

  const copy = locales[language];
  const stageCopy = workflowLabels[language];
  const rawSpec = useMemo(() => (spec ? JSON.stringify(spec, null, 2) : ""), [spec]);
  const normalizedTheme = useMemo(
    () => ({ ...form.theme, custom: safeThemeCustom(form.theme.custom, form.theme.preset) }),
    [form.theme],
  );
  const templateMapping = useMemo(() => buildTemplateMapping(templateMappingDraft), [templateMappingDraft]);
  const templateSelectionRequired = Boolean(templateFile) && !templateMapping;

  function syncThemeDraft(nextTheme: ThemeSpec): void {
    setThemeJsonDraft(JSON.stringify({ ...nextTheme, custom: safeThemeCustom(nextTheme.custom, nextTheme.preset) }, null, 2));
  }

  function clearGeneratedOutput(promptText: string = form.prompt): void {
    setError(null);
    setWarnings([]);
    if (!spec) {
      return;
    }
    setSpec(null);
    setStages({
      input: promptText.trim() ? "success" : "idle",
      llm: "idle",
      render: "idle",
      office: "idle",
    });
  }

  function updateForm(nextForm: FormState): void {
    setForm(nextForm);
    syncThemeDraft(nextForm.theme);
    clearGeneratedOutput(nextForm.prompt);
  }

  function updateField<Key extends keyof FormState>(key: Key, value: FormState[Key]): void {
    if (key === "provider") {
      const nextProvider = value as Provider;
      const nextModel = form.model === defaultModels[form.provider] ? defaultModels[nextProvider] : form.model;
      updateForm({ ...form, provider: nextProvider, model: nextModel });
      return;
    }
    updateForm({ ...form, [key]: value });
  }

  function updateThemePreset(value: ThemePreset): void {
    updateForm({
      ...form,
      theme: {
        preset: value,
        custom: presetThemeDefaults[value],
      },
    });
  }

  function updateThemeCustom<Key extends keyof CustomThemeConfig>(key: Key, value: NonNullable<CustomThemeConfig[Key]>): void {
    updateForm({
      ...form,
      theme: {
        ...form.theme,
        custom: {
          ...safeThemeCustom(form.theme.custom, form.theme.preset),
          [key]: value,
        },
      },
    });
  }

  function updateTemplateMapping(part: TemplatePart, value: string): void {
    setTemplateMappingDraft((current) => ({
      ...current,
      [part]: value === "" ? undefined : Number(value),
    }));
    clearGeneratedOutput();
  }

  async function handleTemplateChange(event: ChangeEvent<HTMLInputElement>): Promise<void> {
    const nextFile = event.target.files?.[0] ?? null;
    setTemplateFile(nextFile);
    setTemplateSlides([]);
    setTemplateMappingDraft({});
    setTemplatePreviewError(null);
    clearGeneratedOutput();

    if (!nextFile) {
      return;
    }

    setIsTemplatePreviewLoading(true);
    try {
      const preview = await previewTemplate(nextFile);
      setTemplateSlides(preview.slides);
    } catch (caught) {
      const message = normalizeStageError(caught instanceof Error ? caught.message : "Unknown error");
      setTemplatePreviewError(message);
    } finally {
      setIsTemplatePreviewLoading(false);
    }
  }

  function clearTemplate(): void {
    setTemplateFile(null);
    setTemplateSlides([]);
    setTemplateMappingDraft({});
    setTemplatePreviewError(null);
    clearGeneratedOutput();
  }

  async function handleGenerateSpec(): Promise<void> {
    if (templateSelectionRequired) {
      setError(copy.templateSelectionRequired);
      return;
    }

    setError(null);
    setWarnings([]);
    setIsGeneratingSpec(true);
    setStages({
      input: form.prompt.trim() ? "success" : "error",
      llm: "running",
      render: "idle",
      office: "idle",
    });
    try {
      const nextSpec = await generateSpec(buildGenerateSpecRequest(form, templateMapping));
      setSpec(nextSpec);
      setStages({
        input: "success",
        llm: "success",
        render: "idle",
        office: "idle",
      });
    } catch (caught) {
      const message = normalizeStageError(caught instanceof Error ? caught.message : "Unknown error");
      setError(message);
      setStages({
        input: form.prompt.trim() ? "success" : "error",
        llm: "error",
        render: "idle",
        office: "idle",
      });
    } finally {
      setIsGeneratingSpec(false);
    }
  }

  async function handleGeneratePpt(): Promise<void> {
    if (!spec) {
      return;
    }
    setError(null);
    setWarnings([]);
    setIsGeneratingPpt(true);
    setStages((current) => ({
      ...current,
      render: "running",
      office: "idle",
    }));
    try {
      const result = await renderPresentation({ filename: form.filename, spec }, templateFile);
      downloadBlob(result.blob, form.filename);
      setWarnings(result.warnings);
      setStages((current) => ({
        ...current,
        render: "success",
        office: "skipped",
      }));
    } catch (caught) {
      const message = normalizeStageError(caught instanceof Error ? caught.message : "Unknown error");
      setError(message);
      setStages((current) => ({
        ...current,
        render: "error",
        office: "idle",
      }));
    } finally {
      setIsGeneratingPpt(false);
    }
  }

  async function handleCopyJson(): Promise<void> {
    if (!rawSpec || !navigator.clipboard) {
      return;
    }
    await navigator.clipboard.writeText(rawSpec);
  }

  function handleExportTheme(): void {
    const blob = new Blob([JSON.stringify(normalizedTheme, null, 2)], { type: "application/json" });
    downloadBlob(blob, "theme.json");
  }

  function handleImportTheme(): void {
    try {
      const nextTheme = parseImportedTheme(themeJsonDraft);
      updateForm({ ...form, theme: nextTheme });
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Invalid theme JSON.");
    }
  }

  const canGenerateSpec =
    Boolean(form.prompt.trim()) &&
    !isGeneratingSpec &&
    !isGeneratingPpt &&
    !isTemplatePreviewLoading &&
    (!templateFile || (Boolean(templateSlides.length) && !templatePreviewError && Boolean(templateMapping)));

  return (
    <main className="app-shell">
      <section className="hero-panel compact-hero">
        <div className="hero-copy">
          <div className="header-row">
            <div>
              <p className="eyebrow">{copy.appTitle}</p>
              <h1>{copy.heroTitle}</h1>
            </div>
            <div className="language-switch" aria-label={copy.language}>
              <button className={language === "en" ? "lang-button active" : "lang-button"} onClick={() => setLanguage("en")}>
                {copy.english}
              </button>
              <button className={language === "zh-CN" ? "lang-button active" : "lang-button"} onClick={() => setLanguage("zh-CN")}>
                {copy.chinese}
              </button>
            </div>
          </div>
          <p className="hero-description">{copy.heroDescription}</p>
        </div>
        <div className="hero-metrics compact-metrics">
          <Metric label={copy.provider} value={copy.providersMetric} />
          <Metric label="Flow" value={copy.flowMetric} />
          <Metric label="Output" value={copy.outputMetric} />
        </div>
      </section>

      <section className="workspace-grid">
        <aside className="sidebar-column">
          <div className="panel stack">
            <header className="panel-header">
              <div>
                <p className="panel-kicker">{copy.configuration}</p>
                <h2>{copy.runtimeControls}</h2>
              </div>
            </header>
            <div className="form-grid">
              <Field label={copy.provider}>
                <select aria-label={copy.provider} value={form.provider} onChange={(event) => updateField("provider", event.target.value as Provider)}>
                  <option value="openai">OpenAI-compatible</option>
                  <option value="ollama">Ollama</option>
                </select>
              </Field>
              <Field label={copy.model}>
                <input aria-label={copy.model} value={form.model} onChange={(event) => updateField("model", event.target.value)} />
              </Field>
              {form.provider === "openai" ? (
                <>
                  <Field label={copy.apiKey}>
                    <input
                      aria-label={copy.apiKey}
                      type="password"
                      value={form.apiKey}
                      onChange={(event) => updateField("apiKey", event.target.value)}
                      placeholder="Only used for this session"
                    />
                  </Field>
                  <Field label={copy.openaiBaseUrl}>
                    <input aria-label={copy.openaiBaseUrl} value={form.openaiBaseUrl} onChange={(event) => updateField("openaiBaseUrl", event.target.value)} />
                  </Field>
                </>
              ) : (
                <Field label={copy.ollamaBaseUrl} fullWidth>
                  <input aria-label={copy.ollamaBaseUrl} value={form.ollamaBaseUrl} onChange={(event) => updateField("ollamaBaseUrl", event.target.value)} />
                </Field>
              )}
            </div>
          </div>

          <div className="panel stack">
            <header className="panel-header">
              <div>
                <p className="panel-kicker">{copy.themeStudio}</p>
                <h2>{copy.themeControls}</h2>
              </div>
            </header>
            <div className="form-grid">
              <Field label={copy.themePreset}>
                <select aria-label={copy.themePreset} value={form.theme.preset} onChange={(event) => updateThemePreset(event.target.value as ThemePreset)}>
                  <option value="default">default</option>
                  <option value="executive">executive</option>
                  <option value="editorial">editorial</option>
                </select>
              </Field>
              <Field label={copy.coverStyle}>
                <select
                  aria-label={copy.coverStyle}
                  value={safeThemeCustom(form.theme.custom, form.theme.preset).cover_style}
                  onChange={(event) => updateThemeCustom("cover_style", event.target.value as "band" | "centered" | "minimal")}
                >
                  <option value="band">band</option>
                  <option value="centered">centered</option>
                  <option value="minimal">minimal</option>
                </select>
              </Field>
              <Field label={copy.primaryColor}>
                <input aria-label={copy.primaryColor} type="color" value={safeThemeCustom(form.theme.custom, form.theme.preset).primary_color} onChange={(event) => updateThemeCustom("primary_color", event.target.value)} />
              </Field>
              <Field label={copy.accentColor}>
                <input aria-label={copy.accentColor} type="color" value={safeThemeCustom(form.theme.custom, form.theme.preset).accent_color} onChange={(event) => updateThemeCustom("accent_color", event.target.value)} />
              </Field>
              <Field label={copy.backgroundColor}>
                <input aria-label={copy.backgroundColor} type="color" value={safeThemeCustom(form.theme.custom, form.theme.preset).background_color} onChange={(event) => updateThemeCustom("background_color", event.target.value)} />
              </Field>
              <Field label={copy.headingFont}>
                <select aria-label={copy.headingFont} value={safeThemeCustom(form.theme.custom, form.theme.preset).heading_font} onChange={(event) => updateThemeCustom("heading_font", event.target.value)}>
                  {headingFontOptions.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label={copy.bodyFont}>
                <select aria-label={copy.bodyFont} value={safeThemeCustom(form.theme.custom, form.theme.preset).body_font} onChange={(event) => updateThemeCustom("body_font", event.target.value)}>
                  {bodyFontOptions.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              </Field>
            </div>

            <div className="subpanel">
              <div className="subpanel-header">
                <strong>{copy.templateTitle}</strong>
                {templateFile ? <span className="template-chip">{templateFile.name}</span> : null}
              </div>
              <Field label={copy.templateFile} fullWidth>
                <input
                  aria-label={copy.templateFile}
                  type="file"
                  accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                  onChange={(event) => void handleTemplateChange(event)}
                />
              </Field>
              <p className="helper-text">{copy.templateHint}</p>
              {templateFile ? (
                <div className="inline-actions">
                  <span className="template-meta">
                    {copy.templateSelected}: {templateFile.name}
                  </span>
                  <button className="ghost-button" onClick={clearTemplate}>
                    {copy.clearTemplate}
                  </button>
                </div>
              ) : null}
            </div>

            <div className="subpanel">
              <div className="subpanel-header">
                <strong>{copy.templatePreview}</strong>
              </div>
              {isTemplatePreviewLoading ? <p className="helper-text">{copy.templatePreviewLoading}</p> : null}
              {templatePreviewError ? <p className="error-banner compact-error">{`${copy.templatePreviewError}: ${templatePreviewError}`}</p> : null}
              {!templateFile && !isTemplatePreviewLoading ? <p className="helper-text">{copy.templatePreviewEmpty}</p> : null}
              {templateSlides.length > 0 ? (
                <div className="template-grid">
                  {templateSlides.map((slide) => (
                    <article key={slide.index} className="template-card">
                      <img src={slide.thumbnailDataUrl} alt={formatTemplateSlideLabel(copy.slideIndex, slide.index, slide.titleText)} />
                      <div className="template-card-body">
                        <strong>{formatTemplateSlideLabel(copy.slideIndex, slide.index, slide.titleText)}</strong>
                        <p>{slide.titleText ?? "-"}</p>
                        <div className="role-badge-row">
                          {slide.placeholderRoles.map((role) => (
                            <span key={`${role}-${slide.index}`} className="role-badge">
                              {role}
                            </span>
                          ))}
                        </div>
                        <div className="role-badge-row selected-badges">
                          {templatePartOrder
                            .filter((part) => templateMappingDraft[part] === slide.index)
                            .map((part) => (
                              <span key={`${part}-${slide.index}`} className="selection-badge">
                                {partLabel(copy, part)}
                              </span>
                            ))}
                        </div>
                      </div>
                    </article>
                  ))}
                </div>
              ) : null}
            </div>

            {templateSlides.length > 0 ? (
              <div className="subpanel">
                <div className="subpanel-header">
                  <strong>{copy.templateMappings}</strong>
                </div>
                <div className="mapping-grid">
                  {templatePartOrder.map((part) => (
                    <Field key={part} label={partLabel(copy, part)} fullWidth>
                      <select value={templateMappingDraft[part] ?? ""} onChange={(event) => updateTemplateMapping(part, event.target.value)}>
                        <option value="">--</option>
                        {templateSlides.map((slide) => (
                          <option key={`${part}-${slide.index}`} value={slide.index}>
                            {formatTemplateSlideLabel(copy.slideIndex, slide.index, slide.titleText)}
                          </option>
                        ))}
                      </select>
                    </Field>
                  ))}
                </div>
                {templateSelectionRequired ? <p className="helper-text">{copy.templateSelectionRequired}</p> : null}
              </div>
            ) : null}

            <Field label={copy.themeJson} fullWidth>
              <textarea aria-label={copy.themeJson} rows={6} value={themeJsonDraft} onChange={(event) => setThemeJsonDraft(event.target.value)} />
            </Field>
            <div className="action-row">
              <button className="secondary-button" onClick={handleExportTheme}>
                {copy.exportThemeJson}
              </button>
              <button className="ghost-button" onClick={handleImportTheme}>
                {copy.importThemeJson}
              </button>
            </div>
            <div className="theme-summary">
              <strong>{copy.themeSummary}</strong>
              <p>{`${form.theme.preset} / ${safeThemeCustom(form.theme.custom, form.theme.preset).cover_style}`}</p>
              {templateFile ? <p>{`${copy.templateSelected}: ${templateFile.name}`}</p> : null}
              {templateMapping ? <p>{templatePartOrder.map((part) => `${partLabel(copy, part)}: ${templateMapping[part] + 1}`).join(" | ")}</p> : null}
            </div>
          </div>

          <div className="panel stack workflow-panel">
            <header className="panel-header">
              <div>
                <p className="panel-kicker">{copy.workflow}</p>
                <h2>{copy.pipelineStatus}</h2>
              </div>
            </header>
            <div className="workflow-list">
              {(Object.keys(stageCopy) as WorkflowStageKey[]).map((key) => (
                <article key={key} className={`workflow-card status-${stages[key]}`}>
                  <div className="workflow-card-top">
                    <div>
                      <h3>{stageCopy[key].title}</h3>
                      <p>{stageCopy[key].subtitle}</p>
                    </div>
                    <StatusPill label={statusLabels[language][stages[key]]} status={stages[key]} />
                  </div>
                  {key === "office" ? <p className="workflow-note">{copy.officePlaceholder}</p> : null}
                </article>
              ))}
            </div>
          </div>
        </aside>

        <section className="main-column">
          <div className="panel stack">
            <header className="panel-header">
              <div>
                <p className="panel-kicker">{copy.input}</p>
                <h2>{copy.describeDeck}</h2>
              </div>
            </header>
            <Field label={copy.prompt} fullWidth>
              <textarea aria-label={copy.prompt} rows={7} value={form.prompt} onChange={(event) => updateField("prompt", event.target.value)} placeholder={copy.promptPlaceholder} />
            </Field>
            <div className="form-grid compact">
              <Field label={copy.filename}>
                <input aria-label={copy.filename} value={form.filename} onChange={(event) => updateField("filename", event.target.value)} />
              </Field>
            </div>
            <div className="action-row">
              <button className="primary-button" onClick={() => void handleGenerateSpec()} disabled={!canGenerateSpec}>
                {isGeneratingSpec ? copy.generatingSpec : copy.generateSpec}
              </button>
              <button className="secondary-button" onClick={() => void handleGeneratePpt()} disabled={!spec || isGeneratingSpec || isGeneratingPpt}>
                {isGeneratingPpt ? copy.renderingPpt : copy.confirmGenerate}
              </button>
              <button className="ghost-button" onClick={() => void handleCopyJson()} disabled={!spec}>
                {copy.copyJson}
              </button>
            </div>
            {error ? <p className="error-banner">{error}</p> : null}
            {warnings.length > 0 ? (
              <div className="warning-list">
                {warnings.map((warning) => (
                  <p key={warning}>{warning}</p>
                ))}
              </div>
            ) : null}
          </div>

          <div className="preview-grid dense-preview">
            <div className="panel stack">
              <header className="panel-header">
                <div>
                  <p className="panel-kicker">{copy.structuredSpec}</p>
                  <h2>{copy.validatedJson}</h2>
                </div>
              </header>
              {spec ? <pre className="json-preview">{rawSpec}</pre> : <EmptyState title={copy.noSpecTitle} description={copy.noSpecDescription} />}
            </div>

            <div className="panel stack">
              <header className="panel-header">
                <div>
                  <p className="panel-kicker">{copy.slidePreview}</p>
                  <h2>{copy.presentationOutline}</h2>
                </div>
              </header>
              {spec ? (
                <div className="slide-grid">
                  {spec.slides.map((slide, index) => (
                    <article key={`${slide.type}-${index}`} className="slide-card">
                      <div className="slide-card-top">
                        <div className="slide-badge">{slide.type}</div>
                        <div className="selection-badge">{`${copy.partPrefix}: ${partLabel(copy, slide.part)}`}</div>
                      </div>
                      <h3>{slide.title}</h3>
                      <ul>
                        {summarizeSlide(slide, copy).map((item) => (
                          <li key={`${item}-${index}`}>{item}</li>
                        ))}
                      </ul>
                    </article>
                  ))}
                </div>
              ) : (
                <EmptyState title={copy.noSlidesTitle} description={copy.noSlidesDescription} />
              )}
            </div>
          </div>
        </section>
      </section>
    </main>
  );
}

function formatTemplateSlideLabel(prefix: string, index: number, titleText?: string | null): string {
  const title = titleText?.trim();
  return title ? `${prefix} ${index + 1}: ${title}` : `${prefix} ${index + 1}`;
}

function partLabel(copy: Record<string, string>, part: TemplatePart): string {
  switch (part) {
    case "opening":
      return copy.templatePartOpening;
    case "agenda":
      return copy.templatePartAgenda;
    case "content":
      return copy.templatePartContent;
    case "closing":
      return copy.templatePartClosing;
  }
}

function Metric({ label, value }: { label: string; value: string }) {
  return (
    <div className="metric-card">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function Field({ label, children, fullWidth = false }: { label: string; children: ReactNode; fullWidth?: boolean }) {
  return (
    <label className={fullWidth ? "field full-width" : "field"}>
      <span>{label}</span>
      {children}
    </label>
  );
}

function StatusPill({ label, status }: { label: string; status: WorkflowStageStatus }) {
  return <span className={`status-pill status-${status}`}>{label}</span>;
}

function EmptyState({ title, description }: { title: string; description: string }) {
  return (
    <div className="empty-state">
      <h3>{title}</h3>
      <p>{description}</p>
    </div>
  );
}
