export type Provider = "openai" | "ollama";
export type Language = "en" | "zh-CN";
export type WorkflowStageStatus = "idle" | "running" | "success" | "error" | "skipped";
export type WorkflowStageKey = "input" | "llm" | "render" | "office";
export type ThemePreset = "default" | "executive" | "editorial";
export type CoverStyle = "band" | "centered" | "minimal";
export type TemplatePart = "opening" | "agenda" | "content" | "closing";

export interface CustomThemeConfig {
  primary_color?: string;
  accent_color?: string;
  background_color?: string;
  heading_font?: string;
  body_font?: string;
  cover_style?: CoverStyle;
}

export interface ThemeSpec {
  preset: ThemePreset;
  custom?: CustomThemeConfig | null;
}

export interface TemplateSelectionSpec {
  opening: number;
  agenda: number;
  content: number;
  closing: number;
}

export interface BaseSlideSpec {
  part: TemplatePart;
  title: string;
}

export interface TitleSlideSpec extends BaseSlideSpec {
  type: "title";
  subtitle?: string | null;
}

export interface SectionSlideSpec extends BaseSlideSpec {
  type: "section";
  subtitle?: string | null;
}

export interface BulletsSlideSpec extends BaseSlideSpec {
  type: "bullets";
  bullets: string[];
}

export interface TwoColumnSlideSpec extends BaseSlideSpec {
  type: "two_column";
  left_title?: string | null;
  left_bullets: string[];
  right_title?: string | null;
  right_bullets: string[];
}

export interface ImageSlideSpec extends BaseSlideSpec {
  type: "image";
  image: string;
  caption?: string | null;
  bullets: string[];
}

export interface TimelineEvent {
  label: string;
  title: string;
  detail?: string | null;
}

export interface TimelineSlideSpec extends BaseSlideSpec {
  type: "timeline";
  events: TimelineEvent[];
}

export interface QuoteSlideSpec extends BaseSlideSpec {
  type: "quote";
  quote: string;
  attribution?: string | null;
  source?: string | null;
}

export interface ComparisonColumn {
  title: string;
  bullets: string[];
}

export interface ComparisonSlideSpec extends BaseSlideSpec {
  type: "comparison";
  left: ComparisonColumn;
  right: ComparisonColumn;
}

export interface SummarySlideSpec extends BaseSlideSpec {
  type: "summary";
  key_points: string[];
  next_steps: string[];
}

export interface TableSlideSpec extends BaseSlideSpec {
  type: "table";
  headers: string[];
  rows: string[][];
}

export type SlideSpec =
  | TitleSlideSpec
  | SectionSlideSpec
  | BulletsSlideSpec
  | TwoColumnSlideSpec
  | ImageSlideSpec
  | TimelineSlideSpec
  | QuoteSlideSpec
  | ComparisonSlideSpec
  | SummarySlideSpec
  | TableSlideSpec;

export interface PresentationSpec {
  title: string;
  theme: ThemeSpec;
  template?: TemplateSelectionSpec | null;
  slides: SlideSpec[];
}

export interface GenerateSpecRequest {
  prompt: string;
  provider: Provider;
  model: string;
  theme: ThemeSpec;
  templateMapping?: TemplateSelectionSpec;
  apiKey?: string;
  openaiBaseUrl?: string;
  ollamaBaseUrl?: string;
}

export interface RenderPresentationRequest {
  filename: string;
  spec: PresentationSpec;
}

export interface TemplatePreviewSlide {
  index: number;
  thumbnailDataUrl: string;
  titleText?: string | null;
  placeholderRoles: string[];
}

export interface TemplatePreviewResponse {
  slides: TemplatePreviewSlide[];
}
