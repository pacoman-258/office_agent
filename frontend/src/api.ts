import type {
  GenerateSpecRequest,
  PresentationSpec,
  RenderPresentationRequest,
  TemplatePreviewResponse,
} from "./types";

async function readError(response: Response): Promise<string> {
  try {
    const payload = (await response.json()) as { detail?: string | Array<{ msg?: string }> };
    if (Array.isArray(payload.detail)) {
      return payload.detail.map((item) => item.msg).filter(Boolean).join("; ") || `Request failed with status ${response.status}`;
    }
    return payload.detail ?? `Request failed with status ${response.status}`;
  } catch {
    return `Request failed with status ${response.status}`;
  }
}

export async function generateSpec(payload: GenerateSpecRequest): Promise<PresentationSpec> {
  const response = await fetch("/api/specs", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await readError(response));
  }
  return (await response.json()) as PresentationSpec;
}

export async function previewTemplate(templateFile: File): Promise<TemplatePreviewResponse> {
  const formData = new FormData();
  formData.append("template", templateFile);
  const response = await fetch("/api/templates/preview", {
    method: "POST",
    body: formData,
  });
  if (!response.ok) {
    throw new Error(await readError(response));
  }
  return (await response.json()) as TemplatePreviewResponse;
}

export interface RenderPresentationResult {
  blob: Blob;
  warnings: string[];
}

export async function renderPresentation(
  payload: RenderPresentationRequest,
  templateFile?: File | null,
): Promise<RenderPresentationResult> {
  const response = await fetch(
    "/api/presentations",
    templateFile
      ? {
          method: "POST",
          body: buildMultipartPayload(payload, templateFile),
        }
      : {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        },
  );
  if (!response.ok) {
    throw new Error(await readError(response));
  }
  const rawWarnings = response.headers.get("X-Office-Agent-Warnings");
  return {
    blob: await response.blob(),
    warnings: rawWarnings ? JSON.parse(decodeURIComponent(rawWarnings)) : [],
  };
}

function buildMultipartPayload(payload: RenderPresentationRequest, templateFile: File): FormData {
  const formData = new FormData();
  formData.append("payload", JSON.stringify(payload));
  formData.append("template", templateFile);
  return formData;
}

export function downloadBlob(blob: Blob, filename: string): void {
  const safeName = filename.includes(".") ? filename : `${filename}.pptx`;
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = safeName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
