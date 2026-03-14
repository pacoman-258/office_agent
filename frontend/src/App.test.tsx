import { fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";

import App from "./App";

const previewResponse = {
  cleanupMode: "preserve_branding" as const,
  slides: [
    { index: 0, thumbnailDataUrl: "data:image/png;base64,aaa", titleText: "Cover", placeholderRoles: ["title", "subtitle"] },
    { index: 1, thumbnailDataUrl: "data:image/png;base64,bbb", titleText: "Content", placeholderRoles: ["title", "body"] },
  ],
};

const specResponse = {
  title: "Demo",
  theme: {
    preset: "editorial" as const,
    custom: {
      primary_color: "#112233",
      accent_color: "#445566",
      background_color: "#F5F5F5",
      heading_font: "Georgia",
      body_font: "Microsoft YaHei",
      cover_style: "minimal" as const,
    },
  },
  template: { opening: 0, agenda: 1, content: 1, closing: 1 },
  slides: [
    { type: "title" as const, part: "opening" as const, title: "Intro", subtitle: "Subtitle" },
    {
      type: "timeline" as const,
      part: "content" as const,
      title: "Roadmap",
      events: [
        { label: "Q1", title: "Plan", detail: "Scope" },
        { label: "Q2", title: "Launch", detail: "Ship" },
      ],
    },
  ],
};

function mockJsonResponse(payload: unknown, init?: ResponseInit): Response {
  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: { "Content-Type": "application/json" },
    ...init,
  });
}

function mockBinaryResponse(init?: ResponseInit): Response {
  return new Response("pptx-bytes", {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "X-Office-Agent-Warnings": encodeURIComponent(JSON.stringify([])),
      "X-Office-Agent-Finalize": encodeURIComponent(
        JSON.stringify({
          enabled: true,
          status: "completed",
          rounds: [{ roundIndex: 1, slidesReviewed: 1, issuesFound: 1, operationsApplied: 1, warnings: [] }],
          issuesFound: 1,
          operationsApplied: 1,
          warnings: [],
        }),
      ),
    },
    ...init,
  });
}

describe("App", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.stubGlobal("fetch", vi.fn());
    Object.defineProperty(globalThis.navigator, "clipboard", {
      value: { writeText: vi.fn() },
      configurable: true,
    });
    Object.defineProperty(globalThis.navigator, "language", {
      value: "en-US",
      configurable: true,
    });
    Object.defineProperty(HTMLAnchorElement.prototype, "click", {
      value: vi.fn(),
      configurable: true,
    });
    vi.stubGlobal("URL", {
      createObjectURL: vi.fn(() => "blob:demo"),
      revokeObjectURL: vi.fn(),
    });
  });

  it("toggles runtime fields when provider changes", async () => {
    render(<App />);
    expect(screen.getByLabelText("API Key")).toBeInTheDocument();
    await userEvent.selectOptions(screen.getByLabelText("Provider"), "ollama");
    expect(screen.queryByLabelText("API Key")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Ollama Base URL")).toBeInTheDocument();
  });

  it("uses manual template mapping by default without requesting preview", async () => {
    render(<App />);
    const file = new File(["pptx-template"], "brand-template.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    await userEvent.upload(screen.getByLabelText("Import PPTX Template"), file);
    expect(screen.getByText("Manual Template Mapping")).toBeInTheDocument();
    expect(screen.getByText("PowerPoint page numbers start at 1.")).toBeInTheDocument();
    expect(fetch).not.toHaveBeenCalled();
  });

  it("can optionally load template previews", async () => {
    vi.mocked(fetch).mockResolvedValueOnce(mockJsonResponse(previewResponse));
    render(<App />);
    const file = new File(["pptx-template"], "brand-template.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    await userEvent.upload(screen.getByLabelText("Import PPTX Template"), file);
    await userEvent.click(screen.getByRole("button", { name: "Preview-assisted Mapping" }));
    await userEvent.click(screen.getByRole("button", { name: "Load Template Preview" }));
    await screen.findByText("Cover");
    expect(screen.getByText(/Template rendering preserves branding elements/i)).toBeInTheDocument();
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it("sends template mapping when generating a spec", async () => {
    vi.mocked(fetch).mockResolvedValueOnce(mockJsonResponse(specResponse));
    render(<App />);
    const file = new File(["pptx-template"], "brand-template.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    await userEvent.upload(screen.getByLabelText("Import PPTX Template"), file);
    await userEvent.type(screen.getByLabelText("Opening"), "1");
    await userEvent.type(screen.getByLabelText("Agenda"), "2");
    await userEvent.type(screen.getByLabelText("Content"), "2");
    await userEvent.type(screen.getByLabelText("Closing"), "2");
    await userEvent.type(screen.getByLabelText("Prompt"), "Create a team update deck");
    await userEvent.click(screen.getByRole("button", { name: "Generate Spec" }));
    await screen.findByText("Roadmap");
    const specCall = vi.mocked(fetch).mock.calls[0];
    const payload = JSON.parse((specCall[1] as RequestInit).body as string);
    expect(payload.templateMapping).toEqual({ opening: 0, agenda: 1, content: 1, closing: 1 });
  });

  it("supports language switching", async () => {
    render(<App />);
    await userEvent.click(screen.getByRole("button", { name: "中文" }));
    expect(screen.getByText("主题、模板映射与可选预览")).toBeInTheDocument();
    expect(
      screen.getByText("默认模式不依赖模板预览。上传文件后可直接填写页码映射；如果你想看缩略图，再切换到预览模式。"),
    ).toBeInTheDocument();
  });

  it("imports theme json with validation", async () => {
    render(<App />);
    fireEvent.change(screen.getByLabelText("Theme JSON"), {
      target: {
        value: JSON.stringify(
          {
            preset: "executive",
            custom: {
              primary_color: "#223344",
              accent_color: "#556677",
              background_color: "#FAFAFA",
              heading_font: "Georgia",
              body_font: "Segoe UI",
              cover_style: "centered",
            },
          },
          null,
          2,
        ),
      },
    });
    await userEvent.click(screen.getByRole("button", { name: "Import Theme JSON" }));
    expect(screen.getByText("executive / centered")).toBeInTheDocument();
  });

  it("uploads a pptx template during render and shows finalization summary", async () => {
    vi.mocked(fetch).mockResolvedValueOnce(mockJsonResponse(specResponse)).mockResolvedValueOnce(mockBinaryResponse());
    render(<App />);
    const file = new File(["pptx-template"], "brand-template.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    await userEvent.upload(screen.getByLabelText("Import PPTX Template"), file);
    await userEvent.type(screen.getByLabelText("Opening"), "1");
    await userEvent.type(screen.getByLabelText("Agenda"), "2");
    await userEvent.type(screen.getByLabelText("Content"), "2");
    await userEvent.type(screen.getByLabelText("Closing"), "2");
    await userEvent.type(screen.getByLabelText("Prompt"), "Create a team update deck");
    await userEvent.click(screen.getByRole("button", { name: "Generate Spec" }));
    await screen.findByText("Roadmap");
    await userEvent.selectOptions(screen.getByLabelText("Enable Visual Review"), "true");
    await userEvent.click(screen.getByRole("button", { name: "Confirm and Generate PPT" }));
    expect(fetch).toHaveBeenCalledTimes(2);
    const renderCall = vi.mocked(fetch).mock.calls[1];
    const requestInit = renderCall[1] as RequestInit;
    expect(requestInit.body).toBeInstanceOf(FormData);
    expect(screen.getByText("Visual Review Summary")).toBeInTheDocument();
  });
});
