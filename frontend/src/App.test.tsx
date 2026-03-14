import { fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";

import App from "./App";

const previewResponse = {
  slides: [
    {
      index: 0,
      thumbnailDataUrl: "data:image/png;base64,aaa",
      titleText: "Cover",
      placeholderRoles: ["title", "subtitle"],
    },
    {
      index: 1,
      thumbnailDataUrl: "data:image/png;base64,bbb",
      titleText: "Content",
      placeholderRoles: ["title", "body"],
    },
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
  template: {
    opening: 0,
    agenda: 1,
    content: 1,
    closing: 1,
  },
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

  it("loads template previews and requires part mapping before generation", async () => {
    vi.mocked(fetch).mockResolvedValueOnce(mockJsonResponse(previewResponse));
    render(<App />);

    const file = new File(["pptx-template"], "brand-template.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    await userEvent.upload(screen.getByLabelText("Import PPTX Template"), file);

    await screen.findByText("Cover");
    expect(screen.getByText("Template Part Mapping")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Generate Spec" })).toBeDisabled();
  });

  it("sends template mapping when generating a spec", async () => {
    vi.mocked(fetch)
      .mockResolvedValueOnce(mockJsonResponse(previewResponse))
      .mockResolvedValueOnce(mockJsonResponse(specResponse));
    render(<App />);

    const file = new File(["pptx-template"], "brand-template.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    await userEvent.upload(screen.getByLabelText("Import PPTX Template"), file);
    await screen.findByText("Cover");

    await userEvent.selectOptions(screen.getByLabelText("Opening"), "0");
    await userEvent.selectOptions(screen.getByLabelText("Agenda"), "1");
    await userEvent.selectOptions(screen.getByLabelText("Content"), "1");
    await userEvent.selectOptions(screen.getByLabelText("Closing"), "1");
    await userEvent.type(screen.getByLabelText("Prompt"), "Create a team update deck");
    await userEvent.click(screen.getByRole("button", { name: "Generate Spec" }));

    await screen.findByText("Roadmap");
    const specCall = vi.mocked(fetch).mock.calls[1];
    const payload = JSON.parse((specCall[1] as RequestInit).body as string);
    expect(payload.templateMapping).toEqual({ opening: 0, agenda: 1, content: 1, closing: 1 });
  });

  it("supports language switching", async () => {
    render(<App />);

    await userEvent.click(screen.getByRole("button", { name: "中文" }));

    expect(screen.getByText("主题、模板预览与版式映射")).toBeInTheDocument();
    expect(screen.getByText("模板预览")).toBeInTheDocument();
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

  it("uploads a pptx template during render", async () => {
    vi.mocked(fetch)
      .mockResolvedValueOnce(mockJsonResponse(previewResponse))
      .mockResolvedValueOnce(mockJsonResponse(specResponse))
      .mockResolvedValueOnce(mockBinaryResponse());
    render(<App />);

    const file = new File(["pptx-template"], "brand-template.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    await userEvent.upload(screen.getByLabelText("Import PPTX Template"), file);
    await screen.findByText("Cover");

    await userEvent.selectOptions(screen.getByLabelText("Opening"), "0");
    await userEvent.selectOptions(screen.getByLabelText("Agenda"), "1");
    await userEvent.selectOptions(screen.getByLabelText("Content"), "1");
    await userEvent.selectOptions(screen.getByLabelText("Closing"), "1");
    await userEvent.type(screen.getByLabelText("Prompt"), "Create a team update deck");
    await userEvent.click(screen.getByRole("button", { name: "Generate Spec" }));
    await screen.findByText("Roadmap");
    await userEvent.click(screen.getByRole("button", { name: "Confirm and Generate PPT" }));

    expect(fetch).toHaveBeenCalledTimes(3);
    const renderCall = vi.mocked(fetch).mock.calls[2];
    const requestInit = renderCall[1] as RequestInit;
    expect(requestInit.body).toBeInstanceOf(FormData);
  });
});
