from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pydantic import ValidationError

from office_agent.config import AppConfig
from office_agent.errors import OfficeAgentError
from office_agent.schema import ThemeSpec
from office_agent.services import generate_spec_from_prompt, render_presentation


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="office-agent", description="Generate PowerPoint files from prompts.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    generate_parser = subparsers.add_parser("generate", help="Generate a .pptx from a prompt.")
    generate_parser.add_argument("--prompt", required=True, help="Natural language prompt for the presentation.")
    generate_parser.add_argument("--out", required=True, help="Output .pptx path.")
    generate_parser.add_argument("--provider", choices=["openai", "ollama"], help="LLM provider.")
    generate_parser.add_argument("--model", help="Model name to call.")
    generate_parser.add_argument(
        "--theme",
        default="default",
        choices=["default", "executive", "editorial"],
        help="Theme preset.",
    )
    generate_parser.add_argument(
        "--debug-spec",
        action="store_true",
        help="Write the validated presentation spec next to the output .pptx.",
    )
    return parser


def write_debug_spec(output_path: Path, spec) -> Path:
    debug_path = output_path.with_suffix(".spec.json")
    debug_path.write_text(spec.model_dump_json(indent=2), encoding="utf-8")
    return debug_path


def run_generate(args: argparse.Namespace) -> int:
    config = AppConfig.from_env().with_overrides(
        provider=args.provider,
        model=args.model,
        theme=ThemeSpec(preset=args.theme),
    )
    config.validate()
    spec = generate_spec_from_prompt(args.prompt, config)
    result = render_presentation(spec, args.out)

    if args.debug_spec:
        debug_path = write_debug_spec(Path(args.out), spec)
        print(f"Spec written to {debug_path}")
    for warning in result.warnings:
        print(f"Warning: {warning}", file=sys.stderr)
    print(f"Presentation written to {result.path}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "generate":
            return run_generate(args)
    except (OfficeAgentError, ValidationError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0
