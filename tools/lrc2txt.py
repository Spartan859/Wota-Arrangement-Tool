from __future__ import annotations

import argparse
import re
from pathlib import Path

TIMESTAMP_PATTERN = re.compile(r"\[\d{1,2}:\d{1,2}(?:[.:]\d{1,3})?\]")
META_TAG_PATTERN = re.compile(r"^\[[a-zA-Z]+:[^\]]*\]$")


def strip_timestamps(text: str, strip_tags: bool = False) -> str:
    """Remove LRC timestamps from text and return plain lyric lines."""
    output_lines: list[str] = []
    for raw_line in text.splitlines():
        line = TIMESTAMP_PATTERN.sub("", raw_line).strip()
        if not line:
            continue
        if strip_tags and META_TAG_PATTERN.fullmatch(line):
            continue
        output_lines.append(line)
    return "\n".join(output_lines)


def convert_lrc_to_txt(input_path: Path, output_path: Path | None = None, strip_tags: bool = False) -> Path:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if not input_path.is_file():
        raise ValueError(f"Input path is not a file: {input_path}")

    target_path = output_path or input_path.with_suffix(".txt")
    text = input_path.read_text(encoding="utf-8")
    cleaned_text = strip_timestamps(text, strip_tags=strip_tags)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_text(cleaned_text + ("\n" if cleaned_text else ""), encoding="utf-8")
    return target_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert .lrc lyrics into plain .txt by removing timestamps.")
    parser.add_argument("input", help="Input .lrc file path")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output .txt path (default: same name as input, with .txt extension)",
    )
    parser.add_argument(
        "--strip-tags",
        action="store_true",
        help="Remove metadata tag lines like [ti:], [ar:], [al:], [by:]",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None

    try:
        target = convert_lrc_to_txt(input_path=input_path, output_path=output_path, strip_tags=args.strip_tags)
    except (FileNotFoundError, ValueError, OSError) as exc:
        parser.exit(status=1, message=f"Error: {exc}\n")

    print(f"Converted: {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
