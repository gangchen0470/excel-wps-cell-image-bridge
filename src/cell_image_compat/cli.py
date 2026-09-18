from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from .core import convert_floating_to_native, convert_native, inspect_workbook, inspection_json, make_compatible_copy


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="cell-image-compat", description="Inspect and convert Excel/WPS cell images")
    sub = parser.add_subparsers(dest="command", required=True)
    inspect_cmd = sub.add_parser("inspect", help="List recognized cell images")
    inspect_cmd.add_argument("workbook", type=Path)
    convert = sub.add_parser("compatible", help="Create a standard floating-image XLSX copy")
    convert.add_argument("workbook", type=Path)
    convert.add_argument("output", type=Path)
    convert.add_argument("--margin", type=int, default=2, help="cell margin in pixels (default: 2)")
    convert.add_argument("--sheet", help="only convert images on this worksheet")
    convert.add_argument("--range", dest="cell_range", help="only convert this A1 range, for example B2:B500")
    convert.add_argument("--log", type=Path, help="write a JSON conversion report")
    native = sub.add_parser("native", help="Convert all recognized cell images to a native target format")
    native.add_argument("workbook", type=Path)
    native.add_argument("output", type=Path)
    native.add_argument("--target", choices=("excel", "wps"), required=True)
    floating = sub.add_parser("floating-native", help="Convert selected floating pictures to native cell images")
    floating.add_argument("workbook", type=Path)
    floating.add_argument("output", type=Path)
    floating.add_argument("--target", choices=("excel", "wps"), required=True)
    floating.add_argument("--sheet", required=True, help="worksheet containing the selected cells")
    floating.add_argument("--range", dest="cell_range", required=True, help="selected A1 cell range")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "inspect":
            print(inspection_json(args.workbook))
        elif args.command == "compatible":
            scanned = len(inspect_workbook(args.workbook))
            images = make_compatible_copy(args.workbook, args.output, args.margin, args.sheet, args.cell_range)
            report = {
                "source": str(args.workbook), "output": str(args.output), "scope": {
                    "sheet": args.sheet or "*", "range": args.cell_range or "*",
                }, "scanned": scanned, "success": len(images), "failed": 0, "skipped": scanned - len(images),
                "completed_at": datetime.now(timezone.utc).isoformat(),
            }
            if args.log:
                args.log.parent.mkdir(parents=True, exist_ok=True)
                args.log.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            print(json.dumps(report, ensure_ascii=False))
        elif args.command == "native":
            images = convert_native(args.workbook, args.output, args.target)
            print(json.dumps({"output": str(args.output), "converted": len(images), "target": args.target}, ensure_ascii=False))
        else:
            images = convert_floating_to_native(args.workbook, args.output, args.target, args.sheet, args.cell_range)
            print(json.dumps({"output": str(args.output), "converted": len(images), "target": args.target, "sheet": args.sheet, "range": args.cell_range}, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
