from pathlib import Path
import argparse
import json

from .engine import WorkbookDrivenRetailROIModel, outputs_to_jsonable


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the retail ROI model from the Excel workbook.")
    parser.add_argument("workbook", help="Path to the .xlsm workbook")
    parser.add_argument("--out", help="Path to save JSON output", default="roi_output.json")
    args = parser.parse_args()

    model = WorkbookDrivenRetailROIModel(args.workbook)
    outputs = model.run()
    payload = outputs_to_jsonable(outputs)

    Path(args.out).write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Saved output to: {args.out}")


if __name__ == "__main__":
    main()
