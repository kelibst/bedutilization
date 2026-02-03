"""
Bed Utilization Workbook Generator
Ghana Health Service - Hohoe Municipal Hospital

Usage:
    python build_workbook.py [--year 2026] [--carry-forward path/to/carry_forward.json]

Prerequisites:
    - Python 3.x with openpyxl and pywin32
    - Excel 2016 or later
    - "Trust access to the VBA project object model" enabled in Excel Trust Center
"""
import argparse
import os
import sys
from datetime import datetime
from config import WorkbookConfig
from phase1_structure import build_structure
from phase2_vba import inject_vba


def main():
    parser = argparse.ArgumentParser(
        description="Generate Bed Utilization Workbook for Ghana Health Service"
    )
    parser.add_argument(
        "--year", type=int, default=datetime.now().year,
        help="Year for the workbook (default: current year)"
    )
    parser.add_argument(
        "--carry-forward", type=str, default=None,
        help="Path to previous year carry-forward JSON file"
    )
    parser.add_argument(
        "--output-dir", type=str, default=".",
        help="Output directory for the workbook (default: current directory)"
    )
    parser.add_argument(
        "--skip-vba", action="store_true",
        help="Skip VBA injection (produces .xlsx without macros)"
    )
    args = parser.parse_args()

    print(f"=" * 60)
    print(f"  Bed Utilization Workbook Generator")
    print(f"  Ghana Health Service - Hohoe Municipal Hospital")
    print(f"  Year: {args.year}")
    print(f"=" * 60)

    # Create config
    config = WorkbookConfig(year=args.year, carry_forward_path=args.carry_forward)

    if args.carry_forward:
        print(f"\nCarry-forward data loaded from: {args.carry_forward}")
        for ward in config.WARDS:
            if ward.prev_year_remaining > 0:
                print(f"  {ward.name}: {ward.prev_year_remaining} patients")

    # Ensure output directory exists
    output_dir = os.path.abspath(args.output_dir)
    os.makedirs(output_dir, exist_ok=True)

    xlsx_path = os.path.join(output_dir, f"Bed_Utilization_{args.year}.xlsx")
    xlsm_path = os.path.join(output_dir, f"Bed_Utilization_{args.year}.xlsm")

    # Phase 1: Build structure with openpyxl
    print(f"\n--- Phase 1: Building workbook structure ---")
    build_structure(config, xlsx_path)

    if args.skip_vba:
        print(f"\nDone (VBA skipped). Open {xlsx_path} in Excel.")
        return

    # Phase 2: Inject VBA with win32com
    print(f"\n--- Phase 2: Injecting VBA macros ---")
    try:
        inject_vba(xlsx_path, xlsm_path, config)
    except Exception as e:
        print(f"\nVBA injection failed: {e}")
        print(f"\nThe .xlsx file was still created at: {xlsx_path}")
        print("You can open it in Excel and add VBA manually if needed.")
        sys.exit(1)

    # Clean up intermediate xlsx
    try:
        os.remove(xlsx_path)
    except:
        pass

    print(f"\n{'=' * 60}")
    print(f"  SUCCESS! Workbook generated:")
    print(f"  {xlsm_path}")
    print(f"{'=' * 60}")
    print(f"\nNext steps:")
    print(f"  1. Open {os.path.basename(xlsm_path)} in Excel")
    print(f"  2. Enable macros when prompted")
    print(f"  3. Use the buttons on the Control sheet to enter data")
    print(f"\nAt year end:")
    print(f"  1. Click 'Export Year-End' to create carry_forward_{args.year}.json")
    print(f"  2. Run: python build_workbook.py --year {args.year + 1} --carry-forward carry_forward_{args.year}.json")


if __name__ == "__main__":
    main()
