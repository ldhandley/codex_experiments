#!/usr/bin/env python3
"""Create an Excel tuition table for Olympic College rates."""

import argparse
import sys


# Scraped once from https://www.olympic.edu/fund-your-education/tuition-fees
# and intentionally hard-coded to avoid live scraping at runtime.
TUITION_RATES = {
    "resident": {
        "lower_first_10": 123.94,
        "lower_after_10": 57.78,
        "upper_first_10": 133.54,
        "upper_after_10": 67.38,
    },
    "non-resident": {
        "lower_first_10": 274.84,
        "lower_after_10": 208.68,
        "upper_first_10": 284.44,
        "upper_after_10": 218.28,
    },
}


def total_cost(credits, first_10_rate, after_10_rate):
    """Return total tuition for a credit count with tiered pricing."""
    if credits <= 10:
        return credits * first_10_rate
    return (10 * first_10_rate) + ((credits - 10) * after_10_rate)


def build_rows(max_credits, rates):
    """Build rows for the spreadsheet."""
    rows = []
    for credits in range(1, max_credits + 1):
        lower_total = total_cost(credits, rates["lower_first_10"], rates["lower_after_10"])
        upper_total = total_cost(credits, rates["upper_first_10"], rates["upper_after_10"])
        rows.append([credits, lower_total, upper_total])
    return rows


def write_xlsx(output_filename, rows):
    """Write the tuition table to an .xlsx file using openpyxl."""
    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise RuntimeError(
            "openpyxl is required to create .xlsx files. Install it with: pip install openpyxl"
        ) from exc

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Tuition"

    worksheet.append([
        "# of Credits",
        "Tuition Cost for Lower-Division",
        "Tuition Cost for Upper-Division",
    ])

    for row in rows:
        worksheet.append(row)

    money_format = "$#,##0.00"
    for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=3):
        for cell in row:
            cell.number_format = money_format

    workbook.save(output_filename)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Create an Excel tuition table for lower/upper division by residency."
    )
    parser.add_argument(
        "residency",
        choices=["resident", "non-resident"],
        help='Residency status: "resident" or "non-resident"',
    )
    parser.add_argument(
        "max_credits",
        type=int,
        help="Maximum number of credits to include (must be >= 1)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Optional output filename (.xlsx). Default: tuition_<residency>_1_to_<N>.xlsx",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if args.max_credits < 1:
        print("Error: max_credits must be at least 1.", file=sys.stderr)
        return 2

    rates = TUITION_RATES[args.residency]
    rows = build_rows(args.max_credits, rates)

    if args.output:
        output_filename = args.output
    else:
        output_filename = f"tuition_{args.residency}_1_to_{args.max_credits}.xlsx"

    try:
        write_xlsx(output_filename, rows)
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Created {output_filename}")
    print(
        "Rates used: "
        f"lower 1-10=${rates['lower_first_10']:.2f}, "
        f"lower 11+=${rates['lower_after_10']:.2f}, "
        f"upper 1-10=${rates['upper_first_10']:.2f}, "
        f"upper 11+=${rates['upper_after_10']:.2f}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
