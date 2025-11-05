#!/usr/bin/env python3
"""
Enhanced Excel to Markdown Converter Test Script

This script demonstrates the excel-mcp-server's ability to convert Excel files
to Markdown format, handling merged cells and providing Excel-style formatting.
"""

import os
import sys
from pathlib import Path
from excel_mcp.workbook import get_workbook_info
from excel_mcp.data import get_sheet_as_markdown_table
from excel_mcp.sheet import get_merged_ranges


def print_separator(title: str = ""):
    """Print a formatted separator line"""
    if title:
        print(f"\n{'='*20} {title} {'='*20}")
    else:
        print("=" * 60)


def test_excel_to_markdown_conversion():
    """
    Enhanced test function that converts Excel to Markdown with detailed analysis
    """
    print_separator("Excel to Markdown Converter Test")

    # Define file paths
    workspace_root = Path(__file__).parent.absolute()
    excel_file_path = workspace_root / "HUA_VVP_RA_INT_05-2025_072339.xlsx"

    # Check if file exists
    if not excel_file_path.exists():
        print(f"❌ Error: Excel file not found at {excel_file_path}")
        return False

    print(f"📁 Reading Excel file: {excel_file_path}")
    print(f"📊 File size: {excel_file_path.stat().st_size:,} bytes")

    try:
        # Get workbook information
        print_separator("Workbook Analysis")
        workbook_info = get_workbook_info(str(excel_file_path), include_ranges=True)

        print(f"📋 Workbook: {workbook_info['filename']}")
        print(f"📄 Total sheets: {len(workbook_info['sheets'])}")
        print(f"📏 File size: {workbook_info['size']:,} bytes")

        # List all sheets
        for i, sheet_name in enumerate(workbook_info["sheets"], 1):
            used_range = workbook_info.get("used_ranges", {}).get(sheet_name, "N/A")
            print(f"  {i}. '{sheet_name}' - Used range: {used_range}")

        # Process each sheet
        for sheet_name in workbook_info["sheets"]:
            print_separator(f"Processing Sheet: {sheet_name}")

            # Check for merged cells
            try:
                merged_ranges = get_merged_ranges(str(excel_file_path), sheet_name)
                if merged_ranges:
                    print(f"🔗 Found {len(merged_ranges)} merged cell ranges:")
                    for merge_range in merged_ranges[:5]:  # Show first 5
                        print(f"  - {merge_range}")
                    if len(merged_ranges) > 5:
                        print(f"  ... and {len(merged_ranges) - 5} more")
                else:
                    print("📝 No merged cells found")
            except Exception as e:
                print(f"⚠️  Could not check merged cells: {e}")

            # Generate markdown table
            print("🔄 Converting to Markdown...")
            markdown_table = get_sheet_as_markdown_table(str(excel_file_path), sheet_name)

            # Save to file
            output_file = workspace_root / f"output_{sheet_name.replace(' ', '_')}.md"
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(f"# Excel Sheet: {sheet_name}\n\n")
                f.write(f"**Source File:** {excel_file_path.name}\n")
                f.write(f"**Generated:** {Path(__file__).name}\n\n")
                f.write("## Table Content\n\n")
                f.write(markdown_table)

            print(f"✅ Markdown saved to: {output_file}")

            # Show preview of first few lines
            lines = markdown_table.split("\n")
            print(f"📖 Preview (first 5 lines):")
            for line in lines[:5]:
                print(f"  {line}")
            if len(lines) > 5:
                print(f"  ... ({len(lines) - 5} more lines)")

        print_separator("Test Completed Successfully")
        print("🎉 All sheets have been successfully converted to Markdown format!")
        return True

    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        import traceback

        traceback.print_exc()
        return False


def main():
    """Main function"""
    print("🚀 Starting Enhanced Excel to Markdown Converter Test")

    success = test_excel_to_markdown_conversion()

    if success:
        print("\n✨ Test completed successfully!")
        sys.exit(0)
    else:
        print("\n💥 Test failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
