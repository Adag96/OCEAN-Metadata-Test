#!/usr/bin/env python3
"""
Aggregate metadata survey results from multiple JSON files.
Reads all JSON files from 'Responses (JSON)' folder and outputs an Excel file
with average votes for each preset and metadata tag.
"""

import json
from pathlib import Path
from collections import defaultdict
from statistics import mean

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# Configuration
RESPONSES_FOLDER = "Responses (JSON)"
OUTPUT_FILE = "aggregated_results.xlsx"

def load_all_responses(folder_path):
    """Load all JSON response files from the specified folder."""
    responses = []
    folder = Path(folder_path)

    if not folder.exists():
        raise FileNotFoundError(f"Folder '{folder_path}' not found")

    json_files = list(folder.glob("*.json"))

    if not json_files:
        raise FileNotFoundError(f"No JSON files found in '{folder_path}'")

    print(f"Found {len(json_files)} JSON file(s)")

    for json_file in json_files:
        with open(json_file, 'r') as f:
            data = json.load(f)
            responses.append(data)
            print(f"  - Loaded {json_file.name}")

    return responses

def aggregate_votes(all_responses):
    """
    Aggregate votes across all respondents.
    Returns a dictionary: {preset_name: {tag: [list of votes]}}
    """
    aggregated = defaultdict(lambda: defaultdict(list))

    for response_data in all_responses:
        for preset_response in response_data['responses']:
            preset_name = preset_response['presetName']

            for vote_key, vote_value in preset_response['votes'].items():
                aggregated[preset_name][vote_key].append(vote_value)

    return aggregated

def calculate_averages(aggregated_data):
    """
    Calculate average votes for each preset and tag.
    Returns a dictionary: {preset_name: {tag: average_vote}}
    """
    averages = {}

    for preset_name, tags in aggregated_data.items():
        averages[preset_name] = {}
        for tag, votes in tags.items():
            averages[preset_name][tag] = round(mean(votes), 2)

    return averages

def write_excel(averages, output_file):
    """Write aggregated results to Excel file with custom columns per preset and bold headers."""

    if not EXCEL_AVAILABLE:
        raise ImportError("openpyxl is required for Excel output. Install with: pip install openpyxl")

    # Create workbook and get active sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Aggregated Results"

    # Sort preset names
    preset_names = sorted(averages.keys())

    # Bold font for headers
    bold_font = Font(bold=True)

    # Current row tracker
    current_row = 1

    # Process each preset
    for preset_name in preset_names:
        preset_data = averages[preset_name]

        # Get sorted tags for this preset
        tags = sorted(preset_data.keys())

        # Write header row for this preset
        header = ['Preset Name'] + tags
        for col_idx, header_value in enumerate(header, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=header_value)
            cell.font = bold_font

        current_row += 1

        # Write data row for this preset
        row = [preset_name] + [preset_data[tag] for tag in tags]
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=current_row, column=col_idx, value=value)

        current_row += 1

        # Write blank spacer row
        current_row += 1

    # Save workbook
    wb.save(output_file)

    print(f"\n✓ Results written to '{output_file}'")
    print(f"  - {len(preset_names)} presets")
    print(f"  - Each preset with its own custom columns")
    print(f"  - Headers are bold for easy reading")

def main():
    """Main execution function."""
    try:
        print("=" * 60)
        print("Metadata Survey Results Aggregator")
        print("=" * 60)
        print()

        # Load all JSON responses
        all_responses = load_all_responses(RESPONSES_FOLDER)

        # Aggregate votes
        print("\nAggregating votes...")
        aggregated = aggregate_votes(all_responses)

        # Calculate averages
        print("Calculating averages...")
        averages = calculate_averages(aggregated)

        # Write to Excel
        print("Writing Excel file...")
        write_excel(averages, OUTPUT_FILE)

        print("\n" + "=" * 60)
        print("✓ Aggregation complete!")
        print("=" * 60)

    except Exception as e:
        print(f"\n✗ Error: {e}")
        return 1

    return 0

if __name__ == "__main__":
    exit(main())
