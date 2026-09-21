import json
import textwrap
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment


def _count_wrapped_lines(value, characters_per_line):
    """Estimate how many lines are required to display the text."""
    text = "" if value is None else str(value)
    source_lines = text.splitlines() or [""]

    return sum(
        max(
            1,
            len(
                textwrap.wrap(
                    line,
                    width=characters_per_line,
                    break_long_words=True
                )
            )
        )
        for line in source_lines
    )


def _set_merged_autofit(
    sheet,
    cell,
    value,
    base_row_heights,
    characters_per_line
):
    """Write wrapped text and increase the row height when required."""
    cell.value = value

    alignment = copy(cell.alignment)
    alignment.wrap_text = True
    alignment.vertical = "center"
    cell.alignment = alignment

    if cell.row not in base_row_heights:
        base_row_heights[cell.row] = (
            sheet.row_dimensions[cell.row].height
            or sheet.sheet_format.defaultRowHeight
            or 15
        )

    base_height = base_row_heights[cell.row]
    required_lines = _count_wrapped_lines(
        value,
        characters_per_line
    )
    required_height = base_height * required_lines

    current_height = (
        sheet.row_dimensions[cell.row].height
        or base_height
    )

    sheet.row_dimensions[cell.row].height = max(
        current_height,
        required_height
    )

def generate_single_psw(form_data):
    """
    Map form data to the PSW template and save the resulting Excel document.
    """
    # Project root directory
    project_root = Path(__file__).resolve().parent.parent

    # Parent directory used for generated documents
    output_root = project_root.parent

    config_path = project_root / "config" / "vda_config.json"
    template_path = project_root / "templates" / "vda_2020.xlsx"

    # Get the part number and report version
    part_number = str(form_data.get("PartNumber", "Missing_PN")).strip()
    report_version = str(form_data.get("ReportVersion", "00")).strip()

    # Output directory: vda/<part number>/<report version>
    output_dir = output_root / "vda" / part_number / report_version

    # Create the output directory if it does not exist
    output_dir.mkdir(parents=True, exist_ok=True)

    # Default output filename
    output_path = output_dir / f"PSW_{part_number}.xlsx"

    # Append the current date and a counter to prevent overwriting an existing file
    file_counter = 1

    current_date = datetime.now().strftime("%Y-%m-%d")

    while output_path.exists():
        new_filename = f"PSW_{part_number}_{current_date}_{file_counter}.xlsx"
        output_path = output_dir / new_filename
        file_counter += 1

    # Load the field mapping and Excel template
    with open(config_path, 'r', encoding='utf-8') as config_file:
        field_mapping = json.load(config_file)

    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.active

    base_row_heights = {}

    # Populate cells according to the field mapping
    for field_name, value in form_data.items():
        if field_name in field_mapping["Fields"]:
            field_config = field_mapping["Fields"][field_name]
            cell = sheet[field_config["cell"]]
            field_type = field_config.get("type", "standard")

            if field_type in ["checkbox", "merged_checkbox"]:
                if value is True:
                    cell.value = "X"
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            elif field_type == "merged_multiline":
                cell.value = value
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            elif field_type == "merged_autofit":
                _set_merged_autofit(sheet, cell, value, base_row_heights, field_config.get("characters_per_line", 24))
            else:
                cell.value = value

    workbook.save(output_path)
    workbook.close()
    return output_path
