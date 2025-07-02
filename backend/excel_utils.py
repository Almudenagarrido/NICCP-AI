from openpyxl import load_workbook
from openpyxl.utils import range_boundaries


CELLS_MAPPING_MODEL_TO_INPUTS = [
    ("Res-Cash", "G16:R16", "Electricity", "D2:O2"),

    ("Res-Cash", "G19:R19", "Electricity", "D4:O4"),
    ("Res-Cash", "G17:R17", "Electricity", "D5:O5"),
    ("Res-Cash", "G20:R20", "Electricity", "D6:O6"),
    ("Res-Cash", "G18:R18", "Electricity", "D7:O7"),
    ("Financial", "G13", "Electricity", "D8"),

    ("Res-Cash", "G23:R23", "Electricity", "D10:O10"),
    ("Res-Cash", "G21:R21", "Electricity", "D11:O11"),
    ("Res-Cash", "G24:R24", "Electricity", "D12:O12"),
    ("Res-Cash", "G22:R22", "Electricity", "D13:O13"),
    ("Financial", "G17", "Electricity", "D14"),

    ("Res-Cash", "G28:R28", "Electricity", "D16:O16"),
    ("Res-Cash", "G29:R29", "Electricity", "D17:O17"),
    ("Res-Cash", "G30:R30", "Electricity", "D18:O18"),
    ("Res-Cash", "G17", "Electricity", "D19"),

    ("Res-Cash", "G34:R34", "LPG", "D2:O2"),

    ("Res-Cash", "G35:R35", "LPG", "D4:O4"),
    ("Res-Cash", "G37:R37", "LPG", "D5:O5"),
    ("Res-Cash", "G36:R36", "LPG", "D6:O6"),
    ("Res-Cash", "G38:R38", "LPG", "D7:O7"),
    ("Financial", "G27", "LPG", "D8"),
    ("Financial", "G29", "LPG", "D9"),

    ("Res-Cash", "G40:R40", "LPG", "D11:O11"),
    ("Res-Cash", "G42:R42", "LPG", "D12:O12"),
    ("Res-Cash", "G41:R41", "LPG", "D13:O13"),
    ("Res-Cash", "G43:R43", "LPG", "D14:O14"),
    ("Res-Cash", "G39:R39", "LPG", "D15:O15"),
    ("Financial", "G32", "LPG", "D16"),

    ("Res-Cash", "G47:R58", "Rest of subsidies or taxes", "D2:O13"),
]

CELLS_MAPPING_MODEL_TO_CARBON = [
    ("Res-Cash", "G10:R10", "Carbon Credits", "C2:N2"),
]

def fill_contents_from_source(source_path: str, destiny_path: str, model_name:str, carbon_flag: bool):
    try:
        src_wb = load_workbook(source_path, data_only=True, read_only=True)
        dst_wb = load_workbook(destiny_path)

        cell_mapping = CELLS_MAPPING_MODEL_TO_INPUTS if not carbon_flag else CELLS_MAPPING_MODEL_TO_CARBON
        for src_sheet, src_range, dst_sheet, dst_range in cell_mapping:
            src_ws = src_wb[src_sheet]
            dst_ws = dst_wb[dst_sheet]

            src_min_col, src_min_row, src_max_col, src_max_row = range_boundaries(src_range)
            dst_min_col, dst_min_row, dst_max_col, dst_max_row = range_boundaries(dst_range)

            if carbon_flag:
                model_row = None
                for row in range(1, dst_ws.max_row + 1):
                    cell_value = dst_ws.cell(row=row, column=1).value
                    if isinstance(cell_value, str) and model_name.lower() in cell_value.strip().lower():
                        model_row = row
                        break
                if model_row is None:
                    raise ValueError(f"Model '{model_name}' not found in destination sheet '{dst_sheet}' column 1.")

                dst_min_row = model_row
                dst_max_row = model_row + (src_max_row - src_min_row)

            src_rows = src_max_row - src_min_row + 1
            src_cols = src_max_col - src_min_col + 1
            dst_rows = dst_max_row - dst_min_row + 1
            dst_cols = dst_max_col - dst_min_col + 1

            if src_rows != dst_rows or src_cols != dst_cols:
                raise ValueError(
                    f"Size mismatch between source {src_range} and destination {dst_range}."
                )

            for i in range(src_rows):
                for j in range(src_cols):
                    value = src_ws.cell(row=src_min_row + i, column=src_min_col + j).value
                    dst_ws.cell(row=dst_min_row + i, column=dst_min_col + j, value=value)

        dst_wb.save(destiny_path)

    except Exception as e:
        raise RuntimeError(f"Error actualizando plantilla con datos del modelo: {str(e)}")