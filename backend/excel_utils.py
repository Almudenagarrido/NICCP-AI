import os
import re
import json
from typing import List
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from copy import deepcopy

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
    ("Financial", "G17", "Electricity", "D19"),

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

OPS_MAP = {
    "subtract": lambda x, y: x - y,
    "multiply": lambda x, y: x * y,
    "multiply_per": lambda x, y: (x * y)/100,
    "gt": lambda x, y: x > y,
    "pos_or_cero": lambda x: max(x, 0),
    "if": lambda condition, true_val, false_val: true_val if condition else false_val,
}


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

def expand_formulas_by_models(formulas_json: dict, models: list) -> dict:
    expanded_json = {}

    for file_path, sheets in formulas_json.items():
        expanded_json[file_path] = {}

        for sheet_name, formulas in sheets.items():
            expanded_formulas = []
            
            for formula in formulas:
                if "{model}" not in formula.get("targets", "") and all("{model}" not in sl for sl in formula.get("source_labels", [])):
                    expanded_formulas.append(formula)
                    continue

                for model in models:
                    new_formula = deepcopy(formula)
                    new_formula["targets"] = new_formula["targets"].replace("{model}", model)
                    new_source_labels = [
                        sl.replace("{model}", model) for sl in new_formula["source_labels"]
                    ]
                    new_formula["source_labels"] = new_source_labels

                    expanded_formulas.append(new_formula)

            expanded_json[file_path][sheet_name] = expanded_formulas

    return expanded_json

def parse_label_reference(default_file: str, default_sheet: str, label: str):
    parts = label.split("::")
    if len(parts) == 3:
        file_path, sheet, label = parts
    elif len(parts) == 2:
        file_path = default_file
        sheet, label = parts
    else:
        file_path = default_file
        sheet = default_sheet
    return os.path.normpath(file_path), sheet, label

def get_numeric_cell_refs_label(wbs: dict, default_file: str, default_sheet: str, full_label: str, label_col: int = 1):
    file_path, sheet_name, row_label = parse_label_reference(default_file, default_sheet, full_label)

    if file_path not in wbs:
        wbs[file_path] = load_workbook(file_path, data_only=True)
    ws = wbs[file_path][sheet_name]
    
    target_row = None
    for row in range(1, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=label_col).value
        if cell_value and row_label in str(cell_value):
            target_row = row
            break

    if target_row is None:
        raise ValueError(f"Label '{row_label}'not found in column {label_col}")

    refs = []
    for col in range(3, ws.max_column + 1):
        val = ws.cell(row=target_row, column=col).value
        if isinstance(val, (int, float)):
            refs.append(ws.cell(row=target_row, column=col).coordinate)
        else:
            if isinstance(val, str):
                try:
                    float(val.replace(",", "."))
                    refs.append(ws.cell(row=target_row, column=col).coordinate)
                except:
                    pass
    return refs

def get_val(operand, values, results):
    if isinstance(operand, int):
        return values[operand]
    elif isinstance(operand, str):
        if operand == "temp":
            temp_keys = [k for k in results.keys() if "temp" in k]
            if temp_keys:
                last_temp_key = temp_keys[-1]
                return results[last_temp_key]
            else:
                return 0
        else:
            return results.get(operand, 0)
    else:
        return 0
    
def get_cell_value_from_ref(wbs, ref, default_file, default_sheet):
        ref_parts = ref.split("::")
        if len(ref_parts) == 2:
            file_path = ref_parts[0]
            sheet_name, cell_coord = ref_parts[1].split("!")
        else:
            file_path = default_file
            sheet_name = default_sheet
            cell_coord = ref

        if file_path not in wbs:
            wbs[file_path] = load_workbook(file_path, data_only=True)

        return wbs[file_path][sheet_name][cell_coord]

def apply_formulas(file_path, formulas_json_path, models):

    with open(formulas_json_path) as f:
        formulas_json = json.load(f)
        formulas_json = expand_formulas_by_models(formulas_json, models)

        file_path_norm = os.path.normpath(file_path)
        json_keys_path = [os.path.normpath(k) for k in formulas_json.keys()]

        if file_path_norm not in json_keys_path:
            return

        json_index = json_keys_path.index(file_path_norm)
        formulas_for_file = list(formulas_json.values())[json_index]

        wbs = {}
        wbs[file_path_norm] = load_workbook(file_path_norm)

        for sheet_name, formulas in formulas_for_file.items():
            if sheet_name not in wbs[file_path_norm].sheetnames:
                continue
            ws = wbs[file_path_norm][sheet_name]

            for formula in formulas:
                targets_label = formula["targets"]
                source_labels = formula["source_labels"]
                formula_steps = formula.get("formula_steps", [])

                targets_cells = get_numeric_cell_refs_label(wbs, file_path_norm, sheet_name, targets_label, 1)
                source_cells_list = []
                for source_label in source_labels:
                    parts = source_label.split("::")
                    
                    if len(parts) == 3:
                        src_path, src_sheet, src_label = parts
                    else:
                        src_path, src_sheet = file_path_norm, sheet_name
                        src_label = source_label

                    src_path_norm = os.path.normpath(src_path)
                    if src_path_norm not in wbs:
                        wbs[src_path_norm] = load_workbook(src_path_norm)                 

                    refs = get_numeric_cell_refs_label(wbs, src_path_norm, src_sheet, src_label, 1)
                    source_cells_list.append((src_path_norm, src_sheet, refs))

                    length = len(targets_cells)
                    if any(len(refs) != length and len(refs) != 1 for (_, _, refs) in source_cells_list):
                        continue

                for i in range(length):
                    values = []
                    for src_path_norm, src_sheet, src_list in source_cells_list:
                        ref = src_list[i] if len(src_list) > 1 else src_list[0]
                        cell = get_cell_value_from_ref(wbs, ref, src_path_norm, src_sheet)
                        val = cell.value
                        if isinstance(val, str):
                            try:
                                val = float(val.replace(",", "."))
                            except:
                                val = 0
                        values.append(val if val is not None else 0)

                    results = {}
                    for step in formula_steps:
                        op = step["op"]
                        operands = step["operands"]
                        if op == "range":
                            res = i + 1
                        elif op == "offset":
                            source_index = operands[0]
                            offset = get_val(operands[1], values, results)

                            if 0 <= source_index < len(source_cells_list):
                                src_path_norm, src_sheet, src_list = source_cells_list[source_index]
                                j = i - offset
                                if 0 <= j < len(src_list):
                                    ref = src_list[j]
                                    cell = get_cell_value_from_ref(wbs, ref, src_path_norm, src_sheet)
                                    val = cell.value
                                    if isinstance(val, str):
                                        try:
                                            val = float(val.replace(",", "."))
                                        except:
                                            val = 0
                                    res = val if val is not None else 0
                                else:
                                    res = 0
                            else:
                                res = 0
                        else:
                            ops_args = [get_val(operand, values, results) for operand in operands]
                            if op in OPS_MAP:
                                res = OPS_MAP[op](*ops_args)
                            else:
                                print(f"Operation {op} not supported by platform")
                                res = 0
    
                        results[step["result"]] = res

                    target_ref = targets_cells[i]
                    ws[target_ref].value = results.get("final", 0)

        wbs[file_path_norm].save(file_path_norm)