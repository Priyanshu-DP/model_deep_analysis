from worker import process_function
from utils import single_excel_file
import pandas as pd
import json
import copy
import os
import shutil

# Load master JSON
with open("config.json", "r") as f:
    master_config = json.load(f)

common = master_config["common"]
file_path = common["rroi_excel_file_path"]
final_file_name = common["output_file_name"]
sheet_names = common["sheet_name"]


# === STEP 1: Create previous & current month files ===
previous_month_file_path = "previous_month_input_file_path.xlsx"
current_month_file_path = "current_month_input_file_path.xlsx"
combined_input_file_path = "combined_input_ROI_file.xlsx"
previous_month_sheet_name = "previous_month_sheet_name"
current_month_sheet_name = "current_month_sheet_name"
combined_file_sheet_name = "final_combined_roi"

def prepare_data(sheet_name, output_path, output_sheet):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df["Model"] = sheet_name
    df = df[["Model"] + [col for col in df.columns if col != "Model"]]
    df.to_excel(output_path, sheet_name=output_sheet, index=False)
    return df

data_1 = prepare_data(sheet_names[0], previous_month_file_path, previous_month_sheet_name)
data_2 = prepare_data(sheet_names[1], current_month_file_path, current_month_sheet_name)


# === STEP 2: Combine sheets if headers match ===
if set(data_1.columns) == set(data_2.columns):
    final_data = pd.concat([data_1[data_1.columns], data_2[data_1.columns]], ignore_index=True)
    final_data.to_excel(combined_input_file_path, sheet_name=combined_file_sheet_name, index=False)
else:
    # Find mismatches
    missing_in_data2 = set(data_1.columns) - set(data_2.columns)
    missing_in_data1 = set(data_2.columns) - set(data_1.columns)

    error_message = "Header mismatch detected!\n"
    if missing_in_data2:
        error_message += f"Columns in data_1 but not in data_2: {missing_in_data2}\n"
    if missing_in_data1:
        error_message += f"Columns in data_2 but not in data_1: {missing_in_data1}\n"

    raise ValueError(error_message)



# === STEP 3: Build final config dynamically ===
final_configs = {}
for config_name, config_values in master_config.items():
    if config_name == "common":
        continue
    merged = copy.deepcopy(common)
    merged.update(config_values)
    final_configs[config_name] = merged

new_config = {}

for key, cfg in final_configs.items():
    cfg["metrics_data"] = cfg.get("all_kpis", []) + ["CPM"] + ["ROI"]
    cfg["granularity_levels_length"] = len(cfg["granularity_levels"])
    cfg["granularity_levels"] = cfg["granularity_levels"] + ["Model", "Year", "Month"]
    cfg['comparision_data_range'] = cfg['pivot_data_range']


    # Decide file paths
    if "rollingmonth" in key.lower():
        # Rolling or comparison configs → use combined file
        cfg["rroi_excel_file_path"] = combined_input_file_path
        cfg["sheet_name"] = combined_file_sheet_name
        cfg["models_name"] = sheet_names  
    else:
        # All other configs → use current month file
        cfg["rroi_excel_file_path"] = current_month_file_path
        cfg["sheet_name"] = current_month_sheet_name
        cfg["models_name"] = [sheet_names[1]]

    cfg["model_length"] = len(cfg.get("models_name", []))
    # Output file paths
    cfg["model_validatoin_file_path"] = f"validation_{key}.xlsx"
    cfg["month_on_month_comparision_file_path"] = f"month_on_month_comparision_{key}.xlsx"
    cfg["combined_data_path"] = f"{key}.xlsx"

    new_config[key] = cfg


# print(json.dumps(new_config, indent=4))



# === STEP 4: Process each config ===
process_data_path = []
for key, value in new_config.items():
    print(f"\n\n------------ Generating {key} sheet ------------")
    process_result = process_function(config=value)
    process_data_path.append(process_result)

# === STEP 5: Merge outputs ===
final_file = single_excel_file(process_data_path, final_file_name)

# === STEP 6: Cleanup temp files ===
for filename in os.listdir("."):
    remove_file_list = process_data_path + [previous_month_file_path] + [current_month_file_path] + [combined_input_file_path]
    if filename.endswith(".xlsx") and any(sub in filename for sub in remove_file_list):
        os.remove(filename)
        print(f"Removed temporary file: {filename}")

if os.path.exists("__pycache__"):
    shutil.rmtree("__pycache__")
    print("__pycache__ has been deleted.")