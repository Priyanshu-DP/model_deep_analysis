import pandas as pd
import numpy as np
import calendar
import os



def rolling_month_comaprarision(df_percent,comparision_data_range_year_month, models_name, required_kpis_list, output_month_on_month_comparision_file_path):
    
    month_list = comparision_data_range_year_month
    model_names = models_name
    metrics = required_kpis_list

    def parse_ym(ym_str):
        """Parse 'YYYY-MM' into (year, month) tuple."""
        return tuple(map(int, ym_str.split('-')))

    changes = {}

    # Each iteration compares a specific month for both models
    for idx, ym in enumerate(month_list):
        year, month = parse_ym(ym)
        month_name = calendar.month_abbr[month]

        # Prepare columns for both models for the same month
        cols_model_1 = [(model_names[0], year, month, metric) for metric in metrics]
        cols_model_2 = [(model_names[1], year, month, metric) for metric in metrics]

        df1 = df_percent[cols_model_1].copy()
        df2 = df_percent[cols_model_2].copy()

        df1.columns = metrics
        df2.columns = metrics

        # Percent change calculation
        percent_change = ((df2 / df1.replace(0, np.nan)) - 1) * 100
        percent_change = percent_change.round(9)

        label = f"{month_name} month % changes for {model_names[1]} to {model_names[0]}"
        changes[label] = percent_change

    # Combine all monthly changes into one DataFrame
    combined = pd.concat(changes.values(), axis=1, keys=changes.keys())

    # Save to Excel
    combined.to_excel(output_month_on_month_comparision_file_path, sheet_name='Month on Month Model Comparison', index=True)
                      
    return output_month_on_month_comparision_file_path


def mom_comparision(df_percent,comparision_data_range_year_month, models_name, required_kpis_list, output_month_on_month_comparision_file_path):
    # comparision_data_range_year_month= config.get('comparision_data_range')
    
    metrics = required_kpis_list

    # List of months to compare
    month_list = comparision_data_range_year_month
    models_name = models_name[0]

    # Helper to create (Year, Month) tuple from 'YYYY-MM'
    def parse_ym(ym):
        return tuple(map(int, ym.split('-')))

    # Create all consecutive month pairs
    month_pairs = [(month_list[i], month_list[i+1]) for i in range(len(month_list) - 1)]

    # Build month-on-month % change for each metric, for each month-pair
    changes = {}  # {comparison_label: DataFrame}
    for ym1, ym2 in month_pairs:
        y1, m1 = parse_ym(ym1)
        y2, m2 = parse_ym(ym2)
        cols_1 = [(models_name, y1, m1, metric) for metric in metrics]
        cols_2 = [(models_name, y2, m2, metric) for metric in metrics]

        df1 = df_percent[cols_1].copy()
        df2 = df_percent[cols_2].copy()
        df1.columns = metrics
        df2.columns = metrics

        # Calculate percent change: ((new / old) - 1) * 100
        ratio_change = ((df2 / df1.replace(0, np.nan)) - 1) * 100
        change = ratio_change.round(9)

        
        ym1_month_number = int(ym1.split('-')[1])
        ym1_month_name = calendar.month_abbr[ym1_month_number]
        ym1_year = ym1.split('-')[0]

        ym2_month_number = int(ym2.split('-')[1])
        ym2_month_name = calendar.month_abbr[ym2_month_number]
        ym2_year = ym2.split('-')[0]

        label = f"% change for {ym1_month_name}'{ym1_year} > {ym2_month_name}'{ym2_year} "
        changes[label] = change

    # Stitch all change DataFrames into a single wide DataFrame.
    combined = pd.concat(changes.values(), axis=1, keys=changes.keys())


    # Save to Excel
    combined.to_excel(output_month_on_month_comparision_file_path, sheet_name='Month on Month Comparision', index=True)


    return output_month_on_month_comparision_file_path


def combine_aggregate_and_comparision_sheet(output_model_validation_file_path, output_month_on_month_comparision_file_path, combined_data_path, granularity_levels_length):

    # Combines aggregatoin and aomparision sheet in a single sheet
    file1 = output_model_validation_file_path
    file2 = output_month_on_month_comparision_file_path
    output_file = combined_data_path

    # Load first pivot (first sheet, raw layout preserved)
    df1 = pd.read_excel(file1, sheet_name=0, header=None)

    # Load second pivot (first sheet)
    df2 = pd.read_excel(file2, sheet_name=0, header=None)

    # Ensure same number of rows (pad with blanks if needed)
    max_rows = max(df1.shape[0], df2.shape[0])
    df1 = df1.reindex(range(max_rows)).fillna("")
    df2 = df2.reindex(range(max_rows)).fillna("")


    num_level = granularity_levels_length
    df2 = df2.iloc[:, num_level:] 

    # add a extra row at the top of df2
    df2.loc[-1] = [''] * df2.shape[1]  # add a new row at the top
    df2.index = df2.index + 1  # shift index
    df2.sort_index(inplace=True)  # sort by index to maintain order
    # add a extra row at the top of df2
    df2.loc[-1] = [''] * df2.shape[1]  # add a new row at the top
    df2.index = df2.index + 1  # shift index
    df2.sort_index(inplace=True)  # sort by index to maintain order

    # Combine side by side
    combined = pd.concat([df1, df2], axis=1)

    # Save into single sheet
    combined.to_excel(output_file, index=False, header=False)
    return output_file


def single_excel_file(process_data_path, file_path):
    # Read all excel in current directory and make one excel filewith sheet name as file name
    excel_files = process_data_path
    with pd.ExcelWriter(file_path) as writer:
        for file in excel_files:
            sheet_name = os.path.splitext(file)[0][:31]  # Excel sheet names max length is 31
            df = pd.read_excel(file, header=None) 
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    return file_path
