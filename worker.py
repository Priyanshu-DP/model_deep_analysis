
## ---------------  Pivot table for all year and months ------------- ##

import pandas as pd # type: ignore
from collections import defaultdict
from utils import *
import wasabi # type: ignore

printer = wasabi.Printer()


def process_function(config):
    excel_file_path = config.get('rroi_excel_file_path')
    sheet_name = config.get('sheet_name')
    granularity_levels = config.get('granularity_levels')
    granularity_levels_length = config.get('granularity_levels_length')
    models_name = config.get('models_name')
    model_length = config.get('model_length')
    required_kpis_list = config.get('specific_kpis')
    all_kpis = config.get('all_kpis')
    metrics_data = config.get('metrics_data')
    pivot_data_range_year_month = config.get('pivot_data_range')
    comparision_data_range_year_month= config.get('comparision_data_range')
    output_model_validation_file_path = config.get('model_validatoin_file_path')
    output_month_on_month_comparision_file_path = config.get('month_on_month_comparision_file_path')
    combined_data_path = config.get('combined_data_path')

    
    sheet_MRF = pd.read_excel(excel_file_path, sheet_name=sheet_name)


    default_agg_func = 'sum'
    agg_dict = {col: default_agg_func for col in all_kpis}


    # Group by specified dimensions
    grouped = sheet_MRF.groupby(granularity_levels, dropna=False).agg(agg_dict).reset_index()



    # Fill NaNs in filter columns only
    for col in granularity_levels:
        if grouped[col].isna().any():
            grouped[col] = grouped[col].fillna('')



    # Rename aggregated columns to include prefix (optional)
    rename_dict = {col: f"{col}" for col in all_kpis}
    grouped.rename(columns=rename_dict, inplace=True)


    # Pivot table creation
    pivot = grouped.pivot_table(
        index=granularity_levels[:-3],
        columns=granularity_levels[-3:],
        values=rename_dict.values(),
        fill_value=0
    )


    # Reorder columns: Model -> Year → Month → Metric
    pivot.columns = pivot.columns.swaplevel(0, 3)
    pivot.columns = pivot.columns.swaplevel(0, 2)
    pivot.columns = pivot.columns.swaplevel(0, 1)
    pivot = pivot.sort_index(axis=1)



    printer.good("Step 1 completed: Pivot table created successfully.")

    ##### Here we are computing "CPM" column as desired.   ######

    # Get all (Year, Month) pairs present in columns
    year_months = sorted(set(col[:3] for col in pivot.columns if col[3] == 'Cost'))



    # Collect new CPM columns in a dict
    new_cols = {}

    for ym in year_months:
        cost_col = (*ym, 'Cost')
        impression_col = (*ym, 'Impression')
        sales_col = (*ym, 'Overall Dollar Sales')
        cpm_col = (*ym, 'CPM')
        roi_col = (*ym, 'ROI')
    

        # Avoid division by zero, compute CPM, `.mask()`` keeps the dtype as numeric (float), while replace(0, pd.NA) may upgrade to object dtype.
        new_cols[cpm_col] = (
            (pivot[cost_col] / pivot[impression_col].mask(pivot[impression_col] == 0)) * 1000
        ).astype(float).round(9)

        new_cols[roi_col] = (
            (pivot[sales_col] / pivot[cost_col].mask(pivot[cost_col] == 0))
        ).astype(float).round(9)


    # Add all new columns at once to avoid fragmentation
    pivot = pd.concat([pivot, pd.DataFrame(new_cols, index=pivot.index)], axis=1)


    printer.good("Step 2 completed: CPM columns calculated successfully.")




    # Calculate Grand Total row (sum across 'Media Type', 'Master Channel', etc.)
    grand_totals = pivot.sum(axis=0)

    # Return string if only one filter, otherwise tuple
    grand_total_key = ('Grand Total') if granularity_levels_length == 1 else tuple(['Grand Total'] + [''] * (granularity_levels_length - 1))

    # grand_total_key = 'Grand Total'
    pivot.loc[grand_total_key, :] = grand_totals

    # Now calculate ROI and CPM for the Grand Total row based on summed metrics
    for ym in year_months:
        cost_col = (*ym, 'Cost')
        impression_col = (*ym, 'Impression')
        sales_col = (*ym, 'Overall Dollar Sales')
        roi_col = (*ym, 'ROI')
        cpm_col = (*ym, 'CPM')

        # Retrieve the grand total values for each metric
        cost = pivot.loc[grand_total_key, cost_col]
        impression = pivot.loc[grand_total_key, impression_col]
        sales = pivot.loc[grand_total_key, sales_col]

        # Calculate ROI = Sales / Cost
        roi = (sales / cost) if cost != 0 else pd.NA
        # Calculate CPM = (Cost / Impression) * 1000
        cpm = ((cost / impression) * 1000) if impression != 0 else pd.NA

        # Assign these values back into the Grand Total row
        pivot.loc[grand_total_key, roi_col] = roi
        pivot.loc[grand_total_key, cpm_col] = cpm

    # Sort columns (optional, for clarity)
    pivot = pivot.sort_index(axis=1)

    ## ---------------  Pivot table for all year and months ------------- ##
    printer.good("Step 3 completed: Grand Total row calculated successfully.")


    # List of selected year-months (from user)
    selected_ym = pivot_data_range_year_month

    # Convert to tuples of integers: [(2024, 12), (2025, 1), (2025, 2)]
    selected_ym_tuples = [(int(ym.split('-')[0]), int(ym.split('-')[1])) for ym in selected_ym]

    # Filter pivot columns based on selected Year-Month combinations
    filtered_columns = [col for col in pivot.columns if (col[1], col[2]) in selected_ym_tuples]
    

    # Filter pivot DataFrame
    filtered_pivot = pivot.loc[:, filtered_columns]



    printer.good("Step 4 completed: Filtered pivot table created successfully.")

    # Arrange the columns (CPM to the last of all column for a year and month)
    grouped = defaultdict(list)
    for col in filtered_pivot.columns:
        year_month = (col[0], col[1], col[2])
        grouped[year_month].append(col)

    # For each group, move 'CPM' to the end
    reordered_cols = []
    for cols in grouped.values():
        cpm_cols = [col for col in cols if col[3] == 'CPM']
        other_cols = [col for col in cols if col[3] != 'CPM']
        reordered_cols.extend(other_cols + cpm_cols)

    # Reorder the DataFrame
    pivot_with_totals = filtered_pivot[reordered_cols]

    df_percent = pivot_with_totals.copy()

    printer.good("Step 5 completed: Columns reordered successfully.")




    # Loop through each (Year, Month)
    for col in pivot_with_totals.columns:
        if col[3] == 'Overall Dollar Sales':
            total = pivot_with_totals[col].sum() / 2 # 
            # Avoid division by zero
            if total != 0:
                df_percent[col] = (pivot_with_totals[col] / total).round(9) * 100
            else:
                df_percent[col] = 0



    if not required_kpis_list:
        required_kpis_list = metrics_data
    else:
        missing_kpis = set(required_kpis_list) - set(metrics_data)
        if missing_kpis:
            raise ValueError(f"These KPIs are not present in given all_kpis list: {missing_kpis}")


    # Filter only those columns from "all_kpis" that in "required_kpis_list"
    filtered_columns = [col for col in df_percent.columns.to_list() if col[3] in required_kpis_list]   

    df_percent = df_percent[filtered_columns]

    df_percent.to_excel(output_model_validation_file_path, sheet_name='RROI Model Validation', index=True)

    printer.good("Step 6 completed: Model validation file saved successfully.")




    ### -------------- Creating month-on-month comparison for two months -------------- ###
    try:
        params = dict(
            df_percent=df_percent,
            comparision_data_range_year_month=comparision_data_range_year_month,
            models_name=models_name,
            required_kpis_list=required_kpis_list,
            output_month_on_month_comparision_file_path=output_month_on_month_comparision_file_path
        )

        if model_length == 2:
            rolling_month_comaprarision(**params)
            printer.good("Step 7 completed: Rolling month comaprarision with model file saved successfully.")


        elif model_length == 1:         
            mom_comparision(**params)
            printer.good("Step 7 completed: Month-on-month comparison file saved successfully.")
        
    except:
            printer.fail("Error in Step 7 : Month-on-month/Rolling month comparison file.")



    # Combines aggregatoin and comparision sheet in a single sheet
    params = dict(
            output_model_validation_file_path = output_model_validation_file_path,
            output_month_on_month_comparision_file_path = output_month_on_month_comparision_file_path,
            combined_data_path = combined_data_path,
            granularity_levels_length = granularity_levels_length
        )
    output_file = combine_aggregate_and_comparision_sheet(**params)
    if output_file:
        printer.good("Step 8 completed: Combined side-by-side pivot saved successfully.")
    else:
        printer.fail("Error in Step 8 : unable to combine aggregate data and comparision sheet in a single sheet")

    return output_file

