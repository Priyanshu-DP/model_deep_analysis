# Model Deep Dive Automation

## 📥 Input  
1. The automation requires an **Excel file** containing two sheets:  
   - `Previous month ROI`  
   - `Current month ROI`  

   Both sheets must follow the **same column headers and naming convention** as the respective `Model_Name`.  
2. The second input is a **`config.json`** file.  

---

## ⚙️ How to Fill `config.json`  

1. `config.json` contains **6 dictionaries** .  
2. The first dictionary is named **`common`**, where you fill all required fields.  
3. The remaining dictionaries like ( **`CP`**, **`CPA`**, **`RollingMonthCPA`**, **`MediaType`**, **`MediaType_model_comparision`** ) defines the output sheets.  
   - You can generate as many sheets as you need like above, with different `granularity_levels` and `kpis`, for that just add more dictionaries with the same structure.
   - If you want to generate a **rolling month comparison** sheet, include the word `RollingMonth` in the key of your new 
   **dictionaries**.  
   - As keys of these **dictionaries** are acutual sheet name in output file, so keep the name of key less than `31` (Standard Limit in Excel).
   - You can add as many as `granularity_levels` and `all_kpis` as you want in any order. 
   - `all_kpis` must contain these value for calcualting the `ROI` and `CPM` : `Cost`, `Impression`, `Overall Dollar Sales`.
   - `ROI` and `CPM` already included in `all_kpis`. You do not need to fill there.

  
4. The field **`specific_kpis`** is optional:  
   - If provided → only the listed KPIs will be generated.  
   - If omitted → all KPIs will be generated as mention in `comman` dictionary.  
5. The order in which you will provide the value in `granularity_levels` and `specific_kpis`, same order will preserve in output sheets.
6. The field **`pivot_data_range`**:  
   - For rolling month → provide a **list of all months** to compare.  
   - Otherwise → provide only the **start and end month**.  

👉 Steps:  
- Fill in the **`config.json`**.  
- Install dependency **`pip install -r requirement.txt`**
- Run **`python main.py`**.  

---


### Sample Config Json Format : 
```json
{
  "common": {
    "brand_name" : "Dove",
    "rroi_excel_file_path": "Dove_B3-INC06_MAY25_RROI - 24-07-2025 Analysis.xlsx",
    "sheet_name":  ["B3-INC05", "B3-INC06"],
    "all_kpis": ["Cost","Impression","Lagged Impression","Overall Dollar Sales","Overall ROI"],
    "output_file_name" : "Dove_INC06.xlsx" 
  },
  "CPA": {
    "granularity_levels" : ["Media Type", "Product Line", "Master Channel", "Channel/Daypart", "Platform", "Influencer Say", "Audience"],
    "pivot_data_range":  ["2025-04","2025-05"],
    "specific_kpis": ["Overall Dollar Sales", "Impression", "Lagged Impression"]
    },

  "CP": {
    "granularity_levels" : ["Media Type", "Product Line", "Master Channel", "Channel/Daypart", "Platform"],
    "pivot_data_range": ["2025-04","2025-05"]
    },

  "RollingMonthCPA": {
    "granularity_levels" : ["Media Type", "Product Line", "Master Channel", "Channel/Daypart", "Platform", "Influencer Say", "Audience"],
    "pivot_data_range": ["2025-03","2025-04"]
    },

  "MediaType": {
    "granularity_levels": ["Media Type"],
    "pivot_data_range": ["2025-04","2025-05"]
    },

  "MediaType_RollingMonth": {
    "granularity_levels": ["Media Type"],
    "pivot_data_range": ["2025-03","2025-04"]
    },

    "Product Line": {
    "granularity_levels": ["Product Line", "Master Channel", "Channel/Daypart"],
    "pivot_data_range": ["2025-04","2025-05"],
    "specific_kpis": ["Overall Dollar Sales", "Impression", "Lagged Impression"]
    }
}
```


## 📤 Output  
The automation produces a **single Excel file** containing the following sheets (corresponding to given sample JOSN):  

- `CP`  
- `CPA`  
- `RollingMonthCPA`  
- `MediaType`  
- `MediaType_RollingMonth`  
- `Product Line`