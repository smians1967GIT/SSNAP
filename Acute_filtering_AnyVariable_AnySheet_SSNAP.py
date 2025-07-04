import numpy as np
import pandas as pd
import os
import re
import gradio as gr

# --- Configuration ---
EXCEL_PATH = r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets\JanMar2025-FullResultsPortfolio.xlsx"
EXPORT_DIR = r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets"
QUARTER = "2025-Q1"

# --- Metric IDs (Patient + Team) ---
METRIC_IDS = [
    "G6.6.3", "H6.6.3", "G6.4", "H6.4", "G6.20", "H6.20", "G6.62", "H6.62", "G9.34", "H9.34",
    "G8.0.3", "H8.0.3", "G14.20", "H14.20", "G7.18.1", "H7.18.1", "G7.4", "H7.4", "J8.11", "K32.11",
    "J12.13", "K24.13", "G16.3", "H16.3", "G16.80", "H16.80", "H16.88", "G16.42", "H16.42",
    "G19.3", "H20.3", "G19.4", "H20.4", "G19.7", "H20.7", "G9.19", "H9.19", "G15.27", "H15.27",
    "G10.27", "H10.27", "G11.27", "H11.27", "G12.24", "H12.24", "J14.6", "J6.12", "K38.3",
    "J6.13", "K38.4", "J3.13", "K34.13", "J3.4", "K34.4", "J4.13", "K35.13", "J4.4", "K35.4",
    "J5.13", "K36.13", "J5.4", "K36.4", "J16.15.1", "K3.15.1", "J26.8", "K6.8", "J17.20", "K12.20",
    "J18.20", "K13.20", "J38.6", "K29.41", "J36.20", "K29.23", "J37.12", "K29.35", "J34.3", "K29.3"
]

# --- Convert HH:MM → MM ---
def hhmm_to_minutes(value):
    try:
        if isinstance(value, str) and re.match(r"^\d{1,2}:\d{2}$", value.strip()):
            h, m = map(int, value.strip().split(":"))
            return h * 60 + m
        return value
    except:
        return value

# --- Get sheet names ---
def get_sheets():
    try:
        xls = pd.ExcelFile(EXCEL_PATH)
        return xls.sheet_names
    except Exception as e:
        return [f"❌ Error loading sheets: {e}"]

# --- Extract one metric ---
def extract_single_metric(sheet_name, metric_id, metadata, df):
    try:
        metric_row = df[df.iloc[:, 0].astype(str).str.contains(metric_id, na=False)]
        if metric_row.empty:
            return None
        metric_label = metric_row.iloc[0, 0]
        metric_values = metric_row.iloc[0, 4:4 + len(metadata)]
        records = []
        for i in range(min(len(metadata), len(metric_values))):
            team = metadata.iloc[i]
            value = metric_values.iloc[i]
            raw_value = str(value).strip()
            if raw_value in ["", " ", "Too few to report", ".", "N/A", "nan"]:
                clean_value = np.nan
            else:
                clean_value = hhmm_to_minutes(raw_value)

            records.append({
                "Quarter": QUARTER,
                "Domain": sheet_name,
                "Team Type": team["Team Type"],
                "Region": team["Region"],
                "Trust": team["Trust"],
                "Team": team["Team"],
                "Metric ID": metric_id,
                "Metric Label": metric_label,
                "Value": clean_value
            })
        return records
    except Exception:
        return None

# --- Extract multiple metrics and export transposed ---
def extract_multiple_metrics(sheet_name, selected_metrics):
    if not selected_metrics:
        return "❌ Please select at least one metric.", None

    try:
        xls = pd.ExcelFile(EXCEL_PATH)
        df = xls.parse(sheet_name, header=None)
        metadata = df.iloc[0:4, 4:].T
        metadata.columns = ['Team Type', 'Region', 'Trust', 'Team']
        metadata = metadata.dropna(subset=['Team']).reset_index(drop=True)
    except Exception as e:
        return f"❌ Error loading or parsing sheet: {e}", None

    all_records = []
    for metric_id in selected_metrics:
        result = extract_single_metric(sheet_name, metric_id, metadata, df)
        if result:
            all_records.extend(result)

    if not all_records:
        return "❌ No valid data found for selected metrics.", None

    df_combined = pd.DataFrame(all_records)

    # Ensure all values are numeric before pivoting
    df_combined["Value"] = pd.to_numeric(df_combined["Value"], errors="coerce")

    # Add column name with ID and label
    df_combined["Metric Header"] = df_combined["Metric ID"] + " - " + df_combined["Metric Label"].astype(str)

    # Pivot and aggregate
    df_pivot = df_combined.pivot_table(
        index=["Quarter", "Domain", "Team Type", "Region", "Trust", "Team"],
        columns="Metric Header",
        values="Value",
        aggfunc="mean"
    ).reset_index()

    df_pivot.columns.name = None

    output_file = os.path.join(EXPORT_DIR, f"Combined_Metrics_{QUARTER}_TRANSPOSED.csv")
    df_pivot.to_csv(output_file, index=False)

    return f"✅ Exported: {output_file}", output_file

# --- "Select All" logic ---
def select_all_metrics():
    return gr.update(value=METRIC_IDS)

# --- Gradio interface handler ---
def gradio_interface(sheet_name, metric_ids):
    return extract_multiple_metrics(sheet_name, metric_ids)

# --- UI ---
sheet_choices = get_sheets()

with gr.Blocks() as demo:
    gr.Markdown("## 🧠 SSNAP Multi-Metric Extractor (Transposed with Metric Labels)")

    sheet_dropdown = gr.Dropdown(choices=sheet_choices, label="Select Sheet")
    metric_checkboxes = gr.CheckboxGroup(choices=METRIC_IDS, label="Select Metric IDs")
    select_all_btn = gr.Button("Select All Metrics")
    export_btn = gr.Button("Export Selected Metrics")

    status_box = gr.Textbox(label="Status")
    file_download = gr.File(label="Download Cleaned Transposed CSV")

    select_all_btn.click(fn=select_all_metrics, outputs=metric_checkboxes)
    export_btn.click(fn=gradio_interface, inputs=[sheet_dropdown, metric_checkboxes], outputs=[status_box, file_download])

# --- Launch ---
if __name__ == "__main__":
    demo.launch(allowed_paths=[r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets"])
