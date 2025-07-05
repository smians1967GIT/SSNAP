import pandas as pd
import numpy as np
import os
import re
import tempfile
import traceback
import gradio as gr

# --- Configuration ---
QUARTER = "2025-Q1"
EXPORT_DIR = r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets"

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

def hhmm_to_minutes(value):
    try:
        if isinstance(value, str) and re.match(r"^\d{1,2}:\d{2}$", value.strip()):
            h, m = map(int, value.strip().split(":"))
            return h * 60 + m
        return value
    except:
        return value

def load_sheet_names(filepath):
    try:
        xls = pd.ExcelFile(filepath)
        return xls.sheet_names
    except Exception as e:
        print("❌ Failed to load sheets:")
        traceback.print_exc()
        return [f"❌ Error loading sheets: {e}"]

def load_sheet_names_with_default(filepath):
    sheets = load_sheet_names(filepath)
    if isinstance(sheets, list) and len(sheets) > 0:
        return gr.update(choices=sheets, value=sheets[0])
    return gr.update(choices=[], value=None)

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
            raw_value = str(value).strip().lower()
            if raw_value in ["", " ", "too few to report", ".", "n/a", "nan"]:
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
    except Exception as e:
        print(f"❌ Failed to extract metric {metric_id} from sheet {sheet_name}")
        traceback.print_exc()
        return None

def extract_multiple_metrics(filepath, sheet_name, selected_metrics, export_dir):
    try:
        if not selected_metrics:
            return "❌ Please select at least one metric.", None

        xls = pd.ExcelFile(filepath)
        df = xls.parse(sheet_name, header=None)

        metadata = df.iloc[0:4, 4:].T
        metadata.columns = ['Team Type', 'Region', 'Trust', 'Team']
        metadata = metadata.dropna(subset=['Team']).reset_index(drop=True)

        all_records = []
        for metric_id in selected_metrics:
            result = extract_single_metric(sheet_name, metric_id, metadata, df)
            if result:
                all_records.extend(result)

        if not all_records:
            return "❌ No valid data found for selected metrics.", None

        df_combined = pd.DataFrame(all_records)
        df_combined["Value"] = pd.to_numeric(df_combined["Value"], errors="coerce")
        df_combined["Metric Header"] = df_combined["Metric ID"] + " - " + df_combined["Metric Label"].astype(str)

        df_pivot = df_combined.pivot_table(
            index=["Quarter", "Domain", "Team Type", "Region", "Trust", "Team"],
            columns="Metric Header",
            values="Value",
            aggfunc="mean"
        ).reset_index()
        df_pivot.columns.name = None

        if not os.path.exists(export_dir):
            os.makedirs(export_dir)

        output_file = os.path.join(export_dir, f"Combined_Metrics_{QUARTER}_TRANSPOSED.csv")
        df_pivot.to_csv(output_file, index=False)

        return f"✅ Exported: {output_file}", output_file

    except Exception as e:
        print("❌ Error during extraction:")
        traceback.print_exc()
        return f"❌ Extraction failed: {str(e)}", None

def filter_metrics_by_type(metric_type):
    if metric_type == "Patient":
        filtered = [m for m in METRIC_IDS if m.startswith("G") or m.startswith("J")]
    elif metric_type == "Team":
        filtered = [m for m in METRIC_IDS if m.startswith("H") or m.startswith("K")]
    else:
        filtered = METRIC_IDS
    return gr.update(choices=filtered, value=[])

def select_all_metrics(metric_type):
    if metric_type == "Patient":
        return [m for m in METRIC_IDS if m.startswith("G") or m.startswith("J")]
    elif metric_type == "Team":
        return [m for m in METRIC_IDS if m.startswith("H") or m.startswith("K")]
    else:
        return METRIC_IDS

def gradio_interface(filepath, sheet_name, metric_ids, export_dir):
    try:
        return extract_multiple_metrics(filepath, sheet_name, metric_ids, export_dir)
    except Exception as e:
        print("❌ Unhandled Gradio interface error:")
        traceback.print_exc()
        return f"❌ Unexpected error: {str(e)}", None

# --- Build Gradio App ---
with gr.Blocks() as demo:
    gr.Markdown("## 🧠 SSNAP Multi-Metric Extractor (Upload + Filtered Export)")

    file_input = gr.File(label="Upload Excel File (.xlsx)", type="filepath")
    sheet_dropdown = gr.Dropdown(label="Select Sheet")

    file_input.change(fn=load_sheet_names_with_default, inputs=file_input, outputs=sheet_dropdown)

    metric_type_dropdown = gr.Dropdown(choices=["Patient", "Team", "Both"], value="Both", label="Select Metric Type")
    metric_checkboxes = gr.CheckboxGroup(choices=METRIC_IDS, label="Select Metric IDs")
    select_all_btn = gr.Button("Select All Metrics")

    export_dir_input = gr.Textbox(label="Export Directory (Full Path)", value=EXPORT_DIR)
    export_btn = gr.Button("Export Selected Metrics")

    status_box = gr.Textbox(label="Status")
    file_download = gr.File(label="Download CSV")

    metric_type_dropdown.change(fn=filter_metrics_by_type, inputs=metric_type_dropdown, outputs=metric_checkboxes)
    select_all_btn.click(fn=select_all_metrics, inputs=metric_type_dropdown, outputs=metric_checkboxes)
    export_btn.click(
        fn=gradio_interface,
        inputs=[file_input, sheet_dropdown, metric_checkboxes, export_dir_input],
        outputs=[status_box, file_download]
    )

if __name__ == "__main__":
    demo.launch(allowed_paths=[EXPORT_DIR, os.getcwd(), tempfile.gettempdir()])
