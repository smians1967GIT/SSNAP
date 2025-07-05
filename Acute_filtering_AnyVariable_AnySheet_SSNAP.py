import gradio as gr
import pandas as pd
import numpy as np
import os
import tempfile
import re

sheet_cache = {}
QUARTER = "2025-Q1"

# Complete SSNAP Metric IDs
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

def handle_file_upload(file):
    if file is None or not os.path.exists(file):
        return gr.update(choices=[], value=None)
    try:
        xls = pd.ExcelFile(file)
        sheet_cache[file] = xls.sheet_names
        return gr.update(choices=xls.sheet_names, value=xls.sheet_names[0])
    except Exception as e:
        return gr.update(choices=[f"❌ {str(e)}"], value=None)

def filter_metrics(metric_type):
    if metric_type == "Patient":
        return [m for m in METRIC_IDS if m.startswith("G") or m.startswith("J")]
    elif metric_type == "Team":
        return [m for m in METRIC_IDS if m.startswith("H") or m.startswith("K")]
    return METRIC_IDS

def extract_single_metric(sheet, metric_id, metadata, df):
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
        clean_value = (np.nan if raw_value in ["", " ", "Too few to report", ".", "N/A", "nan"]
                       else hhmm_to_minutes(raw_value))
        records.append({
            "Quarter": QUARTER,
            "Domain": sheet,
            "Team Type": team["Team Type"],
            "Region": team["Region"],
            "Trust": team["Trust"],
            "Team": team["Team"],
            "Metric ID": metric_id,
            "Metric Label": metric_label,
            "Value": clean_value
        })
    return records

def export_metrics(file, sheet, metric_list):
    if not metric_list:
        return "❌ Select at least one metric", None
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None)
        metadata = df.iloc[0:4, 4:].T
        metadata.columns = ['Team Type', 'Region', 'Trust', 'Team']
        metadata = metadata.dropna(subset=['Team']).reset_index(drop=True)
        all_records = []
        for m in metric_list:
            rows = extract_single_metric(sheet, m, metadata, df)
            if rows:
                all_records.extend(rows)
        if not all_records:
            return "❌ No matching metrics found.", None
        df_all = pd.DataFrame(all_records)
        df_all["Value"] = pd.to_numeric(df_all["Value"], errors="coerce")
        df_all["Metric Header"] = df_all["Metric ID"] + " - " + df_all["Metric Label"].astype(str)
        df_pivot = df_all.pivot_table(
            index=["Quarter", "Domain", "Team Type", "Region", "Trust", "Team"],
            columns="Metric Header", values="Value", aggfunc="mean"
        ).reset_index()
        df_pivot.columns.name = None
        out_path = tempfile.NamedTemporaryFile(delete=False, suffix=".csv").name
        df_pivot.to_csv(out_path, index=False)
        return f"✅ Exported: {out_path}", out_path
    except Exception as e:
        return f"❌ {str(e)}", None

def select_all_metrics(metric_type):
    return filter_metrics(metric_type)

with gr.Blocks() as demo:
    gr.Markdown("## 🧠 SSNAP Metrics Extractor (Safe Sheet Selection)")
    file_input = gr.File(label="Upload Excel File", file_types=[".xlsx"], type="filepath")
    sheet_dropdown = gr.Dropdown(label="Select Sheet", choices=[])
    file_input.change(fn=handle_file_upload, inputs=file_input, outputs=sheet_dropdown)

    metric_type = gr.Dropdown(["Patient", "Team", "Both"], label="Metric Type")
    metric_select = gr.CheckboxGroup(choices=[], label="Choose Metrics")
    select_all_btn = gr.Button("Select All")

    metric_type.change(fn=lambda t: gr.update(choices=filter_metrics(t), value=[]), inputs=metric_type, outputs=metric_select)
    select_all_btn.click(fn=select_all_metrics, inputs=metric_type, outputs=metric_select)

    export_btn = gr.Button("Export to CSV")
    status = gr.Textbox(label="Status")
    file_output = gr.File(label="Download CSV")
    export_btn.click(fn=export_metrics, inputs=[file_input, sheet_dropdown, metric_select], outputs=[status, file_output])

demo.launch()
