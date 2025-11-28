import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import pdfplumber


COLUMNS = [
    "Source File",
    "Patient ID",
    "Scanned ID",
    "Scan Date",
    "Scan Time",
    "Recording #",
    "Date of Birth",
    "Age",
    "Gender",
    "Height (m)",
    "# of Pulses",
    "Pulse Height",
    "Pulse Height Variation (%)",
    "Diastolic Variation (%)",
    "Shape Deviation (%)",
    "Pulse Length Variation (%)",
    "Overall Quality (%)",
    "Peripheral Systolic Pressure (mmHg)",
    "Peripheral Diastolic Pressure (mmHg)",
    "Peripheral Pulse Pressure (mmHg)",
    "Peripheral Mean Pressure (mmHg)",
    "Aortic Systolic Pressure (mmHg)",
    "Aortic Diastolic Pressure (mmHg)",
    "Aortic Pulse Pressure (mmHg)",
    "Heart Rate (bpm)",
    "Pulse Pressure Amplification (%)",
    "Period (ms)",
    "Ejection Duration (ms)",
    "Ejection Duration (%)",
    "Aortic T2 (ms)",
    "P1 Height (mmHg)",
    "Aortic Augmentation (mmHg)",
    "Aortic AIx AP/PP(%)",
    "Aortic AIx P2/P1(%)",
    "Aortic AIx AP/PP @ HR75 (%)",
    "Buckberg SEVR (%)",
    "PTI Systolic (mmHg.s/min)",
    "PTI Diastolic (mmHg.s/min)",
    "End Systolic Pressure (mmHg)",
    "MAP Systolic (mmHg)",
    "MAP Diastolic (mmHg)",
]

DETAILED_REPORT_MARKER = "PWA Detailed Report"
CLINICAL_REPORT_MARKER = "PWA Clinical Report"
CLINICAL_REPORT_MESSAGE = (
    "Recognized as a Clinical Report, only upload the Detailed Reports"
)
UNRECOGNIZED_REPORT_MESSAGE = "Not recognized as a PWA Detailed Report"


def select_input_files() -> tuple[Path, ...]:
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select PWA PDF files",
        filetypes=[("PDF Files", "*.pdf")],
    )
    root.update()
    return tuple(Path(path) for path in file_paths)


def select_output_file() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    output_path = filedialog.asksaveasfilename(
        title="Save Excel file as",
        initialfile="pwa_export.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
    )
    root.update()
    if not output_path:
        return None
    return Path(output_path)


def extract_text(pdf_path: Path) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = [page.extract_text() or "" for page in pdf.pages]
    return "\n".join(pages_text)


def _search(pattern: str, text: str) -> str | None:
    match = re.search(pattern, text, flags=re.IGNORECASE)
    return match.group(1) if match else None


def _to_number(value: str) -> int | float | str:
    normalized = value.strip()
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", normalized):
        return float(normalized) if "." in normalized else int(normalized)
    return value


def _extract_scan_datetime(text: str) -> tuple[str | None, str | None]:
    date_time_match = None
    for date_time_match in re.finditer(
        r"([0-9]{2}/[0-9]{2}/[0-9]{4})\s+([0-9]{2}:[0-9]{2}(?::[0-9]{2})?)",
        text,
    ):
        pass
    if date_time_match:
        return date_time_match.group(1), date_time_match.group(2)
    return None, None


def _derive_patient_id(pdf_path: Path) -> str:
    return pdf_path.stem.split("_", 1)[0]


def parse_report_text(text: str) -> dict[str, object]:
    normalized = re.sub(r"\s+", " ", text)

    patient_id = _search(r"Patient ID:\s*(\S+)", normalized)
    dob = _search(r"Date Of Birth:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", normalized)

    scan_date, scan_time = _extract_scan_datetime(normalized)

    age_gender_match = re.search(r"Age, Gender:\s*([0-9]+),\s*([A-Za-z]+)", normalized, flags=re.IGNORECASE)
    age = age_gender_match.group(1) if age_gender_match else None
    gender = age_gender_match.group(2) if age_gender_match else None

    height_cm = _search(r"Height:\s*([0-9.]+)\s*cm", normalized)
    height_m = round(float(height_cm) / 100, 2) if height_cm else None

    pulses = _search(r"Number Of Pulses:\s*([0-9]+)", normalized)

    heart_rate_period = re.search(r"Heart Rate, Period:\s*([0-9.]+)\s*bpm,\s*([0-9.]+)\s*ms", normalized, flags=re.IGNORECASE)
    heart_rate = heart_rate_period.group(1) if heart_rate_period else None
    period = heart_rate_period.group(2) if heart_rate_period else None

    ejection_match = re.search(r"Ejection Duration \(ED\):\s*([0-9.]+)\s*ms,\s*([0-9.]+)\s*%", normalized, flags=re.IGNORECASE)
    ejection_ms = ejection_match.group(1) if ejection_match else None
    ejection_pct = ejection_match.group(2) if ejection_match else None

    aortic_t2 = _search(r"Aortic T2:\s*([0-9.]+)\s*ms", normalized)
    p1_height = _search(r"P1 Height.*?:\s*([0-9.]+)\s*mmHg", normalized)
    aortic_augmentation = _search(r"Aortic Augmentation.*?:\s*([-+]?[0-9.]+)\s*mmHg", normalized)

    aix_match = re.search(r"Aortic AIx \(AP/PP, P2/P1\):\s*([-+]?[0-9.]+)\s*%,\s*([-+]?[0-9.]+)\s*%", normalized, flags=re.IGNORECASE)
    aortic_aix_ap_pp = aix_match.group(1) if aix_match else None
    aortic_aix_p2_p1 = aix_match.group(2) if aix_match else None

    aix_hr75 = _search(r"Aortic AIx \(AP/PP\) @HR75:\s*([-+]?[0-9.]+)\s*%", normalized)
    buckberg = _search(r"Buckberg SEVR:\s*([0-9.]+)\s*%", normalized)

    pti_match = re.search(r"PTI \(Systole, Diastole\):\s*([0-9.]+),\s*([0-9.]+)\s*mmHg\.s/min", normalized, flags=re.IGNORECASE)
    pti_systolic = pti_match.group(1) if pti_match else None
    pti_diastolic = pti_match.group(2) if pti_match else None

    end_systolic_pressure = _search(r"End Systolic Pressure:\s*([0-9.]+)\s*mmHg", normalized)

    map_match = re.search(r"MAP \(Systole, Diastole\):\s*([0-9.]+),\s*([0-9.]+)\s*mmHg", normalized, flags=re.IGNORECASE)
    map_systolic = map_match.group(1) if map_match else None
    map_diastolic = map_match.group(2) if map_match else None

    pulse_height = _search(r"Pulse Height:\s*([0-9.]+)", normalized)
    pulse_height_variation = _search(r"Pulse Height Variation:\s*([0-9.]+)\s*%", normalized)
    diastolic_variation = _search(r"Diastolic Variation:\s*([0-9.]+)\s*%", normalized)
    shape_deviation = _search(r"Shape Deviation:\s*([0-9.]+)\s*%", normalized)
    pulse_length_variation = _search(r"Pulse Length Variation:\s*([0-9.]+)\s*%", normalized)
    overall_quality = _search(r"Overall Quality:\s*([0-9.]+)\s*%", normalized)

    amplification = _search(r"PP Amplification:\s*([0-9.]+)\s*%", normalized)

    brachial_match = re.search(r"Brachial SYS/DIA:\s*([0-9.]+)/([0-9.]+)", normalized, flags=re.IGNORECASE)
    peripheral_sys = brachial_match.group(1) if brachial_match else None
    peripheral_dia = brachial_match.group(2) if brachial_match else None

    aortic_sys = None
    aortic_dia = None
    peripheral_pp = None
    aortic_pp = None
    peripheral_mean = None
    table_heart_rate = None

    sp_match = re.search(r"SP\s+([0-9.]+)\s+([0-9.]+)", normalized, flags=re.IGNORECASE)
    if sp_match:
        peripheral_sys = peripheral_sys or sp_match.group(1)
        aortic_sys = sp_match.group(2)

    dp_match = re.search(r"DP\s+([0-9.]+)\s+([0-9.]+)", normalized, flags=re.IGNORECASE)
    if dp_match:
        peripheral_dia = peripheral_dia or dp_match.group(1)
        aortic_dia = dp_match.group(2)

    pp_match = re.search(r"PP\s+([0-9.]+)\s+([0-9.]+)", normalized, flags=re.IGNORECASE)
    if pp_match:
        peripheral_pp = pp_match.group(1)
        aortic_pp = pp_match.group(2)

    map_hr_match = re.search(r"MAP HR\s+([0-9.]+)\s+([0-9.]+)", normalized, flags=re.IGNORECASE)
    if map_hr_match:
        peripheral_mean = map_hr_match.group(1)
        table_heart_rate = map_hr_match.group(2)

    if peripheral_sys and peripheral_dia and peripheral_pp is None:
        try:
            peripheral_pp = str(float(peripheral_sys) - float(peripheral_dia))
        except ValueError:
            peripheral_pp = None

    if aortic_sys and aortic_dia and aortic_pp is None:
        try:
            aortic_pp = str(float(aortic_sys) - float(aortic_dia))
        except ValueError:
            aortic_pp = None

    heart_rate = heart_rate or table_heart_rate

    record = {
        "Scanned ID": patient_id,
        "Scan Date": scan_date,
        "Scan Time": scan_time,
        "Date of Birth": dob,
        "Age": age,
        "Gender": gender,
        "Height (m)": height_m,
        "# of Pulses": pulses,
        "Pulse Height": pulse_height,
        "Pulse Height Variation (%)": pulse_height_variation,
        "Diastolic Variation (%)": diastolic_variation,
        "Shape Deviation (%)": shape_deviation,
        "Pulse Length Variation (%)": pulse_length_variation,
        "Overall Quality (%)": overall_quality,
        "Peripheral Systolic Pressure (mmHg)": peripheral_sys,
        "Peripheral Diastolic Pressure (mmHg)": peripheral_dia,
        "Peripheral Pulse Pressure (mmHg)": peripheral_pp,
        "Peripheral Mean Pressure (mmHg)": peripheral_mean,
        "Aortic Systolic Pressure (mmHg)": aortic_sys,
        "Aortic Diastolic Pressure (mmHg)": aortic_dia,
        "Aortic Pulse Pressure (mmHg)": aortic_pp,
        "Heart Rate (bpm)": heart_rate,
        "Pulse Pressure Amplification (%)": amplification,
        "Period (ms)": period,
        "Ejection Duration (ms)": ejection_ms,
        "Ejection Duration (%)": ejection_pct,
        "Aortic T2 (ms)": aortic_t2,
        "P1 Height (mmHg)": p1_height,
        "Aortic Augmentation (mmHg)": aortic_augmentation,
        "Aortic AIx AP/PP(%)": aortic_aix_ap_pp,
        "Aortic AIx P2/P1(%)": aortic_aix_p2_p1,
        "Aortic AIx AP/PP @ HR75 (%)": aix_hr75,
        "Buckberg SEVR (%)": buckberg,
        "PTI Systolic (mmHg.s/min)": pti_systolic,
        "PTI Diastolic (mmHg.s/min)": pti_diastolic,
        "End Systolic Pressure (mmHg)": end_systolic_pressure,
        "MAP Systolic (mmHg)": map_systolic,
        "MAP Diastolic (mmHg)": map_diastolic,
    }

    for key, value in record.items():
        if isinstance(value, str):
            record[key] = _to_number(value)
    return record


def _detect_report_type(text: str) -> str:
    normalized = text.lower()
    if DETAILED_REPORT_MARKER.lower() in normalized:
        return "detailed"
    if CLINICAL_REPORT_MARKER.lower() in normalized:
        return "clinical"
    return "unrecognized"


def _empty_record(message: str, pdf_path: Path) -> dict[str, object]:
    record: dict[str, object] = {column: None for column in COLUMNS}
    record["Source File"] = pdf_path.name
    record["Patient ID"] = message
    return record


def process_pdf(pdf_path: Path) -> dict[str, object]:
    text = extract_text(pdf_path)
    report_type = _detect_report_type(text)

    if report_type == "detailed":
        data = parse_report_text(text)
        data["Source File"] = pdf_path.name
        data["Patient ID"] = _derive_patient_id(pdf_path)
        return data

    if report_type == "clinical":
        return _empty_record(CLINICAL_REPORT_MESSAGE, pdf_path)

    return _empty_record(UNRECOGNIZED_REPORT_MESSAGE, pdf_path)


def _closest_pair_indices(df: pd.DataFrame, fields: list[str]) -> tuple[int, int] | None:
    if len(df) < 2:
        return None

    min_distance = float("inf")
    closest_pair: tuple[int, int] | None = None

    for i, idx_i in enumerate(df.index[:-1]):
        for idx_j in df.index[i + 1 :]:
            diff = df.loc[idx_i, fields] - df.loc[idx_j, fields]
            distance = (diff.pow(2).sum()) ** 0.5
            if distance < min_distance:
                min_distance = distance
                closest_pair = (idx_i, idx_j)

    return closest_pair


def _average_pair_rows(pair_df: pd.DataFrame, excluded_fields: set[str]) -> dict[str, object]:
    averaged: dict[str, object] = {}
    for column in pair_df.columns:
        if column in excluded_fields:
            continue
        if column == "Patient ID":
            averaged[column] = pair_df[column].iloc[0]
            continue

        numeric_values = pd.to_numeric(pair_df[column], errors="coerce")
        if numeric_values.notna().any():
            averaged[column] = numeric_values.mean()
        else:
            non_null = pair_df[column].dropna()
            averaged[column] = non_null.iloc[0] if not non_null.empty else None

    return averaged


def _build_analyzed_data(df: pd.DataFrame) -> pd.DataFrame:
    analysis_fields = [
        "Peripheral Systolic Pressure (mmHg)",
        "Peripheral Diastolic Pressure (mmHg)",
        "Peripheral Mean Pressure (mmHg)",
    ]

    numeric_df = df.copy()
    for field in analysis_fields:
        numeric_df[field] = pd.to_numeric(numeric_df[field], errors="coerce")

    analyzed_records: list[dict[str, object]] = []
    excluded_fields = {"Source File", "Scanned ID", "Scan Date", "Scan Time"}

    for patient_id, group in numeric_df.groupby("Patient ID"):
        valid_group = group.dropna(subset=analysis_fields)
        pair = _closest_pair_indices(valid_group, analysis_fields)
        if pair is None:
            continue

        pair_df = df.loc[list(pair)]
        averaged_record = _average_pair_rows(pair_df, excluded_fields)
        averaged_record["Patient ID"] = patient_id
        analyzed_records.append(averaged_record)

    return pd.DataFrame(analyzed_records)


def save_to_excel(records: list[dict[str, object]], output_path: Path) -> int:
    df = pd.DataFrame(records, columns=COLUMNS)

    df["Special Row"] = df["Patient ID"].isin(
        {CLINICAL_REPORT_MESSAGE, UNRECOGNIZED_REPORT_MESSAGE}
    )

    df.sort_values(
        by=["Special Row", "Patient ID", "Scan Date", "Scan Time"], inplace=True
    )

    df.drop_duplicates(
        subset=["Patient ID", "Scan Time", "PTI Diastolic (mmHg.s/min)"],
        keep="first",
        inplace=True,
    )

    df["Recording #"] = None
    valid_rows = ~df["Special Row"]
    df.loc[valid_rows, "Recording #"] = (
        df[valid_rows].groupby("Patient ID").cumcount() + 1
    )

    df.drop(columns=["Special Row"], inplace=True)

    analyzed_df = _build_analyzed_data(df)

    with pd.ExcelWriter(output_path) as writer:
        df.to_excel(writer, sheet_name="All Data", index=False)
        analyzed_df.to_excel(writer, sheet_name="Analyzed Data", index=False)

    return len(df)


def main() -> None:
    pdf_paths = select_input_files()
    if not pdf_paths:
        messagebox.showinfo("PWA Data Converter", "No PDF files were selected.")
        return

    output_path = select_output_file()
    if not output_path:
        messagebox.showinfo("PWA Data Converter", "No output location selected.")
        return

    records = [process_pdf(path) for path in pdf_paths]
    exported_count = save_to_excel(records, output_path)
    messagebox.showinfo(
        "PWA Data Converter",
        f"Exported {exported_count} record(s) to {output_path}",
    )


if __name__ == "__main__":
    main()
