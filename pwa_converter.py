import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import pdfplumber


COLUMNS = [
    "Source File",
    "Patient ID",
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


def parse_report_text(text: str) -> dict[str, object]:
    normalized = re.sub(r"\s+", " ", text)

    patient_id = _search(r"Patient ID:\s*(\S+)", normalized)
    dob = _search(r"Date Of Birth:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", normalized)

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
    aortic_augmentation = _search(r"Aortic Augmentation.*?:\s*([0-9.]+)\s*mmHg", normalized)

    aix_match = re.search(r"Aortic AIx \(AP/PP, P2/P1\):\s*([0-9.]+)\s*%,\s*([0-9.]+)\s*%", normalized, flags=re.IGNORECASE)
    aortic_aix_ap_pp = aix_match.group(1) if aix_match else None
    aortic_aix_p2_p1 = aix_match.group(2) if aix_match else None

    aix_hr75 = _search(r"Aortic AIx \(AP/PP\) @HR75:\s*([0-9.]+)\s*%", normalized)
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
        "Patient ID": patient_id,
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
        if isinstance(value, str) and value.replace(".", "", 1).isdigit():
            if "." in value:
                record[key] = float(value)
            else:
                record[key] = int(value)
    return record


def process_pdf(pdf_path: Path) -> dict[str, object]:
    text = extract_text(pdf_path)
    data = parse_report_text(text)
    data["Source File"] = pdf_path.name
    return data


def save_to_excel(records: list[dict[str, object]], output_path: Path) -> None:
    df = pd.DataFrame(records, columns=COLUMNS)
    df.to_excel(output_path, index=False)


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
    save_to_excel(records, output_path)
    messagebox.showinfo("PWA Data Converter", f"Exported {len(records)} record(s) to {output_path}")


if __name__ == "__main__":
    main()
