# PWA Data Converter

**Contact:** thomaswhart28@gmail.com

Convert one or many PWA analysis PDFs into a structured Excel workbook. The app guides you through a short startup dialog, lets you review the README in-app, and walks you through selecting files, choosing the output location, and reviewing multi-entry patients before exporting.

## Prerequisites

Install the dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Run the converter:

```bash
python pwa_converter.py
```

2. On startup, use **Read Me** to view this guide in-app or choose **Continue** to proceed.
3. Use the file picker to select one or more PWA PDF reports.
4. Choose where to save the Excel workbook; a timestamped filename is suggested for convenience.
5. Watch the progress dialog as files are analyzed. If a patient has more than two valid entries, you can accept the automatic pairing or launch the **Manual Overview** to pick which records to average.
6. When the export completes, you will see a confirmation with the destination path.

## Key Features

- **In-app README access:** The startup popup offers a **Read Me** button and the window close button behaves the same as clicking **Close** inside the README viewer.
- **PDF validation and processing:** Each PWA Detailed Report is parsed for patient demographics, hemodynamic measurements, waveform quality metrics, and timing values.
- **Automatic and manual record pairing:** Patients with multiple entries are averaged using the configured analysis mode, with an optional manual review step to override the automatic pairing.
- **Organized Excel output:** Three sheets are created—**All Data** (full dataset), **Kept Data** (records included in analysis), and **Averaged Data** (per-patient averages). Dates are normalized, headers and key identifiers are aligned for readability, and analysis flags mark which rows were used.
- **Convenient defaults:** Suggested filenames include timestamps, and progress dialogs keep you informed during analysis and export.

## Extracted Fields

The export preserves a consistent order of fields per PDF, including:

- Source details: Source File and Source Path
- Patient demographics: Patient ID, Date of Birth, Age, Gender, Height (m)
- Signal quality: # of Pulses, Pulse Height, Pulse Height Variation (%), Diastolic Variation (%), Shape Deviation (%), Pulse Length Variation (%), Overall Quality (%)
- Peripheral values: Peripheral Systolic/Diastolic/Mean Pressure (mmHg), Peripheral Pulse Pressure (mmHg)
- Aortic values: Aortic Systolic/Diastolic/Pulse Pressure (mmHg), Aortic AIx (AP/PP, P2/P1, AP/PP @ HR75), Aortic Augmentation (mmHg)
- Cardiac timing: Heart Rate (bpm), Period (ms), Ejection Duration (ms/%), Aortic T2 (ms)
- Additional metrics: Pulse Pressure Amplification (%), P1 Height (mmHg), Buckberg SEVR (%), PTI Systolic/Diastolic (mmHg.s/min), End Systolic Pressure (mmHg), MAP Systolic/Diastolic (mmHg)

The workbook keeps this order across all sheets so downstream analysis remains predictable.
