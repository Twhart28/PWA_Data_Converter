# PWA Data Converter

Upload one or multiple PDF's of PWA analyses and all data will be exported as an Excel file with one row per PDF.

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

2. Use the file picker to select one or more PWA PDF reports.
3. Choose the save location and filename for the output Excel workbook.
4. The script extracts the following fields for each PDF and writes them as a single row:

- Source File
- Patient ID
- Date of Birth
- Age
- Gender
- Height (m)
- # of Pulses
- Pulse Height
- Pulse Height Variation (%)
- Diastolic Variation (%)
- Shape Deviation (%)
- Pulse Length Variation (%)
- Overall Quality (%)
- Peripheral Systolic Pressure (mmHg)
- Peripheral Diastolic Pressure (mmHg)
- Peripheral Pulse Pressure (mmHg)
- Peripheral Mean Pressure (mmHg)
- Aortic Systolic Pressure (mmHg)
- Aortic Diastolic Pressure (mmHg)
- Aortic Pulse Pressure (mmHg)
- Heart Rate (bpm)
- Pulse Pressure Amplification (%)
- Period (ms)
- Ejection Duration (ms)
- Ejection Duration (%)
- Aortic T2 (ms)
- P1 Height (mmHg)
- Aortic Augmentation (mmHg)
- Aortic AIx AP/PP(%)
- Aortic AIx P2/P1(%)
- Aortic AIx AP/PP @ HR75 (%)
- Buckberg SEVR (%)
- PTI Systolic (mmHg.s/min)
- PTI Diastolic (mmHg.s/min)
- End Systolic Pressure (mmHg)
- MAP Systolic (mmHg)
- MAP Diastolic (mmHg)

The output spreadsheet preserves the order above so that each report is consistently structured.
