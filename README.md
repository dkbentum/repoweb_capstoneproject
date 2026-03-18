# Report Card Generator (Single-Page + Python Backend)

## Features
- Single long-scroll frontend (`HTML/CSS/JS`)
- Downloadable Excel template for subject marks
- Upload multiple subject files (`.xlsx`, `.xls`, `.csv`)
- School details customization
- User-defined weighting (`ClassScore%` and `ExamScore%`)
- Backend weighted-score calculation per subject
- Backend ranking (overall position across subjects and per-subject position)
- PDF template customization (3 structures):
  - `Classic Board`
  - `Modern Split`
  - `Minimal Ledger`
- Near-infinite color customization:
  - custom `primary`, `secondary`, and `surface` colors via color pickers + hex values
  - optional random palette generator
- School logo upload and display style options:
  - `Header Badge`
  - `Watermark` (adjustable strength)
  - `Corner Stamp`
  - `Side Mark`
- Per-student PDF report cards
- One compiled PDF containing all students
- ZIP download containing compiled PDF + all individual PDFs

## Expected Subject File Format
Use the template from the app (`Download Excel Template`).

Required columns in each subject file:
- `student ID`
- `Name`
- `ClassScore`
- `ExamScore`

Each file should represent one subject. The file name becomes the subject name in generated report cards.
The frontend requires `ClassScore% + ExamScore% = 100`, and the backend computes:
`weighted final = (ClassScore * ClassScore% / 100) + (ExamScore * ExamScore% / 100)`.

## Run Locally
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open: `http://127.0.0.1:5000`