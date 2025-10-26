
# Invoice Annotation & Extraction Tool

This is a complete offline tool for annotating invoice fields, training a custom NLP model (spaCy), and extracting structured data from new PDF invoices directly into Excel.

## Features
- Annotate invoice PDFs with bounding boxes for fields like Invoice Number, Date, Total Amount, Company, etc.
- Export annotations as training data for a spaCy Named Entity Recognition (NER) model.
- Train the model locally (no internet or API calls).
- Extract data from new invoices using the trained model.
- Export all extracted data to Excel.

## Requirements
- Python 3.9+
- PyQt5
- PyMuPDF (fitz)
- spaCy
- pandas
- openpyxl

Install dependencies with:
```
pip install -r requirements.txt
```

## How to Use

### 1. Annotate Invoices
- Open the app:
```
python app.py
```
- Click **Open PDF** and load your invoice.
- Add labels (Invoice number, Date, etc.) on the right panel.
- Select a label and draw boxes around corresponding text on the invoice.
- Save annotations when done.

### 2. Export Training Data
- Click **Export Training Data** to create a `.jsonl` training file.

### 3. Train Model
- Click **Train Model** and select your exported `.jsonl` file.
- Model will train locally and be saved to `/models` directory.

### 4. Run Extraction
- Click **Run Extraction**, select a folder containing PDF invoices.
- The tool will process them and prompt to save an Excel file with extracted fields.

## Folder Structure
```
project/
│
├── app.py                # Main PyQt app file
├── models/               # Trained spaCy model folder
├── requirements.txt
└── annotations.json      # Saved annotations (optional)
```

## Notes
- Runs fully offline.
- Ensure you have 3–5 labeled samples per field for reliable accuracy.
- Works best with English invoices.

## License
Free to use and modify for personal or commercial projects.
