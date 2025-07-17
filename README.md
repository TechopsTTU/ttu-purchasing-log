# TTU Purchase Orders Log

TTU Purchase Orders Log is a web application built with [Streamlit](https://streamlit.io/) for analyzing purchase order data stored in Excel workbooks. After uploading a log, the tool aggregates the data and displays interactive charts and summary metrics. It can also generate a PDF report containing the plots and tables.

## Features

- Summarizes open order amounts, total orders placed, lines ordered, and most expensive order
- Calculates on‑time delivery percentage by comparing requested and received dates
- Removes outliers from charts using the Interquartile Range (IQR) method
- Interactive graphs powered by Plotly
- Option to export analysis results to a PDF report

## Requirements

- Python 3.9+
- Packages listed in [`requirements.txt`](requirements.txt), including:
  - `pandas`
  - `openpyxl`
  - `plotly`
  - `streamlit`

Install the dependencies with:

```bash
pip install -r requirements.txt
```

## Usage

Run the application with Streamlit:

```bash
streamlit run app.py
```

Upload a purchase order Excel file when prompted. The app will display metrics, charts, and provide an option to download a PDF report.

## License

This project is licensed under the [MIT License](LICENSE).
