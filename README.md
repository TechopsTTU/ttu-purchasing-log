# TTU Purchase Orders Log

This Streamlit application analyzes purchase order data from an uploaded Excel file. It filters and aggregates the dataset to provide metrics such as:

- **Total Open Orders Amount** – sum of `Amt` where `POStatus` is `OPEN`.
- **Total Orders Placed** – number of unique purchase orders.
- **Total Lines Ordered** – total row count after filtering.
- **Most Expensive Order** – highest `Amt` value.
- **On‑time Percentage** – `(on_time / total) * 100` comparing requested and received dates.

Outliers are removed from charts using the Interquartile Range (IQR) method:

```
Q1 = df[column].quantile(0.25)
Q3 = df[column].quantile(0.75)
IQR = Q3 - Q1
lower = Q1 - 1.5 * IQR
upper = Q3 + 1.5 * IQR
filtered_df = df[(df[column] >= lower) & (df[column] <= upper)]
```

## Usage

1. Install dependencies:

```bash
pip install -r requirements.txt
```
The application requires the `kaleido` package for exporting Plotly figures.
It is included in `requirements.txt`, but if you encounter errors related to
image export, ensure Kaleido is installed manually with:

```bash
pip install kaleido
```

2. Launch the app:

```bash
streamlit run app.py
```

Upload an Excel purchase order log to explore the metrics and generate a PDF report.

## License

This project is licensed under the [MIT License](LICENSE).
