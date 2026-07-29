# BrailleSpeakGR Benchmark Analyzer v2.0

## Files

- `benchmark_analyzer_v2.py`: professional benchmark engine.
- `gui_benchmark_analyzer_v2.py`: graphical interface.
- `Benchmark_Master_BrailleSpeakGR.csv`: input dataset.
- `01_pharmaceutical_label.jpg` ... `10_vitamin_box.jpg`: the ten benchmark
  document images used as OCR input for the four evaluated systems.

## Note on Image 10

Image 10 ("Vitamin box") uses a representative vitamin supplement product
box, photographed by the author, matching the general "vitamin box" category
used in the benchmark. It is provided as a representative substitute for
this category.

## Run with GUI

```bash
python gui_benchmark_analyzer_v2.py
```

Then press:

`RUN BENCHMARK v2.0`

## Run without GUI

```bash
python benchmark_analyzer_v2.py Benchmark_Master_BrailleSpeakGR.csv
```

## Outputs

The tool creates:

```text
benchmark_outputs_v2/
  01_summary_tables/
    system_summary.csv
    category_summary.csv
    image_level_summary.csv
    ranking_table.csv

  02_article_text/
    article_main_paragraph.txt
    article_ranking_paragraph.txt
    article_limitations_paragraph.txt

  03_figures/
    01_ocr_success_rate.png
    02_total_known_errors.png
    ...
    16_category_error_matrix.png

  04_reports/
    benchmark_report.docx
    benchmark_report.pdf

  05_excel/
    benchmark_report.xlsx

  06_json/
    benchmark_results.json
```

## Optional libraries

For full output generation:

```bash
pip install matplotlib python-docx openpyxl
```

If a library is missing, the script still produces CSV/TXT/JSON outputs.
