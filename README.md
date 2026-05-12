# Protection Mainstreaming KOI Calculator

A Streamlit app for calculating the Protection Mainstreaming KOI indicator from respondent-level Excel datasets.

The app supports flexible column mapping, response mapping, disaggregation, donor/project filtering, validation checks, and Excel result export.

## Privacy

Use anonymized datasets only.

Before uploading a file, remove personally identifiable data such as:

- names
- phone numbers
- addresses
- ID numbers
- case numbers or other direct identifiers
- any free-text fields that may reveal a person

The uploaded file should contain only the variables needed for the calculation.

## Calculation Method

For each KOI question:

```text
Question result = Positive / (Positive + Negative + Don't know)
```

`No answer`, blank values, and unmapped responses are excluded from the denominator.

The app calculates the four elements:

- `SDH`
- `MEA`
- `ACC`
- `PEM`

Each element is the average of its two question results. The final PM KOI result is the average of the four element scores.

For `MEA. 2`, the standard response set is reversed: `Not really` and `Not at all` are counted as the favourable/positive outcome.

The `People with >=1 negative KOI answer` metric is a diagnostic count. It is not the inverse of the final PM KOI score. A person is counted once if they have any negative answer across the 8 KOI questions, while the final score is calculated from question and element percentages.

## Input File

Upload an anonymized `.xlsx` or `.xlsm` respondent-level dataset.

Each row should represent one respondent/person. The app requires columns for:

- age, unless age disaggregation is disabled
- the 8 KOI question responses:
  - `SDH. 1`
  - `SDH. 2`
  - `MEA. 1`
  - `MEA. 2`
  - `ACC. 1`
  - `ACC. 2`
  - `PEM. 1`
  - `PEM. 2`

Optional columns:

- sex/gender
- disability yes/no
- donor/project filter

The app includes downloadable blank and sample Excel files.

The app shows a short reminder next to each KOI code, for example `SDH. 1 - Safety while accessing assistance`.

## Run Locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the app:

```bash
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Upload this folder to a GitHub repository.
2. Make sure the repository contains:
   - `app.py`
   - `requirements.txt`
   - `README.md`
   - `.streamlit/config.toml`
3. In Streamlit Community Cloud, create a new app from the repository.
4. Set the main file path to:

```text
app.py
```

## Release Checks

Before sharing results, review:

- the mapping table before calculation
- warnings about unmapped or blank responses
- warnings about missing response categories
- the Excel export `Read_me`, `Calculation`, `Summary`, `Mapping`, and `Unmapped responses` sheets

## Troubleshooting

If the app cannot read a file, confirm it is `.xlsx` or `.xlsm`.

If the calculation button is disabled, review the validation message and fix the required mapping issue.

If the result looks too high or too low, check that each response option is mapped to the correct category, especially `MEA. 2` and `ACC. 1`.
