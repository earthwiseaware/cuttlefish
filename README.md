# xlsform
Excel sheets are hard to track as they count as binary files in GIT. This repo gives you a tool for building XLSForms from JSON and digesting XLSForms into JSON so that you can track them properly. 

## Installing
```bash
pip install .
```

## Creating an XLSForm
```bash
xlsform -m create -w examples/created_example.xlsx -f examples/example_json
```

## Digesting an XLSForm
```bash
xlsform -m digest -w examples/example_survey.xlsx -f examples/digested_example_json
```

## Running Tests
```bash
pytest
```
