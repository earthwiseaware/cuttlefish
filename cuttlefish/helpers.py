import json
import os

class SheetHelper(object):
    SHEET_NAME = 'sheet'

    def __init__(self, base_obj, read_func, write_func):
        self.obj = base_obj
        self.read_func = read_func
        self.write_func = write_func

    @classmethod
    def _get_sheet(cls, workbook):
        return workbook[cls.SHEET_NAME]

    @classmethod
    def _write_json(cls, folder, obj):
        with open(os.path.join(folder, cls.SHEET_NAME + '.json'), 'w') as fh:
            json.dump(obj, fh, sort_keys=True, indent=4)

    @classmethod
    def _read_json(cls, folder):
        with open(os.path.join(folder, cls.SHEET_NAME + '.json'), 'r') as fh:
            return json.load(fh)

    def read_json(self, folder):
        self.obj = self._read_json(folder)

    def write_json(self, folder):
        self._write_json(folder, self.obj)

    def read_sheet(self, workbook):
        sheet = self._get_sheet(workbook)
        self.obj = self.read_func(sheet)
                
    def write_sheet(self, workbook):
        sheet = self._get_sheet(workbook)
        self.write_func(sheet, self.obj)

def get_columns(sheet):
    return {
        i: cell.value.strip() for i, cell in enumerate(sheet[1])
    }

def write_survey(sheet, survey, columns=None, next_row=2):
    if not columns:
        columns = {}
    for element in survey:
        for column, value in element.items():
            if column == 'survey':
                continue
            if column not in columns:
                columns[column] = max(columns.values()) + 1 if columns else 1
                sheet.cell(row=1, column=columns[column], value=column)
            sheet.cell(row=next_row, column=columns[column], value=value)
            
        type_definitions = [
            e.strip() for e in element['type'].split(' ') if e.strip() != ''
        ]
        if type_definitions[0] == 'begin':
            next_row = write_survey(sheet, element['survey'], columns, next_row=next_row+1)
            sheet.cell(
                row=next_row, column=columns['type'], 
                value=' '.join(['end'] + type_definitions[1:])
            )
        next_row += 1
    return next_row

def add_survey_element(obj, keys, value):
    if not keys:
        obj.append(value)
        return len(obj) - 1
    return add_survey_element(obj.get(keys[0], []), keys[1:], value)

def read_survey(sheet):
    survey = {}
    current_survey_keys = []
    columns = get_columns(sheet)
    for row in sheet.iter_rows(min_row=2):
        element = {
            columns[i]: cell.value.strip()
            for i, cell in enumerate(row)
        }
        type_definitions = [
            e.strip() for e in element['type'].split(' ') if e.strip() != ''
        ]
        if not type_definitions:
            continue
        if type_definitions[0] == 'end':
            current_survey_keys = current_survey_keys[:-2]
            continue

        index_of_element = add_survey_element(survey, current_survey_keys, element)
        if type_definitions[0] == 'begin':
            current_survey_keys.append(index_of_element)
            current_survey_keys.append('survey')
    return survey

class SurveyHelper(SheetHelper):
    SHEET_NAME = 'survey'

    def __init__(self):
        super().__init__([], read_survey, write_survey)

def write_choices(sheet, choices):
    row = 1
    sheet.cell(row=row, column=1, value='list_name')
    sheet.cell(row=row, column=2, value='name')
    sheet.cell(row=row, column=3, value='label')
    for key, options in choices.items():
        for choice, label in options.items():
            row += 1
            for i, value in enumerate([key, choice, label]):
                sheet.cell(row=row, column=i+1, value=value)

def read_choices(sheet):
    choices = {}
    columns = get_columns(sheet)
    for row in sheet.iter_rows(min_row=2):
        key = row[columns['list_name']].strip()
        if not key:
            continue
        if key not in choices:
            choices[key] = {}
        choices[key][row[columns['name']].strip()] = row[columns['label']].strip()
    return choices

class ChoicesHelper(SheetHelper):
    SHEET_NAME = 'choices'

    def __init__(self):
        super().__init__({}, read_choices, write_choices)

def write_settings(sheet, settings):
    for column, (key, value) in enumerate(settings.items()):
        sheet.cell(row=1, column=column+1, value=key)
        sheet.cell(row=2, column=column+1, value=value)

def read_settings(sheet):
    settings = {}
    for column in sheet.columns:
        key = column[0].strip()
        if not key:
            continue
        settings[key] = next(e.strip() for e in column[1:] if e.strip())
    return settings

class SettingsHelper(SheetHelper):
    SHEET_NAME = 'settings'

    def __init__(self):
        super().__init__({}, read_settings, write_settings)
