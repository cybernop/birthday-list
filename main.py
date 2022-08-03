import warnings
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

EXCEL_FILE = "/mnt/c/Users/Alexander Schulze/Charité - Universitätsmedizin Berlin/Core Unit eHealth & Interoperability - Documents/General/Geburtstage.xlsx"

COLUMN_BIRTHDAY = "Geburtstag"
COLUMN_NAME = "Name"
COLUMN_THIS_YEAR = "ThisYear"
COLUMN_AGE = "Age"

DATEFORMAT_ALTERNATIVE = "%d.%m."
OUTPUT_FORMAT = "{date}\t{name}"


def main():
    df = _read_file(EXCEL_FILE)
    df = _preprocess_dates(df)
    birthdays = _get_this_weeks_birthdays(df)
    _print_birthdays(birthdays)


def _read_file(file) -> pd.DataFrame:
    xlsx_file = Path(file)

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        file = pd.ExcelFile(xlsx_file, engine="openpyxl")

    sheet_name = file.sheet_names[0]
    sheet = pd.read_excel(file, sheet_name=sheet_name)

    return sheet


def _preprocess_dates(df: pd.DataFrame) -> pd.DataFrame:
    df.dropna(inplace=True)

    for index, row in df.iterrows():
        birthday = row[COLUMN_BIRTHDAY]
        now = datetime.now()
        if not isinstance(birthday, datetime):
            birthday = datetime.strptime(birthday, DATEFORMAT_ALTERNATIVE).replace(
                year=now.year
            )
            df.loc[index, COLUMN_BIRTHDAY] = birthday

        this_years = birthday.replace(year=now.year)
        df.loc[index, COLUMN_THIS_YEAR] = this_years

        age = this_years.year - birthday.year
        df.loc[index, COLUMN_AGE] = age if age > 0 else None

    return df


def _get_this_weeks_birthdays(df: pd.DataFrame) -> pd.DataFrame:
    now = datetime.now()
    start = now - timedelta(days=now.weekday())
    end = start + timedelta(days=13)

    birthdays = df[df[COLUMN_THIS_YEAR] >= start]
    birthdays = birthdays[birthdays[COLUMN_THIS_YEAR] <= end]

    return birthdays


def _print_birthdays(df: pd.DataFrame) -> None:
    for _, row in df.iterrows():
        name = row[COLUMN_NAME]

        if age := row[COLUMN_AGE]:
            name += f" ({age:.0f})"

        date = row[COLUMN_BIRTHDAY].strftime(DATEFORMAT_ALTERNATIVE)

        print_values = {
            "name": name,
            "date": date,
        }

        print(OUTPUT_FORMAT.format(**print_values))


if __name__ == "__main__":
    main()
