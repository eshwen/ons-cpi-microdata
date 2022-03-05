# -*- coding: utf-8 -*-
"""Parse spreadsheets of product descriptions and arrange values accordingly."""

__author__ = "eshwen.bhal@ext.ons.gov.uk"
__version__ = "v1.0"

import datetime
from argparse import ArgumentDefaultsHelpFormatter, ArgumentParser
from pathlib import Path

import pandas as pd


class DataframeParser(object):
    """Load and parse dataframe."""

    def __init__(self, filename, out_dir,
                 df_transition, transition_col_map_from, transition_col_map_to,
                 file_out_suffix="_parsed", overwrite=False):
        """Initialise class.

        Attributes:
            filename (pathlib.Path): Path to file.
            out_dir (pathlib.Path): Output directory to write parsed dataframe to.
            df_transition (pandas.DataFrame): Dataframe to map shop codes.
            map_from (str): Column name in self.df_transition to map values from.
            map_to (str): Column name in self.df_transition to map values to.
            file_out_suffix (str, optional): Substring to append to output file.
            overwrite (bool, optional): Overwrite the output files if it already exists.

        """
        self.filename = filename
        self.out_dir = Path(out_dir)
        self.df_transition = df_transition
        self.map_from = transition_col_map_from
        self.map_to = transition_col_map_to
        self.file_out_suffix = file_out_suffix
        self.overwrite = overwrite

        self.df = None
        self.new_col_name = None

    def run(self):
        """Run everything.

        Attributes:
            df (pandas.DataFrame): Dataframe loaded from self.filename.

        """
        self.df = self.load_df(header=None)
        self.parse_df()
        self.run_shop_code_transition()
        self.write_df()

    def load_df(self, **kwargs):
        """Load a dataframe from a file.

        Supports reading Excel spreadsheets by inference from the file extension.
        Otherwise assumes it's a CSV.

        Args:
            kwargs: Arbitrary keyword to be passed to pandas' read_excel/read_csv.

        Returns:
            pandas.DataFrame: Loaded dataframe.

        """
        print(f"Loading input file: {self.filename}")
        if self.is_excel_file:
            return pd.read_excel(self.filename, **kwargs)
        else:
            return pd.read_csv(self.filename, **kwargs)

    def parse_df(self):
        """Extract information from the dataframe and sort into more organised form."""
        print("Parsing dataframe...")
        df_parsed = pd.DataFrame(columns=list(self.column_to_str_index_mapping.keys()))
        raw_str = self.df[0].str  # df only has one column, so use .str accessor on it for vectorised ops

        for column, str_slice in self.column_to_str_index_mapping.items():
            sliced_str = raw_str[str_slice[0]:str_slice[1]]
            df_parsed[column] = sliced_str.str.strip()

        self.df = df_parsed

    def run_shop_code_transition(self):
        """Map incorrectly-stored shop codes to their correct values, if possible.
        
        Use the original, incorrect shop code in the cases it's not possible.

        Attributes:
            new_col_name (str): The new column in self.df to store correct shop codes.

        """
        self.new_col_name = "actual shop"

        self.df[self.map_from] = self.item_key_col

        self._drop_duplicates_on_character_len()
        self._cast_columns(self.map_from)
        self._map_shop_codes(self.map_from, self.map_to)
        self._backfill_unmatched_values()
        self.df.drop(columns=[self.map_from], inplace=True)

    @property
    def item_key_col(self):
        return self.df["shop_code"] + self.df["location"] + self.df["item_id"]

    def _drop_duplicates_on_character_len(self, shop_code_len=4):
        """Drop all duplicates from the self.map_from column where the shop_code is 4 characters."""
        duplicates = self.df.duplicated(subset=self.map_from, keep=False)
        shop_code_is_4_chars = self.df["shop_code"].astype(str).str.len() == shop_code_len
        mask = ~(duplicates & shop_code_is_4_chars)
        self.df = self.df[mask].reset_index(drop=True)

    def _cast_columns(self, col, dtype=str):
        """Cast columns to strings for robust matching."""
        self.df[col] = self.df[col].astype(dtype)
        self.df_transition[col] = self.df_transition[col].astype(dtype)

    def _map_shop_codes(self, map_from, map_to):
        """Map values from 'map_from' column onto 'map_to' column of df, matching on values from 'map_from' column.

        Slow, but most robust way I can find.

        """
        self.df[self.new_col_name] = self.df[map_from]
        self.df[self.new_col_name].replace(
            to_replace=self.df_transition[map_from].values,
            value=self.df_transition[map_to].values,
            inplace=True
        )

    def _backfill_unmatched_values(self, fill_unmatched_with="shop_code"):
        """Backfill unmatched values with old values.

        Args:
            fill_unmatched_with (str): Column name in self.df to backfill with unmappable values.

        """
        mask = self.df[self.new_col_name].isin(self.df[self.map_from])
        self.df.loc[mask, self.new_col_name] = self.df.loc[mask, fill_unmatched_with]

    def write_df(self):
        """Write the output dataframe to a file.

        As with load_df(), the output is saved in Excel format if inferred from the file extension. Otherwise as a CSV.

        """
        if not self.out_file.parent.exists():
            print(f"Output directory {self.out_file.parent} does not exist. Creating...")
            self.out_file.parent.mkdir(parents=True)

        if self.out_file.exists() and not self.overwrite:
            print(f"""The output file {self.out_file} already exists and `force` is False.
Either rerun with the `force` option or select a different output filename.""")
            return

        print(f"Writing output file: {self.out_file}")
        if self.is_excel_file:
            # Hook into excel writer class to prevent issues when reading URLs and formulas
            with pd.ExcelWriter(self.out_file, engine='xlsxwriter',
                options={'strings_to_urls': False, 'strings_to_formulas': False}) as writer:
                self.df.to_excel(writer, index=False)
        else:
            self.df.to_csv(self.out_file, index=False)

    @property
    def is_excel_file(self):
        """Determine whether a file is in Excel format.

        Returns:
            bool: True if the file is in Excel format, False if not.

        """
        return any(self.filename.suffix == ext for ext in [".xls", ".xlsx"])

    @property
    def out_file(self):
        """Return output filename."""
        return self.out_dir / (self.filename.stem + self.file_out_suffix + self.filename.suffix)

    @property
    def column_to_str_index_mapping(self):
        """Return dictionary mapping the string index ranges to a column header.

        Note:
            key: value as <column>: [<index to slice from>, <index to slice to>].
        
        """
        return {
        "quote_date": [0, 6],
        "item_id": [6, 12],
        "location": [12, 17],
        "shop_code": [17, 21],
        "prod_size": [21, 37],
        "prod_measure_id": [37, 39],
        "attribute2": [39, 119],
        "attribute3": [119, 199],
        "attribute4": [199, 279],
        "attribute5": [279, 359],
        "attribute6": [359, 439]
    }


def find_files(dir_in, file_pattern, date_start, date_stop):
    """Search for requested files.

    Args:
        dir_in (pathlib.Path): Directory containing files to parse.
        file_pattern (str): File search pattern within 'dir_in' (non-recursive).
        date_start (datetime.date): Date to parse from.
        date_stop (datetime.date): Date to parse to.

    Returns:
        list of pathlib.Path: Paths to files.

    """
    files_searched = dir_in.glob(file_pattern)
    files_to_clean = []
    for f in files_searched:
        date = f.stem.split("_")[-1]
        
        # Skip over dates that can't be formatted
        try:
            date_fmt = datetime.datetime.strptime(date, "%Y%m")
        except Exception:
            continue

        if date_fmt >= date_start and date_fmt <= date_stop:
            files_to_clean.append(f)

    return files_to_clean


def parse_arguments():
    """Parse CLI arguments.

    Returns:
        argparse.Namespace: Parsed arguments accessible as object attributes.

    """
    parser = ArgumentParser(description=__doc__, formatter_class=ArgumentDefaultsHelpFormatter)
    parser.add_argument(
        "-i", "--dir_in",
        type=str, default=r"\\nsdata4\GSSRPA\Students\CPI Microdata\Product Descriptions",
        help="Directory containing files to parse."
    )
    parser.add_argument(
        "-o", "--dir_out",
        type=str, default="parsed",
        help="Subdirectory under 'dir_in' to store output files."
    )
    parser.add_argument(
        "-p", "--file_pattern",
        type=str, default="Product_descriptions_*.xlsx",
        help="File search pattern within 'dir_in' (non-recursive)."
    )
    parser.add_argument(
        "-s", "--file_out_suffix",
        type=str, default="_parsed",
        help="Substring to append to output files."
    )
    parser.add_argument(
        "--date_start",
        type=str, default="201707",
        help="Date to parse from in YYYYMM format."
    )
    parser.add_argument(
        "--date_stop",
        type=str, default="201902",
        help="Date to parse to in YYYYMM format."
    )
    parser.add_argument(
        "-t", "--transition_workbook",
        type=str, default="Product_descriptions_201707_RowanEdited.xlsx",
        help="Workbook containing table to transition shop codes."
    )
    parser.add_argument(
        "-f", "--force",
        action="store_true",
        help="Overwrite the output files if they already exist."
    )

    args = parser.parse_args()
    return args


def main(dir_in, dir_out, file_pattern, date_start, date_stop, transition_workbook,
         file_out_suffix="_parsed", force=False):
    """Runner of the script.

    Args:
        dir_in (str): Directory containing files to parse.
        dir_out (str): Subdirectory under 'dir_in' to store output files.
        file_pattern (str): File search pattern within 'dir_in' (non-recursive).
        date_start (str): Date to parse from in YYYYMM format.
        date_stop (str): Date to parse to in YYYYMM format.
        transition_workbook (str): Workbook containing table to transition shop codes.
        file_out_suffix (str, optional): Substring to append to output files.
        force (bool, optional): Overwrite the output files if they already exist.

    """
    dir_in = Path(dir_in)
    dir_out = dir_in / dir_out
    date_start = datetime.datetime.strptime(date_start, "%Y%m")
    date_stop = datetime.datetime.strptime(date_stop, "%Y%m")
    transition_workbook = dir_in / transition_workbook

    files_to_clean = find_files(dir_in, file_pattern, date_start, date_stop)

    map_from = "KEY"
    map_to = "Shop Code"
    df_transition = pd.read_excel(
        transition_workbook,
        sheet_name="Transition Table", columns=[map_from, map_to]
    )

    for f in files_to_clean:
        df = DataframeParser(
            filename=f, out_dir=dir_out,
            df_transition=df_transition, transition_col_map_from=map_from, transition_col_map_to=map_to,
            file_out_suffix=file_out_suffix, overwrite=force
        )
        df.run()
    print("Completed")


if __name__ == "__main__":
    args = parse_arguments()
    main(**args.__dict__)
