# -*- coding: utf-8 -*-
"""
Created on Thu Nov  4 10:52:57 2021

@author: bhale
"""

from pathlib import Path
from parse_spreadsheets import DataframeParser

import pandas as pd


root_dir = Path(r"\\nsdata4\GSSRPA\Students\CPI Microdata\Product Descriptions")

parsed_files = list((root_dir / "parsed").glob("*"))

files_2017_half_1 = [
    root_dir / "Product_descriptions_201701_Esh_fixed_merged_rows.xlsx",
    root_dir / "Product_descriptions_201702.xlsx",
    root_dir / "Product_descriptions_201703.xlsx",
    root_dir / "Product_descriptions_201704.xlsx",
    root_dir / "Product_descriptions_201705.xlsx",
    root_dir / "Product_descriptions_201706.xlsx"
]
files_2017_half_2 = [f for f in parsed_files if '_2017' in str(f)]

files_2018_half_1 = [
    root_dir / 'parsed' / 'Product_descriptions_201801_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201802_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201803_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201804_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201805_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201806_parsed.xlsx',    
]

files_2018_half_2 = [
    root_dir / 'parsed' / 'Product_descriptions_201807_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201808_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201809_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201810_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201811_parsed.xlsx',
    root_dir / 'parsed' / 'Product_descriptions_201812_parsed.xlsx',    
]

files_2019_half_1 = [f for f in parsed_files if '_2019' in str(f)]
files_2019_half_1.extend([
    root_dir / "Product_descriptions_201903.xlsx",
    root_dir / "Product_descriptions_201904.xlsx",
    root_dir / "product_descriptions_201905.xlsx",
    root_dir / "product_descriptions_201906.xlsx",
])

files_2019_half_2 = [
    root_dir / "product_descriptions_201907.xlsx",
    root_dir / "product_descriptions_201908.xlsx",
    root_dir / "product_descriptions_201909.xlsx",
    root_dir / "product_desc_201910.xlsx",
    root_dir / "product_description_201911.xlsx",
    root_dir / "product_description_201912.xlsx"
]

files_2020_half_1 = [
    root_dir / "product_description_202001.xlsx",
    root_dir / "product_description_202002.xlsx",
    root_dir / "product_description_202003.xlsx",
    root_dir / "product_desc_202004.xlsx",
    root_dir / "Product_description_202005.xlsx",
    root_dir / "product_description_202006.xlsx"
]
files_2020_half_2 = [
    root_dir / "product_description_202007.xlsx",
    root_dir / "product_desc_202008.xlsx",
    root_dir / "product_desc_202009.xlsx",
    root_dir / "product_desc_202010.xlsx",
    root_dir / "product_desc_202011.xlsx",
    root_dir / "product_desc_202012.xlsx",
]


def load_transition_df(filename):
    map_from = "KEY"
    map_to = "Shop Code"
    print("Loading transition workbook...")
    df_transition = pd.read_excel(
        filename,
        sheet_name="Transition Table", columns=[map_from, map_to]
    )
    return df_transition


def combine_files(files_list, month_start, month_end, run_shop_code_transition=False):
    dfs = []
    for f in files_list:
        print(f"Loading {f}")
        dfs.append(pd.read_excel(f))

    print("Concatenating dataframes")
    final_df = pd.concat(dfs)

    print("Saving concatenated dataframe")
    out_file = root_dir / f"concat_product_descriptions_{month_start}_{month_end}.xlsx"
    with pd.ExcelWriter(
            out_file, engine='xlsxwriter',
            options={'strings_to_urls': False, 'strings_to_formulas': False}
    ) as writer:
        final_df.to_excel(writer, index=False)
    print(f"Written {out_file}")

    if run_shop_code_transition:
        print("Running shop code transition")
        df_parsing = DataframeParser(
            filename=out_file,
            out_dir=root_dir,
            df_transition=df_transition,
            transition_col_map_from="KEY",
            transition_col_map_to="Shop Code",
            file_out_suffix="",
            overwrite=True
        )
        df_parsing.df = final_df
        print("Running shop code transition")
        df_parsing.run_shop_code_transition()
        df_parsing.write_df()


df_transition = load_transition_df(root_dir / "Product_descriptions_201707_RowanEdited.xlsx")

combine_files(files_2017_half_1, '201701', '201706', run_shop_code_transition=True)
combine_files(files_2017_half_2, '201707', '201712')
combine_files(files_2018_half_1, '201801', '201806')
combine_files(files_2018_half_2, '201807', '201812')
combine_files(files_2019_half_1, '201901', '201906')
combine_files(files_2019_half_2, '201907', '201912')
combine_files(files_2020_half_1, '202001', '202006')
combine_files(files_2020_half_2, '202007', '202012')
