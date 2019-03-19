
import pandas as pd


def compare_bom_cost(excel_name, sheet_name, bom_summary_file_name, output_file_name):
    try:
        df = pd.read_excel(excel_name, sheet_name)
        df_cost = pd.read_csv(bom_summary_file_name)
        dfr = pd.merge(left=df, right=df_cost, left_on='物料编号', right_on='top', suffixes=('_l', '_r'), how='left')
        dfr.loc[pd.isna(dfr.total_price_rmb), 'total_price_rmb'] = \
            dfr.loc[pd.isna(dfr.total_price_rmb), '含税人民币价格'] / 1.16
        dfr = dfr.drop(columns=['top', 'top_name', 'top_type', '价格差异', '总成本差异', '价格差异比率'])
        dfr['价格差异'] = dfr.apply(lambda row: row.加权平均成本 - row.total_price_rmb, axis=1)
        dfr['价格差异比率'] = dfr.apply(lambda row: row.价格差异 / row.total_price_rmb if row.total_price_rmb != 0 else 0, axis=1)
        dfr = dfr.rename(index=str, columns={"total_price_rmb": "未税人民币价格"})
        dfr.to_csv(output_file_name, index=False)
    except FileNotFoundError as e:
        print("{0} not found: {1}".format(excel_name, e))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))
