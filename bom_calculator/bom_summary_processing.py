
import pandas as pd


def calculate_bom_summary(input_file_name, output_file_name):
    try:
        df = pd.read_csv(input_file_name)
        dfr = df[df.supply_method == '购买']\
            .groupby(['top', 'top_name', 'top_type'])['total_price_rmb']\
            .sum().reset_index()
        dfr.to_csv(output_file_name, index=False)
    except FileNotFoundError as e:
        print("{0} not found: {1}".format(input_file_name, e))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))
