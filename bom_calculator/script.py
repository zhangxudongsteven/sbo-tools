
import datetime
from bom_calculator.bom_utils import update_usd_ratio, delete_origin_output_if_exists
from bom_calculator.bom_details_processing import calculate_bom_details
from bom_calculator.bom_summary_processing import calculate_bom_summary
from bom_calculator.bom_compare_processing import compare_bom_cost
from bom_calculator.combine_excel_processing import combine_csv_to_excel


USD_RATIO = 6.7
INPUT_FILE_PATH = "input.xlsx"
INPUT_BOM_SHEET = "Sheet1"
INPUT_COST_SHEET = "Sheet2"
MIDDLE_FILE_PATH_BOM = "tmp_1.csv"
MIDDLE_FILE_PATH_BOM_SUMMARY = "tmp_2.csv"
MIDDLE_FILE_PATH_BOM_COST_COMPARE = "tmp_3.csv"
OUTPUT_FILE_PATH = "output.xlsx"
OUTPUT_FILE_SHEET_BOM = "BOM明细"
OUTPUT_FILE_SHEET_BOM_SUMMARY = "BOM汇总"
OUTPUT_FILE_SHEET_BOM_COST_COMPARE = "仓库BOM成本匹配"


if __name__ == '__main__':
    print("Process Start")
    print(str(datetime.datetime.now()))
    print("****************")

    USD_RATIO = update_usd_ratio(USD_RATIO)
    print("****************")

    print("Start to Calculate BOM")
    calculate_bom_details(INPUT_FILE_PATH, INPUT_BOM_SHEET, MIDDLE_FILE_PATH_BOM, USD_RATIO)
    print("Finished")
    print("****************")

    print("Start to Calculate BOM Summary")
    calculate_bom_summary(MIDDLE_FILE_PATH_BOM, MIDDLE_FILE_PATH_BOM_SUMMARY)
    print("Finished")
    print("****************")

    print("Start to Combine Warehouse Item Cost")
    compare_bom_cost(INPUT_FILE_PATH, INPUT_COST_SHEET, MIDDLE_FILE_PATH_BOM_SUMMARY, MIDDLE_FILE_PATH_BOM_COST_COMPARE)
    print("Finished")
    print("****************")

    print("Start to Transfer CSV to Excel")
    combine_csv_to_excel(MIDDLE_FILE_PATH_BOM,
                         MIDDLE_FILE_PATH_BOM_SUMMARY,
                         MIDDLE_FILE_PATH_BOM_COST_COMPARE,
                         OUTPUT_FILE_PATH,
                         OUTPUT_FILE_SHEET_BOM,
                         OUTPUT_FILE_SHEET_BOM_SUMMARY,
                         OUTPUT_FILE_SHEET_BOM_COST_COMPARE)
    print("Process Finished")
    print("****************")

    print("删除文件【{0}】".format(MIDDLE_FILE_PATH_BOM))
    delete_origin_output_if_exists(MIDDLE_FILE_PATH_BOM)
    print("删除文件【{0}】".format(MIDDLE_FILE_PATH_BOM_SUMMARY))
    delete_origin_output_if_exists(MIDDLE_FILE_PATH_BOM_SUMMARY)
    print("删除文件【{0}】".format(MIDDLE_FILE_PATH_BOM_COST_COMPARE))
    delete_origin_output_if_exists(MIDDLE_FILE_PATH_BOM_COST_COMPARE)
    print("Process Finished")
    print("****************")

    input("Press Enter to continue...")
