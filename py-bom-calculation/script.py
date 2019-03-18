
import pandas as pd
import copy
import datetime
import os


build_log = {}
type_map = {}
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


def create_node_from_row(r):
    # 物料编码	物料名称	原厂料号	补给方式	首选供应商	物料分类二	物料分类三	计量单位	上级物料编码	上级物料名称	总消耗量	损耗量	标准损耗率	采购价格	采购货币
    total_price_rmb = r.采购价格 * r.总消耗量 * USD_RATIO if r.采购货币 == 'USD' else r.采购价格 * r.总消耗量
    return {
        "item_code": r.物料编码,
        "item_name": r.物料名称,
        "item_group": r.物料组,
        "level": 1,
        "total_used": r.总消耗量,
        "total_loss": r.损耗量,
        "standard_used": r.标准消耗量,
        "price": r.采购价格,
        "currency": r.采购货币,
        "total_price_rmb": total_price_rmb,
        "origin_item_code": r.原厂料号,
        "supply_method": r.补给方式,
        "first_supplier": r.首选供应商,
        "class2": r.物料分类二,
        "class3": r.物料分类三,
        "unit": r.计量单位,
        "mother_item_code": r.上级物料编码,
        "mother_item_name": r.上级物料名称,
        "mother_item_group": r.上级物料组,
        "loss_ratio": r.损耗量 / r.标准消耗量,
        "standard_loss_ratio": r.标准损耗率
    }


def get_str(a):
    if type(a) == str:
        if '"' in a:
            a = str(a).replace('"', "")
        return '"' + a + '"'
    else:
        return str(a)


def iter_level(top, f, bom, key, level, multi):
    for row in bom[key]:
        obj = copy.copy(row)
        obj["level"] += level
        obj["total_used"] *= multi
        obj["total_loss"] *= multi
        obj["standard_used"] *= multi
        obj["total_price_rmb"] *= multi
        f.write(get_str(top) + "," + get_str(build_log[top]) + "," + get_str(type_map[top]) + ",")
        f.write(','.join(get_str(obj[i]) for i in obj))
        f.write("\n")
        if obj["supply_method"] == '生产':
            iter_level(top, f, bom, obj["item_code"], level + 1, multi * obj["total_used"])


def process_bom_details(excel_name, sheet_name, output_file_name):
    try:
        df = pd.read_excel(excel_name, sheet_name)
        df.apply(lambda r: r['总消耗量'] - r['损耗量'], axis=1)
        df['标准消耗量'] = df.apply(lambda r: r.总消耗量 - r.损耗量, axis=1)
        item_counter = 0
        bom = {}
        produce_set = set()
        bom_set = set()
        # build_log2 = {}
        for index, row in df.iterrows():
            item_counter += 1
            build_log[row["上级物料编码"]] = row["上级物料名称"]
            type_map[row["上级物料编码"]] = row["上级物料组"]
            bom[row["上级物料编码"]] = []
            bom_set.add(row["上级物料编码"])
            if row['补给方式'] == "生产":
                produce_set.add(row['物料编码'])
                # build_log2[row["物料编码"]] = False
            # print(row['上级物料编码'], row['补给方式'])
        print("bom list has {0} items, and totally {1} BOMs.".format(item_counter, len(build_log)))
        # print(item_counter, len(build_log2))
        for index, row in df.iterrows():
            item_counter += 1
            bom[row["上级物料编码"]].append(create_node_from_row(row))
        # 检查生产物料是否都有BOM
        print("check if each production item has bom")
        print(produce_set.difference(bom_set))
        # print(bom_set.difference(produce_set))
        f = open(output_file_name, "w", encoding='utf8')
        f.write("top,top_name,top_type,")
        f.write(','.join(str(i) for i in bom["521D000T0C0"][0]))
        f.write("\n")
        for k in bom:
            # print(k)
            iter_level(k, f, bom, k, 0, 1)
            # break
        f.close()
    except FileNotFoundError as e:
        print("{0} not found: {1}".format(excel_name, e))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))


def calculate_summary(input_file_name, output_file_name):
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


def process_cost(excel_name, sheet_name, bom_summary_file_name, output_file_name):
    try:
        df = pd.read_excel(excel_name, sheet_name)
        df_cost = pd.read_csv(bom_summary_file_name)
        dfr = pd.merge(left=df, right=df_cost, left_on='物料编号', right_on='top', suffixes=('_l', '_r'), how='left')
        dfr.loc[pd.isna(dfr.total_price_rmb), 'total_price_rmb'] = dfr.loc[pd.isna(dfr.total_price_rmb), '含税人民币价格']
        dfr = dfr.drop(columns=['top', 'top_name', 'top_type', '价格差异', '总成本差异', '价格差异比率'])
        dfr['价格差异'] = dfr.apply(lambda row: row.加权平均成本 - row.total_price_rmb, axis=1)
        dfr['价格差异比率'] = dfr.apply(lambda row: row.价格差异 / row.total_price_rmb if row.total_price_rmb != 0 else 0, axis=1)
        dfr = dfr.rename(index=str, columns={"total_price_rmb": "未税人民币价格"})
        dfr.to_csv(output_file_name, index=False)
    except FileNotFoundError as e:
        print("{0} not found: {1}".format(excel_name, e))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))


def summary_to_excel():
    try:
        writer = pd.ExcelWriter(OUTPUT_FILE_PATH, engine='xlsxwriter', options={'strings_to_numbers': True})
        # write sheet 1
        df1 = pd.read_csv(MIDDLE_FILE_PATH_BOM)
        # print(df_new['currency'].describe())
        # print(df_new['currency'].value_counts())
        df1.to_excel(writer, index=False, sheet_name=OUTPUT_FILE_SHEET_BOM)
        # write sheet 2
        try:
            df2 = pd.read_csv(MIDDLE_FILE_PATH_BOM_SUMMARY)
            df2.to_excel(writer, index=False, sheet_name=OUTPUT_FILE_SHEET_BOM_SUMMARY)
        except Exception as e:
            print("未找到文件【{0}】，放弃创建【{1}】页签:{2}"
                  .format(MIDDLE_FILE_PATH_BOM_SUMMARY, OUTPUT_FILE_SHEET_BOM_SUMMARY, e))
        # write sheet 3
        try:
            df3 = pd.read_csv(MIDDLE_FILE_PATH_BOM_COST_COMPARE)
            df3.to_excel(writer, index=False, sheet_name=OUTPUT_FILE_SHEET_BOM_COST_COMPARE)
        except Exception as e:
            print("未找到文件【{0}】，放弃创建【{1}】页签:{2}"
                  .format(MIDDLE_FILE_PATH_BOM_COST_COMPARE, OUTPUT_FILE_SHEET_BOM_COST_COMPARE, e))
        # change excel format
        # workbook = writer.book
        # worksheet = writer.sheets['Sheet1']
        # format1 = workbook.add_format({'num_format': '#,###0.000'})
        # worksheet.set_column('F:F', None, format1)
        # worksheet.set_column('G:G', None, format1)
        # worksheet.set_column('H:H', None, format1)
        # worksheet.set_column('K:K', None, format1)
        writer.save()
    except FileNotFoundError as e:
        print("上一步骤未完成，中断执行：{0}".format(e))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))


def delete_origin_output_if_exists(filepath):
    try:
        if os.path.exists(filepath) and os.path.isfile(filepath):
            os.remove(filepath)
        else:
            print("未找到文件：{0}".format(filepath))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))


def update_usd_ratio():
    try:
        global USD_RATIO
        ratio = input("Press Enter USD Ratio (default 6.7): ")
        USD_RATIO = float(ratio)
        print("更新汇率为： 1美元 = {0}人民币".format(USD_RATIO))
    except Exception as e:
        print("读入美元汇率失败，采用【1美元 = {0}人民币】进行计算:{1}".format(USD_RATIO, e))


if __name__ == '__main__':
    print("Process Start")
    print(str(datetime.datetime.now()))
    print("****************")

    update_usd_ratio()
    print("****************")

    print("Start to Calculate BOM")
    process_bom_details(INPUT_FILE_PATH, INPUT_BOM_SHEET, MIDDLE_FILE_PATH_BOM)
    print("Finished")
    print("****************")

    print("Start to Calculate BOM Summary")
    calculate_summary(MIDDLE_FILE_PATH_BOM, MIDDLE_FILE_PATH_BOM_SUMMARY)
    print("Finished")
    print("****************")

    print("Start to Combine Warehouse Item Cost")
    process_cost(INPUT_FILE_PATH, INPUT_COST_SHEET, MIDDLE_FILE_PATH_BOM_SUMMARY, MIDDLE_FILE_PATH_BOM_COST_COMPARE)
    print("Finished")
    print("****************")

    print("Start to Transfer CSV to Excel")
    summary_to_excel()
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
