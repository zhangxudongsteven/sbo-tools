
import pandas as pd
import copy
import datetime


build_log = {}


def create_node_from_row(r):
    # 物料编码	物料名称	原厂料号	补给方式	首选供应商	物料分类二	物料分类三	计量单位	上级物料编码	上级物料名称	总消耗量	损耗量	标准损耗率	采购价格	采购货币
    return {
        "item_code": r.物料编码,
        "item_name": r.物料名称,
        "level": 1,
        "total_used": r.总消耗量,
        "total_loss": r.损耗量,
        "standard_used": r.标准消耗量,
        "price": r.采购价格,
        "currency": r.采购货币,
        "total_price": r.采购价格 * r.总消耗量,
        "origin_item_code": r.原厂料号,
        "supply_method": r.补给方式,
        "first_supplier": r.首选供应商,
        "class2": r.物料分类二,
        "class3": r.物料分类三,
        "unit": r.计量单位,
        "mother_item_code": r.上级物料编码,
        "mother_item_name": r.上级物料名称,
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
        obj["total_price"] *= multi
        f.write(get_str(top) + "," + get_str(build_log[top]) + ",")
        f.write(','.join(get_str(obj[i]) for i in obj))
        f.write("\n")
        if obj["supply_method"] == '生产':
            iter_level(top, f, bom, obj["item_code"], level + 1, multi * obj["total_used"])


def read_excel_by_pandas(excel_name, sheet_name, output_file_name):
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
            bom[row["上级物料编码"]] = []
            bom_set.add(row["上级物料编码"])
            if row['补给方式'] == "生产":
                produce_set.add(row['物料编码'])
            #     build_log2[row["物料编码"]] = False
            # print(row['上级物料编码'], row['补给方式'])
        print("bom list has " + str(item_counter) + " items, and totally " + str(len(build_log)) + " BOMs.")
        # print(item_counter, len(build_log2))
        for index, row in df.iterrows():
            item_counter += 1
            bom[row["上级物料编码"]].append(create_node_from_row(row))
        # 检查生产物料是否都有BOM
        print("check if each production item has bom")
        print(produce_set.difference(bom_set))
        # print(bom_set.difference(produce_set))
        f = open(output_file_name, "w", encoding='utf8')
        f.write("top,top_name,")
        f.write(','.join(str(i) for i in bom["521D000T0C0"][0]))
        f.write("\n")
        for k in bom:
            # print(k)
            iter_level(k, f, bom, k, 0, 1)
            # break
        f.close()
    except FileNotFoundError:
        print(excel_name + " not found")
    except Exception:
        print("Unknown Exception")


def transfer_csv_to_excel(input_file_path, output_file_name, sheet_name):
    try:
        df_new = pd.read_csv(input_file_path)
        print(df_new['currency'].describe())
        print(df_new['currency'].value_counts())
        writer = pd.ExcelWriter(output_file_name, engine='xlsxwriter', options={'strings_to_numbers': True})
        df_new.to_excel(writer, index=False, sheet_name='Sheet1')
        # workbook = writer.book
        # worksheet = writer.sheets['Sheet1']
        # format1 = workbook.add_format({'num_format': '#,###0.000'})
        # worksheet.set_column('F:F', None, format1)
        # worksheet.set_column('G:G', None, format1)
        # worksheet.set_column('H:H', None, format1)
        # worksheet.set_column('K:K', None, format1)
        writer.save()
    except Exception:
        print("Unknown Exception")


if __name__ == '__main__':
    print("Processing Start")
    print(str(datetime.datetime.now()))
    print("****************")
    print("Start to Calculate BOM")
    read_excel_by_pandas("input.xlsx", "Sheet1", "output.csv")
    print("Finished")
    print("****************")
    print("Start to Transfer CSV to Excel")
    transfer_csv_to_excel('output.csv', "output.xlsx", "Sheet1")
    print("Finished")
    print("****************")
    input("Press Enter to continue...")
