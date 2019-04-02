
import pandas as pd
import copy
from bom_calculator.bom_utils import get_str


def create_node_from_row(r, usd_ratio):
    # 物料编码	物料名称	原厂料号	补给方式	首选供应商	物料分类二	物料分类三	计量单位	上级物料编码	上级物料名称	总消耗量	损耗量	标准损耗率	未税采购价格	采购货币
    total_price_rmb = r.未税采购价格 * r.总消耗量 * usd_ratio if r.采购货币 == 'USD' else r.未税采购价格 * r.总消耗量
    return {
        "item_code": r.物料编码,
        "item_name": r.物料名称,
        "item_group": r.物料组,
        "level": 1,
        "total_used": r.总消耗量,
        "total_loss": r.损耗量,
        "standard_used": r.标准消耗量,
        "price": r.未税采购价格,
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


def iterate_next_level(top, f, bom, key, level, multi, build_log, type_map):
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
            iterate_next_level(top, f, bom, obj["item_code"], level + 1, multi * obj["total_used"], build_log, type_map)


def calculate_bom_details(excel_name, sheet_name, output_file_name, usd_ratio):
    try:
        build_log = {}
        type_map = {}
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
            # item_counter += 1
            bom[row["上级物料编码"]].append(create_node_from_row(row, usd_ratio))
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
            iterate_next_level(k, f, bom, k, 0, 1, build_log, type_map)
            # break
        f.close()
    except FileNotFoundError as e:
        print("{0} not found: {1}".format(excel_name, e))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))
