
import os


def delete_origin_output_if_exists(filepath):
    try:
        if os.path.exists(filepath) and os.path.isfile(filepath):
            os.remove(filepath)
        else:
            print("未找到文件：{0}".format(filepath))
    except Exception as e:
        print("Unknown Exception: {0}".format(e))


def update_usd_ratio(usd_ratio):
    try:
        ratio = input("Press Enter USD Ratio (default 6.7): ")
        usd_ratio = float(ratio)
        print("更新汇率为： 1美元 = {0}人民币".format(usd_ratio))
        return usd_ratio
    except Exception as e:
        print("读入美元汇率失败，采用【1美元 = {0}人民币】进行计算:{1}".format(usd_ratio, e))
        return usd_ratio


def get_str(a):
    if type(a) == str:
        if '"' in a:
            a = str(a).replace('"', "")
        return '"' + a + '"'
    else:
        return str(a)
