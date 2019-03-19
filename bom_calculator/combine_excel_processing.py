import pandas as pd


def combine_csv_to_excel(middle_file_path_bom,
                         middle_file_path_bom_summary,
                         middle_file_path_bom_cost_compare,
                         output_file_path,
                         output_file_sheet_bom,
                         output_file_sheet_bom_summary,
                         output_file_sheet_bom_cost_compare):
    try:
        writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter', options={'strings_to_numbers': True})
        # write sheet 1
        df1 = pd.read_csv(middle_file_path_bom)
        # print(df_new['currency'].describe())
        # print(df_new['currency'].value_counts())
        df1.to_excel(writer, index=False, sheet_name=output_file_sheet_bom)
        # write sheet 2
        try:
            df2 = pd.read_csv(middle_file_path_bom_summary)
            df2.to_excel(writer, index=False, sheet_name=output_file_sheet_bom_summary)
        except Exception as e:
            print("未找到文件【{0}】，放弃创建【{1}】页签:{2}"
                  .format(middle_file_path_bom_summary, output_file_sheet_bom_summary, e))
        # write sheet 3
        try:
            df3 = pd.read_csv(middle_file_path_bom_cost_compare)
            df3.to_excel(writer, index=False, sheet_name=output_file_sheet_bom_cost_compare)
        except Exception as e:
            print("未找到文件【{0}】，放弃创建【{1}】页签:{2}"
                  .format(middle_file_path_bom_cost_compare, output_file_sheet_bom_cost_compare, e))
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
