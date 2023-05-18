import os


# 合并多个result table  单独出多个model name的报告时适用


def Merge_Report_Table(app, mutiresult_url):
    result_names = os.listdir(mutiresult_url)

    multl_result = app.books.add()
    multl_result.save(mutiresult_url + '//Muti_result.xlsx')
    multl_result_sheet = multl_result.sheets('Sheet1')
    try:
        for result_name in result_names:
            result = app.books.open(mutiresult_url + "//" + result_name)
            result_sheet = result.sheets("Sheet0")
            result_rows = result_sheet.used_range.last_cell.row
            multl_result_rows = multl_result_sheet.used_range.last_cell.row
            print(result_rows)
            result_sheet.range('A2:AO%d' % result_rows).copy()
            A1 = multl_result_sheet.range('A1').value
            if A1 is None:
                multl_result_sheet.range('A1').paste()
            else:
                multl_result_sheet.range('A%d' % (multl_result_rows + 1)).paste()
            multl_result.save()
            result.close()
    except Exception as E:
        print("拷贝TestCase时出现Error,请检查,错误信息为：%s" % E)
    finally:
        print("合并完成！")
        multl_result.save()
