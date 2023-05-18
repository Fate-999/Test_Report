# 拷贝测项到模板报告中

def Open_result_template(app,result_url,template_url):
    try:
        result = app.books.open(fullname=result_url)  # 获取result_sheet对象
        template = app.books.open(fullname=template_url)  # 获取template对象
        result_sheet1 = result.sheets("Sheet0")
        test_result = template.sheets("Test result")
        return result, template, result_sheet1, test_result
    except Exception as E:
        print("OpenFile出现Error,请检查文件,错误信息为：%s" % E)
        return None


def Open_multi_result_template(app,mutiresult_url,template_url):
    try:
        result = app.books.open(fullname=mutiresult_url+"//Muti_result.xlsx")  # 获取result_sheet对象
        template = app.books.open(fullname=template_url)  # 获取template对象
        result_sheet1 = result.sheets("Sheet1")
        test_result = template.sheets("Test result")
        return result, template, result_sheet1, test_result
    except Exception as E:
        print("OpenFile出现Error,请检查文件,错误信息为：%s" % E)
        return None


def Copy_multi_test_case(app,report_url,mutiresult_url,template_url):
    result, template, result_sheet1, test_result = Open_multi_result_template(app,mutiresult_url,template_url)
    try:
        info = result_sheet1.used_range
        rows = info.last_cell.row
        test_result.api.PivotTables("樞紐分析表2").Location = "'Test result'!$A$%d" % (rows + 2)
        result_sheet1.range('A1:AO%d' % rows).copy()
        test_result.range('A2').paste()
        template.save(report_url)
        print("TestCase Copy成功！")
    except Exception as e:
        print("拷贝TestCase时出现Error,请检查,错误信息为：%s" % e)
    finally:
        result.close()
        template.close()


def Copy_test_case(app,result_url,report_url,template_url):
    result, template, result_sheet1, test_result = Open_result_template(app,result_url,template_url)
    try:
        info = result_sheet1.used_range
        rows = info.last_cell.row
        test_result.api.PivotTables("樞紐分析表2").Location = "'Test result'!$A$%d" % (rows + 2)
        result_sheet1.range('A2:AO%d' % rows).copy()
        test_result.range('A2').paste()
        template.save(report_url)
        print("TestCase Copy成功！")
    except Exception as e:
        print("拷贝TestCase时出现Error,请检查,错误信息为：%s" % e)
    finally:
        result.close()
        template.close()
