# 刷新全部数据透析表


def Open_File_Report(app,report_url):
    try:
        table = app.books.open(fullname=report_url)
        A = table.sheets("Critical Bug 分佈")
        B = table.sheets("BSP&AP 分佈")
        C = table.sheets("新增Critical Bug 分布")
        D = table.sheets("Test result")
        return A, B, C, D, table
    except Exception as E:
        print("OpenFile出现Error,请检查文件,错误信息为：%s" % E)
        return None


def Refresh_all(app,report_url):
    A, B, C, D, table = Open_File_Report(app,report_url)
    A.api.PivotTables("樞紐分析表1").PivotCache().Refresh()
    B.api.PivotTables("樞紐分析表1").PivotCache().Refresh()
    C.api.PivotTables("樞紐分析表1").PivotCache().Refresh()
    D.api.PivotTables("樞紐分析表2").PivotCache().Refresh()
    print("数据表刷新完成")
    table.save()
    table.close()
