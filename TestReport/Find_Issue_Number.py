import re


# 搜索BugID并贴到All issue
def Open_template(app,report_url):
    try:
        table = app.books.open(fullname=report_url)
        test_result = table.sheets("Test result")
        return test_result, table
    except Exception as E:
        print("OpenFile出现Error,请检查文件,错误信息为：%s" % E)
        return None


def Search_bugid_paste_allissue(app,report_url):
    test_result, table = Open_template(app,report_url)
    try:
        rows = test_result.used_range.last_cell.row
        BugID_values = list(test_result['S2:S%d' % rows].value)
        Comment_values = list(test_result['T2:T%d' % rows].value)
        for i in Comment_values:
            BugID_values.append(i)

        BugId = set(BugID_values)
        BugID_values = list(BugId)
        bug = []
        for i in BugID_values:
            i = str(i)
            b = re.findall(r'\d{8}', i)
            for c in b:
                if len(str(c)) == 8:
                    if c[0] == "2":
                        bug.append(c)
        bug = set(bug)
        bug = list(bug)
        bug.sort()
        all_issue = table.sheets("All issue")
        all_issue.range('A2').options(transpose=True).value = bug
        lenth = len(bug)
        print("BugID 挑出成功！, 共 %s 条 issue" % lenth)
        table.save()
    except Exception as e:
        print("挑BugID出现Error,请检查,错误信息为：%s" % e)
    finally:
        table.close()
