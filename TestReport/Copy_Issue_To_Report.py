import App


# 根据 All issue中的issue号进行在issue表中查询

def open_file(app,issue_url,report_url):
    try:
        issue = app.books.open(fullname=issue_url)
        template = app.books.open(fullname=report_url)
        issue_sheet = template.sheets("All issue")  # 模板 All issue sheet
        issue_table = issue.sheets("report")  # issue table
        return issue_table, issue_sheet, issue, template
    except Exception as E:
        print("OpenFile出现Error,请检查文件,错误信息为：%s" % E)
        return None


def Copy_Case(app,issue_url,template_url):
    issue_table, issue_sheet, issue, template = open_file(app,issue_url,template_url)
    try:
        rows = issue_sheet.used_range.last_cell.row
        BugID_values = list(issue_sheet['A2:A%d' % rows].value)

        BugID = []
        for k in BugID_values:
                k = str(k)
                if k != 'None':
                    z = k[:8]
                    BugID.append(z)

        issue_rows = issue_table.used_range.last_cell.row
        issue_id = list(issue_table['A1:A%d' % issue_rows].value)

        str1 = []
        str2 = []
        for i in BugID:
            for j in issue_id:
                if i == j:
                    a = BugID.index(i)
                    b = issue_id.index(j) + 1
                    str1.append(a)
                    str2.append(b)
                    break

        for num in range(0, len(BugID)):
            issue_table.range('A%d:T%d' % (str2[num], str2[num])).copy()
            issue_sheet.range('A%d' % (str1[num] + 2)).paste()

        print("issue粘贴完成")
        template.save()
        issue.save()
    except Exception as E:
        print("拷贝issue至Report出现Error,请检查,错误信息为：%s" % E)
    finally:
        issue.close()
        template.close()
