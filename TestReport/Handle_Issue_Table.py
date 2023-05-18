
# 处理Issue表，删除多余部分
def Open_issue_table(app,issue_url):
    issue = app.books.open(fullname=issue_url)  # 获取result_sheet对象
    issue_table = issue.sheets("report")  # issue table
    return issue, issue_table


def Handle_issue_table(app,issue_url):
    issue, issue_table = Open_issue_table(app,issue_url)
    A1 = issue_table.range('A1').value
    if not A1.isdigit() or len(A1) != 8:
        issue_table.range(f'{1}:{6}').api.Delete()  # 删除多余行数
        print("issue处理完成")
    else:
        print("issue无需处理")
    issue.save()
    issue.close()
