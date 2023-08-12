# 判断 是否 new issue
import datetime
import re


def Open_File_Report(app, report_url):
    try:
        table = app.books.open(fullname=report_url)
        all_issue = table.sheets("All issue")
        new_issue = table.sheets("此版新增Bug")
        return new_issue, all_issue, table
    except Exception as E:
        print("OpenFile出现Error,请检查文件,错误信息为：%s" % E)
        return None


def Judge(app, Time, report_url, bios_version):
    new_issue, all_issue, table = Open_File_Report(app, report_url)
    rows = all_issue.used_range.last_cell.row
    Time_list = list(all_issue['P2:P%d' % rows].value)
    bios_list = list(all_issue['L2:L%d' % rows].value)
    bios_re_list = []
    for bios in bios_list:
        bios = str(bios)
        bios_re = re.findall(r'\b\d{3}\b', bios)
        if bios_re == []:
            bios_re = ['000']
        for c in bios_re:
            bios_re_list.append(c)
    Time = str(Time).split("/")
    year = int(Time[0])
    month = int(Time[1])
    day = int(Time[2])
    Release_time = datetime.datetime(year, month, day)

    selected_data = []
    # for c in bios_re_list:
    #     for d in Time_list:
    #         if c == bios_version:
    #             if d > Release_time:
    #                 selected_data.append(d)
    for a,b in zip(bios_re_list,Time_list):
        if a == bios_version and b > Release_time:
            selected_data.append((a,b))


    selected_data = list(set(selected_data))
    selected_data_time_list = []
    for i in selected_data:
        x,y=i
        selected_data_time_list.append(y)


    time_list_index = []
    for x in selected_data_time_list:
        time_list_index.append(Time_list.index(x) + 2)

    for num in range(0, len(time_list_index)):
        all_issue.range('A%d:T%d' % (time_list_index[num], time_list_index[num])).copy()
        new_issue.range('A%d' % (num + 2)).paste()

    time_list_index.sort(reverse=True)
    for y in time_list_index:
        all_issue.range('A%d:T%d' % (y, y)).api.EntireRow.Delete()

    print("New Issue 区分完成")
    table.save()
    table.close()
