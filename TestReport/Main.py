# 主程序入口
import App
import Copy_Test_Case
import Find_Issue_Number
import Copy_Issue_To_Report
import Handle_Issue_Table
import Merge_Report_Table
import Judge_New_Issue
import Refresh_All_Table

if __name__ == '__main__':
    app = App.Launch_App()
    for i in range(0, 10):
        strings = input("是否为完整的result.(Y/N):")
        Time = input("请输入BIOS Release日期(用/分隔)：").split("/")
        if strings == "Y" or strings == "y":
            print("正在处理...请稍后...")
            try:
                Handle_Issue_Table.Handle_issue_table(app)
                Copy_Test_Case.Copy_test_case(app)
                Find_Issue_Number.Search_bugid_paste_allissue(app)
                Copy_Issue_To_Report.Copy_Case(app)
                Judge_New_Issue.Judge(app, Time)
                Refresh_All_Table.Refresh_all(app)
            except Exception as e:
                print("Error,请重试,错误信息为：%s" % e)
            finally:
                print("全部完成！")
                app.quit()
            break
        elif strings == "N" or strings == "n":
            print("正在处理...请稍后...")
            try:
                Merge_Report_Table.Merge_Report_Table(app)
                Handle_Issue_Table.Handle_issue_table(app)
                Copy_Test_Case.Copy_multi_test_case(app)
                Find_Issue_Number.Search_bugid_paste_allissue(app)
                Copy_Issue_To_Report.Copy_Case(app)
                Judge_New_Issue.Judge(app, Time)
                Refresh_All_Table.Refresh_all(app)
            except Exception as e:
                print("Error,请重试,错误信息为：%s" % e)
            finally:
                print("全部完成！")
                app.quit()
            break
