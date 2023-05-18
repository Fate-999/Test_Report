import xlwings as xw


def Launch_App():
    app = xw.App(visible=False, add_book=False)  # 界面设置
    app.display_alerts = False  # 关闭提示信息
    app.screen_updating = False  # 关闭显示更新
    return app



