import sys
import json
import random
import oauthhelper
import excel
import runtime

class ExcelTest:
    @staticmethod
    def initGraphSettings():
        fileInfo = oauthhelper.OAuthUtility.getFileAccessInfo(True, "AgaveTest.xlsx")
        requestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo()
        requestUrlAndHeaders.url = fileInfo.fileWorkbookUrl
        requestUrlAndHeaders.headers["Authorization"] = "Bearer " + fileInfo.accessToken
        runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders

    @staticmethod
    def initDevMachine():
        requestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo()
        requestUrlAndHeaders.url = "http://shaozhu-ttvm8.redmond.corp.microsoft.com/th/WacRest.ashx/transport_wopi/Application_Excel/wachost_/Fi_anonymous~AgaveTest.xlsx/ak_1%7CGN=R3Vlc3Q=&SN=Nzk2NzQ5NTk4&IT=NTI0NzU4Mjg1OTA0MzY3MjYxMA==&PU=Nzk2NzQ5NTk4&SR=YW5vbnltb3Vz&TZ=MTExOQ==&SA=RmFsc2U=&LE=RmFsc2U=&AG=VHJ1ZQ==&RH=tcFPR0obGkLwYJoRrzbWPXcQeSWEPdXeZt4RccCYWWQ=/_api"
        runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders

    @staticmethod
    def clearWorksheet(ctx: excel.RequestContext, sheetName: str):
        sheet = ctx.workbook.worksheets.getItem(sheetName)
        ctx.load(sheet.tables)
        ctx.sync()
        for table in sheet.tables.items:
            table.delete()
        sheet.getRange(None).clear(excel.ClearApplyTo.all)
        ctx.sync()

    @staticmethod
    def removeAllCharts(ctx: excel.RequestContext, sheetName: str):
        charts = ctx.workbook.worksheets.getItem(sheetName).charts
        ctx.load(charts, "id")
        ctx.sync()
        for chart in charts.items:
            chart.delete()
        ctx.sync()

    @staticmethod
    def test_Range_SetValueReadValue():
        rangeValue = 12345;
        rangeValuesToSet = [[str(rangeValue), "=A1"], ["=B1", "=A1+B1"]];

        ctx = excel.RequestContext()
        r = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2")
        r.values = rangeValuesToSet
        r.load()
        ctx.sync()
        print(r.values)
        print(r.address)

    @staticmethod
    def test_Worksheet_GetWorksheetCollection():
        ctx = excel.RequestContext()
        ctx.workbook.worksheets.load()
        ctx.sync()
        print("Worksheets")
        for sheet in ctx.workbook.worksheets.items:
            print(sheet.name)

    @staticmethod
    def test_Worksheet_AddDeleteWorksheet():
        ctx = excel.RequestContext()
        random.seed()
        name = "PythonTest" + str(random.randint(1, 3000))
        sheet = ctx.workbook.worksheets.add(name)
        sheet.load()
        ctx.sync()
        print("Created sheet " + sheet.name)
        print("Created sheetIndex " + str(sheet.position))
        sheet.delete()
        ctx.sync()
        print("Deleted sheet")

    @staticmethod
    def test_Table_GetCollection():
        ctx = excel.RequestContext()
        ctx.load(ctx.workbook.tables)
        ctx.sync()
        for table in ctx.workbook.tables.items:
            print(table.name)

    @staticmethod
    def test_Table_CreateTable():
        sheetName = "Tables"
        tableAddress = sheetName + "!A23:B25"
        ctx = excel.RequestContext()
        ExcelTest.clearWorksheet(ctx, sheetName)
        t = ctx.workbook.tables.add(tableAddress, True)
        ctx.load(t)
        ctx.sync()
        print("Created table id=" + str(t.id))
        print("Created table name=" + t.name)

    @staticmethod
    def test_Table_CreateDeleteTable():
        sheetName = "Tables"
        tableAddress = sheetName + "!A23:B25"
        ctx = excel.RequestContext()
        ExcelTest.clearWorksheet(ctx, sheetName)
        t = ctx.workbook.tables.add(tableAddress, True)
        ctx.load(t)
        ctx.sync()
        print("Created table id=" + str(t.id))
        print("Created table name=" + t.name)
        t.delete()
        ctx.sync()
        print("Deleted table")

    @staticmethod
    def test_Chart_CreateChart():
        sheetName = "Charts"
        ctx = excel.RequestContext()
        ExcelTest.removeAllCharts(ctx, sheetName)
        sheet = ctx.workbook.worksheets.getItem(sheetName)
        sourceData = sheet.getRange(sheetName + "!" + "A1:B4")
        chart = sheet.charts.add(excel.ChartType.pie, sourceData, excel.ChartSeriesBy.auto)
        ctx.load(chart)
        ctx.sync()
        print("Created sheet " + chart.name)

    @staticmethod
    def test_Chart_CreateDeleteChart():
        sheetName = "Charts"
        ctx = excel.RequestContext()
        ExcelTest.removeAllCharts(ctx, sheetName)
        sheet = ctx.workbook.worksheets.getItem(sheetName)
        sourceData = sheet.getRange(sheetName + "!" + "A1:B4")
        chart = sheet.charts.add(excel.ChartType.pie, sourceData, excel.ChartSeriesBy.auto)
        ctx.load(chart)
        ctx.sync()
        print("Created sheet " + chart.name)
        chart.delete()
        ctx.sync()
        print("Deleted chart")

if __name__ == "__main__":
    ExcelTest.initGraphSettings()
    ExcelTest.test_Range_SetValueReadValue()
    
    methods = dir(ExcelTest)
    for method in methods:
        if method.startswith("test_"):
            print("invoke " + method)
            func = getattr(ExcelTest, method)
            func()
    
