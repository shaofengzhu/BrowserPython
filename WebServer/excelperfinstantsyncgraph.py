import runtime
import excel
import exceldemolib
import datetime
import oauthhelper
import excelhelper

if __name__ == "__main__":
    accessToken = oauthhelper.OAuthUtility.getAccessToken(True)
    headers = {}
    headers["Authorization"] = "Bearer " + accessToken
    now = datetime.datetime.now()
    filename = "ShaoZhuPython-" + now.strftime("%Y-%m-%d %H-%M-%S") + ".xlsx"
    workbookUrl = excelhelper.ExcelHelper.createBlankExcelFile("https://graph.microsoft.com/v1.0/me/drive/root", filename, headers)
    print("Created file: " + workbookUrl)
    requestUrlAndHeaders = excelhelper.ExcelHelper.createSessionAndBuildRequestUrlAndHeaders(workbookUrl, headers)
    print("Created session");
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders
    print("Warm up")
    context = excel.RequestContext()
    context.workbook.load()
    context.sync()
    print(context.workbook.isNullObject)

    time1 = exceldemolib.ExcelDemoLib.perfTestPopulateAndAnalyzeData(context)
    print(time1)

    print("Clearing workbook")
    exceldemolib.ExcelDemoLib.clearWorkbook(context)

    context = excel.RequestContext()
    context.executionMode = runtime.RequestExecutionMode.instantSync
    time2 = exceldemolib.ExcelDemoLib.perfTestPopulateAndAnalyzeData(context)
    print(time2)

    excelhelper.ExcelHelper.closeSession(requestUrlAndHeaders)
    print("Closed session")
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None
