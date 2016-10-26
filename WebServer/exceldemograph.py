import numbers
import runtime
import excel
import excelhelper
import oauthhelper
import datetime
import json
import exceldemolib

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
    context = excel.RequestContext()
    #ExcelDemo.populateDataSmall(context)
    exceldemolib.ExcelDemoLib.populateData(context)
    print("Populated data");
    exceldemolib.ExcelDemoLib.analyzeData(context)
    print("Analyzed data")
    excelhelper.ExcelHelper.closeSession(requestUrlAndHeaders)
    print("Closed session")
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None

