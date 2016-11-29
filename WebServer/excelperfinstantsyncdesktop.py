import runtime
import excel
import exceldemolib
import datetime


if __name__ == "__main__":
    exceldemolib.ExcelDemoLib.initDesktopContext()
    context = excel.RequestContext()
    print("Clearing workbook")
    exceldemolib.ExcelDemoLib.clearWorkbook(context)

    time1 = exceldemolib.ExcelDemoLib.perfTestPopulateAndAnalyzeData(context)
    print(time1)

    print("Clearing workbook")
    exceldemolib.ExcelDemoLib.clearWorkbook(context)

    context = excel.RequestContext()
    context.executionMode = runtime.RequestExecutionMode.instantSync
    time2 = exceldemolib.ExcelDemoLib.perfTestPopulateAndAnalyzeData(context)
    print(time2)

    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None


