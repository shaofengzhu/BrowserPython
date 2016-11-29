import runtime
import excel
import exceldemolib
if __name__ == "__main__":
    exceldemolib.ExcelDemoLib.initDesktopContext()
    timeSpent = exceldemolib.ExcelDemoLib.perfTest()
    print(timeSpent)
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None

