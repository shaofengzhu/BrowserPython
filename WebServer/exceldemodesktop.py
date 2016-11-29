import runtime
import excel
import exceldemolib

if __name__ == "__main__":
    exceldemolib.ExcelDemoLib.initDesktopContext()
    context = excel.RequestContext()
    print("Clearing workbook")
    exceldemolib.ExcelDemoLib.clearWorkbook(context)
    print("Populating data");
    exceldemolib.ExcelDemoLib.populateData(context)
    print("Populated data");
    exceldemolib.ExcelDemoLib.analyzeData(context)
    print("Analyzed data")
    imageBase64 = exceldemolib.ExcelDemoLib.getChartImage(context)
    print("ImageSize:");
    print(len(imageBase64));
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None

