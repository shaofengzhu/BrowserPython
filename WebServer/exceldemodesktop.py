import numbers
import runtime
import excel
import datetime
import json
import exceldemolib

if __name__ == "__main__":
    requestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo();
    requestUrlAndHeaders.url = "http://localhost:8052";
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders
    context = excel.RequestContext()
    print("Populating data");
    exceldemolib.ExcelDemoLib.populateData(context)
    print("Populated data");
    exceldemolib.ExcelDemoLib.analyzeData(context)
    print("Analyzed data")
    imageBase64 = exceldemolib.ExcelDemoLib.getChartImage(context)
    print("Image:");
    print(imageBase64);
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None

