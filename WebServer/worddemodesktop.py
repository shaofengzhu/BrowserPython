import numbers
import runtime
import word
import datetime
import json
import worddemolib

if __name__ == "__main__":
    requestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo();
    requestUrlAndHeaders.url = "http://localhost:8054";
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders
    context = word.RequestContext()
    print("Insert image");
    worddemolib.WordDemoLib.insertSamplePictureAtEnd(context)
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None
