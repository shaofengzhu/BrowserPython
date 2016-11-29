import numbers
import runtime
import word
import datetime
import json

class WordDemoLib:
    @staticmethod
    def initDesktopContext():
        requestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo()
        #requestUrlAndHeaders.url = "http://localhost:8054"
        requestUrlAndHeaders.url = "pipe://./word/_api"
        runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders
        
    @staticmethod
    def insertPictureAtEnd(context: word.RequestContext, base64ImageData: str):
        context.document.body.insertInlinePictureFromBase64(base64ImageData, word.InsertLocation.end)
        context.sync()
        return

    @staticmethod
    def helloWorld(context: word.RequestContext):
        context.document.body.insertText("Hello, World", word.InsertLocation.end)
        context.sync();
        return
