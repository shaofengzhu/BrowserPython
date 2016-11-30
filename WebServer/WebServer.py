import tornado.ioloop
import tornado.web
import tornado.httpclient
import tornado.websocket
import tornado.gen
import os.path
import json
import threading
import runtime
import excel
import exceldemolib
import httphelper
import time

class MainHandler(tornado.web.RequestHandler):
    async def get(self):
        httpclient = tornado.httpclient.AsyncHTTPClient()
        response = await httpclient.fetch("http://www.bing.com")
        self.write(response.body)

class Constants:
    MessageTypeExecuteAtServer = "ExecuteAtServer"
    MessageTypeExecuteAtServer_Code = "ExecuteAtServer_Code"
    MessageTypeExecuteAtServerResult = "ExecuteAtServerResult"
    MessageTypeExecuteAtClient = "ExecuteAtClient"
    MessageTypeExecuteAtClientResult = "ExecuteAtClientResult"
    Id = "Id"
    Type = "Type"
    Body = "Body"
    SubId = "SubId"

class MessageInfo:
    def __init__(self):
        self.Id = 0
        self.SubId = 0
        self.Type = None
        self.Body = None

class WebSocketRequestExecutor(runtime.IRequestExecutor):
    def __init__(self, msgId: int, wsHandler: tornado.websocket.WebSocketHandler):
        super(self.__class__, self).__init__()
        self._wsHandler = wsHandler
        self._msgId = msgId
        self._msgSubId = 1

    def execute(self, requestInfo: httphelper.RequestInfo):
        self._msgSubId = self._msgSubId + 1
        msgToClient = MessageInfo()
        msgToClient.Id = self._msgId
        msgToClient.SubId = self._msgSubId
        msgToClient.Type = Constants.MessageTypeExecuteAtClient
        msgToClient.Body = requestInfo.body
        textToClient = json.dumps(msgToClient, default = lambda o: o.__dict__)
        self._wsHandler.write_message(textToClient)

        ret = httphelper.ResponseInfo()
        while True:
            msgFromClient = self._wsHandler.getAndRemoveExecuteAtClientResult(msgToClient.Id, msgToClient.SubId)
            if (msgFromClient is not None):
                ret.statusCode = 200
                ret.body = msgFromClient.Body
                break
            time.sleep(1)
        return ret

class EchoSocketHandler(tornado.websocket.WebSocketHandler):
    s_instanceCount = 0

    def __init__(self, application, request, **kwargs):
        EchoSocketHandler.s_instanceCount = EchoSocketHandler.s_instanceCount + 1
        super().__init__(application, request, **kwargs)
        self._queue = []
        self._queueLock = threading.Lock()

    def open(self, *args, **kwargs):
        print('socket open')

    def on_close(self):
        print("Web socket closed")

    def on_message(self, message):
        msg = EchoSocketHandler.parseMessage(message)
        if (msg.Type == Constants.MessageTypeExecuteAtServer):
            self.processExecuteAtServerMessage(msg)
        elif (msg.Type == Constants.MessageTypeExecuteAtServer_Code):
            self.processExecuteCodeAtServerMessage(msg)
        elif (msg.Type == Constants.MessageTypeExecuteAtClientResult):
            self.processExecuteAtClientResultMessage(msg)
        else:
            print("Unknown message" + message)

    def processExecuteAtServerMessage(self, msg: MessageInfo):
        if (msg.Body == "demosimple"):
            threading.Thread(target = EchoSocketHandler.executePopulateDataSimple, args = (msg, self)).start()
        elif (msg.Body == "demo"):
            threading.Thread(target = EchoSocketHandler.executePopulateDataAndAnalyze, args = (msg, self)).start()

    def processExecuteCodeAtServerMessage(self, msg: MessageInfo):
        threading.Thread(target = EchoSocketHandler.executeCode, args = (msg, self)).start()

    def processExecuteAtClientResultMessage(self, msg: MessageInfo):
        with self._queueLock:
            self._queue.append(msg)

    def getAndRemoveExecuteAtClientResult(self, msgId, msgSubId):
        ret = None
        with self._queueLock:
            for i in range(len(self._queue)):
                if self._queue[i].Id == msgId and self._queue[i].SubId == msgSubId:
                    ret = self._queue[i]
                    del self._queue[i]
                    break
        return ret

    @staticmethod
    def parseMessage(messageString: str) -> MessageInfo:
        msg = json.loads(messageString)
        ret = MessageInfo()
        ret.Type = msg.get(Constants.Type)
        ret.Id = msg.get(Constants.Id)
        subId = msg.get(Constants.SubId)
        if (subId is not None):
            ret.SubId = subId
        ret.Body = msg.get(Constants.Body)
        return ret

    @staticmethod
    def executePopulateDataSimple(msg: MessageInfo, wsHandler: tornado.websocket.WebSocketHandler):
        context = excel.RequestContext()
        context.customRequestExecutor = WebSocketRequestExecutor(msg.Id, wsHandler)
        exceldemolib.ExcelDemoLib.populateDataSmall(context)
        EchoSocketHandler.sendExecuteAtServerResultForRequestMessage(msg, wsHandler)

    @staticmethod
    def executePopulateDataAndAnalyze(msg: MessageInfo, wsHandler: tornado.websocket.WebSocketHandler):
        context = excel.RequestContext()
        context.customRequestExecutor = WebSocketRequestExecutor(msg.Id, wsHandler)
        exceldemolib.ExcelDemoLib.clearWorkbook(context)
        exceldemolib.ExcelDemoLib.populateData(context)
        exceldemolib.ExcelDemoLib.analyzeData(context)
        EchoSocketHandler.sendExecuteAtServerResultForRequestMessage(msg, wsHandler)

    @staticmethod
    def executeCode(msg: MessageInfo, wsHandler: tornado.websocket.WebSocketHandler):
        context = excel.RequestContext()
        context.customRequestExecutor = WebSocketRequestExecutor(msg.Id, wsHandler)
        # context.executionMode = runtime.RequestExecutionMode.instantSync
        exec(msg.Body)
        context.sync()
        EchoSocketHandler.sendExecuteAtServerResultForRequestMessage(msg, wsHandler)

    @staticmethod
    def sendExecuteAtServerResultForRequestMessage(msg: MessageInfo, wsHandler: tornado.websocket.WebSocketHandler):
        msgToClient = MessageInfo()
        msgToClient.Id = msg.Id
        msgToClient.Type = Constants.MessageTypeExecuteAtServerResult
        msgToClient.Body = "Done: " + msg.Body
        textToClient = json.dumps(msgToClient, default = lambda o: o.__dict__)
        wsHandler.write_message(textToClient)


def make_app():
    staticPath = os.path.dirname(__file__)
    staticPath = os.path.join(staticPath, "..\\BrowserExecutor")
    settings = {
        "static_path": staticPath
        }
    miniMapsStaticPath = os.path.join(staticPath, "mini-maps")
    return tornado.web.Application([
            (r"/", MainHandler),
            (r"/ws", EchoSocketHandler),
            (r"/mini-maps/(.*)", tornado.web.StaticFileHandler, {"path": miniMapsStaticPath}),
        ], **settings);


if __name__ == "__main__":
    app = make_app()
    app.listen(7080)
    tornado.ioloop.IOLoop.current().start()
