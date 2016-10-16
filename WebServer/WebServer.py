import tornado.ioloop
import tornado.web
import tornado.httpclient
import tornado.websocket
import tornado.gen
import os.path
import json

class MainHandler(tornado.web.RequestHandler):
    async def get(self):
        httpclient = tornado.httpclient.AsyncHTTPClient()
        response = await httpclient.fetch("http://www.bing.com")
        self.write(response.body)

class Constants:
    MessageTypeExecuteAtServer = "ExecuteAtServer"
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


class EchoSocketHandler(tornado.websocket.WebSocketHandler):
    s_instanceCount = 0

    def __init__(self, application, request, **kwargs):
        EchoSocketHandler.s_instanceCount = EchoSocketHandler.s_instanceCount + 1
        return super().__init__(application, request, **kwargs)

    def open(self, *args, **kwargs):
        print('socket open')

    def on_close(self):
        print("Web socket closed")

    def on_message(self, message):
        msg = EchoSocketHandler.parseMessage(message)
        if (msg.Type == Constants.MessageTypeExecuteAtServer):
            self.processExecuteAtServerMessage(msg)
        elif (msg.Type == Constants.MessageTypeExecuteAtClientResult):
            self.processExecuteAtClientResultMessage(msg)
        else:
            print("Unknown message" + message)

    def processExecuteAtServerMessage(self, msg: MessageInfo):
        msgToClient = MessageInfo()
        msgToClient.Id = msg.Id
        msgToClient.Type = Constants.MessageTypeExecuteAtClient
        msgToClient.Body = "Something"
        msgToClient.SubId = 1
        textToClient = json.dumps(msgToClient, default = lambda o: o.__dict__)
        self.write_message(textToClient)

    def processExecuteAtClientResultMessage(self, msg: MessageInfo):
        msgToClient = MessageInfo()
        msgToClient.Id = msg.Id
        msgToClient.Type = Constants.MessageTypeExecuteAtServerResult
        msgToClient.Body = msg.Body
        textToClient = json.dumps(msgToClient, default = lambda o: o.__dict__)
        self.write_message(textToClient)

    @staticmethod
    def parseMessage(messageString: str) -> MessageInfo:
        msg = json.loads(messageString)
        ret = MessageInfo()
        ret.Type = msg.get(Constants.Type)
        ret.Id = msg.get(Constants.Id)
        ret.Body = msg.get(Constants.Body)

        return ret


def make_app():
    staticPath = os.path.dirname(__file__)
    staticPath = os.path.join(staticPath, "..\\BrowserExecutor")
    settings = {
        "static_path": staticPath
        }
    return tornado.web.Application([
            (r"/", MainHandler),
            (r"/ws", EchoSocketHandler),
        ], **settings);


if __name__ == "__main__":
    app = make_app()
    app.listen(8888)
    tornado.ioloop.IOLoop.current().start()
