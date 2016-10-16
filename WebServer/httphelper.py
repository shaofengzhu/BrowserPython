import sys
import json
import enum
import logging
import urllib.request
import urllib


class RequestInfo:
    def __init__(self):
        self.method = "GET"
        self.url = None
        self.headers = {}
        self.body = ""

class ResponseInfo:
    def __init__(self):
        self.statusCode = 200
        self.headers = {}
        self.body = ""

class HttpUtility:
    @staticmethod
    def invoke(requestInfo: RequestInfo) -> ResponseInfo:
        if requestInfo.method is None:
            requestInfo.method = "GET"
        requestInfo.method = requestInfo.method.upper()
        if requestInfo.method == "GET" or requestInfo.method == "DELETE":
            requestInfo.body = None
        bodyData = None
        if requestInfo.body is not None and isinstance(requestInfo.body, str):
            bodyData = requestInfo.body.encode("utf8")
        else:
            bodyData = requestInfo.body
        req = urllib.request.Request(requestInfo.url, method = requestInfo.method, headers = requestInfo.headers, data = bodyData);
        resp = urllib.request.urlopen(req)
        respInfo = ResponseInfo()
        respInfo.headers = {}
        for key, value in resp.info().items():
            respInfo.headers[key] = value
        respInfo.statusCode = resp.status
        charset = resp.info().get_content_charset()
        if charset is None:
            charset = "utf8"
        respInfo.body = resp.read().decode(charset)
        return respInfo

if __name__ == "__main__":
    req = RequestInfo()
    req.url = 'https://www.google.com/'
    response = HttpUtility.invoke(req)
    print(response.body)

