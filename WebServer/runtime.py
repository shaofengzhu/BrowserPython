import sys
import json
import enum
import datetime
import logging
import httphelper
import namedpipeclient

class Constants:
    getItemAt = "GetItemAt"
    id = "Id"
    idPrivate = "_Id"
    index = "_Index"
    items = "_Items"
    iterativeExecutor = "IterativeExecutor"
    localDocument = "http://document.localhost/"
    localDocumentApiPrefix = "http://document.localhost/_api/"
    referenceId = "_ReferenceId"

class RichApiRequestMessageIndex(enum.IntEnum):
    CustomData = 0
    Method = 1
    PathAndQuery = 2
    Headers = 3
    Body = 4
    AppPermission = 5
    RequestFlags = 6

class RichApiResponseMessageIndex(enum.IntEnum):
    StatusCode = 0
    Headers = 1
    Body = 2

class ActionType(enum.IntEnum):
    Instantiate = 1
    Query = 2
    Method = 3
    SetProperty = 4
    Trace = 5

class OperationType(enum.IntEnum):
    Default = 0
    Read = 1

class ObjectPathType(enum.IntEnum):
    GlobalObject = 1
    NewObject = 2
    Method = 3
    Property = 4
    Indexer = 5
    ReferenceId = 6
    NullObject = 7

class ClientRequestFlags(enum.IntEnum):
    NoneValue = 0
    WriteOperation = 1

class ArgumentInfo:
    Arguments = None
    ReferencedObjectPathIds = None

class QueryInfo:
    def __init__(self):
        self.Select = None
        self.Expand = None
        self.Skip = None
        self.Top = None


class ActionInfo:
    def __init__(self):
        self.Id = 0
        self.ActionType = None
        self.Name = None
        self.ObjectPathId = 0
        self.ArgumentInfo = None
        self.QueryInfo = None


class ActionResultInfo:
    def __init__(self):
        self.ActionId = 0
        self.Value = None

class ObjectPathInfo:
    def __init__(self):
        self.Id = 0
        self.ObjectPathType = None
        self.Name = None
        self.ParentObjectPathId = 0
        self.ArgumentInfo = None

class RequestMessageBodyInfo:
    def __init__(self):
        self.Actions = None
        self.ObjectPaths = None

class ErrorInfo:
    def __init__(self):
        self.Code = None
        self.Message = None
        self.Location = None


class ResponseMessageBodyInfo:
    def __init__(self):
        self.Error = None
        self.Results = None
        self.TraceIds = None

class LoadOption:
    def __init__(self):
        self.select = None
        self.expand = None

class IResultHandler:
    def _handleResult(self, value) -> None:
        pass

class IRequestExecutor:
    def execute(self, requestInfo: httphelper.RequestInfo) -> httphelper.ResponseInfo:
        pass

class HttpRequestExecutor(IRequestExecutor):
    def execute(self, requestInfo: httphelper.RequestInfo) -> httphelper.ResponseInfo:
        return httphelper.HttpUtility.invoke(requestInfo)

class PipeExecutor(IRequestExecutor):
    def execute(self, request: httphelper.RequestInfo) -> httphelper.ResponseInfo:
        if request.method is None:
            request.method = "GET"
        request.method = request.method.upper()
        if request.method == "GET" or request.method == "DELETE":
            requestInfo.body = None

        index = request.url.find('/_api')
        if index < 0:
            raise Exception("Invalid url")
        indexPipeNameStart = request.url.rfind('/', 0, index)
        if indexPipeNameStart < 0:
            raise Exception("Cannot find pipename")
        pipeName = request.url[indexPipeNameStart + 1: index]
        index = index + len('/_api')
        pathAndQuery = request.url[index:];

        requestMessage = request.method + " " + pathAndQuery + " HTTP/1.1\n"
        requestMessage = requestMessage + "\n"
        if not Utility.isNullOrEmptyString(request.body):
            requestMessage = requestMessage + request.body

        requestBytes = requestMessage.encode('utf-8')
        pipeClient = namedpipeclient.WinNamedPipeClient(pipeName)
        pipeClient.write(requestBytes)
        responseBytes = pipeClient.read()
        pipeClient.close()
        responseMessage = responseBytes.decode('utf-8')

        ret = httphelper.ResponseInfo()
        headerLines = []
        lastNewLineIndex = -1
        index = responseMessage.find('\n', 0)
        while index > 0:
            if lastNewLineIndex > 0 and lastNewLineIndex + 1 == index:
                # Two consecutive newline
                break
            headerLines.append(responseMessage[lastNewLineIndex + 1 : index])
            lastNewLineIndex = index
            index = responseMessage.find('\n', index + 1)
        if index > 0:
            ret.body = responseMessage[index + 1: ]

        if len(headerLines) == 0:
            raise Utility.createRuntimeError(ErrorCodes.invalidArgument)
        statusLineParts = headerLines[0].split()
        if len(statusLineParts) < 2:
            raise Utility.createRuntimeError(ErrorCodes.invalidArgument)
        ret.statusCode = int(statusLineParts[0])
        return ret

class ClientResultProcessingType(enum.IntEnum):
    none = 0
    date = 1

class ClientResult(IResultHandler):
    def __init__(self, processingType: ClientResultProcessingType = ClientResultProcessingType.none):
        self._value = None
        self._isLoaded = False
        self._processingType = processingType

    @property
    def value(self):
        if not self._isLoaded:
            raise Utility.createRuntimeError(ErrorCodes.propertyNotLoaded)
        return self._value

    def _handleResult(self, value: any) -> None:
        self._isLoaded = True
        if isinstance(value, dict) and value.get("_IsNull"):
            return
        if self._processingType == ClientResultProcessingType.date:
            self._value = Utility.adjustToDateTime(value)
        else:
            self._value = value


class RequestUrlAndHeaderInfo:
    def __init__(self):
        self.url = None
        self.headers = {}

class RequestExecutionMode(enum.IntEnum):
    batch = 0
    instantSync = 1

class ClientRequestContext:
    customRequestExecutor = None
    defaultRequestUrlAndHeaders = None

    def __init__(self, url: str, executionMode: RequestExecutionMode):
        self.__nextId = 1
        self._requestHeaders = {}
        self._url = url
        if Utility.isNullOrEmptyString(self._url):
           if (ClientRequestContext.defaultRequestUrlAndHeaders is not None 
               and not Utility.isNullOrEmptyString(ClientRequestContext.defaultRequestUrlAndHeaders.url)):
               self._url = ClientRequestContext.defaultRequestUrlAndHeaders.url
               if (ClientRequestContext.defaultRequestUrlAndHeaders.headers):
                   for key, value in ClientRequestContext.defaultRequestUrlAndHeaders.headers.items():
                       self._requestHeaders[key] = value

        if Utility.isNullOrEmptyString(self._url):
            self._url = Constants.localDocument

        self._executionMode = executionMode
        if self._executionMode is None:
            self._executionMode = RequestExecutionMode.batch

        self._processingResult = False
        self._customData = Constants.iterativeExecutor
        if (ClientRequestContext.customRequestExecutor is not None):
            self._requestExecutor = ClientRequestContext.customRequestExecutor
        else:
            if self._url.startswith("pipe://"):
                self._requestExecutor = PipeExecutor()
            else:
                self._requestExecutor = HttpRequestExecutor()

        self._rootObject = None
        self.__pendingRequest = None

    @property
    def requestHeaders(self):
        return self._requestHeaders

    def _nextId(self) -> int:
        ret = self.__nextId
        self.__nextId = self.__nextId + 1
        return ret

    @property
    def executionMode(self) -> RequestExecutionMode:
        return self._executionMode

    @executionMode.setter
    def executionMode(self, value : RequestExecutionMode):
        self._executionMode = value

    @property
    def _pendingRequest(self) -> 'ClientRequest':
        if self.__pendingRequest is None:
            self.__pendingRequest = ClientRequest(self)
        return self.__pendingRequest

    def load(self, clientObject: 'ClientObject', option = None):
        Utility.validateContext(self, clientObject)
        queryOption = QueryInfo()
        if isinstance(option, str):
            select = option
            queryOption.Select = self._parseSelectExpand(select)
        elif isinstance(option, list):
            queryOption.Select =  option
        elif isinstance(option, LoadOption):
            loadOption = option
            if isinstance(loadOption.select, str):
                queryOption.Select = self._parseSelectExpand(loadOption.select)
            elif isinstance(loadOption.select, list):
                queryOption.Select = loadOption.select
            elif not Utility.isNullOrUndefined(loadOption.select):
                raise Utility.createInvalidArgumentException("option.select")

            if isinstance(loadOption.expand, str):
                queryOption.Expand = self._parseSelectExpand(loadOption.expand)
            elif isinstance(loadOption.expand, list):
                queryOption.Expand = loadOption.expand
            elif not Utility.isNullOrUndefined(loadOption.expand):
                raise Utility.createInvalidArgumentException("option.expand")

            if isinstance(loadOption.top, int):
                queryOption.Top = loadOption.top
            elif not Utility.isNullOrUndefined(loadOption.top):
                raise Utility.createInvalidArgumentException("option.top")

            if isinstance(loadOption.skip, int):
                queryOption.Skip = loadOption.skip
            elif not Utility.isNullOrUndefined(loadOption.skip):
                raise Utility.createInvalidArgumentException("option.skip")
        elif not Utility.isNullOrUndefined(option):
            raise Utility.createInvalidArgumentException("option")

        action = ActionFactory.createQueryAction(self, clientObject, queryOption)
        self._pendingRequest.addActionResultHandler(action, clientObject)
        if self._executionMode == RequestExecutionMode.instantSync :
            self.sync()

    def trace(self, message: str) -> None:
        ActionFactory.createTraceAction(self, message)

    def _parseSelectExpand(self, select: str) -> list:
        args = []
        if not Utility.isNullOrEmptyString(select):
            propertyNames = select.split(",")
            for propertyName in propertyNames:
                args.append(Utility.trim(propertyName))
        return args

    def sync(self):
        req = self.__pendingRequest
        if req is None:
            return
        if not req.hasActions:
            return

        self.__pendingRequest = None
        msgBody = req.buildRequestMessageBody()
        requestInfo = httphelper.RequestInfo();
        requestInfo.url = Utility.combineUrl(self._url, "ProcessQuery")
        requestInfo.body = json.dumps(msgBody, default = lambda o: o.__dict__)
        requestInfo.method = "POST"
        for key, value in self._requestHeaders.items():
            requestInfo.headers[key] = value
        requestInfo.headers["CONTENT-TYPE"] = "application/json"
        requestExecutor = self.customRequestExecutor
        if (requestExecutor is None):
            requestExecutor = self._requestExecutor
        responseInfo = requestExecutor.execute(requestInfo)
        if responseInfo.statusCode != 200:
            raise Utility.createRuntimeError("NetworkError")
        response = json.loads(responseInfo.body)
        req.processResponse(response)
        return

class Action:
    def __init__(self, actionInfo, isWriteOperation):
        self._actionInfo = actionInfo
        self._isWriteOperation = isWriteOperation
    
    @property
    def actionInfo(self):
        return self._actionInfo

    @property
    def isWriteOperation(self):
        return self._isWriteOperation

class ObjectPath:
    def __init__(self, objectPathInfo: ObjectPathInfo, parentObjectPath: 'ObjectPath', isCollection: bool, isInvalidAfterRequest: bool):
        self._objectPathInfo = objectPathInfo
        self._parentObjectPath = parentObjectPath
        self._isWriteOperation = False
        self._isCollection = isCollection
        self._isInvalidAfterRequest = isInvalidAfterRequest
        self._isValid = True
        self._argumentObjectPaths = None

    @property
    def objectPathInfo(self) -> ObjectPathInfo:
        return self._objectPathInfo

    @property
    def isWriteOperation(self) -> bool:
        return self._isWriteOperation

    @isWriteOperation.setter
    def isWriteOperation(self, value):
        self._isWriteOperation = value

    @property
    def isCollection(self) -> bool:
        return self._isCollection

    @property
    def isInvalidAfterRequest(self) -> bool:
        return self._isInvalidAfterRequest

    @property
    def parentObjectPath(self) -> 'ObjectPath':
        return self._parentObjectPath

    @property
    def argumentObjectPaths(self):
        return self._argumentObjectPaths

    @argumentObjectPaths.setter
    def argumentObjectPaths(self, value):
        self._argumentObjectPaths = value

    @property
    def isValid(self) -> bool:
        return self._isValid

    @isValid.setter
    def isValid(self, value):
        self._isValid = value


    def updateUsingObjectData(self, value) -> None:
        referenceId = value.get(Constants.referenceId, None)
        if not Utility.isNullOrEmptyString(referenceId):
            self._isInvalidAfterRequest = False
            self._isValid = True
            self._objectPathInfo.ObjectPathType = ObjectPathType.ReferenceId
            self._objectPathInfo.Name = referenceId
            self._objectPathInfo.ArgumentInfo = {}
            self._parentObjectPath = None
            self._argumentObjectPaths = None
            return

        if self.parentObjectPath and self.parentObjectPath.isCollection:
            id = value.get(Constants.id, None)
            if Utility.isNullOrUndefined(id):
                id = value.get(Constants.idPrivate, None)

            if not Utility.isNullOrUndefined(id):
                self._isInvalidAfterRequest = False
                self._isValid = True
                self._objectPathInfo.ObjectPathType = ObjectPathType.Indexer
                self._objectPathInfo.Name = ""
                self._objectPathInfo.ArgumentInfo = ArgumentInfo()
                self._objectPathInfo.ArgumentInfo.Arguments = [id]
                self._argumentObjectPaths = None
                return

    def _updateAsNullObject(self) -> None:
        self._isInvalidAfterRequest = False;
        self._isValid = True;
        self._objectPathInfo.ObjectPathType = ObjectPathType.NullObject;
        self._objectPathInfo.Name = "";
        self._objectPathInfo.ArgumentInfo = {};
        self._parentObjectPath = None;
        self._argumentObjectPaths = None;
        return

class ClientRequest:
    def __init__(self, context):
        self._context = context
        self._actions = []
        self._actionResultHandler = {}
        self._referencedObjectPaths = {}
        self._flags = ClientRequestFlags.NoneValue
        self._traceInfos = {}

    @property
    def context(self):
        return self._context


    @property
    def flags(self) -> ClientRequestFlags:
        return self._flags

    @property
    def traceInfos(self):
        return self._traceInfos

    def addAction(self, action: Action):
        if action.isWriteOperation:
            self._flags = self._flags | ClientRequestFlags.WriteOperation
        self._actions.append(action)
    
    @property
    def hasActions(self) -> bool:
        return len(self._actions) > 0


    def addTrace(self, actionId: int, message: str):
        self._traceInfos[actionId] = message

    def addReferencedObjectPath(self, objectPath: ObjectPath):
        if self._referencedObjectPaths.get(objectPath.objectPathInfo.Id):
            return

        if not objectPath.isValid:
            raise Utility.createRuntimeError(ErrorCodes.invalidObjectPath, Utility._getResourceString(ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath)))

        while objectPath:
            if objectPath.isWriteOperation:
                self._flags = self._flags | ClientRequestFlags.WriteOperation

            self._referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath

            if objectPath.objectPathInfo.ObjectPathType == ObjectPathType.Method:
                self.addReferencedObjectPaths(objectPath.argumentObjectPaths)

            objectPath = objectPath.parentObjectPath

    def addReferencedObjectPaths(self, objectPaths: list):
        if objectPaths:
            for objectPath in objectPaths:
                self.addReferencedObjectPath(objectPath)

    def addActionResultHandler(self, action: Action, resultHandler: IResultHandler):
        self._actionResultHandler[action.actionInfo.Id] = resultHandler

    def buildRequestMessageBody(self) -> RequestMessageBodyInfo:
        objectPaths = {}
        for k, v in self._referencedObjectPaths.items():
            objectPaths[k] = v.objectPathInfo

        actions = []
        for action in self._actions:
            actions.append(action.actionInfo)

        ret = RequestMessageBodyInfo()
        ret.Actions = actions
        ret.ObjectPaths = objectPaths

        return ret

    def processResponse(self, msg: dict) -> None:
        if msg and msg.get("Results"):
            for actionResult in msg.get("Results"):
                actionId = actionResult.get("ActionId")
                handler = self._actionResultHandler.get(actionId)
                if (handler):
                    actionValue = actionResult.get("Value")
                    handler._handleResult(actionValue)

    def invalidatePendingInvalidObjectPaths(self):
        for i in self._referencedObjectPaths:
            if self._referencedObjectPaths[i].isInvalidAfterRequest:
                self._referencedObjectPaths[i].isValid = False

class ClientObject(IResultHandler):
    def __init__(self, context: ClientRequestContext, objectPath: ObjectPath):
        Utility.checkArgumentNull(context, "context")
        self._context = context
        self.__objectPath = objectPath
        self.__isLoaded = False
        # set __isNull to be undefined as we do not know the value yet
        self.__isNull = None
        if self.__objectPath:
            if not context._processingResult:
                ActionFactory.createInstantiateAction(context, self)

    @property
    def context(self) -> ClientRequestContext:
        return self._context

    @property
    def isNullObject(self) -> bool:
        if self.__isNull is None and self.context.executionMode == RequestExecutionMode.instantSync:
            self.context.load(self)
        if self.__isNull is None:
            raise Utility.createRuntimeError(ErrorCodes.propertyNotLoaded)
        return self.__isNull

    @property
    def _objectPath(self) -> ObjectPath:
        return self.__objectPath

    @_objectPath.setter
    def _objectPath(self, value: ObjectPath):
        self.__objectPath = value

    @property
    def _isLoaded(self) -> bool:
        return self.__isLoaded

    @property
    def _isNull(self):
        return self.__isNull

    def _setAsNullObject(self):
        if self.__objectPath:
            self.__objectPath._updateAsNullObject()

    def _handleResult(self, value):
        self.__isNull = Utility.isNullOrUndefined(value)
        if self.__isNull and self.__objectPath:
            self.__objectPath._updateAsNullObject()
        self.__isLoaded = True
        Utility.fixObjectPathIfNecessary(self, value)

    def _handleIdResult(self, value):
        self.__isNull = Utility.isNullOrUndefined(value)
        if self.__isNull and self.__objectPath:
            self.__objectPath._updateAsNullObject()
        Utility.fixObjectPathIfNecessary(self, value)


class ActionFactory:
    @staticmethod
    def createSetPropertyAction(context: ClientRequestContext, parent: ClientObject, propertyName: str, value):
        Utility.validateObjectPath(parent)
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.SetProperty
        actionInfo.Name = propertyName
        actionInfo.ObjectPathId = parent._objectPath.objectPathInfo.Id
        actionInfo.ArgumentInfo = ArgumentInfo()
        args = [value]
        referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args)
        Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths)
        ret = Action(actionInfo, True)
        context._pendingRequest.addAction(ret)
        context._pendingRequest.addReferencedObjectPath(parent._objectPath)
        context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths)
        return ret

    @staticmethod
    def createMethodAction(context: ClientRequestContext, parent: ClientObject, methodName: str, operationType, args):
        Utility.validateObjectPath(parent)
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.Method
        actionInfo.Name = methodName
        actionInfo.ObjectPathId = parent._objectPath.objectPathInfo.Id
        actionInfo.ArgumentInfo = ArgumentInfo()
        referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args)
        Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths)
        isWriteOperation = operationType != OperationType.Read
        ret = Action(actionInfo, isWriteOperation)
        context._pendingRequest.addAction(ret)
        context._pendingRequest.addReferencedObjectPath(parent._objectPath)
        context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths)
        return ret

    @staticmethod
    def createQueryAction(context: ClientRequestContext, parent: ClientObject, queryInfo): 
        Utility.validateObjectPath(parent)
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.Query
        actionInfo.Name = ""
        actionInfo.ObjectPathId = parent._objectPath.objectPathInfo.Id
        actionInfo.QueryInfo = queryInfo
        ret = Action(actionInfo, False)
        context._pendingRequest.addAction(ret)
        context._pendingRequest.addReferencedObjectPath(parent._objectPath)
        return ret

    @staticmethod
    def createInstantiateAction(context: ClientRequestContext, clientObject: ClientObject):
        Utility.validateObjectPath(clientObject)
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.Instantiate
        actionInfo.Name = ""
        actionInfo.ObjectPathId = clientObject._objectPath.objectPathInfo.Id
        ret = Action(actionInfo, False)
        context._pendingRequest.addAction(ret)
        context._pendingRequest.addReferencedObjectPath(clientObject._objectPath)
        handler = InstantiateActionResultHandler(clientObject)
        context._pendingRequest.addActionResultHandler(ret, handler)
        return ret

    @staticmethod
    def createTraceAction(context: ClientRequestContext, message: str) -> Action:
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.Trace
        actionInfo.Name = "Trace"
        actionInfo.ObjectPathId = 0
        ret = Action(actionInfo, False)
        context._pendingRequest.addAction(ret)
        context._pendingRequest.addTrace(actionInfo.Id, message)
        return ret


class ObjectPathFactory:
    @staticmethod
    def createGlobalObjectObjectPath(context: ClientRequestContext):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.GlobalObject
        objectPathInfo.Name = ""
        isCollection = False
        isInvalidAfterRequest = False
        return ObjectPath(objectPathInfo, None, isCollection, isInvalidAfterRequest)

    @staticmethod
    def createNewObjectObjectPath(context: ClientRequestContext, typeName: str, isCollection: bool):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.NewObject
        objectPathInfo.Name = typeName
        isInvalidAfterRequest = False
        return ObjectPath(objectPathInfo, None, isCollection, isInvalidAfterRequest)

    @staticmethod
    def createPropertyObjectPath(context: ClientRequestContext, parent: ClientObject, propertyName: str, isCollection: bool, isInvalidAfterRequest: bool) -> ObjectPath:
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Property
        objectPathInfo.Name = propertyName
        objectPathInfo.ParentObjectPathId = parent._objectPath.objectPathInfo.Id
        return ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest)
    

    @staticmethod
    def createIndexerObjectPath(context: ClientRequestContext, parentObject: ClientObject, args: list):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Indexer
        objectPathInfo.Name = ""
        objectPathInfo.ParentObjectPathId = parentObject._objectPath.objectPathInfo.Id
        objectPathInfo.ArgumentInfo = ArgumentInfo()
        objectPathInfo.ArgumentInfo.Arguments = args
        isCollection = False
        isInvalidAfterRequest = False
        return ObjectPath(objectPathInfo, parentObject._objectPath, isCollection, isInvalidAfterRequest)
    

    @staticmethod
    def createIndexerObjectPathUsingParentPath(context: ClientRequestContext, parentObjectPath: ObjectPath, args: list):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Indexer
        objectPathInfo.Name = ""
        objectPathInfo.ParentObjectPathId = parentObjectPath.objectPathInfo.Id
        objectPathInfo.ArgumentInfo = ArgumentInfo()
        objectPathInfo.ArgumentInfo.Arguments = args
        isCollection = False
        isInvalidAfterRequest = False
        return ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest)
    
    @staticmethod
    def createMethodObjectPath(context: ClientRequestContext, parentObject: ClientObject, methodName: str, operationType, args, isCollection: bool, isInvalidAfterRequest: bool):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Method
        objectPathInfo.Name = methodName
        objectPathInfo.ParentObjectPathId = parentObject._objectPath.objectPathInfo.Id
        objectPathInfo.ArgumentInfo = ArgumentInfo()

        argumentObjectPaths = Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args)
        ret = ObjectPath(objectPathInfo, parentObject._objectPath, isCollection, isInvalidAfterRequest)
        ret.argumentObjectPaths = argumentObjectPaths
        ret.isWriteOperation = (operationType != OperationType.Read)
        return ret
    
    @staticmethod
    def createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod: bool, context: ClientRequestContext, parentObject: ClientObject, childItem, index):
        id = childItem.get(Constants.id, None)
        if Utility.isNullOrUndefined(id):
            id = childItem.get(Constants.idPrivate)

        if hasIndexerMethod and not Utility.isNullOrUndefined(id):
            return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parentObject, childItem)
        else:
            return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parentObject, childItem, index)

    @staticmethod
    def createChildItemObjectPathUsingIndexer(context: ClientRequestContext, parentObject: ClientObject, childItem):
        id = childItem.get(Constants.id)
        if Utility.isNullOrUndefined(id):
            id = childItem.get(Constants.idPrivate)

        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id =context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Indexer
        objectPathInfo.Name = ""
        objectPathInfo.ParentObjectPathId = parentObject._objectPath.objectPathInfo.Id
        objectPathInfo.ArgumentInfo = ArgumentInfo()
        objectPathInfo.ArgumentInfo.Arguments = [id]
        isCollection = False
        isInvalidAfterRequest = False
        return ObjectPath(objectPathInfo, parentObject._objectPath, isCollection, isInvalidAfterRequest)
    
    @staticmethod
    def createChildItemObjectPathUsingGetItemAt(context: ClientRequestContext, parentObject: ClientObject, childItem, index):
        indexFromServer = childItem.get(Constants.index)
        if indexFromServer:
            index = indexFromServer

        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Method
        objectPathInfo.Name = Constants.getItemAt
        objectPathInfo.ParentObjectPathId = parent._objectPath.objectPathInfo.Id
        objectPathInfo.ArgumentInfo = ArgumentInfo()
        objectPathInfo.ArgumentInfo.Arguments = [index]
        isCollection = False
        isInvalidAfterRequest = False
        return ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest)

class InstantiateActionResultHandler(IResultHandler):
    def __init__(self, clientObject: ClientObject):
        self._clientObject = clientObject

    def _handleResult(self, value):
        self._clientObject._handleIdResult(value)

class ResourceStrings:
    invalidObjectPath = "InvalidObjectPath"
    propertyNotLoaded = "PropertyNotLoaded"
    invalidRequestContext = "InvalidRequestContext"
    invalidArgument = "InvalidArgument"
    runMustReturnPromise = "RunMustReturnPromise"

class ErrorCodes:
    invalidObjectPath = "InvalidObjectPath"
    propertyNotLoaded = "PropertyNotLoaded"
    invalidRequestContext = "InvalidRequestContext"
    invalidArgument = "InvalidArgument"

class RuntimeException(Exception):
    def __init__(self, code: str, message: str, location = None):
        self.code = code
        self.message = message
        self.location = location
    def __str__(self):
        return repr(self.code) + ":" + repr(self.message)

class Utility:
    @staticmethod
    def checkArgumentNull(value, name: str):
        if (Utility.isNullOrUndefined(value)):
            raise Utility.createInvalidArgumentException(name)

    @staticmethod
    def isUndefined(value) -> bool:
        if (value is None):
            return True
        return False

    @staticmethod
    def isNullOrUndefined(value) -> bool:
        if (value is None):
            return True
        return False

    @staticmethod
    def isNullOrEmptyString(value: str) -> bool:
        if (value is None):
            return True

        if (len(value) == 0):
            return True

        return False

    @staticmethod
    def trim(value: str) -> str:
        return value.strip()

    @staticmethod
    def combineUrl(url1: str, url2: str) -> str:
        if url1 is None:
            return url2
        if url2 is None:
            return url1
        if not url1.endswith("/"):
            url1 = url1 + "/"
        if url2.startswith("/"):
            url2 = url2[1:]
        return url1 + url2

    @staticmethod
    def caseInsensitiveCompareString(str1: str, str2: str) -> bool:
        if (Utility.isNullOrUndefined(str1)):
            return Utility.isNullOrUndefined(str2)
        else:
            if (Utility.isNullOrUndefined(str2)):
                return False
            else:
                return str1.lower() == str2.lower()

    @staticmethod
    def isReadonlyRestRequest(method: str) -> bool:
        return Utility.caseInsensitiveCompareString(method, "GET")

    @staticmethod
    def setMethodArguments(context: ClientRequestContext, argumentInfo: ArgumentInfo, args: list) -> list:
        if (Utility.isNullOrUndefined(args)):
            return None

        referencedObjectPaths = []
        referencedObjectPathIds = []
        hasOne = Utility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds)
        argumentInfo.Arguments = args
        if hasOne:
            argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds
            return referencedObjectPaths

        return None

    @staticmethod
    def collectObjectPathInfos(context: ClientRequestContext, args: list, referencedObjectPaths: list, referencedObjectPathIds: list) -> bool:        
        hasOne = False
        for i in range(0, len(args)):
            arg = args[i]
            if isinstance(arg, ClientObject):
                clientObject = arg
                Utility.validateContext(context, clientObject)
                args[i] = clientObject._objectPath.objectPathInfo.Id
                referencedObjectPathIds.append(clientObject._objectPath.objectPathInfo.Id)
                referencedObjectPaths.append(clientObject._objectPath)
                hasOne = True
            elif isinstance(arg, list):
                childArrayObjectPathIds = []
                childArrayHasOne = Utility.collectObjectPathInfos(context, arg, referencedObjectPaths, childArrayObjectPathIds)
                if childArrayHasOne:
                    referencedObjectPathIds.append(childArrayObjectPathIds)
                    hasOne = True
                else:
                    referencedObjectPathIds.append(0)
            else:
                referencedObjectPathIds.append(0)
        return hasOne

    @staticmethod
    def fixObjectPathIfNecessary(clientObject: ClientObject, value: dict):
        if (clientObject and clientObject._objectPath and value):
            clientObject._objectPath.updateUsingObjectData(value)

    @staticmethod
    def validateObjectPath(clientObject: ClientObject):
        objectPath = clientObject._objectPath
        while (objectPath):
            if (not objectPath.isValid):
                pathExpression = Utility.getObjectPathExpression(objectPath)
                raise Utility.createRuntimeError(ErrorCodes.invalidObjectPath, Utility._getResourceString(ResourceStrings.invalidObjectPath, pathExpression))
            objectPath = objectPath.parentObjectPath

    @staticmethod
    def validateReferencedObjectPaths(objectPaths: list):
        if (objectPaths):
            for objectPathLoop in objectPaths:
                objectPath = objectPathLoop
                while objectPath:
                    if (not objectPath.isValid):
                        pathExpression = Utility.getObjectPathExpression(objectPath)
                        raise Utility.createRuntimeError(ErrorCodes.invalidObjectPath, Utility._getResourceString(ResourceStrings.invalidObjectPath, pathExpression))
                    objectPath = objectPath.parentObjectPath

    @staticmethod
    def validateContext(context: ClientRequestContext, clientObject: ClientObject):
        if (clientObject and clientObject.context != context):
            raise Utility.createRuntimeError(ErrorCodes.invalidRequestContext, Utility._getResourceString(ResourceStrings.invalidRequestContext))


    _logEnabled = False

    @staticmethod
    def log(message: str): 
        if Utility._logEnabled:
            logging.info(message)

    @staticmethod
    def load(clientObj: ClientObject, option):
        clientObj.context.load(clientObj, option)

    _underscoreCharCode = ord('_')

    @staticmethod
    def createRuntimeError(code: str, message: str = None, location: str = None) -> Exception:
        return RuntimeError(code, message, location)
    
    @staticmethod
    def createInvalidArgumentException(name: str) -> Exception:
        return RuntimeError(ErrorCodes.invalidArgument, Utility._getResourceString(ResourceStrings.invalidArgument, name));

    @staticmethod
    def _getResourceString(resourceId: str, arg: str = None) -> str:
        ret = resourceId
        """
        if ((<any>window).Strings && (<any>window).Strings.OfficeOM) {
            var stringName = "L_" + resourceId
            var stringValue = (<any>window).Strings.OfficeOM[stringName]
            if (stringValue) {
                ret = stringValue
            }
        }

        if (!Utility.isNullOrUndefined(arg)) {
            ret = ret.replace("{0}", arg)
        }
        """
        return ret
    
    @staticmethod
    def throwIfNotLoaded(propertyName: str, fieldValue, entityName: str, isNull: bool):
        # isNull could be None, True or False. When we do not know whether the object
        # is null or not, isNull is None
        if (not isNull and fieldValue is None and propertyName[0] != "_"):
            raise Utility.createRuntimeError(ErrorCodes.propertyNotLoaded, Utility._getResourceString(ResourceStrings.propertyNotLoaded, propertyName))

    @staticmethod
    def loadIfInstantSyncExecutionMode(clientObject: ClientObject, propertyName: str, fieldValue) -> None:
        if clientObject.context.executionMode == RequestExecutionMode.instantSync and clientObject._isLoaded == False:
            clientObject.load();

    @staticmethod
    def syncIfInstantSyncExecutionMode(clientObject: ClientObject) -> None:
        if clientObject.context.executionMode == RequestExecutionMode.instantSync:
            clientObject.context.sync()

    @staticmethod
    def getObjectPathExpression(objectPath: ObjectPath) -> str:
        ret = ""
        while (objectPath) :
            if objectPath.objectPathInfo.ObjectPathType == ObjectPathType.GlobalObject:
                ret = ret
            elif objectPath.objectPathInfo.ObjectPathType == ObjectPathType.NewObject:
                if (len(ret) > 0):
                    ret = "." + ret
                ret = "new()" + ret
            elif objectPath.objectPathInfo.ObjectPathType == ObjectPathType.Method:
                if (len(ret) > 0):
                    ret = "." + ret
                ret = Utility.normalizeName(objectPath.objectPathInfo.Name) + "()" + ret
            elif objectPath.objectPathInfo.ObjectPathType == ObjectPathType.Property:
                if len(ret) > 0:
                    ret = "." + ret
                ret = Utility.normalizeName(objectPath.objectPathInfo.Name) + ret
            elif objectPath.objectPathInfo.ObjectPathType == ObjectPathType.Indexer:
                if len(ret) > 0:
                    ret = "." + ret
                ret = "getItem()" + ret
            elif objectPath.objectPathInfo.ObjectPathType == ObjectPathType.ReferenceId:
                if len(ret) > 0:
                    ret = "." + ret
                ret = "_reference()" + ret

            objectPath = objectPath.parentObjectPath

        return ret
    
    @staticmethod
    def _addActionResultHandler(clientObj: ClientObject, action: Action, resultHandler: IResultHandler):
        clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler)

    @staticmethod
    def normalizeName(name: str) -> str:
        return name[0:1].lower() + name[1:]

    @staticmethod
    def adjustToDateTime(value):
        if isinstance(value, str):
            return datetime.datetime.strptime(value, "%Y-%m-%dT%H:%M:%S.%fZ")
        elif isinstance(value, list):
            return map(Utility.adjustToDateTime, value)
        return value


