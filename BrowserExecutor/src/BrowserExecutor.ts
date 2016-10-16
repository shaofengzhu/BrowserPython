import {Utility} from './Utility'
import {OfficeExtension} from './../lib/OfficeExtension'

interface RequestInfo{
    id: number;
    promiseResolveFunc: any;
    promiseRejectFunc: any;
}

interface Message {
    Type: string;
    Id: number;
    SubId: number;
    Body: any;
}

export class BrowserExecutor {
    private static s_nextMessageId = 1;

    private m_requestMap: {[index:number]: RequestInfo} = {};    

    private m_socket: WebSocket;
    private m_url: string;
    private static MessageTypeExecuteAtServer = "ExecuteAtServer";
    private static MessageTypeExecuteAtServerResult = "ExecuteAtServerResult";
    private static MessageTypeExecuteAtClient = "ExecuteAtClient";
    private static MessageTypeExecuteAtClientResult = "ExecuteAtClientResult";
    constructor(url: string) {
        this.m_url = url;
        this.createWebSocket();
    }

    private static getNextMessageId(): number{
        var ret = BrowserExecutor.s_nextMessageId;
        BrowserExecutor.s_nextMessageId++;
        return ret;
    }

    public sendRequest(message: string): OfficeExtension.IPromise<string>{        
        return new OfficeExtension.Promise((resolve, reject) =>{
            var requestInfo: RequestInfo = {
                    id: BrowserExecutor.getNextMessageId(),
                    promiseResolveFunc: resolve,
                    promiseRejectFunc: reject
                };

            this.m_requestMap[requestInfo.id] = requestInfo;
            var msg : Message = {
                Id: requestInfo.id,
                SubId: 0,
                Type: BrowserExecutor.MessageTypeExecuteAtServer,
                Body: message
            };

            this.m_socket.send(JSON.stringify(msg));
        });
    }

    private processReceivedMessage(messageString: string): void {
        Utility.log("Received: " + messageString);
        if (Utility.isNullOrEmptyString(messageString)) {
            Utility.log("Received empty message");
            return;
        }        

        var msg: Message = JSON.parse(messageString);
        if (Utility.isNullOrEmptyString(msg.Type)) {
            Utility.log("Unknown message: " + messageString);
            return;
        }

        if (msg.Type === BrowserExecutor.MessageTypeExecuteAtClient) {
            this.processExecuteAtClient(msg);
        }
        else if (msg.Type == BrowserExecutor.MessageTypeExecuteAtServerResult) {
            Utility.log("ExecuteAtServerResult:" + msg.Body);
            var requestInfo = this.m_requestMap[msg.Id];
            requestInfo.promiseResolveFunc(msg.Body);
        }
    }

    private processExecuteAtClient(msg: Message): void {
        setTimeout(
            () => {
                var str = new Date().toISOString();
                Utility.log("Sending Executing at client result for " + msg.Id);
                Utility.log("Message Body is " + str);

                var respMsg: Message = {
                    Id: msg.Id,
                    SubId: msg.SubId,
                    Type: BrowserExecutor.MessageTypeExecuteAtClientResult,
                    Body: str
                };

                this.m_socket.send(JSON.stringify(respMsg));
            },
            2000);
    }

    private createWebSocket(): void {
        Utility.log("creating socket " + this.m_url);
        this.m_socket = new WebSocket(this.m_url);
        this.m_socket.onclose = () => {
            Utility.log("socket closed: " + this.m_url);
        }

        this.m_socket.onopen = () => {
            Utility.log("socket opened: " + this.m_url);
        }

        this.m_socket.onmessage = (evt) => {
            this.processReceivedMessage(evt.data);
        }
    }
}

