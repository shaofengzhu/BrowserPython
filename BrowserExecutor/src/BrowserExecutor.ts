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
    Body: string;
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
                Utility.log("Sending Executing at client result for " + msg.Id);
                Utility.log("Message Body is " + TestData.response);

                var respMsg: Message = {
                    Id: msg.Id,
                    SubId: msg.SubId,
                    Type: BrowserExecutor.MessageTypeExecuteAtClientResult,
                    Body: TestData.response
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

class TestData{
    static request = "{\"Actions\": [{\"Id\": 2, \"ObjectPathId\": 1, \"ArgumentInfo\": null, \"ActionType\": 1, \"QueryInfo\": null, \"Name\": \"\"}, {\"Id\": 4, \"ObjectPathId\": 3, \"ArgumentInfo\": null, \"ActionType\": 1, \"QueryInfo\": null, \"Name\": \"\"}, {\"Id\": 6, \"ObjectPathId\": 5, \"ArgumentInfo\": null, \"ActionType\": 1, \"QueryInfo\": null, \"Name\": \"\"}, {\"Id\": 8, \"ObjectPathId\": 7, \"ArgumentInfo\": null, \"ActionType\": 1, \"QueryInfo\": null, \"Name\": \"\"}, {\"Id\": 10, \"ObjectPathId\": 9, \"ArgumentInfo\": null, \"ActionType\": 1, \"QueryInfo\": null, \"Name\": \"\"}, {\"Id\": 11, \"ObjectPathId\": 9, \"ArgumentInfo\": {\"Arguments\": [null]}, \"ActionType\": 3, \"QueryInfo\": null, \"Name\": \"Clear\"}, {\"Id\": 12, \"ObjectPathId\": 9, \"ArgumentInfo\": {\"Arguments\": [[[\"Bellevue\", \"Redmond\"], [1234, \"=A2 + 100\"]]]}, \"ActionType\": 4, \"QueryInfo\": null, \"Name\": \"Values\"}, {\"Id\": 13, \"ObjectPathId\": 9, \"ArgumentInfo\": null, \"ActionType\": 2, \"QueryInfo\": {\"Expand\": null, \"Skip\": null, \"Select\": null, \"Top\": null}, \"Name\": \"\"}], \"ObjectPaths\": {\"7\": {\"Id\": 7, \"ObjectPathType\": 3, \"ParentObjectPathId\": 5, \"Name\": \"GetCell\", \"ArgumentInfo\": {\"Arguments\": [0, 0]}}, \"1\": {\"Id\": 1, \"ObjectPathType\": 1, \"ParentObjectPathId\": 0, \"Name\": \"\", \"ArgumentInfo\": null}, \"3\": {\"Id\": 3, \"ObjectPathType\": 4, \"ParentObjectPathId\": 1, \"Name\": \"Worksheets\", \"ArgumentInfo\": null}, \"5\": {\"Id\": 5, \"ObjectPathType\": 5, \"ParentObjectPathId\": 3, \"Name\": \"\", \"ArgumentInfo\": {\"Arguments\": [\"Sheet1\"]}}, \"9\": {\"Id\": 9, \"ObjectPathType\": 3, \"ParentObjectPathId\": 7, \"Name\": \"GetResizedRange\", \"ArgumentInfo\": {\"Arguments\": [1, 1]}}}}";

    static response = "{\"Results\":[{\"ActionId\":2,\"Value\":{}},{\"ActionId\":4,\"Value\":{}},{\"ActionId\":6,\"Value\":{\"Id\":\"{00000000-0001-0000-0000-000000000000}\"}},{\"ActionId\":8,\"Value\":{}},{\"ActionId\":10,\"Value\":{}},{\"ActionId\":11,\"Value\":null},{\"ActionId\":12,\"Value\":null},{\"ActionId\":13,\"Value\":{\"_ReferenceId\":null,\"Address\":\"Sheet1!A1:B2\",\"AddressLocal\":\"Sheet1!A1:B2\",\"CellCount\":4,\"ColumnCount\":2,\"ColumnHidden\":false,\"ColumnIndex\":0,\"Formulas\":[[\"Bellevue\",\"Redmond\"],[1234,\"=A2 + 100\"]],\"FormulasLocal\":[[\"Bellevue\",\"Redmond\"],[1234,\"=A2 + 100\"]],\"FormulasR1C1\":[[\"Bellevue\",\"Redmond\"],[1234,\"=RC[-1] + 100\"]],\"Hidden\":false,\"NumberFormat\":[[\"General\",\"General\"],[\"General\",\"General\"]],\"RowCount\":2,\"RowHidden\":false,\"RowIndex\":0,\"Text\":[[\"Bellevue\",\"Redmond\"],[\"1234\",\"1334\"]],\"Values\":[[\"Bellevue\",\"Redmond\"],[1234,1334]],\"ValueTypes\":[[\"String\",\"String\"],[\"Double\",\"Double\"]]}}]}";
}
