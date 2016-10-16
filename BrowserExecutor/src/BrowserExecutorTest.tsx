import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {BrowserExecutor} from './BrowserExecutor'
import {Utility} from './Utility'

export class BrowserExecutorTest extends React.Component<any, any>{
    private m_executor:BrowserExecutor;
    constructor(props: any){
        super(props);
        this.m_executor = new BrowserExecutor("ws://localhost:8888/ws");
    }

    private onButtonClick(evt){
        var message = (this.refs["TxtMessage"] as any).value;
        this.m_executor.sendRequest(message)
            .then((resp: string)=>{
                Utility.log(resp);
            });
    }

    public render(){
        return (
        <div>
            <input type="text" ref="TxtMessage"/>
            <input type="button" value="Send" onClick={this.onButtonClick.bind(this)} />
        </div>
        );
    }
}