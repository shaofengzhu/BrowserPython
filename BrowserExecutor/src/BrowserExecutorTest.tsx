import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {BrowserExecutor} from './BrowserExecutor'
import {Utility} from './Utility'

export interface BrowserExecutorTestProps{
    webSocketUrl: string
}

export class BrowserExecutorTest extends React.Component<BrowserExecutorTestProps, any>{
    private m_executor:BrowserExecutor;
    constructor(props: BrowserExecutorTestProps){
        super(props);
        this.m_executor = new BrowserExecutor(props.webSocketUrl);
    }

    public sendRequest(name: string){
        this.m_executor.sendRequest(name)
        .then((resp: string) => {
            Utility.log(resp);
        });
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