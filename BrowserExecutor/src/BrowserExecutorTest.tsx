import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {BrowserExecutor} from './BrowserExecutor'
import {Utility} from './Utility'

export interface BrowserExecutorTestProps{
    webSocketUrl: string
}

export interface BrowserExecutorTestState{
    message?: string,
    code?: string
}

export class BrowserExecutorTest extends React.Component<BrowserExecutorTestProps, BrowserExecutorTestState>{
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

    private onButtonClick(evt: MouseEvent){
        var message = (this.refs["TxtMessage"] as any).value;
        this.m_executor.sendRequest(message)
            .then((resp: string)=>{
                Utility.log(resp);
            });
    }

    private onExecuteButtonClick(evt: MouseEvent){
        var code = this.state.code;
        this.m_executor.executeCode(code)
            .then((resp: string)=>{
                Utility.log(resp);
            });
    }

    private handleMessageChange(evt){
        this.setState({message: evt.target.value as string});
    }

    private handleCodeChange(evt){
        this.setState({code: evt.target.value as string});
    }

    public render(){
        return (
        <div>
            Cmd:
            <input type="text" ref="TxtMessage" onChange={this.handleMessageChange.bind(this)} />
            <input type="button" value="Send" onClick={this.onButtonClick.bind(this)} />
            <br />
            Code:
            <br />
            <textarea ref="TxtCode" rows="10" cols="50" onChange={this.handleCodeChange.bind(this)} />
            <br />
            <input type="button" value="Execute" onClick={this.onExecuteButtonClick.bind(this)} />
        </div>
        );
    }
}