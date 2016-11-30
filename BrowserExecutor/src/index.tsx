import * as React from 'react'
import * as ReactDOM from 'react-dom'
import {Hello} from "./Hello"
import {BrowserExecutorTest} from "./BrowserExecutorTest"
import MonacoEditor from 'react-monaco-editor';

var elem = ReactDOM.render(
    <BrowserExecutorTest webSocketUrl="ws://localhost:7080/ws" />,
    document.getElementById("main-content")
);
 
window["runDemo"] = function runDemo(){
    (elem as any).sendRequest("demo");
}
