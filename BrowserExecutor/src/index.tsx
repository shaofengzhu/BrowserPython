import * as React from 'react'
import * as ReactDOM from 'react-dom'
import {Hello} from "Hello"
import {BrowserExecutorTest} from "./BrowserExecutorTest"

var elem =React.createElement(BrowserExecutorTest, {webSocketUrl: "ws://localhost:7080/ws"});
ReactDOM.render(
    elem,
    document.getElementById("main-content")
);
 