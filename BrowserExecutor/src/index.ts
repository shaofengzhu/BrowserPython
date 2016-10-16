import * as React from 'react'
import * as ReactDOM from 'react-dom'
import {Hello} from "Hello"
import {BrowserExecutorTest} from "./BrowserExecutorTest"

ReactDOM.render(
    React.createElement(BrowserExecutorTest, null),
    document.getElementById("main-content")
);
