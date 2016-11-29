import runtime
import excel
context = excel.RequestContext("http://localhost:8052")
context.executionMode = runtime.RequestExecutionMode.instantSync