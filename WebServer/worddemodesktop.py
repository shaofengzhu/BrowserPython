import runtime
import word
import worddemolib

if __name__ == "__main__":
    worddemolib.WordDemoLib.initDesktopContext()
    context = word.RequestContext()
    worddemolib.WordDemoLib.helloWorld(context)
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None
