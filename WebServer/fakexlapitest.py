import runtime
import fakexlapi

class FakeExcelTest:
    @staticmethod
    def test_basic():
        ctx = fakexlapi.RequestContext()
        ctx.application.activeWorkbook.sheets.load()
        ctx.sync()
        print("Sheets:")
        for sheet in ctx.application.activeWorkbook.sheets.items:
            print(sheet.name)
    
    @staticmethod
    def test_workbook():
        ctx = fakexlapi.RequestContext()
        workbook = ctx.application.activeWorkbook
        ctx.load(workbook);
        range = workbook.activeWorksheet.range("A1")
        result1 = range.replaceValue("Hello");
        result2 = range.replaceValue("HelloWorld");
        ctx.load(range, "Value");
        range.activate();
        ctx.sync()
        print(result1.value)
        print(result2.value)
        print(range.value)

    @staticmethod
    def test_updateValue():
        ctx = fakexlapi.RequestContext()
        workbook = ctx.application.activeWorkbook
        range = workbook.activeWorksheet.range("A1")
        range.value = "123"
        ctx.load(range, "Value")
        ctx.sync()
        print(range.value)

    @staticmethod
    def test_updateText():
        ctx = fakexlapi.RequestContext()
        workbook = ctx.application.activeWorkbook
        range = workbook.activeWorksheet.range("A1")
        range.text = "abc"
        ctx.load(range, "Text")
        print(range.text)

    @staticmethod
    def test_arrayValue():
        ctx = fakexlapi.RequestContext()
        sheet = ctx.application.activeWorkbook.activeWorksheet
        range = sheet.range("A1")
        range.value = ['Hello', 123, True]
        ctx.load(range)
        ctx.sync()
        print("Range.Value=")
        print(range.value)

    @staticmethod
    def test_Value2DArray():
        ctx = fakexlapi.RequestContext()
        sheet = ctx.application.activeWorkbook.activeWorksheet
        range = sheet.range("A1")
        range.valueArray = [['Hello', 123, True], ['World', 456, False]]
        ctx.load(range)
        ctx.sync()
        print("Range.ValueArray=")
        print(range.valueArray)

if __name__ == "__main__":
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo()
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders.url = "pipe://./fakeexcel/_api"

    methods = dir(FakeExcelTest)
    for method in methods:
        if method.startswith("test_"):
            print("")
            print("------invoke " + method + " ------")
            func = getattr(FakeExcelTest, method)
            func()
