import numbers
import runtime
import datetime
import excel
import excelhelper
import oauthhelper
import datetime
import json

class ExcelDemoLib:
    @staticmethod
    def initDesktopContext():
        requestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo()
        #requestUrlAndHeaders.url = "http://localhost:8052"
        requestUrlAndHeaders.url = "pipe://./excel/_api"
        runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders

    @staticmethod
    def populateData(context: excel.RequestContext):
        values = [["Size rank 2014","City","7/1/2014 population estimate","7/1/2013 population estimate","4/1/2010 census population","7/1/2005 population estimate","4/1/2000 census population","4/1/1990 census population","Size rank 1990","Size rank 2000","Size rank 2010","Size rank 2013"],[1,"New York, N.Y.",8491079,8405837,8175133,8143197,8008278,7322564,1,1,1,1],[2,"Los Angeles, Calif.",3928864,3884307,3792621,3844829,3694820,3485398,2,2,2,2],[3,"Chicago, Ill.",2722389,2718782,2695598,2842518,2896016,2783726,3,3,3,3],[4,"Houston, Tex.",2239558,2195914,2100263,2016582,1953631,1630553,4,4,4,4],[5,"Philadelphia, Pa.",1560297,1553165,1526006,1463281,1517550,1585577,5,5,5,5],[6,"Phoenix, Ariz.",1537058,1513367,1445632,1461575,1321045,983403,10,6,6,6],[7,"San Antonio, Tex.",1436697,1409019,1327407,1256509,1144646,935933,9,9,7,7],[8,"San Diego, Calif.",1381069,1355896,1307402,1255540,1223400,1110549,6,7,8,8],[9,"Dallas, Tex.",1281047,1257676,1197816,1213825,1188580,1006877,8,8,9,9],[10,"San Jose, Calif.",1015785,998537,945942,912332,894943,782248,11,11,10,10],[11,"Austin, Tex.",912791,885400,790390,690252,656562,465622,25,16,14,11],[12,"Jacksonville, Fla.",853382,842583,821784,782623,735617,635230,15,14,11,13],[13,"San Francisco , Calif.",852469,837442,805235,739426,776733,723959,14,13,13,14],[14,"Indianapolis, Ind.",848788,843393,820445,784118,781870,741952,13,12,12,12],[15,"Columbus, Ohio",835957,822553,787033,730657,711470,632910,16,15,15,15],[16,"Fort Worth , Tex.",812238,792727,741206,624067,534694,447619,29,27,16,17],[17,"Charlotte, N.C.",809958,792862,731424,610949,540828,395934,33,26,17,16],[18,"Detroit, Mich.",680250,688701,713777,886671,951270,1027974,7,10,18,18],[19,"El Paso, Tex.",679036,674443,649121,598590,563662,515342,22,23,19,19],[20,"Seattle , Wash.",668342,652405,608660,573911,563374,516259,21,24,23,21],[21,"Denver , Colo.",663862,649495,600158,557917,554636,467610,28,25,26,22],[22,"Washington, DC",658893,646449,601723,550521,572059,606900,19,21,24,23],[23,"Memphis, Tenn.","656,860",653450,646889,672277,650100,610337,18,18,20,20],[24,"Boston, Mass.",655884,645966,617594,559034,589141,574283,20,20,22,24],[25,"Nashville-Davidson, Tenn.1",644014,634464,601222,549110,545524,510784,26,22,25,25],[26,"Baltimore, Md.",622793,622104,620961,635815,651154,736014,12,17,21,26],[27,"Oklahoma City, Okla.",620602,610613,579999,531324,506132,444719,30,29,31,27],[28,"Portland , Ore.",619360,609456,583776,533427,529121,437319,27,28,29,29],[29,"Las Vegas , Nev.",613599,603488,583756,545147,478434,258295,63,32,30,30],[30,"Louisville-Jefferson County, Ky.2",612780,609893,597337,556429,256231,269063,58,67,27,28],[31,"Milwaukee, Wis.",599642,599164,594833,578887,596974,628088,17,19,28,31],[32,"Albuquerque, N.M.",557169,556495,545852,494236,448607,384736,40,35,32,32],[33,"Tucson, Ariz.",527972,526116,520116,515526,486699,405390,34,32,33,33],[34,"Fresno, Calif.",515986,509924,494665,461116,427652,354202,48,37,34,34],[35,"Sacramento, Calif.",485199,479686,466488,456441,407018,369365,37,40,35,35],[36,"Long Beach, Calif.",473577,469428,462257,474014,461522,429433,32,34,36,36],[37,"Kansas City, Mo.",470800,467007,459787,444965,441545,435146,31,36,37,37],[38,"Mesa, Ariz.",464704,457587,439041,442780,396375,288091,53,42,38,38],[39,"Atlanta , Ga.",456002,447841,420003,470688,416474,394017,38,39,40,40],[40,"Virginia Beach, Va.",450980,448479,437994,438415,425257,393069,39,38,39,39],[41,"Omaha , Nebr.",446599,434353,408958,414521,390007,335795,47,44,42,42],[41,"Colorado Springs, Colo.",445830,439886,416427,369815,360890,281140,54,48,41,41],[43,"Raleigh, N.C.",439896,431746,403892,"-","-","-","-","-",43,43],[44,"Miami, Fla.",430332,417650,399457,386417,362470,358548,46,47,44,44],[45,"Oakland, Calif.",413775,406253,390724,395274,399484,372242,35,41,47,45],[46,"Minneapolis, Minn.",407207,400070,382578,372811,382618,368383,43,45,48,46],[47,"Tulsa, Okla.",399682,398121,391906,382457,393049,367302,44,43,46,47],[48,"Cleveland, Ohio",389521,390113,396815,452208,478403,505616,23,33,45,48],[49,"Wichita, Kans.",388413,386552,382368,353823,344284,"-","-",50,49,49],[50,"New Orleans, La.",384320,378715,343829,455188,484674,495080,24,38,51,51],[51,"Arlington, Tex.",383204,379577,365438,362805,332969,261721,62,54,50,50]]
        range = context.workbook.worksheets.getItem("Sheet1").getCell(0, 0).getResizedRange(len(values) - 1, len(values[0]) - 1)
        range.clear()
        range.values = values
        range.worksheet.getRange("C:H").numberFormat = [["#,##0"]]
        range.worksheet.getRange("A:H").format.autofitColumns()
        table = range.worksheet.tables.add(range, True)
        table.name = "PopulationTable"
        context.sync()
        return

    @staticmethod
    def analyzeData(context: excel.RequestContext):
        table = context.workbook.tables.getItem("PopulationTable")
        nameColumn = table.columns.getItem("City")
        latestPopulationColumn = table.columns.getItem("7/1/2014 population estimate")
        earliestCensusColumn = table.columns.getItem("4/1/1990 census population")
        
        nameColumn.load("values")
        latestPopulationColumn.load("values")
        earliestCensusColumn.load("values")
    
        context.sync()
        cityData = []
    
        for i in range(1, len(nameColumn.values)):
            # A couple of the cities don't have data for 1990,
            # so skip over those.
            # Note that because the "values" is a 2D array (even though,
            # in this particular case, it's just a single column),
            # need to extract out the 0th element of each row.
            population1990 = earliestCensusColumn.values[i][0]
    
            if not isinstance(population1990, numbers.Number):
                # Skip this iteration of the loop, and move
                # to the next one.
                continue
    
            populationLatest = latestPopulationColumn.values[i][0]
            if not isinstance(populationLatest, numbers.Number):
                # Skip this iteration of the loop, and move
                # to the next one.
                continue

            # Otherwise, push the data into the in-memory store
            cityData.append((nameColumn.values[i][0], populationLatest - population1990))
    
        sortedCityData = sorted(cityData, key = lambda d: d[1], reverse = True)
        top10 = sortedCityData[0:10]

        # Now that we've computed the data, create a new worksheet
        # for the output
        outputSheet = context.workbook.worksheets.add("Top 10 Growing Cities")
    
        sheetHeader = outputSheet.getRange("B2:D2")
        sheetHeader.values = [["Top 10 Growing Cities", "", ""]]
        # sheetHeader.merge()
        sheetHeader.format.font.bold = True
        sheetHeader.format.font.size = 14
    
        tableHeader = outputSheet.getRange("B4:D4")
        tableHeader.values = [["Rank", "City", "Population Growth"]]
        table = outputSheet.tables.add("B4:D4", True)

        for i in range(len(top10)):
            table.rows.add(None, [[i + 1, top10[i][0], top10[i][1]]])
    
        # Auto-fit the column widths, and set uniform
        # thousands-separator number formatting on the
        # "Population" column of the table.
        table.getRange().getEntireColumn().format.autofitColumns()
        table.getDataBodyRange().getLastColumn().numberFormat = [["#,##"]]
    
    
        # Finally, with the table in place, add a chart:
    
        fullTableRange = table.getRange();
    
        # For the chart, no need to show the "Rank", so only use the
        # city's name and population delta
        dataRangeForChart = fullTableRange.getColumn(1).getBoundingRect(fullTableRange.getLastColumn())
    
        # A note on the function call above:
        # Range.getBoundingRect can be thought of like a 
        # "get range between" function, creating a new range spanning
        # between this object (in our case, the column at index 1,
        # which is the "City" column -- remember that all indexes in 
        # Office.js is zero-indexed!), and the last column of the table 
        # ("Population Growth").
    
        chart = outputSheet.charts.add(excel.ChartType.columnClustered, dataRangeForChart, excel.ChartSeriesBy.columns)
        chart.name = "PopulationGrowthChart"
        chart.title.text = "Population Growth between 1990 and 2014"
    
        #4 -- remember that we're 0-indexed */
        # the table header 
        tableEndRow = 3 + 1  + len(top10)
        chartPositionStart = outputSheet.getRange("F2")
        # 19 rows down, i.e., 20 rows in total 
        # 9 columns to the right, so 10 in total
        chart.setPosition(chartPositionStart, chartPositionStart.getOffsetRange(19, 9))
    
        # outputSheet.activate();
        context.sync()
        return

    @staticmethod
    def populateDataSmall(context: excel.RequestContext):
        values = [["Bellevue", "Redmond"], [1234, "=A2 + 100"]]
        range = context.workbook.worksheets.getItem("Sheet1").getCell(0, 0).getResizedRange(len(values) - 1, len(values[0]) - 1)
        range.clear()
        range.values = values
        context.load(range)
        context.sync()
        print(json.dumps(range.values, default = lambda o: o.__dict__))
        return

    @staticmethod
    def getChartImage(context: excel.RequestContext) -> str:
        chart = context.workbook.worksheets.getItem("Top 10 Growing Cities").charts.getItem("PopulationGrowthChart")
        image = chart.getImage()
        context.sync()
        return image.value

    @staticmethod
    def perfTest():
        rowCount = 100;
        colCount = 20;
        startTime = datetime.datetime.utcnow()
        for row in range(rowCount):
            for col in range(colCount):
                ctx = excel.RequestContext()
                r = ctx.workbook.worksheets.getItem("Sheet1").getCell(row, col);
                r.values = (row + 1) * (col + 1);
                ctx.sync()
        endTime = datetime.datetime.utcnow()
        diff = endTime - startTime
        return diff.seconds * 1000 + diff.microseconds / 1000

    @staticmethod
    def perfTestPopulateAndAnalyzeData(context: excel.RequestContext):
        startTime = datetime.datetime.utcnow()
        ExcelDemoLib.populateData(context)
        ExcelDemoLib.analyzeData(context)
        endTime = datetime.datetime.utcnow()
        diff = endTime - startTime
        return diff.seconds * 1000 + diff.microseconds / 1000

    @staticmethod
    def clearWorkbook(context: excel.RequestContext) -> None:
        sheet1 = context.workbook.worksheets.getItemOrNullObject("Sheet1")
        sheet2 = context.workbook.worksheets.getItemOrNullObject("Top 10 Growing Cities")
        r = sheet1.getUsedRange()
        r.clear()
        r = sheet2.getUsedRange()
        r.clear()
        sheet2.delete()
        context.sync()

