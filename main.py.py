#############################################################
#                                                           #
#   Functionalities to transfer data from Mechanical        #
#   to Excel                                                #
# --------------------------------------------------------- #
#                                                           #
#   Programmer:         Josef Behal                         #
#   Company:            SVS FEM --> www.svsfem.cz           #
#   Ansys version:      2020 R1                             #
#                                                           #
# --------------------------------------------------------- #
#                                                           #
#   Changes:                                                #
#       25.03.2020 - Initial version                        #
#                                                           #
#############################################################

#
#   IMPORTS
#
import clr                
clr.AddReference("Microsoft.Office.Interop.Excel") 
import Microsoft.Office.Interop.Excel as ExcelNS
import units as UnitsLibrary

#
#   MECHANICAL
#
class MechanicalActions(object):
    '''
    Reading Mechanical Tabular data
    '''

    @staticmethod
    def GetMechData(sheetName = "ANSYS_Export"):
        sheetData = {}
        objData = {} 
        aos = Tree.ActiveObjects

        # Get tabular data
        for ao in aos:
            ao.Activate()
            pane  = ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
            table = pane.ControlUnknown                
            colls = []                            
            for mCol in range(2, table.ColumnsCount+1):                    
                isActive = table.Cell(1,mCol).CheckStateChecked
                col      = [isActive] + [table.Cell(mRow,mCol).Text for mRow in range(1, table.RowsCount+1)]
                colls   += [col]
            
            objData[ao.ObjectId] = colls                
            del(pane)
        
            type  = str(ao.GetType()).Split(".")[-1]
            uLab = ""
            try:    uLab = ao.MinimumTotal.Unit
            except: pass
            try:    uLab = ao.Minimum.Unit
            except: pass
            
            units = UnitsLibrary.Ansys.Core.Units.UnitsManager.GetQuantityNamesForUnit(uLab)            
            for unit in units: 
                if type.Contains(unit): break                        
            unit = "Displacement" if unit.ToLower() == "length" else unit
            aoUnit = "{} [{}]".format(unit,uLab)
        
            try:    sheetData[sheetName] += [[ao.Name, aoUnit, objData[ao.ObjectId]]]
            except: sheetData[sheetName]  = [[ao.Name, aoUnit, objData[ao.ObjectId]]]
            
        return sheetData

#   
#   EXCEL
#       
class ExcelActions(object):
    '''
    MS Excel operations
    '''
    worksheets = []
    @staticmethod
    def CreateNewWB(maximaze = True, visible = True, scrUpdate = True):
            """
            wb = ExcelActions.CreateNewWB(maximaze = True, visible = True, scrUpdate = True) --> excel workbook
            """
            excel = ExcelNS.ApplicationClass()            
            wb    = excel.Workbooks.Add()
            
            if maximaze:
                window             = excel.Windows(1)
                window.WindowState = ExcelNS.XlWindowState.xlMaximized
            excel.Visible        = visible       
            excel.ScreenUpdating = scrUpdate
            
            ExcelActions.currentWB = wb
            ExcelActions.currentEX = excel
            return wb
            
    @staticmethod
    def CreateSheet(wb = None, name = "New Sheet"):
        if wb == None: wb = ExcelActions.CreateNewWB()
        ws      = wb.Worksheets.Add()
        ws.Name = name
        ExcelActions.worksheets += [ws]
        return ws
        
    class Table(object):
        def __init__(self,ws):
            self.ws       = ws
    
        def CreateTable(self,                    
                        title      = "",
                        units      = "",
                        pivot      = [], 
                        columns    = [],
                        actives    = [],
                        initCell   = [3,2]):
            """
            tab = Table()
            tab.AddData(title = [["","",""]], pivot = [["a","b","c"]], columns = [[1,2,3],[4,5,6],[7,8]])
            """
            self.units = units
            self.name  = title
            
            titleRng = self.AddDataSeries(dataLists = [[""] for i in range(len(pivot))],   initCell = [initCell[0]-2,initCell[1]])
            self.Format(titleRng, merge = True, bgColor = ExcelNS.XlRgbColor.rgbBlack , hAlign   = "xlCenter", vAlign   = "xlCenter", fntColor = 53758, fntSize  = 13, fntBold  = True,  wrapText = True)
            titleRng.Value2 = title
            self.titleRng   = titleRng
            
            self.pivotNames = pivot
            pivotRng = self.AddDataSeries(dataLists = [[title + ": " + p] for p in pivot],   initCell = [initCell[0]-1,initCell[1]], resizeDelim = ":")
            self.Format(pivotRng, bgColor = 53758 , hAlign   = "xlCenter", vAlign   = "xlCenter", fntBold  = True,  wrapText = True)
            self.pivotRng = pivotRng
            
            dataRng  = self.AddDataSeries(dataLists = columns, initCell = initCell)
            self.Format(dataRng, bgColor = ExcelNS.XlRgbColor.rgbWhite , hAlign   = "xlCenter", vAlign   = "xlCenter")
            self.dataRng = dataRng
            self.actives = actives
            
            return self
        
        def Format(self,
                   obj,
                   excelNS  = "ExcelNS",
                   merge    = False,
                   bgColor  = None, # 53758 (svsyellow), ExcelNS.XlRgbColor.rgbBlack, ...
                   hAlign   = None, # "xlCenter", "xlBottom", ...
                   vAlign   = None, # "xlCenter", "xlBottom", ...
                   fntColor = None, # 53758 (svsyellow), ExcelNS.XlRgbColor.rgbBlack, ...
                   fntSize  = None, # 13
                   fntBold  = False,
                   wrapText = False):
                
            if merge: obj.Merge()                
            if bgColor  != None: obj.Interior.Color      = bgColor
            if hAlign   != None: obj.HorizontalAlignment = eval("{}.Constants.{}".format(excelNS,hAlign))
            if vAlign   != None: obj.VerticalAlignment   = eval("{}.Constants.{}".format(excelNS,vAlign))                
            if fntColor != None: obj.Font.Color          = fntColor
            if fntSize  != None: obj.Font.Size           = fntSize
            obj.Font.Bold           = fntBold
            obj.WrapText            = wrapText

        def AddDataSeries(self, dataLists, initCell, resizeDelim = None):         
            """
            dataList = [["aaaaaaaaaaaaaaaaaaaaaaa","b","c"]]
            initCell = [1,1]            
            table    = Table()
            t = table.AddDataSeries(dataList, initCell)           
            """
            
            rowInd, colInd = initCell
            for mCol, data in enumerate(dataLists):
                colS     = self.ws.Cells(rowInd, colInd)
                colE     = self.ws.Cells(rowInd + len(data), colInd)
                colRange = self.ws.Range(colS,colE)
                
                colVals  = colRange.Value2                
                for rCol, val in enumerate(data): colVals.SetValue(val, rCol+1, 1)
                
                if resizeDelim:
                    column = self.ws.Columns(colInd)
                    width = column.ColumnWidth

                    valLen = max([len(word) for values in data for word in values.Split(resizeDelim)]) + 3                
                    if width < valLen: column.ColumnWidth = valLen
                
                colRange.Value2 = colVals                 
                colInd += 1

            maxCol = max([len(data) for data in dataLists])
            dataS  = self.ws.Cells(initCell[0],initCell[1])
            dataE  = self.ws.Cells(rowInd + maxCol -1, colInd - 1)
            
            dataRange = self.ws.Range(dataS, dataE)
            
            return dataRange

#
# UTILS
#
def Msg(msg):ExtAPI.Log.WriteMessage(msg)

def ExportData():
    '''
    main function for data export
    '''
    sheetData = MechanicalActions.GetMechData()
    wb = ExcelActions.CreateNewWB()
    for sheetName, sheetTabs in sorted(sheetData.iteritems()):
        ws = ExcelActions.CreateSheet(wb, name = sheetName)

        tabs = []
        iRow, iCol = 5 , 2
        for name, shObjUnit, tabData in sheetTabs:

            pivot   = [col[1] for col in tabData]
            actives = [col[0] for col in tabData]
            columns = [col[2:] for col in tabData]

            tab = ExcelActions.Table(ws)
            tab.CreateTable(title = name, pivot = pivot, columns = columns, actives = actives, initCell = [iRow, iCol], units = shObjUnit)
            iCol += (len(pivot)+1)
            tabs += [tab]

#
#   CALLS
#
ExportData()
