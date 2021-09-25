VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Used to move documents
Private Const PDF_FILEPATH As String = "C:\Repos\Software\Bexar VBA\BexarVBA\100\002 Request Form\pdfs\"
Private Const LOG_FILEPATHNAME As String = "C:\Repos\Software\Bexar VBA\BexarVBA\100\002 Request Form\Request Log.xlsx"
Private Const LOG_PASSWORD As String = "Z45KUWkf37Drdy*2"

'Used to access Content Controls
Private Const ccrTitle = "ccrTitle"
Private Const ccrInitiator = "ccrInitiator"
Private Const ccrNumber = "ccrNumber"
Private Const ccrRequest = "ccrRequest"
Private Const ccrReason = "ccrReason"
Private Const ccrProducts = "ccrProducts"
Private Const ccrAdverse = "ccrAdverse"
Private Const ccrJustification = "ccrJustification"

'Used to access data array
Enum formInputEnums
    iUser
    iComputer
    iDatetime
    iTitle
    iInitiator
    iNumber
    iRequest
    iReason
    iProducts
    iAdverse
    iJustification
    iPDF
End Enum

Private Sub submitButton_Click()
'Nolan Manteufel
'15JUL2020

'Logs form in LOG_FILEPATHNAME excel file.
'Protects the LOG_FILEPATHNAME excel file with LOG_PASSWORD.
'Saves pdf of form in PDF_FILEPATH folder.

Application.StatusBar = "Beginning form submission..."

'Variables
Dim formDoc As Document
Set formDoc = ActiveDocument
Dim cc As ContentControls

Dim excelApp As Excel.Application
Dim logWB As Excel.Workbook
Dim logSheet As Excel.WOrksheet

'User
Dim formUser As String
formUser = Environ$("username")
formUser = LCase(formUser)

'Computer
Dim formComputer As String
formComputer = Environ$("computername")
formComputer = LCase(formComputer)

'Datetime
Dim dateTime As String
Dim timeOnly As String
timeOnly = Format(Time, "hhmmss")
dateTime = Format(Date, "yyyymmdd")
dateTime = dateTime & timeOnly

'Inputs array
Dim formInputs(12) As String

'Populate metadata
formInputs(iUser) = formUser
formInputs(iComputer) = formComputer
formInputs(iDatetime) = dateTime

'Populate data
Set cc = formDoc.SelectContentControlsByTag(ccrTitle)
formInputs(iTitle) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(ccrInitiator)
formInputs(iInitiator) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(ccrNumber)
formInputs(iNumber) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(ccrRequest)
formInputs(iRequest) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(ccrReason)
formInputs(iReason) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(ccrProducts)
formInputs(iProducts) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(ccrAdverse)
formInputs(iAdverse) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(ccrJustification)
formInputs(iJustification) = cc(1).Range.Text

formInputs(iPDF) = PDF_FILEPATH & dateTime & "_" & formUser & ".pdf"

Application.StatusBar = "Logging form..."

'Launch excel
Set excelApp = New Excel.Application

'The workbook
Set logWB = excelApp.Workbooks.Open(FileName:=LOG_FILEPATHNAME)

'The worksheet
Set logSheet = logWB.Worksheets("Observations")

'Unprotect the log
logSheet.Unprotect (LOG_PASSWORD)

'Log the entry
Dim primaryKey As Long
primaryKey = logSheet.UsedRange.Rows.Count + 1
logSheet.Cells(primaryKey, 1).Value = primaryKey
logSheet.Cells(primaryKey, 2).Value = formInputs(iDatetime)
logSheet.Cells(primaryKey, 3).Value = formInputs(iUser)
logSheet.Cells(primaryKey, 4).Value = formInputs(iComputer)
logSheet.Cells(primaryKey, 5).Value = formInputs(iTitle)
logSheet.Cells(primaryKey, 6).Value = formInputs(iInitiator)
logSheet.Cells(primaryKey, 7).Value = formInputs(iNumber)
logSheet.Cells(primaryKey, 8).Value = formInputs(iRequest)
logSheet.Cells(primaryKey, 9).Value = formInputs(iReason)
logSheet.Cells(primaryKey, 10).Value = formInputs(iProducts)
logSheet.Cells(primaryKey, 11).Value = formInputs(iAdverse)
logSheet.Cells(primaryKey, 12).Value = formInputs(iJustification)
logSheet.Cells(primaryKey, 13).Value = formInputs(iPDF)

'Protect the log
logSheet.Protect (LOG_PASSWORD)

'Close log
logWB.Close savechanges:=True

'Release object references
Set cc = Nothing
Set logSheet = Nothing
Set logWB = Nothing
Set excelApp = Nothing

Application.StatusBar = "Saving PDF..."

'Save as PDF
formDoc.ExportAsFixedFormat _
     OutputFileName:=formInputs(iPDF), _
     ExportFormat:=wdExportFormatPDF, _
     OpenAfterExport:=False

'Confirmation message
MsgBox ("Awesome, " & formUser & ". Your request has been submitted successfully.")

Application.StatusBar = ""

End Sub


