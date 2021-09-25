VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Nolan Manteufel
'Bexar VBA
'Document No. 100.003
'Document Rev. 00
'15SEP2020
'VBA Password = 9hvmq6KchrZ+V&&j

'Used to move documents
Private Const PDF_FILEPATH As String = "C:\Repos\Software\Bexar VBA\BexarVBA\100\003 Ultrasound Transducer Test Form\data\pdfs\"
Private Const LOG_FILEPATHNAME As String = "C:\Repos\Software\Bexar VBA\BexarVBA\100\003 Ultrasound Transducer Test Form\data\Log.xlsx"
Private Const LOG_PASSWORD As String = "Z45KUWkf37Drdy*2"

'Used to access CC form fields
Private Const ccSN = "ccSN"
Private Const ccCondition = "ccCondition"
Private Const ccTester = "ccTester"
Private Const ccUC = "ccUC"
Private Const ccTicket = "ccTicket"
Private Const ccIssue = "ccIssue"
Private Const ccT1OpMode = "ccT1OpMode"
Private Const ccT1Param = "ccT1Param"
Private Const ccT1Start = "ccT1Start"
Private Const ccT1Stop = "ccT1Stop"
Private Const ccT1Marker = "ccT1Marker"
Private Const ccT2OpMode = "ccT2OpMode"
Private Const ccT2Frequency = "ccT2Frequency"
Private Const ccYmax = "ccYmax"
Private Const ccYmin = "ccYmin"
Private Const ccMarkerReading = "ccMarkerReading"
Private Const ccVSWR = "ccVSWR"
Private Const ccConclusion = "ccConclusion"

'Used to access CC document stamps
Private Const ccUser = "ccUser"
Private Const ccComputer = "ccComputer"
Private Const ccDatetime = "ccDatetime"

Private Const NUMBER_OF_DATA_FIELDS = 21

'Used to access data array and determine data locations in log
Enum formInputEnums
    iUser
    iComputer
    iDatetime
    iSN
    iCondition
    iTester
    iUC
    iTicket
    iIssue
    iT1OpMode
    iT1Param
    iT1Start
    iT1Stop
    iT1Marker
    iT2OpMode
    iT2Frequency
    iYmax
    iYmin
    iMarkerReading
    iVSWR
    iConclusion
    iPDF
End Enum

Public WithEvents wordApp As Word.Application
Attribute wordApp.VB_VarHelpID = -1
Private Sub wordApp_DocumentBeforePrint(ByVal Doc As Document, Cancel As Boolean)

'Nolan Manteufel
'15SEP2020

'Logs the form in LOG_FILEPATHNAME excel file.
'Protects the LOG_FILEPATHNAME excel file with LOG_PASSWORD.
'Saves PDF of form in PDF_FILEPATH folder.

'Exit if error
On Error GoTo ErrHandler

'Variables
Dim formDoc As Document
Set formDoc = ActiveDocument
Dim cc As ContentControls

Dim excelApp As Excel.Application
Dim logWB As Excel.Workbook
Dim logSheet As Excel.Worksheet

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
Dim formInputs(NUMBER_OF_DATA_FIELDS) As String
Dim dataColumnIndex As Integer

'Populate metadata
formInputs(iUser) = formUser
formInputs(iComputer) = formComputer
formInputs(iDatetime) = dateTime

'Populate data
formInputs(iSN) = formDoc.SelectContentControlsByTag(ccSN)(1).Range.Text
formInputs(iCondition) = formDoc.SelectContentControlsByTag(ccCondition)(1).Range.Text
formInputs(iTester) = formDoc.SelectContentControlsByTag(ccTester)(1).Range.Text
formInputs(iUC) = formDoc.SelectContentControlsByTag(ccUC)(1).Range.Text
formInputs(iTicket) = formDoc.SelectContentControlsByTag(ccTicket)(1).Range.Text
formInputs(iIssue) = formDoc.SelectContentControlsByTag(ccIssue)(1).Range.Text
formInputs(iT1OpMode) = formDoc.SelectContentControlsByTag(ccT1OpMode)(1).Range.Text
formInputs(iT1Param) = formDoc.SelectContentControlsByTag(ccT1Param)(1).Range.Text
formInputs(iT1Start) = formDoc.SelectContentControlsByTag(ccT1Start)(1).Range.Text
formInputs(iT1Stop) = formDoc.SelectContentControlsByTag(ccT1Stop)(1).Range.Text
formInputs(iT1Marker) = formDoc.SelectContentControlsByTag(ccT1Marker)(1).Range.Text
formInputs(iT2OpMode) = formDoc.SelectContentControlsByTag(ccT2OpMode)(1).Range.Text
formInputs(iT2Frequency) = formDoc.SelectContentControlsByTag(ccT2Frequency)(1).Range.Text
formInputs(iYmax) = formDoc.SelectContentControlsByTag(ccYmax)(1).Range.Text
formInputs(iYmin) = formDoc.SelectContentControlsByTag(ccYmin)(1).Range.Text
formInputs(iMarkerReading) = formDoc.SelectContentControlsByTag(ccMarkerReading)(1).Range.Text
formInputs(iVSWR) = formDoc.SelectContentControlsByTag(ccVSWR)(1).Range.Text
'formInputs(iConclusion) left empty
formInputs(iPDF) = PDF_FILEPATH & dateTime & "_" & formUser & ".pdf"

''Request for confirmation
'Dim intResponse As Integer
'intResponse = MsgBox("Are you sure you wanted to print?", vbYesNo)
'If intResponse = vbNo Then Cancel = True

'Launch excel
Set excelApp = New Excel.Application

'The workbook
Set logWB = excelApp.Workbooks.Open(filename:=LOG_FILEPATHNAME)

'The worksheet
Set logSheet = logWB.Worksheets("log")

'Unprotect the log
logSheet.Unprotect (LOG_PASSWORD)

'Log the entry
Dim logPKey As Long
logPKey = logSheet.UsedRange.Rows.Count + 1
logSheet.Cells(logPKey, 1).Value = logPKey

For dataColumnIndex = 0 To NUMBER_OF_DATA_FIELDS
logSheet.Cells(logPKey, dataColumnIndex + 2).Value = formInputs(dataColumnIndex)
Next

'Protect the log
logSheet.Protect (LOG_PASSWORD)

'Close log
logWB.Close savechanges:=True

'Release object references
Set cc = Nothing
Set logSheet = Nothing
Set logWB = Nothing
Set excelApp = Nothing

'' STAMPS REMOVED
''Set form field stamps
'formDoc.SelectContentControlsByTag(ccUser)(1).LockContents = False
'formDoc.SelectContentControlsByTag(ccUser)(1).Range.Text = formInputs(iUser)
'formDoc.SelectContentControlsByTag(ccComputer)(1).LockContents = False
'formDoc.SelectContentControlsByTag(ccComputer)(1).Range.Text = formInputs(iComputer)
'formDoc.SelectContentControlsByTag(ccDatetime)(1).LockContents = False
'formDoc.SelectContentControlsByTag(ccDatetime)(1).Range.Text = formInputs(iDatetime)

'Save as PDF
formDoc.ExportAsFixedFormat _
     OutputFileName:=formInputs(iPDF), _
     ExportFormat:=wdExportFormatPDF, _
     OpenAfterExport:=False

''Clear form field stamps
'formDoc.SelectContentControlsByTag(ccUser)(1).Range.Text = " "
'formDoc.SelectContentControlsByTag(ccUser)(1).LockContents = True
'formDoc.SelectContentControlsByTag(ccComputer)(1).Range.Text = " "
'formDoc.SelectContentControlsByTag(ccComputer)(1).LockContents = True
'formDoc.SelectContentControlsByTag(ccDatetime)(1).Range.Text = " "
'formDoc.SelectContentControlsByTag(ccDatetime)(1).LockContents = True

''Confirmation message
'MsgBox ("Awesome, " & UCase(formUser) & "! The form has been saved successfully.")

ErrHandler:

End Sub
