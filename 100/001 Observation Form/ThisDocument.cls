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
Private Const PDF_FILEPATH As String = "C:\Repos\Software\Bexar VBA\BexarVBA\100\001 Observation Form\Observation Log\PDFs\"
Private Const OBSERVATION_LOG_FILEPATHNAME As String = "C:\Repos\Software\Bexar VBA\BexarVBA\100\001 Observation Form\Observation Log\Observation Log.xlsx"
Private Const OBSERVATION_LOG_PASSWORD As String = "Z45KUWkf37Drdy*2"

'Used to access Content Controls
Private Const oTitle = "observationTitle"
Private Const oType = "observationType"
Private Const oObservation = "observation"
Private Const oCaption1 = "caption1"
Private Const oCaption2 = "caption2"

'Used to access data array
Enum formInputEnums
    iUser
    iComputer
    iDatetime
    iTitle
    iType
    iObservation
    iCaption1
    iCaption2
    iPDF
End Enum

Private Sub submitButton_Click()
'Nolan Manteufel
'10JUL2020

'Logs form in OBSERVATION_LOG_FILEPATHNAME excel file.
'Protects the OBSERVATION_LOG_FILEPATHNAME excel file with OBSERVATION_LOG_PASSWORD.
'Saves pdf of form in PDF_FILEPATH folder.

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
Dim formInputs(8) As String

'Populate metadata
formInputs(iUser) = formUser
formInputs(iComputer) = formComputer
formInputs(iDatetime) = dateTime

'Populate data
Set cc = formDoc.SelectContentControlsByTag(oTitle)
formInputs(iTitle) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(oType)
formInputs(iType) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(oObservation)
formInputs(iObservation) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(oCaption1)
formInputs(iCaption1) = cc(1).Range.Text

Set cc = formDoc.SelectContentControlsByTag(oCaption2)
formInputs(iCaption2) = cc(1).Range.Text

formInputs(iPDF) = PDF_FILEPATH & dateTime & "_" & formUser & ".pdf"

'Launch excel
Set excelApp = New Excel.Application

'The workbook
Set logWB = excelApp.Workbooks.Open(FileName:=OBSERVATION_LOG_FILEPATHNAME)

'The worksheet
Set logSheet = logWB.Worksheets("Observations")

'Unprotect the log
logSheet.Unprotect (OBSERVATION_LOG_PASSWORD)

'Log the entry
Dim observationPKey As Long
observationPKey = logSheet.UsedRange.Rows.Count + 1
logSheet.Cells(observationPKey, 1).Value = observationPKey
logSheet.Cells(observationPKey, 2).Value = formInputs(iDatetime)
logSheet.Cells(observationPKey, 3).Value = formInputs(iUser)
logSheet.Cells(observationPKey, 4).Value = formInputs(iComputer)
logSheet.Cells(observationPKey, 5).Value = formInputs(iType)
logSheet.Cells(observationPKey, 6).Value = formInputs(iTitle)
logSheet.Cells(observationPKey, 7).Value = formInputs(iObservation)
logSheet.Cells(observationPKey, 8).Value = formInputs(iCaption1)
logSheet.Cells(observationPKey, 9).Value = formInputs(iCaption2)
logSheet.Cells(observationPKey, 10).Value = formInputs(iPDF)

'Protect the log
logSheet.Protect (OBSERVATION_LOG_PASSWORD)

'Close log
logWB.Close savechanges:=True

'Release object references
Set cc = Nothing
Set logSheet = Nothing
Set logWB = Nothing
Set excelApp = Nothing

'Save as PDF
formDoc.ExportAsFixedFormat _
     OutputFileName:=formInputs(iPDF), _
     ExportFormat:=wdExportFormatPDF, _
     OpenAfterExport:=False

'Confirmation message
MsgBox ("Awesome, " & formUser & ". Your observation has been submitted successfully.")

End Sub
