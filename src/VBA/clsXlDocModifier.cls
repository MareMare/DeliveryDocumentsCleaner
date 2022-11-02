VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXlDocModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private m_xlApp As Excel.Application
Private m_xlBook As Excel.Workbook

Private Sub Class_Initialize()
    Set m_xlApp = New Excel.Application
    Call SetBusy(True)
    Set m_xlBook = Nothing
End Sub

Private Sub Class_Terminate()
    If Not m_xlBook Is Nothing Then
        With m_xlBook
            Call .Close(SaveChanges:=False)
        End With
        Set m_xlBook = Nothing
    End If
    If Not m_xlApp Is Nothing Then
        Call SetBusy(False)
        With m_xlApp
            .Quit
        End With
        Set m_xlApp = Nothing
    End If
End Sub

Public Sub OpenBook(ByVal filePath As String)
    If m_xlApp Is Nothing Then
        Exit Sub
    End If
    Set m_xlBook = m_xlApp.Workbooks.Open(fileName:=filePath)
End Sub

Public Sub CloseBook()
    If Not m_xlBook Is Nothing Then
        With m_xlBook
            Call .Close(SaveChanges:=False)
        End With
        Set m_xlBook = Nothing
    End If
End Sub

Public Sub ClearHighlightColor(ByVal fullColorCode As Long)
    If m_xlApp Is Nothing Then
        Exit Sub
    End If
    If m_xlBook Is Nothing Then
        Exit Sub
    End If
    
    With m_xlApp.FindFormat.Interior
        .PatternColorIndex = xlAutomatic
        .Color = fullColorCode
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With m_xlApp.ReplaceFormat.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    m_xlApp.Cells.Replace _
        What:="", _
        Replacement:="", _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=True, _
        ReplaceFormat:=True, _
        FormulaVersion:=xlReplaceFormula2
End Sub

Public Sub ClearPersonalInfo()
    If m_xlBook Is Nothing Then
        Exit Sub
    End If
    
    With m_xlBook
        Call .RemoveDocumentInformation(xlRDIAll)
    
        '' RemovePersonalInformation = True Ç∆ÇµÇƒÇ©ÇÁï€ë∂Ç∑ÇÈÇ±Ç∆Ç≈å¬êlèÓïÒÇçÌèúÇµÇ‹Ç∑ÅB
        .RemovePersonalInformation = True
        Call .Save
    End With
End Sub

Private Sub SetBusy(ByVal isBusy As Boolean)
    With m_xlApp
        .DisplayAlerts = Not isBusy
        .DisplayStatusBar = Not isBusy
        .EnableEvents = Not isBusy
        .PrintCommunication = Not isBusy
        .ScreenUpdating = Not isBusy
    End With
End Sub
