VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWdDocModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private m_wdApp As Word.Application
Private m_wdDoc As Word.Document

Private Sub Class_Initialize()
    Set m_wdApp = New Word.Application
    With m_wdApp
        .DisplayAlerts = wdAlertsNone
        .Visible = False
    End With
    Set m_wdDoc = Nothing
End Sub

Private Sub Class_Terminate()
    If Not m_wdDoc Is Nothing Then
        With m_wdDoc
            Call .Close(SaveChanges:=False)
        End With
        Set m_wdDoc = Nothing
    End If
    If Not m_wdApp Is Nothing Then
        With m_wdApp
            Call .Quit(SaveChanges:=False)
        End With
        Set m_wdApp = Nothing
    End If
End Sub

Public Sub OpenDoc(ByVal filePath As String)
    Set m_wdDoc = m_wdApp.Documents.Open(fileName:=filePath)
End Sub

Public Sub CloseDoc()
    If m_wdDoc Is Nothing Then
        Exit Sub
    End If
    
    With m_wdDoc
        Call .Close
    End With
    Set m_wdDoc = Nothing
End Sub

Public Sub ClearHighlightColorAll()
    If m_wdDoc Is Nothing Then
        Exit Sub
    End If

    Dim objRange As Word.Range
    Set objRange = m_wdApp.ActiveDocument.Range(0, 0)
    With objRange.Find
        Call .ClearFormatting
        Call .Replacement.ClearFormatting
        
        .Highlight = True
        .Replacement.Highlight = True
        
        .Text = ""
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
        
        Call .Execute(Replace:=wdReplaceAll)
    End With
End Sub

Public Sub ClearHighlightColor(ByVal wdColorIndexKeyPrefix As String, ByRef wdColorIndexesMapping As Scripting.Dictionary)
    If m_wdDoc Is Nothing Then
        Exit Sub
    End If
    
    Dim tempKey As Variant
    Dim objRange As Word.Range
    Set objRange = m_wdApp.ActiveDocument.Range(0, 0)
    With objRange.Find
        Call .ClearFormatting
        Call .Replacement.ClearFormatting
        
        .Highlight = True
        
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop '' wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
    End With

    Do While objRange.Find.Execute(Replace:=wdReplaceNone) = True
        tempKey = wdColorIndexKeyPrefix & VBA.Format$(objRange.HighlightColorIndex)
        If wdColorIndexesMapping.Exists(tempKey) = True Then
            '' 蛍光ペンを解除します。
            objRange.HighlightColorIndex = wdNoHighlight
        End If

        objRange.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set objRange = Nothing
End Sub

Public Sub ClearPersonalInfo()
    If m_wdDoc Is Nothing Then
        Exit Sub
    End If
    
    With m_wdDoc
        Call .RemoveDocumentInformation(wdRDIAll)
    
        '' RemovePersonalInformation = True としてから保存することで個人情報を削除します。
        .RemovePersonalInformation = True
        Call .Save
    End With
End Sub

