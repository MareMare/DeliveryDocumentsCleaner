VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DocModifierForm 
   Caption         =   "納品ドキュメントの調整"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7305
   OleObjectBlob   =   "DocModifierForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DocModifierForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Base 0
Option Explicit

Private WithEvents m_docModifier As clsDocModifier
Attribute m_docModifier.VB_VarHelpID = -1
Private m_wdColorMapping As Scripting.Dictionary

Private Sub UserForm_Initialize()
    Set m_docModifier = New clsDocModifier
    Set m_wdColorMapping = New Scripting.Dictionary
        
    Call m_wdColorMapping.Add(Me.WdColorCheckBox01, m_docModifier.WdColorIndexesOf(0))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox02, m_docModifier.WdColorIndexesOf(1))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox03, m_docModifier.WdColorIndexesOf(2))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox04, m_docModifier.WdColorIndexesOf(3))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox05, m_docModifier.WdColorIndexesOf(4))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox06, m_docModifier.WdColorIndexesOf(5))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox07, m_docModifier.WdColorIndexesOf(6))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox08, m_docModifier.WdColorIndexesOf(7))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox09, m_docModifier.WdColorIndexesOf(8))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox10, m_docModifier.WdColorIndexesOf(9))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox11, m_docModifier.WdColorIndexesOf(10))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox12, m_docModifier.WdColorIndexesOf(11))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox13, m_docModifier.WdColorIndexesOf(12))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox14, m_docModifier.WdColorIndexesOf(13))
    Call m_wdColorMapping.Add(Me.WdColorCheckBox15, m_docModifier.WdColorIndexesOf(14))

    Me.XlColorLabel.Tag = ""
    Me.XlColorLabel.Caption = "---"
    Me.XlColorImage.BackColor = &H8000000F
    
    Call SetEnabled(False)
End Sub

Private Sub FolderSelectionButton_Click()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = 0 Then
            folderPath = ""
        Else
            folderPath = .SelectedItems(1)
        End If
    End With
    Me.FolderTextBox.Text = folderPath
    
    Dim fileNames() As String
    Dim index As Long
    m_docModifier.DocFolderPath = folderPath
    fileNames = m_docModifier.BrowseDocFiles()
    Call Me.FilesListBox.Clear
    If (Not fileNames) <> -1 Then
        For index = LBound(fileNames) To UBound(fileNames) Step 1
            Call Me.FilesListBox.AddItem(fileNames(index))
        Next index
    End If
    
    Call SetEnabled(Me.FilesListBox.ListCount > 0)
End Sub

Private Sub WdColorAllOnCheckBox_Click()
    Call SetWdColorCheckedValue(WdColorAllOnCheckBox.Value)
End Sub

Private Sub ChooseXlColorButton_Click()
    Dim fullColorCode As Long
    
    fullColorCode = GetCurrentXlColorCode()
    fullColorCode = ChooseXlColorCode(fullColorCode)
    If fullColorCode < 0 Then
        Me.XlColorLabel.Tag = ""
        Me.XlColorLabel.Caption = "---"
        Me.XlColorImage.BackColor = &H8000000F
    Else
        Me.XlColorLabel.Tag = fullColorCode
        Me.XlColorLabel.Caption = ConvertToColorHex(fullColorCode)
        Me.XlColorImage.BackColor = fullColorCode
    End If
End Sub

Private Sub ExecuteButton_Click()
    Me.ExecuteButton.Enabled = False
    Me.CurrentWorkingLabel.Caption = ""
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = Me.FilesListBox.ListCount
    Me.ProgressBar1.Value = 0
    Call VBA.DoEvents
    
    With m_docModifier
        Call .SetWdColorIndexesToClear(GetWdColorCheckedIndexes())
        Call .SetXlColorCodeToClear(GetCurrentXlColorCode())
        Call .ExecuteToClear
    End With
    
    Me.ExecuteButton.Enabled = True
End Sub

Private Sub m_docModifier_Starting(ByVal completedCount As Long, ByVal totalCount As Long, ByVal fileName As String, ByVal elapsedTime As Date)
    Me.CurrentWorkingLabel.Caption = "(" & VBA.Format$(completedCount) & "/" & VBA.Format$(totalCount) & ") " & fileName _
        & vbCrLf & "経過時間: " & VBA.Format$(elapsedTime, "HH:mm:ss")
    Me.ProgressBar1.Value = completedCount
    Call VBA.DoEvents
End Sub

Private Sub m_docModifier_Completed(ByVal completedCount As Long, ByVal totalCount As Long, ByVal fileName As String, ByVal elapsedTime As Date)
    Me.CurrentWorkingLabel.Caption = "(" & VBA.Format$(completedCount) & "/" & VBA.Format$(totalCount) & ") " & fileName _
        & vbCrLf & "経過時間: " & VBA.Format$(elapsedTime, "HH:mm:ss")
    Me.ProgressBar1.Value = completedCount
    Call VBA.DoEvents
End Sub

Private Sub SetEnabled(ByVal isEnabled As Boolean)
    Me.WdColorFrame.Enabled = isEnabled
    Me.XlColorFrame.Enabled = isEnabled
    Me.ExecuteButton.Enabled = isEnabled
End Sub

Private Function GetWdColorCheckedIndexes() As WdColorIndex()
    Dim index As Long
    Dim ctrl As Variant
    Dim wdColorIndexes() As WdColorIndex

    index = 0
    For Each ctrl In m_wdColorMapping.Keys
        If ctrl.Value = True Then
            ReDim Preserve wdColorIndexes(index)
            wdColorIndexes(index) = m_wdColorMapping(ctrl)
            index = index + 1
        End If
    Next ctrl

    GetWdColorCheckedIndexes = wdColorIndexes
End Function

Private Sub SetWdColorCheckedValue(ByVal isChecked As Boolean)
    Dim ctrl As Variant

    For Each ctrl In m_wdColorMapping.Keys
        ctrl.Value = IIf(isChecked, Checked, Unchecked)
    Next ctrl
End Sub

Private Function GetCurrentXlColorCode() As Long
    Dim tagText As String
    Dim fullColorCode As Long
    
    tagText = VBA.Trim$(VBA.Format$(Me.XlColorLabel.Tag))
    fullColorCode = IIf(VBA.Len(tagText) > 0, VBA.Val(tagText), -1)
    
    GetCurrentXlColorCode = fullColorCode
End Function

Private Function ChooseXlColorCode(ByVal fullColorCode As Long) As Long
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer

    If fullColorCode < 0 Then
        fullColorCode = IIf(Application.Dialogs(xlDialogEditColor).Show(1) = True, _
            ActiveWorkbook.Colors(1), -1)
    Else
        Call ConvertToRGB(fullColorCode, intRed, intGreen, intBlue)
        fullColorCode = IIf(Application.Dialogs(xlDialogEditColor).Show(1, intRed, intGreen, intBlue) = True, _
            ActiveWorkbook.Colors(1), -1)
    End If

    ChooseXlColorCode = fullColorCode
End Function

Private Sub ConvertToRGB(ByVal fullColorCode As Long, ByRef intRed As Integer, ByRef intGreen As Integer, ByRef intBlue As Integer)
    '' Get the RGB value for each color (possible values 0 - 255)
    intRed = fullColorCode Mod 256
    intGreen = (fullColorCode \ 256) Mod 256
    intBlue = fullColorCode \ 65536
End Sub

Private Function ConvertToColorHex(ByVal fullColorCode As Long) As String
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer

    Call ConvertToRGB(fullColorCode, intRed, intGreen, intBlue)
    ConvertToColorHex = "#" _
        & VBA.Right$("000" & VBA.Hex$(intRed), 2) _
        & VBA.Right$("000" & VBA.Hex$(intGreen), 2) _
        & VBA.Right$("000" & VBA.Hex$(intBlue), 2)
End Function


