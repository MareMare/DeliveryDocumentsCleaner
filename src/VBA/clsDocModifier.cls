VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDocModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Base 0
Option Explicit

Public Event Starting(ByVal completedCount As Long, ByVal totalCount As Long, ByVal fileName As String, ByVal elapsedTime As Date)
Public Event Completed(ByVal completedCount As Long, ByVal totalCount As Long, ByVal fileName As String, ByVal elapsedTime As Date)
Private Enum DocKind
    Excel
    Word
End Enum

Private Const m_wdColorIndexKeyPrefix As String = "WdColorIndex="
Private Const m_wdColorIndexesMaxCount As Long = 15
Private m_docFolderPath As String
Private m_targetFileNamesMapping As Scripting.Dictionary
Private m_wdColorIndexes() As WdColorIndex
Private m_wdColorIndexesToClear() As WdColorIndex
Private m_wdColorIndexesIsAllClear As Boolean
Private m_xlColorCodeToClear As Long

Private Sub Class_Initialize()
    ReDim m_wdColorIndexes(0 To m_wdColorIndexesMaxCount - 1)
    m_wdColorIndexes(0) = WdColorIndex.wdYellow
    m_wdColorIndexes(1) = WdColorIndex.wdBrightGreen
    m_wdColorIndexes(2) = WdColorIndex.wdTurquoise
    m_wdColorIndexes(3) = WdColorIndex.wdPink
    m_wdColorIndexes(4) = WdColorIndex.wdBlue
    m_wdColorIndexes(5) = WdColorIndex.wdRed
    m_wdColorIndexes(6) = WdColorIndex.wdDarkBlue
    m_wdColorIndexes(7) = WdColorIndex.wdTeal
    m_wdColorIndexes(8) = WdColorIndex.wdGreen
    m_wdColorIndexes(9) = WdColorIndex.wdViolet
    m_wdColorIndexes(10) = WdColorIndex.wdDarkRed
    m_wdColorIndexes(11) = WdColorIndex.wdDarkYellow
    m_wdColorIndexes(12) = WdColorIndex.wdGray50
    m_wdColorIndexes(13) = WdColorIndex.wdGray25
    m_wdColorIndexes(14) = WdColorIndex.wdBlack
    
    m_docFolderPath = ""
    Set m_targetFileNamesMapping = New Scripting.Dictionary
    Erase m_wdColorIndexesToClear
    m_wdColorIndexesIsAllClear = False
    m_xlColorCodeToClear = -1
End Sub

Public Property Get DocFolderPath() As String
    DocFolderPath = m_docFolderPath
End Property
Public Property Let DocFolderPath(ByVal Value As String)
    If m_docFolderPath <> Value Then
        m_docFolderPath = Value
    End If
End Property

Public Property Get WdColorIndexesOf(ByVal index As Long) As WdColorIndex
    WdColorIndexesOf = m_wdColorIndexes(index)
End Property

Public Property Get XlColorCodeToClear() As Long
    XlColorCodeToClear = m_xlColorCodeToClear
End Property

Public Sub SetWdColorIndexesToClear(ByRef wdColorIndexes() As WdColorIndex)
    Dim mapping As Scripting.Dictionary
    Dim tempKey As Variant
    Dim index As Long
    
    Set mapping = CreateWdColorIndexesMapping(wdColorIndexes)
    
    Erase m_wdColorIndexesToClear
    If mapping.Count = m_wdColorIndexesMaxCount Then
        m_wdColorIndexesIsAllClear = True
    Else
        index = 0
        m_wdColorIndexesIsAllClear = False
        For Each tempKey In mapping.Keys
            ReDim Preserve m_wdColorIndexesToClear(index)
            m_wdColorIndexesToClear(index) = mapping(tempKey)
            index = index + 1
        Next tempKey
    End If
End Sub

Public Sub SetXlColorCodeToClear(ByVal Value As Long)
    If m_xlColorCodeToClear <> Value Then
        m_xlColorCodeToClear = Value
    End If
End Sub

Public Function BrowseDocFiles() As String()
    Dim fso As New FileSystemObject
    Dim mapping As New Scripting.Dictionary
    With mapping
        Call .Add("docx", DocKind.Word)
        Call .Add("docm", DocKind.Word)
        Call .Add("xlsx", DocKind.Excel)
        Call .Add("xlsm", DocKind.Excel)
    End With
    Dim extensionName As String
    Dim tempKey As String
    Dim objFile As file
    Dim fileNames() As String
    Dim index As Long
    
    index = 0
    Call m_targetFileNamesMapping.RemoveAll
    
    If m_docFolderPath <> "" Then
        With fso.GetFolder(m_docFolderPath)
            For Each objFile In .Files
                extensionName = fso.GetExtensionName(objFile.Path)
                tempKey = VBA.LCase$(extensionName)
                If mapping.Exists(tempKey) Then
                    Debug.Print mapping(tempKey), objFile.Name

                    Call m_targetFileNamesMapping.Add(objFile.Name, mapping(tempKey))
                    ReDim Preserve fileNames(index)
                    fileNames(index) = objFile.Name
                    index = index + 1
                End If
            Next objFile
        End With
    End If
    BrowseDocFiles = fileNames
    Set fso = Nothing
End Function

Public Sub ExecuteToClear()
    If m_docFolderPath = "" Then
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    Dim fileName As String
    Dim fileKind As DocKind
    Dim filePath As String
    Dim index As Long
    Dim startDateTime As Date
    startDateTime = VBA.Now()

    For index = 0 To m_targetFileNamesMapping.Count - 1 Step 1
        fileName = m_targetFileNamesMapping.Keys(index)
        fileKind = m_targetFileNamesMapping.Items(index)
        filePath = fso.BuildPath(m_docFolderPath, fileName)
        
        RaiseEvent Starting(index, m_targetFileNamesMapping.Count, fileName, VBA.Now() - startDateTime)
        
        Select Case fileKind
        Case DocKind.Excel
            Call ExecuteToClearExcel(filePath)
        Case DocKind.Word
            Call ExecuteToClearWord(filePath)
        Case Else
            ''
        End Select
        
        RaiseEvent Completed(index + 1, m_targetFileNamesMapping.Count, fileName, VBA.Now() - startDateTime)
    Next index
    Set fso = Nothing
End Sub

Private Function CreateWdColorIndexesMapping(ByRef wdColorIndexes() As WdColorIndex) As Scripting.Dictionary
    Dim mapping As New Scripting.Dictionary
    Dim tempColorIndex As WdColorIndex
    Dim tempKey As Variant
    Dim index As Long
    
    If (Not wdColorIndexes) <> -1 Then
        For index = LBound(wdColorIndexes) To UBound(wdColorIndexes) Step 1
            tempColorIndex = wdColorIndexes(index)
            tempKey = m_wdColorIndexKeyPrefix & VBA.Format$(tempColorIndex)
            If mapping.Exists(tempKey) Then
                Call mapping.Remove(tempKey)
            End If
            Call mapping.Add(tempKey, tempColorIndex)
        Next index
    End If
    Set CreateWdColorIndexesMapping = mapping
End Function

Private Sub ExecuteToClearExcel(ByVal filePath As String)
    Dim docModifier As New clsXlDocModifier
    With docModifier
        Call .OpenBook(filePath)
        If m_xlColorCodeToClear <> -1 Then
            Call .ClearHighlightColor(m_xlColorCodeToClear)
        End If
        Call .ClearPersonalInfo
        Call .CloseBook
    End With
End Sub

Private Sub ExecuteToClearWord(ByVal filePath As String)
    Dim wdColorIndexesMapping As Scripting.Dictionary
    Set wdColorIndexesMapping = CreateWdColorIndexesMapping(m_wdColorIndexesToClear)

    Dim docModifier As New clsWdDocModifier
    With docModifier
        Call .OpenDoc(filePath)
        If m_wdColorIndexesIsAllClear = True Then
            Call .ClearHighlightColorAll
        ElseIf wdColorIndexesMapping.Count > 0 Then
            Call .ClearHighlightColor(m_wdColorIndexKeyPrefix, wdColorIndexesMapping)
        End If
        Call .ClearPersonalInfo
        Call .CloseDoc
    End With
    Set docModifier = Nothing
End Sub

