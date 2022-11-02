Attribute VB_Name = "SetupAddinModule"
'@Folder("VBAProject")
Option Base 0
Option Explicit

Private Const AddinMenuTitle As String = "納品ドキュメントの調整"

Public Sub InstallAddinMenu()
    Dim xlRootMenuBar As CommandBar
    Dim xlMenuButton As CommandBarButton
    Dim xlCommandControlTemp As CommandBarControl

    '' メニューバーを取得します。
    Set xlRootMenuBar = Application.CommandBars("Worksheet Menu Bar")
    
    '' 既存メニューを削除します。
    For Each xlCommandControlTemp In xlRootMenuBar.Controls
        If xlCommandControlTemp.Caption = AddinMenuTitle Then
            Call xlCommandControlTemp.Delete
        End If
    Next xlCommandControlTemp
    
    '' メニューバーへ追加します。
    Set xlMenuButton = xlRootMenuBar.Controls.Add(Type:=msoControlButton)
    xlMenuButton.BeginGroup = True
    With xlMenuButton
        .Caption = AddinMenuTitle
        .TooltipText = AddinMenuTitle
        .Style = msoButtonCaption
        .OnAction = "PerformDeliveryDocumentsCleaner"
    End With
End Sub

Public Sub UninstallAddinMenu()
    Dim xlRootMenuBar As CommandBar
    Dim xlMenuButton As CommandBarButton
    Dim xlCommandControlTemp As CommandBarControl

    '' メニューバーを取得します。
    Set xlRootMenuBar = Application.CommandBars("Worksheet Menu Bar")
    
    '' 既存メニューを削除します。
    For Each xlCommandControlTemp In xlRootMenuBar.Controls
        If xlCommandControlTemp.Caption = AddinMenuTitle Then
            Call xlCommandControlTemp.Delete
        End If
    Next xlCommandControlTemp
End Sub

Private Sub PerformDeliveryDocumentsCleaner()
    If Excel.Application.ActiveWorkbook Is Nothing Then
        MsgBox "アクティブな Excel ワークブックがありません。" _
            & vbCrLf & "この機能の実行には少なくとも1つのワークブックが必要です。" _
            & vbCrLf _
            & vbCrLf & "Ctrl + N キーで空白のブックを作成することが可能です。" _
            , vbQuestion, AddinMenuTitle
        Exit Sub
    End If
    If Not Excel.Application.ActiveWorkbook Is Nothing Then
        Call Excel.Application.ActiveWorkbook.ActiveSheet.Range("A1").Select
    End If
    
    Dim frm As New DocModifierForm
    Call frm.Show(False)
    Set frm = Nothing
End Sub
