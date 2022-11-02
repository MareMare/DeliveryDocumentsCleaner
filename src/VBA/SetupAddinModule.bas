Attribute VB_Name = "SetupAddinModule"
'@Folder("VBAProject")
Option Base 0
Option Explicit

Private Const AddinMenuTitle As String = "�[�i�h�L�������g�̒���"

Public Sub InstallAddinMenu()
    Dim xlRootMenuBar As CommandBar
    Dim xlMenuButton As CommandBarButton
    Dim xlCommandControlTemp As CommandBarControl

    '' ���j���[�o�[���擾���܂��B
    Set xlRootMenuBar = Application.CommandBars("Worksheet Menu Bar")
    
    '' �������j���[���폜���܂��B
    For Each xlCommandControlTemp In xlRootMenuBar.Controls
        If xlCommandControlTemp.Caption = AddinMenuTitle Then
            Call xlCommandControlTemp.Delete
        End If
    Next xlCommandControlTemp
    
    '' ���j���[�o�[�֒ǉ����܂��B
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

    '' ���j���[�o�[���擾���܂��B
    Set xlRootMenuBar = Application.CommandBars("Worksheet Menu Bar")
    
    '' �������j���[���폜���܂��B
    For Each xlCommandControlTemp In xlRootMenuBar.Controls
        If xlCommandControlTemp.Caption = AddinMenuTitle Then
            Call xlCommandControlTemp.Delete
        End If
    Next xlCommandControlTemp
End Sub

Private Sub PerformDeliveryDocumentsCleaner()
    If Excel.Application.ActiveWorkbook Is Nothing Then
        MsgBox "�A�N�e�B�u�� Excel ���[�N�u�b�N������܂���B" _
            & vbCrLf & "���̋@�\�̎��s�ɂ͏��Ȃ��Ƃ�1�̃��[�N�u�b�N���K�v�ł��B" _
            & vbCrLf _
            & vbCrLf & "Ctrl + N �L�[�ŋ󔒂̃u�b�N���쐬���邱�Ƃ��\�ł��B" _
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
