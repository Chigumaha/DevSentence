Attribute VB_Name = "Module2"
Sub Select_Folder()
'****************************************************
'*                  �t�H���_�I��                    *
'****************************************************
    With Application.FileDialog(msoFileDialogFolderPicker)

        If ActiveSheet.Cells(1, 2) <> "" Then
            .InitialFileName = ActiveSheet.Cells(1, 2) & "\"
            ActiveSheet.Cells(1, 2).ClearContents    '������̃Z�����N���A
        Else
            .InitialFileName = ActiveSheet.Cells(2, 3)
        End If
        If .Show = True Then
            ActiveSheet.Cells(1, 2) = .SelectedItems(1)    '�I�������t�H���_�̐�΃p�X���Z���ɏ�����
        End If

    End With
'    Call FileList
End Sub
