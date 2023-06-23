Attribute VB_Name = "Module2"
Sub Select_Folder()
'****************************************************
'*                  フォルダ選択                    *
'****************************************************
    With Application.FileDialog(msoFileDialogFolderPicker)

        If ActiveSheet.Cells(1, 2) <> "" Then
            .InitialFileName = ActiveSheet.Cells(1, 2) & "\"
            ActiveSheet.Cells(1, 2).ClearContents    '書込先のセルをクリア
        Else
            .InitialFileName = ActiveSheet.Cells(2, 3)
        End If
        If .Show = True Then
            ActiveSheet.Cells(1, 2) = .SelectedItems(1)    '選択したフォルダの絶対パスをセルに書込む
        End If

    End With
'    Call FileList
End Sub
