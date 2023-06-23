Attribute VB_Name = "Module1"
Sub output()
    Const str1 As String = " at each application then self-recovered after each application"
    Const str2 As String = " during test"
    Const strAnd As String = " and"
    Const str3 As String = " after test"
    Const str4 As String = "."
    
    With Sheets(1)
    
        'initial process
        .Cells(7, 2).Clear
        
'        Testing Number = .Cells(5, 2)
        appType = .Cells(5, 3)
        funcName = .Cells(5, 4)
        duringObservation = " " & .Cells(5, 5)
        afterObservation = " " & .Cells(5, 6)
        ratedBehavior = " (" & .Cells(5, 7) & " behavior)"
        
        Select Case appType
        Case "cycles"
            devSentence = funcName & duringObservation & str1 & str2 & strAnd & afterObservation & str3 & str4 & ratedBehavior
        End Select
        
        Select Case appType
        Case "one cycle"
            devSentence = funcName & duringObservation & str1 & str2 & strAnd & afterObservation & str3 & str4 & ratedBehavior
        End Select
        
'        Select Case appType
'        Case "dropout"
'            devSentence = funcName & duringObservation & str1 & str2 & strAnd & afterObservation & str3 & str4 & ratedBehavior
'        End Select
                
        Select Case appType
        Case "N/A"
            devSentence = funcName & duringObservation & str2 & strAnd & afterObservation & str3 & str4 & ratedBehavior
        End Select
'        'output
        .Cells(7, 3) = devSentence


    End With
    
    'end process
    MsgBox ("Completed")
    
End Sub
