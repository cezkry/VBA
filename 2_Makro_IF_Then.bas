Option Explicit

Sub PytanieODługoscFilmu()
    Dim wyb1 As VbMsgBoxResult
    Dim wyb2 As VbMsgBoxResult
    Dim dlugosc As String
    wyb1 = MsgBox("Czy chcesz pójść na ten film", vbQuestion + vbYesNoCancel, "Odpowiedz na pytanie")
    If wyb1 = vbYes Then
        wyb2 = MsgBox("Wybrany film to: " & ActiveCell.Value, vbInformation + vbOKCancel, "Potwierdzenie wyboru")
        If wyb2 = vbOK Then
            dlugosc = ActiveCell.Offset(0, 2).Value
            If IsNumeric(dlugosc) Then
            
                If dlugosc < 90 Then
                    MsgBox " ten film jest krótki"
                ElseIf dlugosc >= 90 And dlugosc < 100 Then
                    MsgBox " Film jest sredni"
                Else
                    MsgBox " Film ten jest dlugi"
                End If
                
            Else
                MsgBox "Długość nie jest znana"
            End If
        Else
            MsgBox "wybierz jeszcze raz i naciśnij przycisk"
        End If
    ElseIf wyb1 = vbNo Then
        MsgBox "wybierz inny film i naciśnij znów przycisk"
    Else
        MsgBox "wciśnięto anuluj"
    End If
End Sub