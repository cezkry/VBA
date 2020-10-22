Option Explicit

Sub zadanie4()
    Dim strnowadata As String
    Dim datnowadata As Date
    Dim nowysprzedawca As String
    Dim nowyprodukt As String
    Dim nowyzysk As Integer
    strnowadata = InputBox("Podaj date")
    datnowadata = CDate(strnowadata)
    nowysprzedawca = InputBox("Podaj nazwę sprzedawcy")
    nowyprodukt = InputBox("Podaj nazwę produktu")
    nowyzysk = InputBox("Podaj wartość zysku")
    Range("A1").End(xlDown).Offset(1, 0).Select
    ActiveCell.Value = datnowadata
    ActiveCell.Offset(0, 1).Value = nowysprzedawca
    ActiveCell.Offset(0, 2).Value = nowyprodukt
    ActiveCell.Offset(0, 3).Value = nowyzysk
    'ActiveCell.Value.NumberFormat = "#,##0.00 [$USD]"
    'ActiveCell.Value.Select
    'Selection.NumberFormat = "[$kr-414] #,##0.00"
    ActiveCell.Offset(0, 3).Select
    Selection.NumberFormat = "#,##0.00 [$zł-415]"

End Sub
