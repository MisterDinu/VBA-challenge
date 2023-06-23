Sub ticker()

'declarar variables

Dim ticker As String
Dim ticker2 As String

Dim op As Double
Dim fin As Double

Dim i As Integer
Dim j As Integer



i = 2
j = 2
'copiar primer valor de open
op = Cells(i, 3)
Range("O7") = op

'poner en nueva fila el ticker correspondiente

ticker = Cells(i, 1).Value
Cells(j, 10) = ticker


'si ticker corresponde con el siguiente, aumentar una fila
    While Cells(i, 1).Value <> ""
        ticker2 = Cells(i + 1, 1).Value
        
        'si no, copiar el valor a celda (j, 10)
            If ticker <> ticker2 And ticker2 <> "" Then
                j = j + 1
                ticker = ticker2
                Cells(j, 10) = ticker
            End If
        fin = Cells(i, 6).Value
        Range("O8").Value = fin
        i = i + 1
    Wend
    
End Sub



