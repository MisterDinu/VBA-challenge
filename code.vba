Sub Elbueno()

'declarar variables

Dim ticker As String

Dim op As Double
Dim cl As Double
Dim dif As Double
Dim por As Double

Dim i As Double
Dim j As Double
Dim x As Double

Dim total As Double

'para recorrer todas las hojas
Dim ws As Worksheet


'asignar valor a i y a j,
'i debe incrementar si la celda i, 1 es igual a la celda i+1, 1
'y si no, debe detenerse y obtener el valor de la celda i, 6, que corresponde al close del ticker correspondiente

For Each ws In ThisWorkbook.Worksheets

    i = 2
    j = 2
    l = 2
    
    'copiar primer valor de open
    op = ws.Cells(i, 3).Value
    x = 1
    
    'total debe ir sumando todos los valores que correspondan a un mismo ticker
    total = ws.Cells(i, 7).Value
    
    'poner en nueva fila el ticker correspondiente
    
    ticker = ws.Cells(i, 1).Value
    
    '----------------------------------------------
     While ws.Cells(i, 1) <> ""
                If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                    total = total + ws.Cells(i + 1, 7).Value
                
                Else
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) And ws.Cells(i + 1, 1) <> 0 Then
                        cl = ws.Cells(i, 6).Value
                        ws.Range("K" & (x + 1)) = cl - op
                        por = ws.Range("K" & (x + 1)).Value / op
                        ws.Range("L" & l) = por
                        ws.Range("J" & j) = ticker
                        ws.Range("M" & (x + 1)) = total
                        
                        ' Multiplicar y dividir por "op" en cada iteración
                                    
                        x = x + 1
                        j = j + 1
                        l = l + 1
                        
                        total = ws.Cells(i + 1, 7).Value
                        ticker = ws.Cells(i + 1, 1).Value
                        op = ws.Cells(i + 1, 3).Value
                        End If
                End If
            i = i + 1
        Wend
        
    ws.Range("J" & j) = ticker
    ws.Range("K" & (x + 1)) = cl - op
    ws.Range("M" & (x + 1)) = total
    
    por = ws.Range("K" & (x + 1)).Value / op
    ws.Range("L" & l) = por
    
    Dim o As Integer
    
    o = 2
    
    Do While Not IsEmpty(ws.Range("K" & o))
        If ws.Range("K" & (o)).Value < 0 Then
        ws.Range("K" & (o)).Interior.Color = RGB(255, 0, 0)
        Else
            If ws.Range("K" & (o)).Value >= 0 Then
            ws.Range("K" & (o)).Interior.Color = RGB(0, 255, 0)
            End If
        End If
        o = o + 1
    Loop
    
    o = 2
    While ws.Range("L" & (o)) <> ""
        ws.Range("L" & (o)).NumberFormat = "0.000%"
        o = o + 1
    Wend
    
    Dim great As Double
    ticker = ws.Range("J2")
    great = ws.Range("L2")
    
    o = 2
    While ws.Range("L" & (o)) <> ""
        If ws.Range("L" & (o + 1)) > great Then
            great = ws.Range("L" & (o + 1)).Value
            ticker = ws.Range("J" & (o + 1)).Value
        End If
        o = o + 1
    Wend
    
    ws.Range("P2") = ticker
    ws.Range("Q2") = great
    ws.Range("Q2").NumberFormat = "0.000%"
    
    Dim lower As Double
    ticker = ws.Range("J2")
    lower = ws.Range("L2")
    
    o = 2
    While ws.Range("L" & (o)) <> ""
        If ws.Range("L" & (o + 1)) < lower Then
            lower = ws.Range("L" & (o + 1)).Value
            ticker = ws.Range("J" & (o + 1)).Value
        End If
        
        o = o + 1
    Wend
    
    ws.Range("P3") = ticker
    ws.Range("Q3") = lower
    ws.Range("Q3").NumberFormat = "0.000%"
    
    Dim stock As Double
    ticker = ws.Range("J2")
    stock = ws.Range("M2")
    
    o = 2
    While ws.Range("M" & (o)) <> ""
        If ws.Range("M" & (o + 1)) > stock Then
            stock = ws.Range("M" & (o + 1)).Value
            ticker = ws.Range("J" & (o + 1)).Value
        End If
        o = o + 1
    Wend
    
    ws.Range("P4") = ticker
    ws.Range("Q4") = stock

Next ws


End Sub

