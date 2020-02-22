Private Sub Texto_Change()
    NumeroDatos = Hoja2.Range("B" & Rows.Count).End(xlUp).Row
    Me.Lista = Clear
    Me.Lista.RowSource = Clear
    
    y = 0
    
    For fila = 2 To NumeroDatos
        descrip = Hoja3.Cells(fila, 3).Value
            If descrip Like "*" & Me.txtCeco.Value & "*" Then
                Me.Lista.AddItem
                Me.Lista.List(y, 0) = Hoja3.Cells(fila, 0)
            End If
    Next
End Sub
