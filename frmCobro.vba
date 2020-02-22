Private Sub btnCalcular_Click()
'mediante un for capturar para cada valor unitario el valor del servicio, el for va en el boton y le asignara el valor a una variable
Dim valorTotalUno, valorTotalDos, valorTotalTres, valorTotalCuatro, valorTotalCinco, valorTotalSeis As Double
Dim valorUnitarioUno As Double
Dim listaServicios() As String
Dim largoTipoServicio As Integer
largoTipoServicio = Application.CountA(Worksheets("BD").Columns("A")) - 1

ReDim listaServicios(3, largoTipoServicio)
For fila = 1 To largoTipoServicio
    listaServicios(1, fila) = Worksheets("BD").Cells(fila + 1, 1).Value
    listaServicios(2, fila) = Worksheets("BD").Cells(fila + 1, 2).Value
    listaServicios(3, fila) = Worksheets("BD").Cells(fila + 1, 3).Value
Next

For i = 1 To largoTipoServicio
    If cbxTipoServicio1.Value = listaServicios(1, i) Then
        valorUnitarioUno = listaServicios(2, i)
        tipoImpresionUno = listaServicios(3, i)
    End If
    
    If cbxTipoServicio2.Value = listaServicios(1, i) Then
        valorUnitarioDos = listaServicios(2, i)
        tipoImpresionDos = listaServicios(3, i)
    End If
    
    If cbxTipoServicio3.Value = listaServicios(1, i) Then
        valorUnitarioTres = listaServicios(2, i)
        tipoImpresionTres = listaServicios(3, i)
    End If
    
    If cbxTipoServicio4.Value = listaServicios(1, i) Then
        valorUnitarioCuatro = listaServicios(2, i)
        tipoImpresionCuatro = listaServicios(3, i)
    End If
    
    If cbxTipoServicio5.Value = listaServicios(1, i) Then
        valorUnitarioCinco = listaServicios(2, i)
        tipoImpresionCinco = listaServicios(3, i)
    End If
    If cbxTipoServicio6.Value = listaServicios(1, i) Then
        valorUnitarioSeis = listaServicios(2, i)
        tipoImpresionSeis = listaServicios(3, i)
    End If
Next


If cbxTipoServicio1.Value <> "" And txtCantidad1.Value <> "" Then
    If txtValorUnitario1 = "" Then
        valorTotalUno = valorUnitarioUno * txtCantidad1.Value
        txtValorUnitario1 = valorUnitarioUno
    Else
        valorTotalUno = txtValorUnitario1.Value * txtCantidad1.Value
    End If
    txtValorTotal1 = FormatNumber(valorTotalUno, 0, , vbFalse)
    
End If

If cbxTipoServicio2.Value <> "" And txtCantidad2.Value <> "" Then
    If txtValorUnitario2 = "" Then
        valorTotalDos = valorUnitarioDos * txtCantidad2.Value
        txtValorUnitario2 = valorUnitarioDos
    Else
        valorTotalDos = txtValorUnitario2.Value * txtCantidad2.Value
    End If
    txtValorTotal2 = FormatNumber(valorTotalDos, 0, , vbFalse)
End If

If cbxTipoServicio3.Value <> "" And txtCantidad3.Value <> "" Then
    If txtValorUnitario3 = "" Then
        valorTotalTres = valorUnitarioTres * txtCantidad3.Value
        txtValorUnitario3 = valorUnitarioTres
    Else
        valorTotalTres = txtValorUnitario3.Value * txtCantidad3.Value
    End If
    txtValorTotal3 = FormatNumber(valorTotalTres, 0, , vbFalse)
End If

If cbxTipoServicio4.Value <> "" And txtCantidad4.Value <> "" Then
    If txtValorUnitario4 = "" Then
        valorTotalCuatro = valorUnitarioCuatro * txtCantidad4.Value
        txtValorUnitario4 = valorUnitarioCuatro
    Else
        valorTotalCuatro = txtValorUnitario4.Value * txtCantidad4.Value
    End If
    txtValorTotal4 = FormatNumber(valorTotalCuatro, 0, , vbFalse)
End If

If cbxTipoServicio5.Value <> "" And txtCantidad5.Value <> "" Then
    If txtValorUnitario5 = "" Then
        valorTotalCinco = valorUnitarioCinco * txtCantidad5.Value
        txtValorUnitario5 = valorUnitarioCinco
    Else
        valorTotalCinco = txtValorUnitario5.Value * txtCantidad5.Value
    End If
    txtValorTotal5 = FormatNumber(valorTotalCinco, 0, , vbFalse)
End If

If cbxTipoServicio6.Value <> "" And txtCantidad6.Value <> "" Then
    If txtValorUnitario6 = "" Then
        valorTotalSeis = valorUnitarioSeis * txtCantidad6.Value
        txtValorUnitario6 = valorUnitarioSeis
    Else
        valorTotalSeis = txtValorUnitario6.Value * txtCantidad6.Value
    End If
    txtValorTotal6 = FormatNumber(valorTotalSeis, 0, , vbFalse)
End If

valorTotalServicio = 0
valorTotalServicio = valorTotalUno + valorTotalDos + valorTotalTres + valorTotalCuatro + valorTotalCinco + valorTotalSeis
txtTotalServicio.Value = FormatNumber(valorTotalServicio, 0, , vbFalse)

End Sub

Private Sub btnGuardar_Click()

'Desproteger la macro
Application.ScreenUpdating = False
Application.DisplayAlerts = False
ActiveSheet.Unprotect "tata1302"

'llevar los datos a la hoja Data
Const ubicacionTipoServicio = 1
Const ubicacionCantidad = 2
Const ubicacionValorUnitario = 3
Const ubicacionValorTotal = 4
Const ubicacionCeco = 5
Const ubicacionNombreCeCo = 6
Const ubicacionFecha = 7
Const ubicacionHora = 8
Const ubicacionResponsable = 9
Const ubicacionTipoPago = 10
Const ubicacionFechaVoucher = 11
Const ubicacionTipoImpresion = 12
Const ubicacionModelo = 13

Dim valorTotalUno, valorTotalDos, valorTotalTres, valorTotalCuatro, valorTotalCinco, valorTotalSeis As Double
Dim valorUnitarioUno As Double
Dim listaServicios() As String
Dim largoTipoServicio As Integer
largoTipoServicio = Application.CountA(Worksheets("BD").Columns("A")) - 1

ReDim listaServicios(3, largoTipoServicio)
For fila = 1 To largoTipoServicio
    listaServicios(1, fila) = Worksheets("BD").Cells(fila + 1, 1).Value
    listaServicios(2, fila) = Worksheets("BD").Cells(fila + 1, 2).Value
    listaServicios(3, fila) = Worksheets("BD").Cells(fila + 1, 3).Value
Next

For i = 1 To largoTipoServicio
    If cbxTipoServicio1.Value = listaServicios(1, i) Then
        valorUnitarioUno = listaServicios(2, i)
        tipoImpresionUno = listaServicios(3, i)
    End If
    
    If cbxTipoServicio2.Value = listaServicios(1, i) Then
        valorUnitarioDos = listaServicios(2, i)
        tipoImpresionDos = listaServicios(3, i)
    End If
    
    If cbxTipoServicio3.Value = listaServicios(1, i) Then
        valorUnitarioTres = listaServicios(2, i)
        tipoImpresionTres = listaServicios(3, i)
    End If
    
    If cbxTipoServicio4.Value = listaServicios(1, i) Then
        valorUnitarioCuatro = listaServicios(2, i)
        tipoImpresionCuatro = listaServicios(3, i)
    End If
    
    If cbxTipoServicio5.Value = listaServicios(1, i) Then
        valorUnitarioCinco = listaServicios(2, i)
        tipoImpresionCinco = listaServicios(3, i)
    End If
    If cbxTipoServicio6.Value = listaServicios(1, i) Then
        valorUnitarioSeis = listaServicios(2, i)
        tipoImpresionSeis = listaServicios(3, i)
    End If
Next


If cbxTipoServicio1.Value <> "" And txtCantidad1.Value <> "" Then
    If txtValorUnitario1 = "" Then
        valorTotalUno = valorUnitarioUno * txtCantidad1.Value
        txtValorUnitario1 = valorUnitarioUno
    Else
        valorTotalUno = txtValorUnitario1.Value * txtCantidad1.Value
    End If
    txtValorTotal1 = valorTotalUno
End If

If cbxTipoServicio2.Value <> "" And txtCantidad2.Value <> "" Then
    If txtValorUnitario2 = "" Then
        valorTotalDos = valorUnitarioDos * txtCantidad2.Value
        txtValorUnitario2 = valorUnitarioDos
    Else
        valorTotalDos = txtValorUnitario2.Value * txtCantidad2.Value
    End If
    txtValorTotal2 = valorTotalDos
End If

If cbxTipoServicio3.Value <> "" And txtCantidad3.Value <> "" Then
    If txtValorUnitario3 = "" Then
        valorTotalTres = valorUnitarioTres * txtCantidad3.Value
        txtValorUnitario3 = valorUnitarioTres
    Else
        valorTotalTres = txtValorUnitario3.Value * txtCantidad3.Value
    End If
    txtValorTotal3 = valorTotalTres
End If

If cbxTipoServicio4.Value <> "" And txtCantidad4.Value <> "" Then
    If txtValorUnitario4 = "" Then
        valorTotalCuatro = valorUnitarioCuatro * txtCantidad4.Value
        txtValorUnitario4 = valorUnitarioCuatro
    Else
        valorTotalCuatro = txtValorUnitario4.Value * txtCantidad4.Value
    End If
    txtValorTotal4 = valorTotalCuatro
End If

If cbxTipoServicio5.Value <> "" And txtCantidad5.Value <> "" Then
    If txtValorUnitario5 = "" Then
        valorTotalCinco = valorUnitarioCinco * txtCantidad5.Value
        txtValorUnitario5 = valorUnitarioCinco
    Else
        valorTotalCinco = txtValorUnitario5.Value * txtCantidad5.Value
    End If
    txtValorTotal5 = valorTotalCinco
End If

If cbxTipoServicio6.Value <> "" And txtCantidad6.Value <> "" Then
    If txtValorUnitario6 = "" Then
        valorTotalSeis = valorUnitarioSeis * txtCantidad6.Value
        txtValorUnitario6 = valorUnitarioSeis
    Else
        valorTotalSeis = txtValorUnitario6.Value * txtCantidad6.Value
    End If
    txtValorTotal6 = valorTotalSeis
End If

Dim filaVacia, filaVacia2, filaVacia3, filaVacia4, filaVacia5, filaVacia6  As Long
Dim dDate As Date
Dim dDateVoucher As Date

'Worksheets("Data").Active
filaVacia = Application.WorksheetFunction.CountA(Range("A:A")) + 1
filaVacia2 = filaVacia + 1
filaVacia3 = filaVacia + 2
filaVacia4 = filaVacia + 3
filaVacia5 = filaVacia + 4
filaVacia6 = filaVacia + 5
'Corregir el formato de fecha, para
dDate = DateSerial(Year(Date), Month(Date), Day(Date))
dDate = txtFecha.Value
dDateVoucher = DateSerial(Year(Date), Month(Date), Day(Date))
If txtFechaVoucher.Value <> "" Then
    dDateVoucher = txtFechaVoucher.Value
End If

If (txtCecoSeleccionado.Text = "") And (txtNombreCecoSeleccionado.Text = "") And (cbxTipoPago.Text = "Efectivo") Then
    cecoSeleccionado = 3118238
    nombreCecoSeleccionado = "Pago en efectivo centro de copiado"
Else
    cecoSeleccionado = txtCecoSeleccionado.Text
    nombreCecoSeleccionado = txtNombreCecoSeleccionado.Text
End If

Cells(filaVacia, ubicacionTipoServicio) = cbxTipoServicio1.Text
Cells(filaVacia, ubicacionCantidad) = txtCantidad1.Text
Cells(filaVacia, ubicacionValorUnitario) = txtValorUnitario1.Text
Cells(filaVacia, ubicacionValorTotal) = valorTotalUno
Cells(filaVacia, ubicacionCeco) = cecoSeleccionado
Cells(filaVacia, ubicacionNombreCeCo) = nombreCecoSeleccionado
Cells(filaVacia, ubicacionFecha) = dDate
Cells(filaVacia, ubicacionHora) = txtHora.Text
Cells(filaVacia, ubicacionResponsable) = txtResponsable.Text
Cells(filaVacia, ubicacionTipoPago) = cbxTipoPago.Text
Cells(filaVacia, ubicacionFechaVoucher) = dDateVoucher
Cells(filaVacia, ubicacionTipoImpresion) = tipoImpresionUno
Cells(filaVacia, ubicacionModelo) = cbxModelo1.Text

If txtCantidad2.Value <> "" And txtValorUnitario2.Value <> "" Then
    Cells(filaVacia2, ubicacionTipoServicio) = cbxTipoServicio2.Text
    Cells(filaVacia2, ubicacionCantidad) = txtCantidad2.Text
    Cells(filaVacia2, ubicacionValorUnitario) = txtValorUnitario2.Text
    Cells(filaVacia2, ubicacionValorTotal) = valorTotalDos
    Cells(filaVacia2, ubicacionCeco) = cecoSeleccionado
    Cells(filaVacia2, ubicacionNombreCeCo) = nombreCecoSeleccionado
    Cells(filaVacia2, ubicacionFecha) = dDate
    Cells(filaVacia2, ubicacionHora) = txtHora.Text
    Cells(filaVacia2, ubicacionResponsable) = txtResponsable.Text
    Cells(filaVacia2, ubicacionTipoPago) = cbxTipoPago.Text
    Cells(filaVacia2, ubicacionFechaVoucher) = dDateVoucher
    Cells(filaVacia2, ubicacionTipoImpresion) = tipoImpresionDos
    Cells(filaVacia2, ubicacionModelo) = cbxModelo2.Text
End If

If txtCantidad3.Value <> "" And txtValorUnitario3.Value <> "" Then
    Cells(filaVacia3, ubicacionTipoServicio) = cbxTipoServicio3.Text
    Cells(filaVacia3, ubicacionCantidad) = txtCantidad3.Text
    Cells(filaVacia3, ubicacionValorUnitario) = txtValorUnitario3.Text
    Cells(filaVacia3, ubicacionValorTotal) = valorTotalTres
    Cells(filaVacia3, ubicacionCeco) = cecoSeleccionado
    Cells(filaVacia3, ubicacionNombreCeCo) = nombreCecoSeleccionado
    Cells(filaVacia3, ubicacionFecha) = dDate
    Cells(filaVacia3, ubicacionHora) = txtHora.Text
    Cells(filaVacia3, ubicacionResponsable) = txtResponsable.Text
    Cells(filaVacia3, ubicacionTipoPago) = cbxTipoPago.Text
    Cells(filaVacia3, ubicacionFechaVoucher) = dDateVoucher
    Cells(filaVacia3, ubicacionTipoImpresion) = tipoImpresionTres
    Cells(filaVacia3, ubicacionModelo) = cbxModelo3.Text
End If

If txtCantidad4.Value <> "" And txtValorUnitario4.Value <> "" Then
    Cells(filaVacia4, ubicacionTipoServicio) = cbxTipoServicio4.Text
    Cells(filaVacia4, ubicacionCantidad) = txtCantidad4.Text
    Cells(filaVacia4, ubicacionValorUnitario) = txtValorUnitario4.Text
    Cells(filaVacia4, ubicacionValorTotal) = valorTotalCuatro
    Cells(filaVacia4, ubicacionCeco) = cecoSeleccionado
    Cells(filaVacia4, ubicacionNombreCeCo) = nombreCecoSeleccionado
    Cells(filaVacia4, ubicacionFecha) = dDate
    Cells(filaVacia4, ubicacionHora) = txtHora.Text
    Cells(filaVacia4, ubicacionResponsable) = txtResponsable.Text
    Cells(filaVacia4, ubicacionTipoPago) = cbxTipoPago.Text
    Cells(filaVacia4, ubicacionFechaVoucher) = dDateVoucher
    Cells(filaVacia4, ubicacionTipoImpresion) = tipoImpresionCuatro
    Cells(filaVacia4, ubicacionModelo) = cbxModelo4.Text
End If

If txtCantidad5.Value <> "" And txtValorUnitario5.Value <> "" Then
    Cells(filaVacia5, ubicacionTipoServicio) = cbxTipoServicio5.Text
    Cells(filaVacia5, ubicacionCantidad) = txtCantidad5.Text
    Cells(filaVacia5, ubicacionValorUnitario) = txtValorUnitario5.Text
    Cells(filaVacia5, ubicacionValorTotal) = valorTotalCinco
    Cells(filaVacia5, ubicacionCeco) = cecoSeleccionado
    Cells(filaVacia5, ubicacionNombreCeCo) = nombreCecoSeleccionado
    Cells(filaVacia5, ubicacionFecha) = dDate
    Cells(filaVacia5, ubicacionHora) = txtHora.Text
    Cells(filaVacia5, ubicacionResponsable) = txtResponsable.Text
    Cells(filaVacia5, ubicacionTipoPago) = cbxTipoPago.Text
    Cells(filaVacia5, ubicacionFechaVoucher) = dDateVoucher
    Cells(filaVacia5, ubicacionTipoImpresion) = tipoImpresionCinco
    Cells(filaVacia5, ubicacionModelo) = cbxModelo5.Text
End If

If txtCantidad6.Value <> "" And txtValorUnitario6.Value <> "" Then
    Cells(filaVacia6, ubicacionTipoServicio) = cbxTipoServicio6.Text
    Cells(filaVacia6, ubicacionCantidad) = txtCantidad6.Text
    Cells(filaVacia6, ubicacionValorUnitario) = txtValorUnitario6.Text
    Cells(filaVacia6, ubicacionValorTotal) = valorTotalSeis
    Cells(filaVacia6, ubicacionCeco) = cecoSeleccionado
    Cells(filaVacia6, ubicacionNombreCeCo) = nombreCecoSeleccionado
    Cells(filaVacia6, ubicacionFecha) = dDate
    Cells(filaVacia6, ubicacionHora) = txtHora.Text
    Cells(filaVacia6, ubicacionResponsable) = txtResponsable.Text
    Cells(filaVacia6, ubicacionTipoPago) = cbxTipoPago.Text
    Cells(filaVacia6, ubicacionFechaVoucher) = dDateVoucher
    Cells(filaVacia6, ubicacionTipoImpresion) = tipoImpresionSeis
    Cells(filaVacia6, ubicacionModelo) = cbxModelo6.Text
End If

Unload Me

'Volver a proteger la macro
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub



Private Sub cbxTipoPago_Change()

End Sub

Private Sub cbxTipoServicio1_Change()

End Sub

Private Sub UserForm_Activate()
txtFecha = Date
txtHora = Format(Now, "hh:mm:ss")
filasTipoServicio = Worksheets("BD").Cells(Rows.Count, 1).End(xlUp).Row
filasModelo = Worksheets("Modelos").Cells(Rows.Count, 1).End(xlUp).Row

cbxTipoPago.AddItem "Voucher"
cbxTipoPago.AddItem "Efectivo"

For i = 2 To filasTipoServicio

    cbxTipoServicio1.AddItem (Worksheets("BD").Cells(i, 1))
    cbxTipoServicio2.AddItem (Worksheets("BD").Cells(i, 1))
    cbxTipoServicio3.AddItem (Worksheets("BD").Cells(i, 1))
    cbxTipoServicio4.AddItem (Worksheets("BD").Cells(i, 1))
    cbxTipoServicio5.AddItem (Worksheets("BD").Cells(i, 1))
    cbxTipoServicio6.AddItem (Worksheets("BD").Cells(i, 1))
    
Next i

For j = 2 To filasModelo
    cbxModelo1.AddItem (Worksheets("Modelos").Cells(j, 1))
    cbxModelo2.AddItem (Worksheets("Modelos").Cells(j, 1))
    cbxModelo3.AddItem (Worksheets("Modelos").Cells(j, 1))
    cbxModelo4.AddItem (Worksheets("Modelos").Cells(j, 1))
    cbxModelo5.AddItem (Worksheets("Modelos").Cells(j, 1))
    cbxModelo6.AddItem (Worksheets("Modelos").Cells(j, 1))
Next j

End Sub


'****************************************************************************************************************************************************
'BUSQUEDA POR PALABRA CLAVE EN NOMBRE DE CECO

'1)Al iniciar
Private Sub UserForm_Initialize()
    Me.Height = 450
End Sub

'2)Al escribir texto en el TextBox
Private Sub TextBox1_Change()

    If Me.TextBox1.Value = "" Or Me.TextBox1.Value = " " Then
        Me.Height = 450

    Else
        Me.Height = 450
        Dim rng As Range, e
        Set Lista = Range("CECOS")
        With Me
            .ListBox1.Clear

            For Each i In Lista.Value
                If (i <> "") * (LCase(i) Like "*" & LCase(.TextBox1.Value) & "*") Then
                    .ListBox1.AddItem i
                End If
            Next i

        End With
    End If
End Sub

'3)Aceptar el valor elegido y capturarlo en la celda activa
Private Sub CommandButton2_Click()
    Cuenta = Me.ListBox1.ListCount

    For i = 0 To Cuenta - 1

        If Me.ListBox1.Selected(i) = True Then
            nombreBuscado = Me.ListBox1.List(i)
            txtNombreCecoSeleccionado.Value = Me.ListBox1.List(i)
        End If

    Next i
    If nombreBuscado <> "" Then
    txtCecoSeleccionado.Value = Application.WorksheetFunction.VLookup(nombreBuscado, Range("CECOSBUSQUEDA"), 5, 0)
    End If
    'Unload Me

End Sub

'4)Cerrar el formulario
Private Sub CommandButton1_Click()
    Unload Me
End Sub


'******************************************************************************************************************************************
'BUSQUEDA POR CECO

'2)Al escribir texto en el TextBox
Private Sub TextBox2_Change()

    If Me.TextBox2.Value = "" Or Me.TextBox2.Value = " " Then

    Else
        Dim rng As Range, e
        Set Lista = Range("CODIGOCECOS")
        With Me
            .ListBox2.Clear

            For Each i In Lista.Value
                If (i <> "") * (LCase(i) Like "*" & LCase(.TextBox2.Value) & "*") Then
                    .ListBox2.AddItem i
                End If
            Next i

        End With
    End If
End Sub

'3)Aceptar el valor elegido y capturarlo en la celda activa
Private Sub CommandButton3_Click()
    Cuenta = Me.ListBox2.ListCount
    For i = 0 To Cuenta - 1

        If Me.ListBox2.Selected(i) = True Then
            cecoBuscado = Me.ListBox2.List(i)
            txtCecoSeleccionado.Value = Me.ListBox2.List(i)
        End If

    Next i
        
    If cecoBuscado <> "" Then
        cecoBuscado = CLng(cecoBuscado)
        txtNombreCecoSeleccionado.Value = Application.WorksheetFunction.VLookup(cecoBuscado, Range("CODIGOBUSQUEDA"), 2, 0)
    End If
    'Unload Me

End Sub
