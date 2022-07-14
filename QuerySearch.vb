Private Sub ExitForm_Click()
    QuerySearch.Hide
End Sub

Private Sub RunQuery_Click()
    DatosIngresados = False
    DatosEncontrados = False
    NumeroDeDatos = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    UserForm1.ListBox1.Clear

    'mensaje de alerta para error
    If Nombre = "" Then
        If Autor = "" Then
            If RatingDeUsuarios = "" Then
                If Año = "" Then
                    If Genero = "" Then
                        MsgBox "Datos no ingresados, favor de intentarlo nuevamente"
                        Datos = False
                    End If
                End If
            End If
        End If
    End If
    
    'creacion de la Query
    y = 0
    UserForm1.ListBox1.AddItem
    UserForm1.ListBox1.List(y, 0) = ActiveSheet.Cells(3, 2).Value
    UserForm1.ListBox1.List(y, 1) = ActiveSheet.Cells(3, 3).Value
    UserForm1.ListBox1.List(y, 2) = ActiveSheet.Cells(3, 4).Value
    UserForm1.ListBox1.List(y, 3) = ActiveSheet.Cells(3, 5).Value
    UserForm1.ListBox1.List(y, 4) = ActiveSheet.Cells(3, 6).Value
    UserForm1.ListBox1.List(y, 5) = ActiveSheet.Cells(3, 7).Value
    UserForm1.ListBox1.List(y, 6) = ActiveSheet.Cells(3, 8).Value
    y = y + 1
    
    If Nombre <> "" Then
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 2).Value
            If UCase(Descripcion) Like "*" & UCase(Me.Nombre.Value) & "*" Then
                UserForm1.ListBox1.AddItem
                UserForm1.ListBox1.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                UserForm1.ListBox1.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                UserForm1.ListBox1.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                UserForm1.ListBox1.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                UserForm1.ListBox1.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                UserForm1.ListBox1.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                UserForm1.ListBox1.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
                DatosEncontrados = True
            End If
        Next
        DatosIngresados = True
    End If
    
    If Autor <> "" Then
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 3).Value
            If UCase(Descripcion) Like "*" & UCase(Me.Autor.Value) & "*" Then
                UserForm1.ListBox1.AddItem
                UserForm1.ListBox1.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                UserForm1.ListBox1.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                UserForm1.ListBox1.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                UserForm1.ListBox1.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                UserForm1.ListBox1.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                UserForm1.ListBox1.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                UserForm1.ListBox1.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
                DatosEncontrados = True
            End If
        Next
        DatosIngresados = True
    End If
    
    If RatingDeUsuarios <> "" Then
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 4).Value
            If UCase(Descripcion) Like "*" & UCase(Me.RatingDeUsuarios.Value) & "*" Then
                UserForm1.ListBox1.AddItem
                UserForm1.ListBox1.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                UserForm1.ListBox1.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                UserForm1.ListBox1.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                UserForm1.ListBox1.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                UserForm1.ListBox1.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                UserForm1.ListBox1.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                UserForm1.ListBox1.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
                DatosEncontrados = True
            End If
        Next
        DatosIngresados = True
    End If
    
    If Año <> "" Then
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 7).Value
            If UCase(Descripcion) Like "*" & UCase(Me.Año.Value) & "*" Then
                UserForm1.ListBox1.AddItem
                UserForm1.ListBox1.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                UserForm1.ListBox1.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                UserForm1.ListBox1.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                UserForm1.ListBox1.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                UserForm1.ListBox1.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                UserForm1.ListBox1.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                UserForm1.ListBox1.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
                DatosEncontrados = True
            End If
        Next
        DatosIngresados = True
    End If
    
    If Genero <> "" Then
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 8).Value
            If UCase(Descripcion) Like "*" & UCase(Me.Genero.Value) & "*" Then
                UserForm1.ListBox1.AddItem
                UserForm1.ListBox1.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                UserForm1.ListBox1.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                UserForm1.ListBox1.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                UserForm1.ListBox1.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                UserForm1.ListBox1.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                UserForm1.ListBox1.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                UserForm1.ListBox1.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
                DatosEncontrados = True
            End If
        Next
        DatosIngresados = True
    End If
    
    'mensaje de alerta por falta de coincidencias
    If DatosIngresados = True Then
        If DatosEncontrados = False Then
            MsgBox "Su búsqueda no regreso ningún resultado"
        Else
            UserForm1.Show
            QuerySearch.Hide
        End If
    End If
    
End Sub