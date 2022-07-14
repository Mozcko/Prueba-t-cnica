Private Sub ExitForm_Click()
    QuerySearch.Hide
End Sub

Private Sub RunQuery_Click()
    Results = True
    NumeroDeDatos = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    Me.LISTA.Clear


    'mensaje de alerta para error
    If Nombre = "" Then
        If Autor = "" Then
            If RatingDeUsuarios = "" Then
                If Año = "" Then
                    If Genero = "" Then
                        MsgBox "Datos no ingresados, favor de intentarlo nuevamente"
                    End If
                End If
            End If
        End If
    End If
    
    
    'creacion de la Query
    If Nombre <> "" Then
        y = 0
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 2).Value
            If Descripcion Like "*" & Me.Nombre.Value & "*" Then
                Me.LISTA.AddItem
                Me.LISTA.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                Me.LISTA.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                Me.LISTA.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                Me.LISTA.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                Me.LISTA.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                Me.LISTA.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                Me.LISTA.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
            Else
                Resluts = False
            End If
        Next
    End If
    
    If Autor <> "" Then
        y = 0
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 3).Value
            If Descripcion Like "*" & Me.Autor.Value & "*" Then
                Me.LISTA.AddItem
                Me.LISTA.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                Me.LISTA.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                Me.LISTA.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                Me.LISTA.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                Me.LISTA.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                Me.LISTA.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                Me.LISTA.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
            Else
                Resluts = False
            End If
        Next
    End If
    
    If RatingDeUsuarios <> "" Then
        y = 0
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 4).Value
            If Descripcion Like "*" & Me.RatingDeUsuarios.Value & "*" Then
                Me.LISTA.AddItem
                Me.LISTA.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                Me.LISTA.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                Me.LISTA.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                Me.LISTA.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                Me.LISTA.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                Me.LISTA.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                Me.LISTA.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
            Else
                Resluts = False
            End If
        Next
    End If
    
    If Año <> "" Then
        y = 0
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 7).Value
            If Descripcion Like "*" & Me.Año.Value & "*" Then
                Me.LISTA.AddItem
                Me.LISTA.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                Me.LISTA.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                Me.LISTA.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                Me.LISTA.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                Me.LISTA.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                Me.LISTA.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                Me.LISTA.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
            Else
                Resluts = False
            End If
        Next
    End If
    
    If Genero <> "" Then
        y = 0
        For Fila = 4 To NumeroDeDatos
            Descripcion = ActiveSheet.Cells(Fila, 8).Value
            If Descripcion Like "*" & Me.Genero.Value & "*" Then
                Me.LISTA.AddItem
                Me.LISTA.List(y, 0) = ActiveSheet.Cells(Fila, 2).Value
                Me.LISTA.List(y, 1) = ActiveSheet.Cells(Fila, 3).Value
                Me.LISTA.List(y, 2) = ActiveSheet.Cells(Fila, 4).Value
                Me.LISTA.List(y, 3) = ActiveSheet.Cells(Fila, 5).Value
                Me.LISTA.List(y, 4) = ActiveSheet.Cells(Fila, 6).Value
                Me.LISTA.List(y, 5) = ActiveSheet.Cells(Fila, 7).Value
                Me.LISTA.List(y, 6) = ActiveSheet.Cells(Fila, 8).Value
                
                y = y + 1
            Else
                Resluts = False
            End If
        Next
    End If
    
    'mensaje de alerta por falta de coincidencias
    If Resluts = False Then
        MsgBox "Su búsqueda no regreso ningún resultado"
    End If
    
    
End Sub
