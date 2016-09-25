Private Sub commandbutton1_Click()
rpta = MsgBox("Esta seguro que desea eliminar al proveedor? ", vbYesNo + vbQuestion, "PREGUNTA")
If rpta = vbYes Then
proveed = ComboBox1.Text
fila = 2
While Sheets("Proveedores").Cells(fila, 1) <> Empty
fila = fila + 1
Wend
columna = 2
While Sheets("Proveedores").Cells(1, columna) <> Empty
columna = columna + 1
Wend

For i = 1 To fila
If Sheets("Proveedores").Cells(i, 1) = proveed Then
       
        For j = 1 To columna
        Sheets("Proveedores").Cells(i, j) = ""
        Next j
              

End If
Next i


h = fila - 1
For i = 2 To h
ComboBox1.AddItem (Sheets("Proveedores").Cells(h, 1))
Next i




End If


End Sub

Private Sub CommandButton2_Click()

Unload EliminarProveedores

End Sub

Private Sub Userform_Activate()
fila = 2
While Sheets("Proveedores").Cells(fila, 1) <> Empty
fila = fila + 1
Wend
j = fila - 1
For i = 2 To j
ComboBox1.AddItem (Sheets("Proveedores").Cells(i, 1))
Next i

End Sub
