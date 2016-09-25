Private Sub CommandButton2_Click()
Unload Eliminarpedido
End Sub

Private Sub commandbutton1_Click()
rpta = MsgBox("Esta seguro que desea eliminar la solicitud? ", vbYesNo + vbQuestion, "PREGUNTA")
If rpta = vbYes Then
solped = ComboBox1.Text
fila = 2
While Sheets("Solicitud").Cells(fila, 1) <> Empty
fila = fila + 1
Wend
columna = 1
While Sheets("Solicitud").Cells(1, columna) <> Empty
columna = columna + 1
Wend

For i = 1 To fila
If Sheets("Solicitud").Cells(i, 1) = ComboBox1.Text Then
       
        For j = 1 To columna
        Sheets("Solicitud").Cells(i, j) = ""
        Next j
              

End If
Next i


h = fila - 1
For i = 2 To h
ComboBox1.AddItem (Sheets("Solicitud").Cells(h, 1))
Next i




End If


End Sub


Private Sub Userform_Activate()
fila = 2
While Sheets("Solicitud").Cells(fila, 1) <> Empty
fila = fila + 1
Wend
j = fila - 1
For i = 2 To j
ComboBox1.AddItem (Sheets("Solicitud").Cells(i, 1))
Next i

End Sub
