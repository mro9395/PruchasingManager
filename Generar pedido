Private Sub CommandButton2_Click()
Unload GenerarPedido
End Sub


Private Sub commandbutton1_Click()
rpta = MsgBox("Esta seguro que desea generar la solicitud ? ", vbQuestion + vbYesNo, "PREGUNTA")

If rpta = vbYes Then
  fila = 2
While Sheets("Solicitud").Cells(fila, 1) <> Empty
  fila = fila + 1
Wend


Sheets("Solicitud").Cells(fila, 1) = "SOLPED-" & (fila - 1)
Sheets("Solicitud").Cells(fila, 2) = ComboBox1.Text
Sheets("Solicitud").Cells(fila, 3) = ComboBox2.Text
Sheets("Solicitud").Cells(fila, 4) = TextBox1.Text

End If


ComboBox1.Enabled = False
ComboBox2.Enabled = False
TextBox1.Enabled = False



End Sub


Private Sub CommandButton3_Click()
ComboBox1.Enabled = True
ComboBox2.Enabled = True
TextBox1.Enabled = True

ComboBox1.SetFocus

End Sub


Private Sub Userform_Activate()


ComboBox1.Enabled = False
ComboBox2.Enabled = False
TextBox1.Enabled = False

fila = 2
While Sheets("Productos").Cells(fila, 1) <> Empty
  fila = fila + 1
Wend
j = fila - 1
For i = 2 To j
  ComboBox1.AddItem (Sheets("Productos").Cells(i, 4))
Next i
For i = 2 To j
  ComboBox2.AddItem (Sheets("Proveedores").Cells(i, 1))
Next i
End Sub
