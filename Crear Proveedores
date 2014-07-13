Private Sub CommandButton2_Click()
rpta = MsgBox("Esta seguro que desea crear al proveedor ? ", vbQuestion + vbYesNo, "PREGUNTA")
If rpta = vbYes Then
fila = 2
While Sheets("Proveedores").Cells(fila, 1) <> Empty
fila = fila + 1
Wend


Sheets("Proveedores").Cells(fila, 1) = nom.Text
Sheets("Proveedores").Cells(fila, 2) = "00-" & (fila - 1)
Sheets("Proveedores").Cells(fila, 3) = rucc.Text
Sheets("Proveedores").Cells(fila, 4) = direc1.Text
Sheets("Proveedores").Cells(fila, 5) = direc2.Text
Sheets("Proveedores").Cells(fila, 6) = fono.Text
Sheets("Proveedores").Cells(fila, 7) = cel.Text
Sheets("Proveedores").Cells(fila, 8) = faxx.Text
Sheets("Proveedores").Cells(fila, 9) = emaill.Text
Sheets("Proveedores").Cells(fila, 9) = contacto.Text
Sheets("Proveedores").Cells(fila, 9) = cuenta.Text

End If
nom.Text = ""
rucc.Text = ""
direc1.Text = ""
direc2.Text = ""
fono.Text = ""
cel.Text = ""
contacto.Text = ""
emaill.Text = ""
cuenta.Text = ""
faxx.Text = ""

nom.Enabled = False
rucc.Enabled = False
direc1.Enabled = False
direc2.Enabled = False
fono.Enabled = False
cel.Enabled = False
contacto.Enabled = False
emaill.Enabled = False
cuenta.Enabled = False
faxx.Enabled = False



End Sub


Private Sub CommandButton3_Click()
Unload CrearProveedores

End Sub

Private Sub CommandButton4_Click()
nom.Enabled = True
rucc.Enabled = True
direc1.Enabled = True
direc2.Enabled = True
fono.Enabled = True
cel.Enabled = True
contacto.Enabled = True
emaill.Enabled = True
cuenta.Enabled = True
faxx.Enabled = True
nom.Text = ""
rucc.Text = ""
direc1.Text = ""
direc2.Text = ""
fono.Text = ""
cel.Text = ""
contacto.Text = ""
emaill.Text = ""
cuenta.Text = ""
faxx.Text = ""
nom.SetFocus

End Sub

Private Sub Label22_Click()

End Sub

Private Sub nom_Change()

End Sub

Private Sub Userform_Activate()
nom.Text = ""
rucc.Text = ""
direc1.Text = ""
direc2.Text = ""
fono.Text = ""
cel.Text = ""
contacto.Text = ""
emaill.Text = ""
cuenta.Text = ""
faxx.Text = ""

nom.Enabled = False
rucc.Enabled = False
direc1.Enabled = False
direc2.Enabled = False
fono.Enabled = False
cel.Enabled = False
contacto.Enabled = False
emaill.Enabled = False
cuenta.Enabled = False
faxx.Enabled = False

End Sub
