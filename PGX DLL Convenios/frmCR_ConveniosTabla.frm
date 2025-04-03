VERSION 5.00
Begin VB.Form frmCR_ConveniosTabla 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla para Cálculo de Créditos de Convenios"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7635
   Icon            =   "frmCR_ConveniosTabla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "Nombre de la Persona"
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdCalcular 
      Appearance      =   0  'Flat
      Caption         =   "&Calcular"
      Height          =   855
      Left            =   6360
      Picture         =   "frmCR_ConveniosTabla.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtEspecialDisp 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtComercialDisp 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtEspecialMora 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtComercialMora 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtEspecialSaldos 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtComercialSaldos 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtMontoSolicitado 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "Monto Solicitado"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtPlazo 
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtInteres 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtCuota 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Cuota Calculada"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtCedula 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Cédula de la Persona"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblEstado 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   960
      TabIndex        =   24
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label Label6 
      Caption         =   "Estado"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disponible"
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   21
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mora"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   20
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldos"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   19
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Especial"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Comercial"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   7440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Monto"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Plazo"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Interes"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Cuota"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Cédula"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmCR_ConveniosTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcular_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPasa As Boolean, vMembresia As Currency

On Error GoTo vError

'Procedimiento:
' 1. Buscar Saldos de Creditos Activos
' 2. Buscar Montos Solicitados de Creditos En Tramites y Sumarlos a los Saldos
' 3. Buscar Moras de Creditos Activos (Intereses Corrientes y Moratorios)
' 4. Sacar Membresia y Monto Segun Tabla
' 5. Restar al monto segun Membresias Saldos y Moras Cargadas...
' Fin


'1.
strSQL = "select C.tipo,coalesce(sum(R.saldo),0) as Saldo" _
       & " from reg_creditos R inner join convenios_codigos C on R.codigo = C.codigo" _
       & " where R.estado = 'A' and R.saldo > 0 and cedula = '" & txtCedula & "'" _
       & " group by C.tipo"
rs.Open strSQL, glogon.Conection, adOpenStatic

txtComercialSaldos = 0
txtEspecialSaldos = 0

txtComercialMora.ForeColor = vbBlue
txtEspecialMora.ForeColor = vbBlue
txtComercialMora = "Al Día"
txtEspecialMora = "Al Día"


If rs.BOF And rs.EOF Then
    vPasa = False
    txtComercialSaldos = 0
    txtEspecialSaldos = 0
Else
    vPasa = True
    Do While Not rs.EOF
     If rs!Tipo = "C" Then
        txtComercialSaldos = Format(rs!Saldo, "Standard")
     Else
        txtEspecialSaldos = Format(rs!Saldo, "Standard")
     End If
     rs.MoveNext
    Loop
End If
rs.Close


'2.
strSQL = "select C.tipo,coalesce(sum(R.montosol),0) as Monto" _
       & " from reg_creditos R inner join convenios_codigos C on R.codigo = C.codigo" _
       & " where R.estadosol in('R','A','P') and R.estado is null and cedula = '" & txtCedula & "'" _
       & " group by C.tipo"
rs.Open strSQL, glogon.Conection, adOpenStatic


If Not rs.BOF And Not rs.EOF Then
    Do While Not rs.EOF
     If rs!Tipo = "C" Then
        txtComercialSaldos = Format(CCur(txtComercialSaldos) + rs!Monto, "Standard")
     Else
        txtEspecialSaldos = Format(CCur(txtEspecialSaldos) + rs!Monto, "Standard")
     End If
     rs.MoveNext
    Loop
End If
rs.Close



'3.

 strSQL = "select coalesce(count(*),0) as Existe " _
        & " from reg_creditos R inner join Morosidad M on R.id_solicitud = M.id_solicitud" _
        & " where R.cedula = '" & txtCedula & "' and R.estado = 'A' and R.saldo > 0" _
        & " and M.estado = 'A'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe > 0 Then
    txtComercialMora.ForeColor = vbRed
    txtEspecialMora.ForeColor = vbRed
    txtComercialMora = "Moroso"
    txtEspecialMora = "Moroso"
 End If
 rs.Close


''If vPasa Then 'Busca Mora
''   'Convenios Comerciales
''    strSQL = "select coalesce(sum(intc),0) as intc,coalesce(sum(intm),0) as intM" _
''           & " from morosidad where estado = 'A' and id_solicitud in(" _
''           & " select id_solicitud " _
''           & " from reg_creditos R inner join convenios_codigos C on R.codigo = C.codigo" _
''           & " where R.estado = 'A' and R.saldo > 0 and cedula = '" & txtCedula & "'" _
''           & " and Tipo = 'C')"
''    rs.Open strSQL, glogon.Conection, adOpenStatic
''    If rs.EOF And rs.BOF Then
''      txtComercialMora = 0
''    Else
''      txtComercialMora = Format(rs!intc + rs!intm, "Standard")
''    End If
''    rs.Close
''
''
''   'Convenios Especiales
''    strSQL = "select coalesce(sum(intc),0) as intc,coalesce(sum(intm),0) as intM" _
''           & " from morosidad where estado = 'A' and id_solicitud in(" _
''           & " select id_solicitud " _
''           & " from reg_creditos R inner join convenios_codigos C on R.codigo = C.codigo" _
''           & " where R.estado = 'A' and R.saldo > 0 and cedula = '" & txtCedula & "'" _
''           & " and Tipo = 'E')"
''    rs.Open strSQL, glogon.Conection, adOpenStatic
''    If rs.EOF And rs.BOF Then
''      txtEspecialMora = 0
''    Else
''      txtEspecialMora = Format(rs!intc + rs!intm, "Standard")
''    End If
''    rs.Close
''
''End If

'Buscar Disponible Con Membresia  (Tabla - Saldos y Mora)
strSQL = "select datediff(mm, fechaingreso , getdate()) as Membresia" _
       & " from socios where cedula = '" & txtCedula & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  vMembresia = 0
Else
  vMembresia = rs!membresia
End If
rs.Close

strSQL = "select monto from convenios_tablas where tipo = 'C'" _
       & " and " & vMembresia & " between desde and hasta"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  txtComercialDisp = 0
Else
  txtComercialDisp = rs!Monto - (CCur(txtComercialSaldos))
End If
rs.Close


strSQL = "select monto from convenios_tablas where tipo = 'E'" _
       & " and " & vMembresia & " between desde and hasta"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  txtEspecialDisp = 0
Else
  txtEspecialDisp = rs!Monto - (CCur(txtEspecialSaldos))
End If
rs.Close

txtComercialDisp = Format(txtComercialDisp, "Standard")
txtEspecialDisp = Format(txtEspecialDisp, "Standard")

If CCur(txtComercialDisp) > 0 Then
   txtComercialDisp.ForeColor = vbBlue
Else
   txtComercialDisp.ForeColor = vbRed
End If

If CCur(txtEspecialDisp) > 0 Then
   txtEspecialDisp.ForeColor = vbBlue
Else
   txtEspecialDisp.ForeColor = vbRed
End If

lblEstado.Caption = ""

If Not IsNumeric(txtMontoSolicitado) Then txtMontoSolicitado = 0

If CCur(txtComercialDisp) >= CCur(txtMontoSolicitado) Then
  lblEstado.Caption = ">> La Solicitud se puede APROBAR en Convenio Comercial..."
Else
  lblEstado.Caption = ">> La Solicitud es RECHAZADA en Convenio Comercial..."
End If


If CCur(txtEspecialDisp) >= CCur(txtMontoSolicitado) Then
  lblEstado.Caption = lblEstado.Caption & vbCrLf & ">> La Solicitud se puede APROBAR en Convenio Especial..."
Else
  lblEstado.Caption = lblEstado.Caption & vbCrLf & ">> La Solicitud es RECHAZADA en Convenio Especial..."
End If


Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
'En los saldos incluir las que estan en tramite en el ultimo mes
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Cedula"
  gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
  gBusquedas.Orden = "Cedula"
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select nombre from socios where cedula = '" & txtCedula & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
 MsgBox "No se encontró ningun registro de esta cédula...", vbExclamation
Else
 txtNombre = rs!Nombre & ""
End If
rs.Close
End Sub

Private Sub txtInteres_Change()
On Error GoTo vError
If CCur(IIf((txtInteres = ""), 0, txtInteres)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMontoSolicitado = ""), 0, txtMontoSolicitado)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMontoSolicitado), CCur(txtPlazo), CCur(txtInteres))
End If
vError:

End Sub

Private Sub txtMontoSolicitado_Change()
On Error GoTo vError
If CCur(IIf((txtInteres = ""), 0, txtInteres)) > 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMontoSolicitado = ""), 0, txtMontoSolicitado)) > 0 Then
 txtCuota = fxCalcula_Cuota(CCur(txtMontoSolicitado), CCur(txtPlazo), CCur(txtInteres))
End If

vError:
End Sub

Private Sub txtInteres_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  On Error Resume Next
    If CCur(IIf((txtInteres = ""), 0, txtInteres)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
        And CCur(IIf((txtMontoSolicitado = ""), 0, txtMontoSolicitado)) > 0 Then
      txtCuota = fxCalcula_Cuota(CCur(txtMontoSolicitado), CCur(txtPlazo), CCur(txtInteres))
    End If
   cmdCalcular.SetFocus
End If
End Sub

Private Sub txtInteres_LostFocus()
On Error GoTo vError
If CCur(IIf((txtInteres = ""), 0, txtInteres)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMontoSolicitado = ""), 0, txtMontoSolicitado)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMontoSolicitado), CCur(txtPlazo), CCur(txtInteres))
End If
vError:
End Sub

Private Sub txtMontoSolicitado_GotFocus()
On Error GoTo vError
txtMontoSolicitado = CCur(txtMontoSolicitado)
vError:
End Sub

Private Sub txtMontoSolicitado_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtMontoSolicitado_LostFocus()
txtMontoSolicitado = Format(txtMontoSolicitado, "Standard")
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
  gBusquedas.Orden = "Nombre"
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoSolicitado.SetFocus

End Sub

Private Sub txtPlazo_Change()
On Error GoTo vError
If CCur(IIf((txtInteres = ""), 0, txtInteres)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMontoSolicitado = ""), 0, txtMontoSolicitado)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMontoSolicitado), CCur(txtPlazo), CCur(txtInteres))
End If

vError:
End Sub


Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtInteres.SetFocus
End Sub


