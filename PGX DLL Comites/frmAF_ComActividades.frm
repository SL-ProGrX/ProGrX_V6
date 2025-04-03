VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAF_ComActividades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Actividades"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   7110
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   14
      ToolTipText     =   "Código del Préstamo"
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cboEvento 
      Height          =   315
      ItemData        =   "frmAF_ComActividades.frx":0000
      Left            =   1200
      List            =   "frmAF_ComActividades.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   11
      ToolTipText     =   "Código del Préstamo"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtDia 
      Height          =   315
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   9
      ToolTipText     =   "Código del Préstamo"
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cboEstado 
      Height          =   315
      ItemData        =   "frmAF_ComActividades.frx":0020
      Left            =   4680
      List            =   "frmAF_ComActividades.frx":002A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtNotas 
      Height          =   1035
      Left            =   1200
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Código del Préstamo"
      Top             =   2280
      Width           =   5775
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4680
      MaxLength       =   60
      TabIndex        =   3
      ToolTipText     =   "Código del Préstamo"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1920
      MaxLength       =   35
      TabIndex        =   1
      ToolTipText     =   "Descripción del código de la línea"
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Código del Préstamo"
      Top             =   480
      Width           =   735
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Style           =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   4
      Left            =   1200
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Activación"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Evento"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Día"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   1200
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Estado"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Monto"
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   855
   End
End
Attribute VB_Name = "frmAF_ComActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vEdita As Boolean

Private Sub cboevento_Click()
If Mid(cboEvento.Text, 1, 1) = "U" Then
   txtAnio.Enabled = False
Else
   txtAnio.Enabled = True
End If

End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 1
 
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox Err.Description, vbExclamation
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
        Call sbToolBar(tlb, "nuevo")
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_actividad,descripcion from afi_comactividades"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       vCodigo = gBusquedas.Resultado
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.Text = gBusquedas.Resultado2
       txtDescripcion.SetFocus
       Call sbConsulta(vCodigo)
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select


End Sub

Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub txtCodigo_GotFocus()
txtCodigo = UCase(txtCodigo)
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
  
End If

If KeyCode = vbKeyF4 Then

  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_actividad"
  gBusquedas.Orden = "cod_actividad"
  gBusquedas.Consulta = "select cod_actividad,descripcion from afi_comactividades"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
  
  txtDescripcion.SetFocus
  
End If

End Sub

Private Sub txtContacto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
txtCodigo = UCase(txtCodigo)
End Sub

Private Sub txtDescripcion_GotFocus()
txtDescripcion = UCase(txtDescripcion)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_actividad,descripcion from afi_comactividades"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
  
End If

 If KeyCode = vbKeyReturn Then cboEvento.SetFocus
 
End Sub
Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCodigo = ""

cboEstado.Clear
cboEstado.AddItem "ACTIVO"
cboEstado.AddItem "INACTIVO"
cboEstado.Text = "ACTIVO"

cboEvento.Clear
cboEvento.AddItem "PERIODICO"
cboEvento.AddItem "UNICO"
cboEvento.Text = "PERIODICO"

txtNotas = ""
txtDescripcion = ""
txtMes = ""
txtMonto = ""
txtDia = ""
txtAnio = ""

End Sub



Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from afi_comactividades where cod_actividad = '" & xCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_actividad
  txtCodigo = rs!cod_actividad
 
  txtDescripcion = rs!Descripcion
  txtMes = rs!Mes
  txtAnio = IIf(Not IsNull(rs!Anio), rs!Anio, "")
  txtDia = rs!dia
  txtMonto = Format(rs!Monto, "Standard")
      
  If rs!Tipo = "P" Then
    cboEvento.Text = "PERIODICO"
  Else
    cboEvento.Text = "UNICO"
  End If
  
  If rs!Estado = "A" Then
    cboEstado.Text = "ACTIVO"
  Else
    cboEstado.Text = "INACTIVO"
  End If
  txtNotas = rs!notas

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripción del comite no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update afi_comactividades set descripcion = '" & UCase(Trim(txtDescripcion)) & "'" _
         & ",notas = '" & txtNotas & "',estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "',dia = '" & txtDia & "',mes = '" & txtMes & "',anio = '" & txtAnio & "'" _
         & ",tipo ='" & Mid(cboEvento.Text, 1, 1) & "',monto = " & CCur(txtMonto) & "" _
         & " where cod_actividad = '" & vCodigo & "'"
  glogon.Conection.Execute strSQL
  Call Bitacora("Modifica", "Actividad : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert afi_comactividades(cod_actividad,descripcion,estado,notas,dia" _
          & ",mes,anio,tipo,monto)" _
          & " values('" & vCodigo & "','" & txtDescripcion & "','" & Mid(cboEstado.Text, 1, 1) & "','" _
          & txtNotas & "','" & txtDia & "','" & txtMes & "','" & txtAnio & "','" & Mid(cboEvento.Text, 1, 1) & "'," & CCur(txtMonto) & ")"
          
   glogon.Conection.Execute strSQL
    
   Call Bitacora("Registra", "Actividad: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete afi_comactividades where cod_actividad = '" & vCodigo & "'"
  glogon.Conection.Execute strSQL
  
  Call Bitacora("Elimina", "Actividad : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub





Private Sub txtDescripcion_LostFocus()
txtDescripcion = UCase(txtDescripcion)
End Sub

Private Sub txtDia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtMes.SetFocus
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAnio.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto = CCur(txtMonto)
vError:

End Sub

Private Sub txtMOnto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto = Format(txtMonto, "Standard")

vError:


End Sub

Private Sub txtNotas_GotFocus()
txtNotas = UCase(txtNotas)
End Sub

Private Sub txtNotas_LostFocus()
txtNotas = UCase(txtNotas)
End Sub
