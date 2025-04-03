VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAF_ComDirRegional 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de direcciones Regionales"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7185
   Begin VB.ComboBox cboProvincia 
      Height          =   315
      ItemData        =   "frmAF_ComDirRegional.frx":0000
      Left            =   1200
      List            =   "frmAF_ComDirRegional.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   2040
      MaxLength       =   35
      TabIndex        =   1
      ToolTipText     =   "Descripción del código de la línea"
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Código del Préstamo"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtTelefono 
      Height          =   315
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   5
      ToolTipText     =   "Código del Préstamo"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtDirector 
      Height          =   315
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   3
      ToolTipText     =   "Código del Préstamo"
      Top             =   1200
      Width           =   5895
   End
   Begin VB.TextBox txtFax 
      Height          =   315
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   6
      ToolTipText     =   "Código del Préstamo"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtadministrador 
      Height          =   315
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   2
      ToolTipText     =   "Código del Préstamo"
      Top             =   840
      Width           =   5895
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7185
      _ExtentX        =   12674
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Director"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fax"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Administrador"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmAF_ComDirRegional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vEdita As Boolean



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
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_direccion,descripcion from afi_comdirregional"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
        If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
       txtDescripcion.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select


End Sub

Private Sub txtadministrador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDirector.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_direccion"
  gBusquedas.Orden = "cod_direccion"
  gBusquedas.Consulta = "select cod_direccion,descripcion from afi_comdirregional"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
  txtDescripcion.SetFocus
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_direccion,descripcion from afi_comdirregional"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtadministrador.SetFocus

End Sub
Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = ""
txtCodigo = ""

cboProvincia.Clear

cboProvincia.AddItem "SAN JOSE"
cboProvincia.AddItem "ALAJUELA"
cboProvincia.AddItem "CARTAGO"
cboProvincia.AddItem "HEREDIA"
cboProvincia.AddItem "GUANACASTE"
cboProvincia.AddItem "PUNTARENAS"
cboProvincia.AddItem "LIMON"


For i = 0 To 6
  cboProvincia.ListIndex = i
  cboProvincia.ItemData(cboProvincia.ListIndex) = i + 1
  
Next i

cboProvincia.Text = "SAN JOSE"

txtFax = ""
txtTelefono = ""
txtDescripcion = ""
txtadministrador = ""
txtDirector = ""


End Sub



Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from afi_comdirregional where cod_direccion = '" & xCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_direccion
  txtCodigo = rs!cod_direccion
 
  txtDescripcion = rs!Descripcion
      
  Select Case rs!provincia
     Case 1
        cboProvincia.Text = "SAN JOSE"
     Case 2
        cboProvincia.Text = "ALAJUELA"
     Case 3
        cboProvincia.Text = "CARTAGO"
     Case 4
        cboProvincia.Text = "HEREDIA"
     Case 5
        cboProvincia.Text = "GUANACASTE"
     Case 6
        cboProvincia.Text = "PUNTARENAS"
     Case 7
        cboProvincia.Text = "LIMON"
  End Select
  
  txtTelefono = rs!telefono
  txtFax = rs!fax
  txtDirector = rs!director
  txtadministrador = rs!administrador

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
  strSQL = "update afi_comdirregional set descripcion = '" & UCase(Trim(txtDescripcion)) & "'" _
         & ",director = '" & txtDirector & "',administrador = '" & txtadministrador _
         & "',telefono = '" & txtTelefono & "',fax = '" & txtFax & "',provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) & "" _
         & " where cod_direccion = '" & vCodigo & "'"
  glogon.Conection.Execute strSQL
  Call Bitacora("Modifica", "Dir. Regional : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert afi_comdirregional(cod_direccion,descripcion,director,administrador,telefono" _
          & ",fax,provincia)" _
          & " values('" & vCodigo & "','" & txtDescripcion & "','" & txtDirector & "','" _
          & txtadministrador & "','" & txtTelefono & "','" & txtFax & "'," & cboProvincia.ItemData(cboProvincia.ListIndex) & ")"
          
   glogon.Conection.Execute strSQL
    
   Call Bitacora("Registra", "Dir. Regional: " & vCodigo)
 
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
  strSQL = "delete afi_comdirregional where cod_direccion = '" & vCodigo & "'"
  glogon.Conection.Execute strSQL
  
  Call Bitacora("Elimina", "Comite : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtDirector_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus

End Sub

