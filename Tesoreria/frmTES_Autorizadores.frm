VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmTES_Autorizadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios Autorizadores de Documentos"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   Icon            =   "frmTES_Autorizadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9225
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   7440
      TabIndex        =   3
      Top             =   960
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtRngFirmasDesde 
      Height          =   312
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   4800
      Width           =   2772
      _Version        =   1310723
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRngFirmasHasta 
      Height          =   312
      Left            =   5160
      TabIndex        =   5
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   5160
      Width           =   2772
      _Version        =   1310723
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRngAutoDesde 
      Height          =   312
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3360
      Width           =   2772
      _Version        =   1310723
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRngAutoHasta 
      Height          =   312
      Left            =   5160
      TabIndex        =   9
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3720
      Width           =   2772
      _Version        =   1310723
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   12
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Width           =   2052
      _Version        =   1310723
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   1800
      TabIndex        =   13
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2040
      Width           =   2052
      _Version        =   1310723
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1800
      TabIndex        =   14
      Top             =   2400
      Width           =   2052
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1032
      Left            =   3960
      TabIndex        =   15
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Width           =   3972
      _Version        =   1310723
      _ExtentX        =   7006
      _ExtentY        =   1820
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Autorizadores"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1800
      TabIndex        =   20
      Top             =   240
      Width           =   4692
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notas:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   6720
      TabIndex        =   19
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   600
      TabIndex        =   18
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   600
      TabIndex        =   17
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   16
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4320
      TabIndex        =   11
      Top             =   3720
      Width           =   732
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4320
      TabIndex        =   10
      Top             =   3360
      Width           =   732
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   5160
      Width           =   732
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   4320
      TabIndex        =   6
      Top             =   4800
      Width           =   732
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rango General de Autorización de Impresión de Firmas Electrónicas: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Cuenta de Inventarios para la Bodega"
      Top             =   4320
      Width           =   7332
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rango General de Autorización de Emisión de Solicitudes: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Cuenta de Inventarios para la Bodega"
      Top             =   2880
      Width           =   5532
   End
End
Attribute VB_Name = "frmTES_Autorizadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select nombre from tes_autorizaciones"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where nombre > '" & txtCodigo.Text & "' order by nombre asc"
    Else
       strSQL = strSQL & " where nombre < '" & txtCodigo.Text & "' order by nombre desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Nombre
      Call sbConsulta(txtCodigo.Text)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 9
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 9
 
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""

txtCodigo.Text = ""

cbo.Clear
cbo.AddItem "Activo"
cbo.AddItem "Inactivo"
cbo.Text = "Activo"

txtClave.Text = ""
txtNotas.Text = ""

txtRngAutoDesde.Text = "0.00"
txtRngAutoHasta.Text = "0.00"

txtRngFirmasDesde.Text = "0.00"
txtRngFirmasHasta.Text = "0.00"


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
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select nombre,notas from tes_Autorizaciones"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo.Text = gBusquedas.Resultado
       txtClave.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from tes_Autorizaciones where nombre = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!Nombre
  txtCodigo = rs!Nombre
 
  If rs!Estado = "A" Then
    cbo.Text = "Activo"
  Else
    cbo.Text = "Inactivo"
  End If
       
  txtNotas = rs!Notas
  txtClave = ""


txtRngAutoDesde.Text = Format(rs!rango_gen_Inicio, "Standard")
txtRngAutoHasta.Text = Format(rs!rango_gen_corte, "Standard")

txtRngFirmasDesde.Text = Format(rs!firmas_gen_inicio, "Standard")
txtRngFirmasHasta.Text = Format(rs!firmas_gen_corte, "Standard")

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtCodigo = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Usuario no es válido ..."
If Not IsNumeric(txtRngAutoDesde.Text) Then vMensaje = vMensaje & vbCrLf & " - El Rango de Autorización de Emisión [DESDE] no es válido..."
If Not IsNumeric(txtRngAutoHasta.Text) Then vMensaje = vMensaje & vbCrLf & " - El Rango de Autorización de Emisión [HASTA] no es válido..."
If Not IsNumeric(txtRngFirmasDesde.Text) Then vMensaje = vMensaje & vbCrLf & " - El Rango de Autorización de Firmas [DESDE] no es válido..."
If Not IsNumeric(txtRngFirmasHasta.Text) Then vMensaje = vMensaje & vbCrLf & " - El Rango de Autorización de Firmas [HASTA] no es válido..."

If Len(vMensaje) = 0 Then
  If CCur(txtRngAutoDesde.Text) > CCur(txtRngAutoHasta.Text) Then vMensaje = vMensaje & vbCrLf & " - El Rango de Autorización de Emisión [DESDE es Mayor que HASTA]"
  If CCur(txtRngFirmasDesde.Text) > CCur(txtRngFirmasHasta.Text) Then vMensaje = vMensaje & vbCrLf & " - El Rango de Autorización de Firmas [DESDE es Mayor que HASTA]"

End If
  
If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

'   RANGO_GEN_INICIO  DEC(18,2) DEFAULT 0
'  ,RANGO_GEN_CORTE   DEC(18,2) DEFAULT 0
'  ,FIRMAS_GEN_INICIO DEC(18,2) DEFAULT 0
'  ,FIRMAS_GEN_CORTE  DEC(18,2) DEFAULT 0

If vEdita Then
  strSQL = "update tes_autorizaciones set Notas = '" & txtNotas & "'" _
         & ",estado = '" & Mid(cbo.Text, 1, 1) & "',Clave = '" & fxTESCifrado(txtClave) _
         & "', RANGO_GEN_INICIO = " & CCur(txtRngAutoDesde.Text) & ", RANGO_GEN_CORTE = " & CCur(txtRngAutoHasta.Text) _
         & ", FIRMAS_GEN_INICIO = " & CCur(txtRngFirmasDesde.Text) & ",FIRMAS_GEN_CORTE = " & CCur(txtRngFirmasHasta.Text) _
         & " where nombre = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Usuario Autorizador : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into tes_autorizaciones(nombre,clave,notas,estado,RANGO_GEN_INICIO,RANGO_GEN_CORTE" _
          & ",FIRMAS_GEN_INICIO,FIRMAS_GEN_CORTE)" _
          & " values('" & vCodigo & "','" & fxTESCifrado(txtClave) & "','" & txtNotas _
          & "','" & Mid(cbo.Text, 1, 1) & "'," & CCur(txtRngAutoDesde.Text) & "," & CCur(txtRngAutoHasta.Text) _
          & "," & CCur(txtRngFirmasDesde.Text) & "," & CCur(txtRngFirmasHasta.Text) & ")"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Usuario Autorizador: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete tes_autorizaciones where nombre = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Usuario Autorizador : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtClave.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select nombre,Notas from tes_autorizaciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRngAutoDesde.SetFocus
End Sub


Private Sub txtRngAutoDesde_GotFocus()
On Error GoTo vError
  txtRngAutoDesde.Text = CCur(txtRngAutoDesde.Text)
vError:
End Sub

Private Sub txtRngAutoDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRngAutoHasta.SetFocus
End Sub


Private Sub txtRngAutoDesde_LostFocus()
On Error GoTo vError
  txtRngAutoDesde.Text = Format(CCur(txtRngAutoDesde.Text), "Standard")
vError:
End Sub

Private Sub txtRngAutoHasta_GotFocus()
On Error GoTo vError
  txtRngAutoHasta.Text = CCur(txtRngAutoHasta.Text)
vError:
End Sub

Private Sub txtRngAutoHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRngFirmasDesde.SetFocus
End Sub

Private Sub txtRngAutoHasta_LostFocus()
On Error GoTo vError
  txtRngAutoHasta.Text = Format(CCur(txtRngAutoHasta.Text), "Standard")
vError:
End Sub

Private Sub txtRngFirmasDesde_GotFocus()
On Error GoTo vError
  txtRngFirmasDesde.Text = CCur(txtRngFirmasDesde.Text)
vError:
End Sub

Private Sub txtRngFirmasDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRngFirmasHasta.SetFocus
End Sub


Private Sub txtRngFirmasDesde_LostFocus()
On Error GoTo vError
  txtRngFirmasDesde.Text = Format(CCur(txtRngFirmasDesde.Text), "Standard")
vError:
End Sub

Private Sub txtRngFirmasHasta_GotFocus()
On Error GoTo vError
  txtRngFirmasHasta.Text = CCur(txtRngFirmasHasta.Text)
vError:
End Sub

Private Sub txtRngFirmasHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub

Private Sub txtRngFirmasHasta_LostFocus()
On Error GoTo vError
  txtRngFirmasHasta.Text = Format(CCur(txtRngFirmasHasta.Text), "Standard")
vError:
End Sub
