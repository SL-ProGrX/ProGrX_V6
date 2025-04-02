VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCpr_Proveedores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proveedores para el Proceso de Compras"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbImportar 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Importar de Cuentas por Pagar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnImportar 
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   480
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Importar Proveedores Activos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCpr_Proveedores.frx":0000
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8880
      TabIndex        =   0
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   6375
      _Version        =   1441793
      _ExtentX        =   11239
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   330
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtObservacion 
      Height          =   1035
      Left            =   1320
      TabIndex        =   11
      Top             =   2400
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   1826
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   6720
      TabIndex        =   13
      Top             =   1080
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCedJur 
      Height          =   330
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTelMovil 
      Height          =   330
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   1110
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Observación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. Móvil"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmCpr_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean

Private Sub btnImportar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCPR_Proveedores_Importar"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Proveedores Sincronizados/Importados Satisfactoriamente!", vbInformation
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedJur.SetFocus
End Sub

'Private Sub cboClasificacion_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
'End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelMovil.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 PROVEEDOR_CODIGO from CPR_PROVEEDORES_TEMPO"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where PROVEEDOR_CODIGO > " & IIf(txtCodigo = "", 0, txtCodigo) & " order by PROVEEDOR_CODIGO asc"
    Else
       strSQL = strSQL & " where PROVEEDOR_CODIGO < " & IIf(txtCodigo = "", 0, txtCodigo) & " order by PROVEEDOR_CODIGO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!PROVEEDOR_CODIGO)
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
vModulo = 35
End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 35

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 

cbo.Clear
cbo.AddItem "Persona Física"
cbo.AddItem "Entidad Juridica"


'strSQL = "select Rtrim(Cod_clasificacion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
'       & " from cxp_prov_clas order by descripcion"
'Call sbCbo_Llena_New(cboClasificacion, strSQL, False, True)

 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()


vCodigo = 0
txtCodigo.Text = ""


cbo.Text = "Entidad Juridica"

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "InActivo"
cboEstado.Text = "Activo"

txtNombre.Text = ""
txtObservacion.Text = ""

txtCedJur.Text = ""

txtEmail.Text = ""

txtTelMovil.Text = ""

txtCodigo.Enabled = True

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNombre.SetFocus
      txtCodigo.Enabled = False
      
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
         gBusquedas.Columna = "descripcion"
         gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select PROVEEDOR_CODIGO, cedjur, descripcion from CPR_PROVEEDORES_TEMPO"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select P.*" _
       & " from CPR_PROVEEDORES_TEMPO P" _
       & " where P.PROVEEDOR_CODIGO = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  vEdita = True
  vCodigo = rs!PROVEEDOR_CODIGO
  txtCodigo.Text = CStr(rs!PROVEEDOR_CODIGO)
  
    txtNombre.Text = rs!Descripcion & ""
    txtObservacion.Text = rs!observacion & ""
    
    txtCedJur.Text = rs!CEDJUR & ""
    
    cboEstado.Clear
    cboEstado.AddItem "Activo"
    cboEstado.AddItem "InActivo"
    
    Select Case rs!Estado
        Case "A"
            cboEstado.Text = "Activo"
        Case "I"
            cboEstado.Text = "InActivo"
        Case "S"
            cboEstado.Clear
            cboEstado.AddItem "Suspendido"
            cboEstado.Text = "Suspendido"
    End Select
    
    
    Select Case rs!Tipo
      Case "P", "F"
          cbo.Text = "Persona Física"
      Case "E", "J"
          cbo.Text = "Entidad Juridica"
    End Select
    
    txtEmail.Text = rs!Email & ""
    
    txtTelMovil.Text = Trim(rs!telefono & "")
   
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
Dim vCuenta As String, vDivisa As String
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Verifica que exista ningun otro proveedor con la misma cedula juridica
strSQL = "select isnull(count(*),0) as Existe from CPR_PROVEEDORES_TEMPO" _
       & " where PROVEEDOR_CODIGO not in(" & vCodigo & ") and cedJur = '" _
       & Trim(txtCedJur) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   vMensaje = vMensaje & vbCrLf & " - Existe ya un Proveedor registrado con la misma Cédula Jurídica ..."
End If
rs.Close

txtEmail.Text = Trim(txtEmail.Text)

If Not fxEmail_Valida(txtEmail.Text) Then
    vMensaje = vMensaje & " - El Email principal no es válido!" & vbCrLf
End If

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()

On Error GoTo vError


If vEdita Then
  strSQL = "update CPR_PROVEEDORES_TEMPO set descripcion = '" & Trim(txtNombre) & "', CedJur = '" & txtCedJur _
         & "', Tipo = '" & Mid(cbo.Text, 1, 1) & "',observacion = '" & txtObservacion & "', Estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "', email = '" & txtEmail & "', telefono = '" & txtTelMovil _
         & "', MODIFICA_FECHA = dbo.MyGetdate(), MODIFICA_USUARIO = '" & glogon.Usuario & "'" _
         & "  where PROVEEDOR_CODIGO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  
  
  Call Bitacora("Modifica", "Proveedor Compras Cod: " & vCodigo)

Else
   strSQL = "select isnull(max(PROVEEDOR_CODIGO), 10000) as ultimo from CPR_PROVEEDORES_TEMPO"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo = rs!ultimo + 1
     vCodigo = txtCodigo
   rs.Close
   
   strSQL = "insert into CPR_PROVEEDORES_TEMPO(PROVEEDOR_CODIGO, Tipo, Descripcion, Observacion" _
          & ", estado , Telefono, Email, CedJur,  REGISTRO_FECHA, REGISTRO_USUARIO) values(" _
          & vCodigo & ", '" & Mid(cbo.Text, 1, 1) & "','" & txtNombre.Text & "', '" & txtObservacion _
          & "','" & Mid(cboEstado.Text, 1, 1) & "', '" & txtTelMovil.Text & "','" & txtEmail _
          & "', '" & txtCedJur & "', dbo.MyGetdate(), '" & glogon.Usuario & "'" & ")"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Proveedor Compras Cod: " & vCodigo)
    
   txtCodigo.Enabled = True
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CPR_PROVEEDORES_TEMPO where PROVEEDOR_CODIGO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Proveedor Compras Cod: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCedJur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = "Id. Real"
  
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedjur"
  gBusquedas.Orden = "cedjur"
  gBusquedas.Consulta = "select PROVEEDOR_CODIGO, Descripcion, CedJur from CPR_PROVEEDORES_TEMPO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If


End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "PROVEEDOR_CODIGO"
  gBusquedas.Orden = "PROVEEDOR_CODIGO"
  gBusquedas.Consulta = "select PROVEEDOR_CODIGO,cedjur,descripcion from CPR_PROVEEDORES_TEMPO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select PROVEEDOR_CODIGO,cedjur,descripcion from CPR_PROVEEDORES_TEMPO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub


Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtNombre.SetFocus
End If
End Sub

Private Sub txtTelMovil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub
