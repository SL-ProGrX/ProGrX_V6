VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_DiferidosPlantilla 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diferidos : Plantillas"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1815
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   11775
      _Version        =   1310723
      _ExtentX        =   20770
      _ExtentY        =   3201
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   795
         Left            =   1200
         TabIndex        =   14
         Top             =   720
         Width           =   10575
         _Version        =   1310723
         _ExtentX        =   18653
         _ExtentY        =   1402
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
      Begin XtremeSuiteControls.FlatEdit txtCAsiento 
         Height          =   330
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   975
         _Version        =   1310723
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtDAsiento 
         Height          =   330
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   4215
         _Version        =   1310723
         _ExtentX        =   7435
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   8160
         TabIndex        =   20
         Top             =   360
         Width           =   3615
         _Version        =   1310723
         _ExtentX        =   6376
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Asiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Diferido"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   6600
         TabIndex        =   18
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   945
      End
   End
   Begin XtremeSuiteControls.CheckBox chkAsientoXPlantilla 
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Width           =   9495
      _Version        =   1310723
      _ExtentX        =   16748
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Genera Asientos de Diferidos por Plantilla?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   7215
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario que Crea la Plantilla"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Fecha y Hora de la Creación"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "DIFERIDOS"
            TextSave        =   "DIFERIDOS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   11775
      _Version        =   524288
      _ExtentX        =   20770
      _ExtentY        =   5741
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_DiferidosPlantilla.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodPlantilla 
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   600
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtConsecutivo 
      Height          =   315
      Left            =   9000
      TabIndex        =   12
      Top             =   600
      Width           =   1575
      _Version        =   1310723
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDesPlantilla 
      Height          =   315
      Left            =   2160
      TabIndex        =   13
      Top             =   600
      Width           =   6015
      _Version        =   1310723
      _ExtentX        =   10610
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtDebito 
      Height          =   315
      Left            =   9600
      TabIndex        =   21
      Top             =   6720
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
      _ExtentY        =   556
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCredito 
      Height          =   315
      Left            =   10560
      TabIndex        =   22
      Top             =   6720
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
      _ExtentY        =   556
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDiferencia 
      Height          =   315
      Left            =   6480
      TabIndex        =   10
      Top             =   6720
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   192
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos Asiento de Plantilla Base "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   6720
      Width           =   3975
   End
   Begin VB.Label lsblr 
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   6720
      Width           =   915
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   2
      Top             =   6720
      Width           =   765
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Plantilla"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1005
   End
End
Attribute VB_Name = "frmCntX_DiferidosPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type xUltimos
    Tipo         As String
    Unidad       As String
    UnidadDesc   As String
    Divisa       As String
    DivisaDesc   As String
    CC           As String
    CCDesc        As String
End Type
Dim vEdita As Boolean, vBusca As Integer, vUltimos As xUltimos
Dim vScroll As Boolean


Private Sub sbLimpiezaParcial(iCodigo As Integer)
vGrid.MaxRows = 0
vGrid.MaxRows = 1

Select Case iCodigo
  Case 1 'Cambia el Tipo de Asiento
    txtDAsiento = ""
   
End Select

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodPlantilla.Text) Then
    txtCodPlantilla.Text = "0"
End If
       
If vScroll Then
    strSQL = "select Top 1 cod_diferido from CntX_Diferidos" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_diferido > " & txtCodPlantilla.Text & " order by cod_diferido asc"
    Else
       strSQL = strSQL & " and cod_diferido < " & txtCodPlantilla.Text & " order by cod_diferido desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsultaPlantilla(rs!cod_diferido)
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

Private Sub Form_Load()

vGrid.AppearanceStyle = fxGridStyle

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
Call sbToolBarIconos(tlb)

vEdita = False
Call sbToolBar(tlb, "activo")
Call sbLimpiaPantalla

If gCntX_Arbol.ArbolActivo Then
  Call sbConsultaPlantilla(Val(gCntX_Arbol.AsientoNumr))
End If
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub

Private Function fxVerificaAsiento() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, lng As Long

'Verificar Periodo
'Tipo de Asiento
'CntX_Cuentas (En el Detalle)

fxVerificaAsiento = True
vMensaje = ""

strSQL = "select isnull(count(*),0) as existe from CntX_Tipos_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & txtCAsiento & "'"
Call OpenRecordSet(rsX, strSQL, 0)
  If rsX!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de Asiento Indicano no existe..."
rsX.Close


If CCur(txtDiferencia) <> 0 Then vMensaje = vMensaje & vbCrLf & "- El Asiento No se encuentra Balanceado..."

If CCur(txtDebito) <> 100 Then vMensaje = vMensaje & vbCrLf & "- Los débitos no estan al 100%"
If CCur(txtCredito) <> 100 Then vMensaje = vMensaje & vbCrLf & "- Los créditos no estan al 100%"


For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 1
 If vGrid.Text <> "" Then
   vGrid.Col = 5
   If vGrid.Text = "" Then
      vGrid.Col = 1
      vMensaje = vMensaje & vbCrLf & "- Cuenta " & vGrid.Text & " No Existe"
   End If
 End If
Next lng

If Len(vMensaje) > 0 Then
  fxVerificaAsiento = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbLimpiaPantalla()
vBusca = 1

txtCodPlantilla = ""
txtDesPlantilla = ""

txtCAsiento = ""
txtDAsiento = ""

txtCredito = 0
txtDebito = 0
txtDiferencia = 0

txtNotas = ""

txtConsecutivo = ""

chkAsientoXPlantilla.Value = vbUnchecked


cboTipo.Clear
cboTipo.AddItem "Ingresos Diferidos (Plantilla)"
cboTipo.AddItem "Gastos Diferidos (Plantilla)"
cboTipo.Text = "Gastos Diferidos (Plantilla)"

vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 7

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(tlb, "edicion")
    
      txtDesPlantilla.SetFocus
    
    Case "MODIFICAR", "EDITAR"
        vEdita = True
        txtDesPlantilla.SetFocus
        Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
        Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case "DESHACER"
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
    
    Case "CONSULTAR"
       Select Case vBusca
         Case 1, 2 'Tipo de ASiento
            If vBusca = 1 Then
                gBusquedas.Columna = "Tipo_Asiento"
                gBusquedas.Orden = "Tipo_Asiento"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
            frmBusquedas.Show vbModal
            txtCAsiento = gBusquedas.Resultado
            txtDAsiento = gBusquedas.Resultado2
            
         Case 3, 4 'Codigo o Descripcion  de Plantilla
            If vBusca = 3 Then
                gBusquedas.Columna = "cod_diferido"
                gBusquedas.Orden = "cod_diferido"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            gBusquedas.Consulta = "select cod_diferido,descripcion from CntX_Diferidos"
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            frmBusquedas.Show vbModal
            txtCodPlantilla = gBusquedas.Resultado
            txtDesPlantilla = gBusquedas.Resultado2
            txtCodPlantilla.SetFocus
       
       End Select

    Case "REPORTES"
      
'      strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
'             & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & txtCAsiento & "' AND " _
'             & " {Cntx_Asientos.NUM_ASIENTO} = '" & txtNAsiento & "'"
'
'      Call sbCntX_Reportes("ASIENTO", strSQL)
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    
    Case "CERRAR"
      UnLoad Me
End Select

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
      Case 1
        vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs!cod_cuenta))
      Case 2
        vGrid.TextTip = TextTipFixed
        vGrid.CellNote = rs!UniDes
        vGrid.Text = CStr(rs!cod_unidad)
      Case 3
        vGrid.Text = CStr(rs!cod_centro_costo)
      Case 4
        vGrid.Text = CStr(rs!cod_Divisa)
      Case 5
        vGrid.Text = CStr(rs!Descripcion)
      Case 6
        vGrid.Text = CStr(rs!porc_debito)
      Case 7
        vGrid.Text = CStr(rs!porc_credito)
    End Select
 
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub sbConsultaPlantilla(pPlantilla As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from CntX_Diferidos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_diferido = " & pPlantilla
       
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  'llenar datos en pantalla
  
  txtCAsiento.Text = rs!tipo_asiento
  txtDAsiento.Text = fxCntX_TiposAsientos("D", rs!tipo_asiento)
  
  txtCodPlantilla = rs!cod_diferido
  txtDesPlantilla = rs!Descripcion & ""
  
  txtNotas = rs!observacion
  
  txtConsecutivo = rs!Consecutivo
  
  Select Case UCase(rs!Tipo)
    Case "I"
        cboTipo.Text = "Ingresos Diferidos (Plantilla)"
    Case "G"
        cboTipo.Text = "Gastos Diferidos (Plantilla)"
  End Select
 
  sBar.Panels(1) = rs!user_crea
  sBar.Panels(2) = rs!fecha_crea
  
strSQL = "select A.cod_cuenta,B.descripcion,porc_debito,porc_credito,linea" _
       & ",A.cod_unidad,U.descripcion as UniDes,A.cod_divisa,A.cod_centro_costo" _
       & " from CntX_Diferidos_detalle A inner join CntX_Cuentas B on A.cod_cuenta = B.cod_cuenta and A.cod_contabilidad = B.cod_contabilidad" _
       & " inner join CntX_Unidades U on A.cod_unidad = U.cod_unidad and A.cod_contabilidad = U.cod_contabilidad" _
       & " left join CntX_Centro_Costos C on A.cod_centro_costo = C.cod_centro_costo and A.cod_contabilidad = C.cod_contabilidad" _
       & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and A.cod_diferido = " & rs!cod_diferido _
       & " order by linea"
   
  Call sbCargaGridLocal(vGrid, 7, strSQL)
 
  Call sbSumaDebitosCreditos

End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset, lng As Long

On Error GoTo vError

If fxVerificaAsiento Then
    
    If vEdita Then
      
      strSQL = "update CntX_Diferidos set descripcion = '" & UCase(txtDesPlantilla) _
             & "',tipo_asiento = '" & txtCAsiento _
             & "',observacion = '" & Trim(txtNotas) _
             & "',tipo = '" & Mid(cboTipo.Text, 1, 1) _
             & "', ASIENTO_RESUMEN = " & chkAsientoXPlantilla.Value & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and cod_diferido = " & txtCodPlantilla
      Call ConectionExecute(strSQL, 0)
     
    
      Call Bitacora("Modifica", "Plantilla Diferido : " & txtCodPlantilla & " Conta." & gCntX_Parametros.CodigoConta)
    
    
    Else 'Inserta
      
      'Saca Consecutivo de Plantilla
       strSQL = "select isnull(max(cod_diferido),0) as Ultimo from CntX_Diferidos" _
              & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
       Call OpenRecordSet(rs, strSQL, 0)
         
       strSQL = "insert into CntX_Diferidos(cod_diferido,tipo_asiento,cod_contabilidad," _
              & "descripcion,consecutivo,observacion,user_crea,fecha_crea,tipo,ASIENTO_RESUMEN) values(" _
              & (rs!ultimo + 1) & ",'" & UCase(txtCAsiento) & "'," & gCntX_Parametros.CodigoConta _
              & ",'" & UCase(txtDesPlantilla) & "',0,'" & txtNotas & "','" & glogon.Usuario _
              & "',getdate(),'" & Mid(cboTipo.Text, 1, 1) & "'," & chkAsientoXPlantilla.Value & ")"
       Call ConectionExecute(strSQL, 0)
            
       txtCodPlantilla = (rs!ultimo + 1)
           
       rs.Close
       
        Call Bitacora("Registra", "Plantilla Diferido : " & txtCodPlantilla & " Conta." & gCntX_Parametros.CodigoConta)
        
    End If 'Si Inserta o Actualiza

      'Actualiza el Detalle
      strSQL = "delete CntX_Diferidos_detalle where cod_contabilidad = " _
             & gCntX_Parametros.CodigoConta & " and cod_diferido = " & txtCodPlantilla
    
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.Col = 1
        If vGrid.Text <> "" Then
            strSQL = strSQL & Space(10) & "insert into CntX_Diferidos_detalle(cod_diferido,cod_contabilidad" _
                   & ",linea,cod_cuenta,cod_unidad,cod_centro_costo,cod_divisa,porc_debito,porc_credito,tipo_asiento" _
                   & ") values(" & txtCodPlantilla & "," & gCntX_Parametros.CodigoConta & "," & lng & ",'"
            vGrid.Row = lng
            vGrid.Col = 1
            strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
            vGrid.Col = 2
            strSQL = strSQL & vGrid.Text & "','"
            vGrid.Col = 3
            strSQL = strSQL & vGrid.Text & "','"
            vGrid.Col = 4
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.Col = 6
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.Col = 7
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ",'" & txtCAsiento & "')"
         End If 'vgrid.Text <> ""
       
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL, 0)
            strSQL = ""
        End If
       
       Next lng
       
        'Procesa Ultimo Lote
        If Len(strSQL) > 0 Then
            Call ConectionExecute(strSQL, 0)
            strSQL = ""
        End If

        Call sbToolBar(tlb, "activo")
        Call sbConsultaPlantilla(txtCodPlantilla)
        
        vEdita = True
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation


End If 'Verificacion del Asiento

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
  strSQL = "delete CntX_Diferidos_detalle where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_diferido = " & txtCodPlantilla
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete CntX_Diferidos where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_diferido = " & txtCodPlantilla
  Call ConectionExecute(strSQL, 0)
  

  Call Bitacora("Elimina", "Plantilla Diferidos : " & txtCodPlantilla & " Conta." _
                  & gCntX_Parametros.CodigoConta)

  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCAsiento_Change()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close

End Sub

Private Sub txtCAsiento_GotFocus()
vBusca = 1
End Sub

Private Sub txtCAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtCAsiento_LostFocus()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close

End Sub

Private Sub txtCodPlantilla_GotFocus()
vBusca = 3
End Sub

Private Sub txtDAsiento_GotFocus()
vBusca = 2
End Sub

Private Sub txtDAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtDescripcion_GotFocus()
vBusca = 4
End Sub

Private Sub txtDesPlantilla_GotFocus()
vBusca = 4
End Sub

Private Sub txtDesPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCAsiento.SetFocus
End Sub

Private Sub txtCodPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 Call sbConsultaPlantilla(txtCodPlantilla)
 txtDesPlantilla.SetFocus
End If
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
Exit Sub
vError:
  Call sbLimpiaPantalla
End Sub


Private Function fxVerificaCuenta(pCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select isnull(count(*),0) as Existe from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & pCuenta & "' and acepta_movimientos = 1"

Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function

Private Sub sbSumaDebitosCreditos()
Dim x As Long, curValor As Currency

  txtDebito = 0
  txtCredito = 0
   
  For x = 1 To vGrid.MaxRows
    vGrid.Row = x
    vGrid.Col = 1
    If vGrid.Text <> "" Then
      vGrid.Col = 6
      txtDebito = CCur(txtDebito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
      vGrid.Col = 7
      txtCredito = CCur(txtCredito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
    End If 'vGrid.text <> ""
      
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(7) As Variant, x As Integer
Dim vTempo As String

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = vGrid.MaxCols
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  
  
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.Col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
  
End If

'Consulta Cuentas
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

'Consulta Unidad
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 2 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
  gBusquedas.Consulta = "select cod_unidad,descripcion from CntX_Unidades"
  frmBusquedas.Show vbModal
    
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.TextTip = TextTipFixed
  vGrid.CellNote = gBusquedas.Resultado2
  
End If


'Consulta Centro de Costo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 2
  vTempo = vGrid.Text
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_centro_costo in(select cod_centro_costo" _
                    & " from cntX_unidades_cc where cod_unidad = '" & vTempo & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
  gBusquedas.Consulta = "select cod_centro_costo,descripcion from CntX_Centro_Costos"
  frmBusquedas.Show vbModal
    
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.CellTag = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  vGrid.CellNote = vGrid.CellTag
  
End If


'Consulta Divisa
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
  gBusquedas.Consulta = "select cod_divisa,descripcion from CntX_Divisas"
  frmBusquedas.Show vbModal
    
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.CellTag = gBusquedas.Resultado2
End If


If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text)
        i = fxCntX_CuentaFormato(False, vGrid.Text)
        If fxVerificaCuenta(CStr(i)) Then
          vGrid.Col = 5
          vGrid.Text = fxCntX_Cuenta("D", CStr(i))
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE Cuentas", vbCritical
        End If
        
      Case 2 'Verificar Unidad
        If fxCntx_UnidadVerifica(vGrid.Text) Then
          vGrid.TextTip = TextTipFixed
          vGrid.CellNote = fxCntX_Unidad("D", vGrid.Text)
          vUltimos.Unidad = vGrid.Text
          vUltimos.UnidadDesc = vGrid.CellNote
        Else
          MsgBox "La cod_unidad de negocio no es válida : " & vbCrLf & " - No Existe...", vbCritical
        End If
        
      Case 3 'Verificar el Centro de Costo
        vGrid.Col = 2
        vTempo = vGrid.Text
        vGrid.Col = 3
        
        If fxCntX_CentroCostoVerifica(vGrid.Text, vTempo) Then
          vGrid.TextTip = TextTipFixed
          vGrid.CellTag = fxCntX_CentroCosto("D", vGrid.Text)
          vGrid.CellNote = vGrid.CellTag
          
          vUltimos.CC = vGrid.Text
          vUltimos.CCDesc = vGrid.CellTag
        Else
          MsgBox "El Centro de Costo no es válido y no puede ser utilizada por esta unidad: " & vbCrLf & " - No Existe...", vbCritical
        End If
        
      Case 4 'Divisa
      
        If fxCntX_DivisaVerifica(vGrid.Text) Then
          vGrid.CellTag = fxCntX_Divisas("D", vGrid.Text)
          vUltimos.Divisa = vGrid.Text
          vUltimos.DivisaDesc = vGrid.CellTag
          'Tipo de Cambio
        Else
          MsgBox "La DIVISA no es válida : " & vbCrLf & " - No Existe...", vbCritical
        End If
        
      Case 6 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 7 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
        
        End If
      
        If vGrid.MaxRows = vGrid.Row Then
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Col = 2
            vGrid.Text = vUltimos.Unidad
            vGrid.TextTip = TextTipFixed
            vGrid.CellNote = vUltimos.UnidadDesc
        End If
    
    End Select
End If


If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 2
    vGrid.Text = vUltimos.Unidad
    vGrid.TextTip = TextTipFixed
    vGrid.CellNote = vUltimos.UnidadDesc
    
    vGrid.Col = 3
    vGrid.Text = vUltimos.CC
    vGrid.TextTip = TextTipFixed
    vGrid.CellTag = vUltimos.CCDesc
    vGrid.CellNote = vGrid.CellTag
    
    vGrid.Col = 4
    vGrid.Text = vUltimos.Divisa
    vGrid.CellTag = vUltimos.DivisaDesc
End If


End Sub


