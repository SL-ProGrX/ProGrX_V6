VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmFNDOperadoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operadoras"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "frmFNDOperadoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "frmFNDOperadoras.frx":030A
   ScaleHeight     =   6240
   ScaleWidth      =   8925
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5175
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   8655
      _Version        =   1441792
      _ExtentX        =   15266
      _ExtentY        =   9128
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "General"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "chkActiva"
      Item(0).Control(1)=   "txtNotas"
      Item(0).Control(2)=   "Label5(1)"
      Item(0).Control(3)=   "gbCuentas"
      Item(0).Control(4)=   "txtMultaTope"
      Item(0).Control(5)=   "Label5(3)"
      Item(1).Caption =   "Planes vinculados"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4815
         Left            =   -70000
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   8655
         _Version        =   1441792
         _ExtentX        =   15266
         _ExtentY        =   8493
         _StockProps     =   77
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkActiva 
         Height          =   252
         Left            =   6840
         TabIndex        =   6
         Top             =   360
         Width           =   1212
         _Version        =   1441792
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activa?   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1032
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   6732
         _Version        =   1441792
         _ExtentX        =   11874
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
      Begin XtremeSuiteControls.GroupBox gbCuentas 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   8415
         _Version        =   1441792
         _ExtentX        =   14838
         _ExtentY        =   4466
         _StockProps     =   79
         Caption         =   "Cuentas por Omisión"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCuentaFondo 
            Height          =   312
            Left            =   1200
            TabIndex        =   10
            Top             =   720
            Width           =   1812
            _Version        =   1441792
            _ExtentX        =   3196
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaFondoDesc 
            Height          =   312
            Left            =   3000
            TabIndex        =   11
            Top             =   720
            Width           =   5292
            _Version        =   1441792
            _ExtentX        =   9334
            _ExtentY        =   550
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaIngreso 
            Height          =   312
            Left            =   1200
            TabIndex        =   12
            Top             =   1440
            Width           =   1812
            _Version        =   1441792
            _ExtentX        =   3196
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaIngresoDesc 
            Height          =   312
            Left            =   3000
            TabIndex        =   13
            Top             =   1440
            Width           =   5292
            _Version        =   1441792
            _ExtentX        =   9334
            _ExtentY        =   550
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaRetiros 
            Height          =   312
            Left            =   1200
            TabIndex        =   14
            Top             =   2160
            Width           =   1812
            _Version        =   1441792
            _ExtentX        =   3196
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaRetirosDesc 
            Height          =   312
            Left            =   3000
            TabIndex        =   15
            Top             =   2160
            Width           =   5292
            _Version        =   1441792
            _ExtentX        =   9334
            _ExtentY        =   550
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de Ingresos por Retiros anticipados o multas"
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
            Left            =   240
            TabIndex        =   18
            Top             =   1200
            Width           =   5412
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta Transito para Retiros"
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
            Left            =   240
            TabIndex        =   17
            Top             =   1920
            Width           =   3132
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta por Omisión de una Plan no configurado"
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
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   4212
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtMultaTope 
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Top             =   2040
         Width           =   1815
         _Version        =   1441792
         _ExtentX        =   3201
         _ExtentY        =   556
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tope para Multas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   852
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8280
      TabIndex        =   1
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   1212
      _Version        =   1441792
      _ExtentX        =   2138
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   5532
      _Version        =   1441792
      _ExtentX        =   9758
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
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
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "frmFNDOperadoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean


Private Sub sbPlanes_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim pContrato As Long, pTotal As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass
   
lsw.ListItems.Clear

pContrato = 0
pTotal = 0

strSQL = "select * from vFnd_Operadoras_Rsm where cod_Operadora = " & vCodigo
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_Plan)
     itmX.SubItems(1) = rs!Plan_Desc
     itmX.SubItems(2) = rs!cod_Divisa
     itmX.SubItems(3) = Format(rs!Contratos, "###,###,##0")
     itmX.SubItems(4) = Format(rs!Total * fxSys_Tipo_Cambio_Apl(rs!Tipo_Cambio), "Standard")
     itmX.SubItems(5) = Format(rs!Total, "Standard")
     
     pContrato = pContrato + rs!Contratos
     pTotal = pTotal + rs!Total * fxSys_Tipo_Cambio_Apl(rs!Tipo_Cambio)
 rs.MoveNext
Loop
rs.Close

 Set itmX = lsw.ListItems.Add(, , "")
     itmX.SubItems(1) = ""
     itmX.SubItems(2) = "Totales:"
     itmX.SubItems(3) = Format(pContrato, "###,###,##0")
     itmX.SubItems(4) = Format(pTotal, "Standard")
    
     itmX.Bold = True
     
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbConsulta(pCodigo As Long)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vFnd_Operadoras where cod_Operadora = " & pCodigo
Call OpenRecordSet(rs, strSQL)

Call sbLimpia

If Not rs.BOF And Not rs.EOF Then
   Call sbToolBar(Me.tlb, "activo")
   vEdita = True
   vCodigo = rs!cod_Operadora
   
   txtCodigo = rs!cod_Operadora
   txtDescripcion = Trim(rs!Descripcion)
   
   chkActiva.Value = rs!Activa
   txtNotas.Text = rs!Notas & ""
   
   txtCuentaFondo.Text = rs!CtaPlan
   txtCuentaFondoDesc.Text = rs!CtaPlanDesc
   
   txtCuentaIngreso.Text = rs!CtaIng
   txtCuentaIngresoDesc.Text = rs!CtaIngDesc
   
   txtCuentaRetiros.Text = rs!CtaRet
   txtCuentaRetirosDesc.Text = rs!CtaRetDesc
   
   txtMultaTope.Text = Format(rs!MULTA_MNT_TOPE, "Standard")
   
Else
    txtDescripcion.SetFocus
    Call sbToolBar(Me.tlb, "nuevo")
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbLimpia()
vCodigo = 0

tcMain.Item(0).Selected = True

chkActiva.Value = xtpChecked

txtCodigo.Text = ""
txtDescripcion.Text = ""


txtNotas.Text = ""
txtCuentaFondo.Text = ""
txtCuentaFondoDesc.Text = ""
txtCuentaIngreso.Text = ""
txtCuentaIngresoDesc.Text = ""
txtCuentaRetiros.Text = ""
txtCuentaRetirosDesc.Text = ""


txtMultaTope.Text = "999999999.00"

End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
End If

If vScroll Then
    strSQL = "select Top 1 cod_Operadora from fnd_Operadoras"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " Where cod_Operadora > " & txtCodigo.Text & " order by cod_Operadora asc"
    Else
       strSQL = strSQL & " Where cod_Operadora < " & txtCodigo.Text & " order by cod_Operadora desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_Operadora)
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
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()

vModulo = 18 'Fondo de Inversion

vEdita = True
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Plan", 1200
    .Add , , "Descripción", 3600
    .Add , , "Divisa", 1000, vbCenter
    .Add , , "Contratos:", 1200, vbRightJustify
    .Add , , "Total:", 1400, vbRightJustify
    .Add , , "Importe Real:", 1400, vbRightJustify
End With
 
vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

 
 
Call sbToolBarIconos(Me.tlb)
Call sbToolBar(Me.tlb, "nuevo")
Call sbLimpia

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxValida() As Boolean
Dim pMensaje As String, pReturn As Boolean

pMensaje = ""
pReturn = False

If Len(txtDescripcion.Text) = 0 Then
    pMensaje = pMensaje & " - La descripción de la Operadora no es válida!" & vbCrLf
End If

If Not fxCntX_CuentaValida(txtCuentaFondo.Text) Then
    pMensaje = pMensaje & " - La Cuenta Default para Fondos no es válida!" & vbCrLf
End If

If Not fxCntX_CuentaValida(txtCuentaIngreso.Text) Then
    pMensaje = pMensaje & " - La Cuenta Default para Ingresos no es válida!" & vbCrLf
End If

If Not fxCntX_CuentaValida(txtCuentaRetiros.Text) Then
    pMensaje = pMensaje & " - La Cuenta Default para Retiros no es válida!" & vbCrLf
End If

If Not IsNumeric(txtMultaTope.Text) Then
    pMensaje = pMensaje & " - El Monto para Tope de las Multas por Retiros no es valida!" & vbCrLf
End If



If Len(pMensaje) > 0 Then
  pReturn = False
  MsgBox pMensaje, vbExclamation
Else
   pReturn = True
End If

fxValida = pReturn

End Function


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset


If Not fxValida Then
    Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

If vEdita Then
  strSQL = "Update FND_Operadoras Set Descripcion='" & Trim(txtDescripcion) & "'," _
          & "Cta_Fondo='" & fxCntX_CuentaFormato(False, txtCuentaFondo.Text, 0) & "'," _
          & "Cta_Retiros='" & fxCntX_CuentaFormato(False, txtCuentaRetiros.Text, 0) & "'," _
          & "Cta_Ingresos='" & fxCntX_CuentaFormato(False, txtCuentaIngreso.Text, 0) & "'," _
          & "Notas = '" & txtNotas.Text & "', Activa = " & chkActiva.Value & ", MULTA_MNT_TOPE = " & CCur(txtMultaTope.Text) _
          & " Where cod_Operadora = " & vCodigo
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Datos Operadora:" & vCodigo & "-" & UCase(Trim(txtDescripcion)))

Else
   strSQL = "insert FND_Operadoras(Descripcion,Activa, Notas,Cta_Fondo,Cta_Retiros,cta_ingresos,MULTA_MNT_TOPE)" _
          & " values('" & Trim(txtDescripcion) & "'," & chkActiva.Value & ",'" & txtNotas.Text _
          & "','" & fxCntX_CuentaFormato(False, txtCuentaFondo.Text, 0) _
          & "','" & fxCntX_CuentaFormato(False, txtCuentaRetiros.Text, 0) _
          & "','" & fxCntX_CuentaFormato(False, txtCuentaIngreso.Text, 0) & "'," & CCur(txtMultaTope.Text) & ")"
   Call ConectionExecute(strSQL)
      
   strSQL = "Select max(Cod_Operadora) as Codigo from Fnd_Operadoras"
   Call OpenRecordSet(rs, strSQL)
   txtCodigo.Text = rs!Codigo
   rs.Close
   
   Call Bitacora("Registra", "Operadora:" & Trim(txtCodigo) & "-" & UCase(Trim(txtDescripcion)))
   
End If

vCodigo = Trim(txtCodigo)
vEdita = True
Call sbToolBar(Me.tlb, "activo")

Me.MousePointer = vbDefault

MsgBox "Información guardada satisfactoriamente...", vbInformation
txtDescripcion.SetFocus

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError
i = MsgBox("Esta Seguro que desea borrar esta operadora", vbYesNo)

If i = vbYes Then
  strSQL = "delete Fnd_Operadoras where cod_operadora = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Borra", "Operadora: " & vCodigo & "-" & UCase(Trim(txtDescripcion)))
  
  Call sbLimpia
  Call sbToolBar(Me.tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
    Call sbPlanes_Load
End If
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbToolBar(Me.tlb, "edicion")
      Call sbLimpia
      txtDescripcion.SetFocus
            
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      Call sbToolBar(Me.tlb, "edicion")
      txtDescripcion.SetFocus
      
    Case "BORRAR"
      Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
      
    Case "DESHACER"
      Call sbToolBar(Me.tlb, "nuevo")
      Call sbLimpia
      
      txtCodigo.SetFocus
      vEdita = True
    
    Case "CONSULTAR"
      Call sbConsulta(vCodigo)
      
    Case "REPORTES"
      frmContenedor.Crt.Connect = glogon.ConectRPT
      frmContenedor.Crt.ReportFileName = SIFGlobal.fxPathReportes("Fondos_Operadoras.rpt")
      frmContenedor.Crt.Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
      frmContenedor.Crt.Formulas(1) = "Usuario='" & Trim(glogon.Usuario) & "'"
      frmContenedor.Crt.Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
      frmContenedor.Crt.PrintReport
       
    Case "CERRAR"
       Unload Me
End Select

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "cod_operadora"
  gBusquedas.Orden = "cod_operadora"

  gBusquedas.Filtro = ""
  gBusquedas.Consulta = "select cod_operadora,descripcion from fnd_Operadoras"
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()

If Trim(txtCodigo) <> "" And txtCodigo.Locked = False Then
   Call sbConsulta(txtCodigo)
Else
  If vEdita = True And txtCodigo.Locked = False Then
   Call sbToolBar(Me.tlb, "nuevo")
   Call sbLimpia
  End If
End If

End Sub


Private Sub txtCuentaRetiros_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaRetirosDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaRetiros.Text = gCuenta
   txtCuentaRetirosDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaRetiros.Text = fxgCntCuentaFormato(True, txtCuentaRetiros, 0)
End If

End Sub

Private Sub txtCuentaRetiros_LostFocus()
   txtCuentaRetirosDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaRetiros, 0))
   txtCuentaRetiros.Text = fxgCntCuentaFormato(True, txtCuentaRetiros, 0)
End Sub

Private Sub txtCuentaRetirosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub



Private Sub txtCuentaFondo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaFondoDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaFondo.Text = gCuenta
   txtCuentaFondoDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaFondo.Text = fxgCntCuentaFormato(True, txtCuentaFondo, 0)
End If

End Sub

Private Sub txtCuentaFondo_LostFocus()
   txtCuentaFondoDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaFondo, 0))
   txtCuentaFondo.Text = fxgCntCuentaFormato(True, txtCuentaFondo, 0)
End Sub

Private Sub txtCuentaFondoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIngreso.SetFocus
End Sub

Private Sub txtCuentaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIngresoDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaIngreso.Text = gCuenta
   txtCuentaIngresoDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaIngreso.Text = fxgCntCuentaFormato(True, txtCuentaIngreso, 0)
End If

End Sub

Private Sub txtCuentaIngreso_LostFocus()
   txtCuentaIngresoDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaIngreso, 0))
   txtCuentaIngreso.Text = fxgCntCuentaFormato(True, txtCuentaIngreso, 0)
End Sub

Private Sub txtCuentaIngresoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaRetiros.SetFocus
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    tcMain.Item(0).Selected = True
    txtNotas.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"

  gBusquedas.Filtro = ""
  gBusquedas.Consulta = "select cod_operadora,descripcion from fnd_Operadoras"
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
  
  Call sbConsulta(txtCodigo.Text)
End If

End Sub


Private Sub txtMultaTope_GotFocus()
On Error GoTo vError

txtMultaTope.Text = CCur(txtMultaTope.Text)

vError:

End Sub

Private Sub txtMultaTope_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaFondo.SetFocus
End Sub

Private Sub txtMultaTope_LostFocus()
On Error GoTo vError

txtMultaTope.Text = Format(CCur(txtMultaTope.Text), "Standard")

vError:
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMultaTope.SetFocus
End Sub
