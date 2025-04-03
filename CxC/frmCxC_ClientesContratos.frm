VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCxC_ClientesContratos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Suscripción de Contratos"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   9255
      _Version        =   1310723
      _ExtentX        =   16325
      _ExtentY        =   6588
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
      ItemCount       =   3
      Item(0).Caption =   "General"
      Item(0).ControlCount=   17
      Item(0).Control(0)=   "chkActivo"
      Item(0).Control(1)=   "cboContratoTipo"
      Item(0).Control(2)=   "dtpContratoVence"
      Item(0).Control(3)=   "txtNotas"
      Item(0).Control(4)=   "Label1(0)"
      Item(0).Control(5)=   "Label1(4)"
      Item(0).Control(6)=   "Label1(5)"
      Item(0).Control(7)=   "Label1(3)"
      Item(0).Control(8)=   "Label1(2)"
      Item(0).Control(9)=   "Label1(6)"
      Item(0).Control(10)=   "Label1(7)"
      Item(0).Control(11)=   "Label1(8)"
      Item(0).Control(12)=   "Label1(9)"
      Item(0).Control(13)=   "txtContratoNum"
      Item(0).Control(14)=   "txtPlazo"
      Item(0).Control(15)=   "txtTasaCorriente"
      Item(0).Control(16)=   "txtTasaMora"
      Item(1).Caption =   "Pagadores Autorizados"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswP"
      Item(2).Caption =   "Cargos de Suscripción"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswC"
      Begin XtremeSuiteControls.ListView lswC 
         Height          =   3135
         Left            =   -69880
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1310723
         _ExtentX        =   15901
         _ExtentY        =   5530
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lswP 
         Height          =   3135
         Left            =   -69880
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1310723
         _ExtentX        =   15901
         _ExtentY        =   5530
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   600
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activo? "
         BackColor       =   -2147483633
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboContratoTipo 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4473924
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
      Begin XtremeSuiteControls.DateTimePicker dtpContratoVence 
         Height          =   315
         Left            =   4920
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   795
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Width           =   7335
         _Version        =   1310723
         _ExtentX        =   12933
         _ExtentY        =   1397
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
      Begin XtremeSuiteControls.FlatEdit txtContratoNum 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   2400
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasaCorriente 
         Height          =   315
         Left            =   2760
         TabIndex        =   10
         Top             =   2760
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasaMora 
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         Top             =   3120
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "(Tasa Anual para Días de Atraso de la cancelación)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   20
         Top             =   3120
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "(Tasa Anual para uso en sustitución de facturas)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   19
         Top             =   2760
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa Mora"
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
         Index           =   5
         Left            =   1200
         TabIndex        =   18
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo (Dias)"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa Corriente"
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
         Index           =   2
         Left            =   1200
         TabIndex        =   16
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
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
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Contrato"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
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
         Index           =   9
         Left            =   3480
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   5175
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario de Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario que Actualiza"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha de Actualización"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8640
      TabIndex        =   22
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   23
      Top             =   960
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   3000
      TabIndex        =   24
      Top             =   960
      Width           =   5535
      _Version        =   1310723
      _ExtentX        =   9763
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1200
      TabIndex        =   25
      Top             =   600
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3000
      TabIndex        =   26
      Top             =   600
      Width           =   5535
      _Version        =   1310723
      _ExtentX        =   9763
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblContrato 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
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
      Left            =   120
      TabIndex        =   28
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblCedulla 
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula"
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
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmCxC_ClientesContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean, vPaso As Boolean
Dim vTasaCor As Currency, vTasaMor As Currency, vPlazo As Long
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbLimpia()

tcMain.Item(0).Selected = True
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False

txtDescripcion.Text = ""
txtNotas.Text = ""
txtContratoNum.Text = ""

chkActivo.Value = vbChecked

txtPlazo.Text = vPlazo
txtTasaCorriente.Text = Format(vTasaCor, "Standard")
txtTasaMora.Text = Format(vTasaMor, "Standard")

StatusBarX.Panels.Item(1).Text = ""
StatusBarX.Panels.Item(2).Text = ""
StatusBarX.Panels.Item(3).Text = ""
StatusBarX.Panels.Item(4).Text = ""


End Sub


Private Sub cboContratoTipo_Change()
If Mid(cboContratoTipo.Text, 1, 1) = "I" Then
   dtpContratoVence.Enabled = False
Else
   dtpContratoVence.Enabled = True
End If
End Sub

Private Sub FlatScrollBar_Change()


On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_contrato from CxC_Personas_Contratos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_contrato > '" & txtCodigo.Text & "' and cedula = '" & txtCedula.Text & "' order by cod_contrato asc"
    Else
       strSQL = strSQL & " where cod_contrato < '" & txtCodigo.Text & "' and cedula = '" & txtCedula.Text & "' order by cod_contrato desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_CONTRATO
      Call sbConsulta(txtCodigo.Text)
    End If
End If

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()

vModulo = 31

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
 
txtCedula.Text = GLOBALES.gTag
txtNombre.Text = GLOBALES.gTag2

vTasaCor = 0
vTasaMor = 0
vPlazo = 0

cboContratoTipo.AddItem "Definido"
cboContratoTipo.AddItem "Indefinido"
cboContratoTipo.Text = "Definido"

vEdita = True

Call sbToolBarIconos(Me.tlb)
Call sbToolBar(Me.tlb, "nuevo")

With lswP.ColumnHeaders
    .Clear
    .Add , , "Id Pagador", 2000
    .Add , , "Nombre", 3500
    .Add , , "Registro", 3500, vbCenter
End With


With lswC.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Cargo", 3000
    .Add , , "Tipo", 1000, vbCenter
    .Add , , "Frecuencia", 1000, vbCenter
    .Add , , "Valor", 1400, vbRightJustify
    .Add , , "Frec.Días", 1000, vbCenter
    .Add , , "Recaudado", 2100, vbRightJustify
    .Add , , "Pago Ult.", 1400, vbCenter
    .Add , , "Pago Prox.", 1400, vbCenter
    .Add , , "Modifica?", 1400, vbCenter
    
End With



Call sbLimpia

Call Formularios(Me)
Call RefrescaTags(Me)

'Activa Replica de Seguridad a Componentes Asociados la Accion de Editar:
lswC.Enabled = tlb.Buttons.Item(1).Enabled
lswP.Enabled = tlb.Buttons.Item(1).Enabled


End Sub


Private Sub sbConsulta(pCodigo As String)


On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

strSQL = "select P.Descripcion,C.* " _
       & " from CxC_Contratos P inner join CxC_Personas_Contratos C on P.cod_contrato = C.cod_contrato" _
       & " where C.cedula = '" & txtCedula.Text & "' and C.cod_contrato = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(Me.tlb, "activo")
  vEdita = True
  vCodigo = Trim(rs!COD_CONTRATO)
  txtCodigo.Text = rs!COD_CONTRATO
  txtDescripcion.Text = rs!Descripcion
  txtNotas.Text = rs!Notas
  txtContratoNum.Text = rs!Contrato_Num
  dtpContratoVence.Value = rs!Contrato_Vence
  
  Select Case rs!Contrato_Tipo
    Case "I"
      cboContratoTipo.Text = "Indefinido"
    Case "D"
      cboContratoTipo.Text = "Definido"
  End Select
  
  
  txtPlazo.Text = rs!Plazo
  txtTasaCorriente.Text = Format(rs!Tasa_Corriente, "Standard")
  txtTasaMora.Text = Format(rs!Tasa_Mora, "Standard")

  StatusBarX.Panels.Item(1).Text = rs!Registro_Usuario
  StatusBarX.Panels.Item(2).Text = rs!Registro_Fecha
  StatusBarX.Panels.Item(3).Text = rs!Actualiza_usuario & ""
  StatusBarX.Panels.Item(4).Text = rs!Actualiza_fecha & ""
  
  
  tcMain.Item(1).Enabled = True
  tcMain.Item(2).Enabled = True
  
Else
  'Busca Datos del Cliente Unicamente
    strSQL = "select *,dbo.MyGetdate() as 'Fecha' from CxC_Contratos where cod_contrato = '" & pCodigo & "' and activo = 1"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.BOF And Not rs.EOF Then
      vTasaCor = rs!Tasa_Corriente
      vTasaMor = rs!Tasa_Mora
      vPlazo = rs!Plazo
      dtpContratoVence.Value = rs!fecha
      
      Call sbLimpia
      
      txtDescripcion.Text = rs!Descripcion

    End If


End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

 
End Sub

Private Sub sbGuardar()

On Error GoTo vError


If vEdita Then
  strSQL = "Update CxC_Personas_Contratos Set Cedula ='" & Trim(txtCedula.Text) & "', Notas = '" & Trim(txtNotas.Text) _
         & "',Activo = " & chkActivo.Value & ", Plazo = " & txtPlazo.Text & ", Tasa_Corriente = " & CCur(txtTasaCorriente.Text) _
         & ",Tasa_Mora = " & CCur(txtTasaMora.Text) & ",Actualiza_Usuario = '" & glogon.Usuario _
         & "',Actualiza_Fecha = dbo.MyGetdate(),Contrato_Num = '" & txtContratoNum.Text _
         & "', Contrato_Tipo = '" & Mid(cboContratoTipo.Text, 1, 1) & "',Contrato_Vence = '" & Format(dtpContratoVence.Value, "yyyy/mm/dd") _
         & "' where cedula = '" & txtCedula.Text & "' and cod_contrato = '" & txtCodigo.Text & "'"
  Call ConectionExecute(strSQL)
  
  
  Call Bitacora("Modifica", "Suscripción: Ced." & txtCedula.Text & " Cnt: " & Trim(txtCodigo.Text))

Else
   strSQL = "insert CxC_Personas_Contratos(cedula,cod_contrato,Notas,Activo,Plazo,Tasa_Corriente,Tasa_Mora" _
          & ",registro_usuario,registro_fecha,Contrato_Num,Contrato_Tipo,Contrato_Vence)" _
          & " values('" & txtCedula.Text & "','" & Trim(txtCodigo.Text) & "','" & Trim(txtNotas.Text) & "'," & chkActivo.Value _
          & "," & txtPlazo.Text & "," & CCur(txtTasaCorriente.Text) & "," & CCur(txtTasaMora.Text) _
          & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'" & txtContratoNum.Text _
          & "','" & Mid(cboContratoTipo.Text, 1, 1) & "','" & Format(dtpContratoVence.Value, "yyyy/mm/dd") & "')"
   Call ConectionExecute(strSQL)
  
   Call Bitacora("Registra", "Suscripción: Ced." & txtCedula.Text & " Cnt: " & Trim(txtCodigo.Text))
   
End If

vCodigo = Trim(txtCodigo.Text)
vEdita = True

Call sbToolBar(Me.tlb, "activo")

txtNombre.SetFocus

tcMain.Item(1).Enabled = True
tcMain.Item(2).Enabled = True



MsgBox "Información guardada satisfactoriamente...", vbInformation


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CxC_Personas_Contratos where cedula = '" & txtCedula.Text & "' and cod_contrato = '" & txtCodigo.Text & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Elimina", "Suscripción: Ced." & txtCedula.Text & " Cnt: " & Trim(txtCodigo.Text))
  
  Call sbLimpia
  Call sbToolBar(Me.tlb, "nuevo")
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub lswC_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vMovimiento As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If Item.Checked Then
   vMovimiento = "Registra"
   strSQL = "insert CxC_Personas_Contratos_Suscripciones(cod_contrato,cedula,cod_cargo,Tipo,valor,frecuencia_Tipo,frecuencia_dias,recaudado,pago_ultimo" _
          & ",pago_proximo,modifica,registro_Fecha,Registro_Usuario) values('" & txtCodigo.Text & "','" & txtCedula.Text & "','" & Item.Text _
          & "','" & Mid(Item.SubItems(2), 1, 1) & "'," & CCur(Item.SubItems(3)) & ",'" & Mid(Item.SubItems(4), 1, 1) & "'," & CLng(Item.SubItems(5)) _
          & "," & CCur(Item.SubItems(6)) & ",'" & Format(Item.SubItems(7), "yyyy/mm/dd") _
          & "','" & Format(Item.SubItems(8), "yyyy/mm/dd") & "'," & Item.SubItems(9) & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   vMovimiento = "Elimina"
   strSQL = "Delete CxC_Personas_Contratos_Suscripciones where cod_contrato = '" & txtCodigo.Text & "' and cedula = '" & txtCedula.Text _
          & "' and cod_cargo = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Call Bitacora(vMovimiento, "Suscripción: Ced." & txtCedula.Text & " Cnt: " & Trim(txtCodigo.Text) & " Cargo.:" & Item.Text)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswP_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vMovimiento As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If Item.Checked Then
   vMovimiento = "Registra"
   strSQL = "insert CxC_Personas_Contratos_Pagadores(cod_contrato,cedula,cedula_pagador,registro_fecha,registro_usuario) values('" _
          & txtCodigo.Text & "','" & txtCedula.Text & "','" & Item.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   vMovimiento = "Elimina"
   strSQL = "Delete CxC_Personas_Contratos_Pagadores where cod_contrato = '" & txtCodigo.Text & "' and cedula = '" & txtCedula.Text _
          & "' and cedula_pagador = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Call Bitacora(vMovimiento, "Suscripción: Ced." & txtCedula.Text & " Cnt: " & Trim(txtCodigo.Text) & " Pagador.:" & Item.Text)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

Select Case Item.Index
   Case 1 'Pagadores
        
        lswP.ListItems.Clear
        
        'Pagadores Asignados
        strSQL = "select P.nombre, C.*" _
             & " from CxC_Personas P inner join CxC_Personas_Contratos_Pagadores C on P.cedula = C.cedula_pagador" _
             & " where C.cod_contrato = '" & txtCodigo.Text & "' and C.cedula = '" & txtCedula.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            Set itmX = lswP.ListItems.Add(, , rs!cedula_pagador)
                itmX.SubItems(1) = rs!Nombre
                itmX.SubItems(2) = rs!Registro_Usuario & "..." & rs!Registro_Fecha
                itmX.Checked = True
            rs.MoveNext
        Loop
        rs.Close
        
        'Pagadores No - Asignados
        strSQL = "select P.nombre, C.*" _
             & " from CxC_Personas P inner join CxC_Contratos_Pagadores C on P.cedula = C.cedula" _
             & " where C.cod_contrato = '" & txtCodigo.Text & "' and C.cedula <> '" & txtCedula.Text & "' and C.cedula not in(select Cedula_Pagador" _
             & " from CxC_Personas_Contratos_Pagadores" _
             & " Where Cedula = '" & txtCedula.Text & "' and Cod_Contrato = '" & txtCodigo.Text & "')"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            Set itmX = lswP.ListItems.Add(, , rs!Cedula)
                itmX.SubItems(1) = rs!Nombre
            rs.MoveNext
        Loop
        rs.Close
  
  
   Case 2 'Cargos de Suscripcion
        lswC.ListItems.Clear
        'Cargos asignados
        strSQL = " select C.descripcion,S.*" _
               & " from CxC_Cargos C inner join CxC_Personas_Contratos_Suscripciones S on C.cod_cargo = S.cod_cargo" _
               & " where S.cod_contrato = '" & txtCodigo.Text & "' and S.cedula = '" & txtCedula.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            Set itmX = lswC.ListItems.Add(, , rs!COD_CARGO)
                itmX.SubItems(1) = rs!Descripcion
                
                Select Case rs!Tipo
                   Case "P"
                     itmX.SubItems(2) = "Porcentual"
                   Case "M"
                     itmX.SubItems(2) = "Monto"
                End Select
                
                
                Select Case rs!Frecuencia_Tipo
                  Case "O"
                    itmX.SubItems(4) = "Operación"
                  Case "D"
                    itmX.SubItems(4) = "Días"
                End Select
                
                itmX.SubItems(3) = Format(rs!Valor, "Standard")
                itmX.SubItems(5) = rs!Frecuencia_dias
                itmX.SubItems(6) = Format(rs!Recaudado, "Standard")
                itmX.SubItems(7) = Format(rs!Pago_Ultimo, "dd/mm/yyyy")
                itmX.SubItems(8) = Format(rs!Pago_Proximo, "dd/mm/yyyy")
                itmX.SubItems(9) = rs!Modifica
                itmX.Checked = True
            rs.MoveNext
        Loop
        rs.Close

        
        'Cargos No asignados
        strSQL = " select C.descripcion,S.*, dbo.MyGetdate() as 'Pago_Ultimo', dateadd(d,S.frecuencia_dias,dbo.MyGetdate()) as 'Pago_Proximo', 0 as 'Recaudado'" _
               & " from CxC_Cargos C inner join CxC_Contratos_Cargos S on C.cod_cargo = S.cod_cargo" _
               & " where S.cod_contrato = '" & txtCodigo.Text & "' and S.cod_cargo not in(select cod_cargo" _
               & " from CxC_Personas_Contratos_Suscripciones where cod_contrato = '" & txtCodigo.Text _
               & "' and cedula = '" & txtCedula & "')"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            Set itmX = lswC.ListItems.Add(, , rs!COD_CARGO)
                itmX.SubItems(1) = rs!Descripcion
                
                Select Case rs!Tipo
                   Case "P"
                     itmX.SubItems(2) = "Porcentual"
                   Case "M"
                     itmX.SubItems(2) = "Monto"
                End Select
                
                
                Select Case rs!Frecuencia_Tipo
                  Case "O"
                    itmX.SubItems(4) = "Operación"
                  Case "D"
                    itmX.SubItems(4) = "Días"
                End Select
                
                itmX.SubItems(3) = Format(rs!Valor, "Standard")
                itmX.SubItems(5) = rs!Frecuencia_dias
                itmX.SubItems(6) = Format(rs!Recaudado, "Standard")
                itmX.SubItems(7) = Format(rs!Pago_Ultimo, "dd/mm/yyyy")
                itmX.SubItems(8) = Format(rs!Pago_Proximo, "dd/mm/yyyy")
                itmX.SubItems(9) = rs!Modifica
          rs.MoveNext
        Loop
        rs.Close
        
       
End Select

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpia
      Call sbToolBar(Me.tlb, "edicion")
      vCodigo = ""
      txtCodigo.Text = ""
      txtCodigo.SetFocus

    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(Me.tlb, "edicion")
      
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
     
    Case "REPORTES"

    Case "CERRAR"
        Unload Me
End Select

End Sub





Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_contrato"
  gBusquedas.Orden = "cod_contrato"
  gBusquedas.Consulta = "select cod_contrato,descripcion from CxC_Contratos"
  gBusquedas.Filtro = " and activo = 1"
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()

If Trim(txtCodigo.Text) <> "" Then
   Call sbConsulta(Trim(txtCodigo))
End If


End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_Contrato,descripcion from CxC_Contratos"
  gBusquedas.Filtro = " and activo = 1"
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
  txtCodigo.SetFocus
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTasaCorriente.SetFocus
End Sub

Private Sub txtTasaCorriente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTasaMora.SetFocus
End Sub
