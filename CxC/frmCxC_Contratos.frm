VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxC_Contratos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Contratos"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   9030
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4815
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   8493
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
      ItemCount       =   5
      Item(0).Caption =   "General"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "chkActivo"
      Item(0).Control(1)=   "txtNotas"
      Item(0).Control(2)=   "txtPlazo"
      Item(0).Control(3)=   "txtTasaCorriente"
      Item(0).Control(4)=   "txtTasaMora"
      Item(0).Control(5)=   "Label1(6)"
      Item(0).Control(6)=   "Label1(2)"
      Item(0).Control(7)=   "Label1(3)"
      Item(0).Control(8)=   "Label1(5)"
      Item(0).Control(9)=   "Label1(4)"
      Item(0).Control(10)=   "Label1(7)"
      Item(0).Control(11)=   "GroupBox1"
      Item(1).Caption =   "Suscripciones"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswS"
      Item(2).Caption =   "Pagadores"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswP"
      Item(3).Caption =   "Cargos Contractuales"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lswC"
      Item(4).Caption =   "Conceptos"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "lswConceptos"
      Begin XtremeSuiteControls.ListView lswC 
         Height          =   4215
         Left            =   -69880
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   8775
         _Version        =   1441793
         _ExtentX        =   15478
         _ExtentY        =   7435
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
      Begin XtremeSuiteControls.ListView lswConceptos 
         Height          =   4215
         Left            =   -69880
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   8775
         _Version        =   1441793
         _ExtentX        =   15478
         _ExtentY        =   7435
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
      Begin XtremeSuiteControls.ListView lswS 
         Height          =   4215
         Left            =   -69880
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   8775
         _Version        =   1441793
         _ExtentX        =   15478
         _ExtentY        =   7435
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
         Height          =   4215
         Left            =   -69880
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   8775
         _Version        =   1441793
         _ExtentX        =   15478
         _ExtentY        =   7435
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
         Left            =   7320
         TabIndex        =   9
         Top             =   435
         Width           =   1215
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   795
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   7335
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         Top             =   1800
         Width           =   975
         _Version        =   1441793
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
         TabIndex        =   12
         Top             =   2160
         Width           =   975
         _Version        =   1441793
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
         TabIndex        =   13
         Top             =   2520
         Width           =   975
         _Version        =   1441793
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1695
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Width           =   8295
         _Version        =   1441793
         _ExtentX        =   14631
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Clientes y Pagadores (General) "
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkSuscripcionAbierta 
            Height          =   252
            Left            =   960
            TabIndex        =   22
            Top             =   480
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Suscripción Abierta ( Aplica para todos los clientes activos)"
            BackColor       =   -2147483633
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
         Begin XtremeSuiteControls.CheckBox chkPagadoresAbierto 
            Height          =   252
            Left            =   960
            TabIndex        =   23
            Top             =   840
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Participan todos los pagadores registrados"
            BackColor       =   -2147483633
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
         Index           =   7
         Left            =   3960
         TabIndex        =   19
         Top             =   2520
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
         TabIndex        =   18
         Top             =   2160
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
         TabIndex        =   17
         Top             =   2520
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
         TabIndex        =   16
         Top             =   1800
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
         TabIndex        =   15
         Top             =   2160
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
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Contratos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Contratos.frx":3492
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Contratos.frx":6924
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Contratos.frx":6A42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9030
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinWidth1       =   1800
      MinHeight1      =   330
      Width1          =   1800
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   1110
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   8025
         TabIndex        =   5
         Top             =   30
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   7635
         _ExtentX        =   13467
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
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5775
      Width           =   9030
      _ExtentX        =   15928
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
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8400
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Width           =   1815
      _Version        =   1441793
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   480
      Width           =   5535
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmCxC_Contratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset



Private Sub sbLimpia()

tcMain.Item(0).Selected = True
tlbAux.Buttons.Item(1).Enabled = False 'Nuevo
tlbAux.Buttons.Item(3).Enabled = False 'Borrar


txtDescripcion.Text = ""
txtNotas.Text = ""

chkActivo.Value = xtpChecked
chkPagadoresAbierto.Value = xtpUnchecked
chkSuscripcionAbierta.Value = xtpUnchecked

txtPlazo.Text = "30"
txtTasaCorriente.Text = "0.00"
txtTasaMora.Text = "0.00"

StatusBarX.Panels.Item(1).Text = ""
StatusBarX.Panels.Item(2).Text = ""
StatusBarX.Panels.Item(3).Text = ""
StatusBarX.Panels.Item(4).Text = ""


End Sub


Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_contrato from CxC_Contratos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_contrato > '" & txtCodigo & "' order by cod_contrato asc"
    Else
       strSQL = strSQL & " where cod_contrato < '" & txtCodigo & "' order by cod_contrato desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_Contrato
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

vEdita = True

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 

Call sbToolBarIconos(Me.tlb)
Call sbToolBar(Me.tlb, "nuevo")



With lswS.ColumnHeaders
    .Clear
    .Add , , "Cédula", 2000
    .Add , , "Nombre", 3500
    .Add , , "Registro", 3000
    .Add , , "Plazo", 1200, vbRightJustify
    .Add , , "Tasa Cor.", 1400, vbRightJustify
    .Add , , "Tasa. Mor.", 1400, vbRightJustify
    .Add , , "Notas", 3000
End With
With lswP.ColumnHeaders
    .Clear
    .Add , , "Cédula", 2000
    .Add , , "Nombre", 3500
    .Add , , "Registro", 3000
End With
With lswC.ColumnHeaders
    .Clear
    .Add , , "Código", 1400
    .Add , , "Cargo", 3500
    .Add , , "Valor", 1800, vbRightJustify
    .Add , , "Tipo", 1400, vbCenter
    .Add , , "Frecuencia", 1600, vbCenter
    .Add , , "Frec.Días", 1600, vbCenter
    .Add , , "Modifica?", 1200, vbCenter
    .Add , , "Registro", 3000
End With
With lswConceptos.ColumnHeaders
    .Clear
    .Add , , "Código", 1400
    .Add , , "Descripción", 3500
    .Add , , "Reg. Usuario", 2100, vbCenter
    .Add , , "Reg. Fecha", 2100, vbCenter
End With







Call sbLimpia

Call Formularios(Me)
Call RefrescaTags(Me)

'Activa Replica de Seguridad a Componentes Asociados la Accion de Editar:
lswC.Enabled = tlb.Buttons.Item(1).Enabled
lswP.Enabled = tlb.Buttons.Item(1).Enabled
lswS.Enabled = tlb.Buttons.Item(1).Enabled
lswConceptos.Enabled = tlb.Buttons.Item(1).Enabled


End Sub


Private Sub sbConsulta(pContrato As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

strSQL = "select * from CxC_Contratos where cod_contrato = '" & pContrato & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(Me.tlb, "activo")
  vEdita = True
  vCodigo = Trim(rs!COD_Contrato)
  txtCodigo = Trim(rs!COD_Contrato)
  txtDescripcion = Trim(rs!Descripcion)
  txtNotas.Text = rs!Notas

  txtPlazo.Text = rs!Plazo
  txtTasaCorriente.Text = Format(rs!Tasa_Corriente, "Standard")
  txtTasaMora.Text = Format(rs!Tasa_Mora, "Standard")

  chkSuscripcionAbierta.Value = rs!SUSCRIPCION_ABIERTA
  chkPagadoresAbierto.Value = rs!PAGADORES_ABIERTO

  StatusBarX.Panels.Item(1).Text = rs!Registro_Usuario
  StatusBarX.Panels.Item(2).Text = rs!Registro_Fecha
  StatusBarX.Panels.Item(3).Text = rs!Actualiza_usuario & ""
  StatusBarX.Panels.Item(4).Text = rs!Actualiza_fecha & ""
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
  strSQL = "Update CxC_Contratos Set descripcion='" & Trim(txtDescripcion) & "', Notas = '" & Trim(txtNotas.Text) _
         & "',Activo = " & chkActivo.Value & ", Plazo = " & txtPlazo.Text & ", Tasa_Corriente = " & CCur(txtTasaCorriente.Text) _
         & ",Tasa_Mora = " & CCur(txtTasaMora.Text) & ",Actualiza_Usuario = '" & glogon.Usuario & "',Actualiza_Fecha = dbo.MyGetdate()" _
         & ", SUSCRIPCION_ABIERTA = " & chkSuscripcionAbierta.Value & ", PAGADORES_ABIERTO = " & chkPagadoresAbierto.Value _
         & " where cod_contrato = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  
  Call Bitacora("Modifica", "Contrato: " & Trim(txtCodigo.Text))

Else
   strSQL = "insert CxC_Contratos(cod_contrato,descripcion,Notas,Activo,Plazo,Tasa_Corriente,Tasa_Mora" _
          & ",SUSCRIPCION_ABIERTA,PAGADORES_ABIERTO, registro_usuario,registro_fecha)" _
          & " values('" & txtCodigo.Text & "','" & Trim(txtDescripcion) & "','" & Trim(txtNotas.Text) & "'," & chkActivo.Value _
          & "," & txtPlazo.Text & "," & CCur(txtTasaCorriente.Text) & "," & CCur(txtTasaMora.Text) _
          & "," & chkSuscripcionAbierta.Value & "," & chkPagadoresAbierto.Value _
          & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)
  
   Call Bitacora("Registra", "Contrato: " & Trim(txtCodigo.Text))
   
End If

vCodigo = Trim(txtCodigo)

vEdita = True

Call sbToolBar(Me.tlb, "activo")

txtDescripcion.SetFocus

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
  strSQL = "delete CxC_Contratos where cod_contrato = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Borra", "Contrato: " & txtCodigo.Text)
  
  Call sbLimpia
  Call sbToolBar(Me.tlb, "nuevo")
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub





Private Sub lswConceptos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError


If Item.Checked Then
   strSQL = "insert CxC_Conceptos_Contratos(cod_concepto,cod_contrato,registro_usuario,registro_fecha) values('" _
          & Item.Text & "','" & vCodigo & "','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = "delete CxC_Conceptos_Contratos where cod_concepto = '" & Item.Text & "' and cod_contrato = '" & vCodigo & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Dim itmX As ListViewItem

If vCodigo = "" Then
  
  tcMain.Item(0).Selected = True
  tlbAux.Buttons.Item(1).Enabled = False 'Nuevo
  tlbAux.Buttons.Item(3).Enabled = False 'Borrar
 
  Exit Sub
End If

Me.MousePointer = vbHourglass

If Item.Index > 0 Then
  tlbAux.Buttons.Item(1).Enabled = True 'Nuevo
  tlbAux.Buttons.Item(3).Enabled = True 'Borrar
End If

vPaso = True


Select Case Item.Index
   Case 1 'Suscripciones
       tlbAux.Buttons.Item(1).ToolTipText = "Nueva : Suscripción"
       tlbAux.Buttons.Item(3).ToolTipText = "Borra : Suscripción"
       
       lswS.ListItems.Clear
       strSQL = " select P.nombre,C.*" _
              & " from CxC_Personas P inner join CxC_Personas_Contratos C on P.cedula = C.cedula" _
              & " where C.cod_contrato = '" & vCodigo & "'"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswS.ListItems.Add(, , rs!Cedula)
             itmX.SubItems(1) = rs!Nombre
             itmX.SubItems(2) = rs!Registro_Usuario & "..." & rs!Registro_Fecha
             itmX.SubItems(3) = rs!Plazo
             itmX.SubItems(4) = Format(rs!Tasa_Corriente, "Standard")
             itmX.SubItems(5) = Format(rs!Tasa_Mora, "Standard")
             itmX.SubItems(6) = Format(rs!Notas, "Standard")
          rs.MoveNext
       Loop
       rs.Close
   
   
   Case 2 'Pagadores
       
       tlbAux.Buttons.Item(1).ToolTipText = "Nuevo : Pagador"
       tlbAux.Buttons.Item(3).ToolTipText = "Borra : Pagador"
       
       lswP.ListItems.Clear
       strSQL = " select P.nombre, C.*" _
              & " from CxC_Personas P inner join CxC_Contratos_Pagadores C on P.cedula = C.cedula" _
              & " where C.cod_contrato = '" & vCodigo & "'"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswP.ListItems.Add(, , rs!Cedula)
             itmX.SubItems(1) = rs!Nombre
             itmX.SubItems(2) = rs!Registro_Usuario & "..." & rs!Registro_Fecha
         rs.MoveNext
       Loop
       rs.Close
   
   
   Case 3 'Cargos
      
       tlbAux.Buttons.Item(1).ToolTipText = "Nuevo : Cargo Contractual"
       tlbAux.Buttons.Item(3).ToolTipText = "Borra : Cargo Contractual"
       
       lswC.ListItems.Clear
       strSQL = " select C.descripcion,S.*" _
              & " from CxC_Cargos C inner join CxC_Contratos_Cargos S on C.cod_cargo = S.cod_cargo" _
              & " where S.cod_contrato = '" & vCodigo & "'"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
          Set itmX = lswC.ListItems.Add(, , rs!COD_CARGO)
              itmX.SubItems(1) = rs!Descripcion
              itmX.SubItems(3) = Format(rs!Valor, "Standard")
        
            If rs!Tipo = "P" Then
               itmX.SubItems(2) = "Porcentual"
            Else
               itmX.SubItems(2) = "Monto"
            End If
           
           
            If rs!Frecuencia_Tipo = "P" Then
               itmX.SubItems(4) = "Operación"
            Else
               itmX.SubItems(4) = "Días"
            End If
           
            itmX.SubItems(5) = rs!Frecuencia_dias
            itmX.SubItems(6) = IIf((rs!Modifica = 1), "Sí", "No")
            itmX.SubItems(7) = rs!Registro_Usuario & "..." & rs!Registro_Fecha & ""
    
         rs.MoveNext
       Loop
       rs.Close
       
       
   Case 4 'Conceptos
       tlbAux.Buttons.Item(1).Enabled = False 'Nuevo
       tlbAux.Buttons.Item(3).Enabled = True 'Borrar
       tlbAux.Buttons.Item(1).ToolTipText = "Nuevo : Concepto Nuevo"
       tlbAux.Buttons.Item(3).ToolTipText = "Borra : Concepto Asociado"
       
       lswConceptos.ListItems.Clear
       strSQL = " select C.cod_concepto as 'Codigo',C.descripcion,S.*" _
              & " from CxC_Conceptos C left join CxC_Conceptos_Contratos S on C.cod_Concepto = S.cod_Concepto" _
              & " and S.cod_contrato = '" & vCodigo & "'"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswConceptos.ListItems.Add(, , rs!Codigo)
             itmX.SubItems(1) = rs!Descripcion
             itmX.SubItems(2) = rs!Registro_Usuario & ""
             itmX.SubItems(3) = rs!Registro_Fecha & ""
         
         If IsNull(rs!Registro_Fecha) Then
            itmX.Checked = False
         Else
            itmX.Checked = True
         End If
         
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


Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long

GLOBALES.gTag = vCodigo
GLOBALES.gTag2 = txtDescripcion.Text

Select Case Button.Key
 Case "nuevo"
    Select Case tcMain.SelectedItem
       Case 1 'Suscripciones
            Call sbFormsCall("frmCxC_ContratosSuscripciones", 1, , , False)
       Case 2 'Pagadores
            Call sbFormsCall("frmCxC_ContratosPagadores", 1, , , False)
       Case 3 'Cargos Contractuales
            Call sbFormsCall("frmCxC_ContratosCargos", 1, , , False)
    End Select
 
 
 Case "borrar"
    Select Case tcMain.SelectedItem
       Case 1 'Suscripciones
          With lswS.ListItems
          For i = 1 To .Count
            If .Item(i).Checked Then
               strSQL = "delete CxC_Personas_Contratos_Pagadores where cod_contrato = '" & vCodigo _
                      & "' and cedula = '" & .Item(i).Text & "'"
               Call ConectionExecute(strSQL)
               
               strSQL = "delete CxC_Personas_Contratos_Suscripciones where cod_contrato = '" & vCodigo _
                      & "' and cedula = '" & .Item(i).Text & "'"
               Call ConectionExecute(strSQL)
               
               strSQL = "delete CxC_Personas_Contratos where cod_contrato = '" & vCodigo _
                      & "' and cedula = '" & .Item(i).Text & "'"
               Call ConectionExecute(strSQL)
               
            End If
          Next i
        End With
        
       Case 2 'Pagadores
       
          With lswP.ListItems
          For i = 1 To .Count
            If .Item(i).Checked Then
               strSQL = "delete CxC_Contratos_Pagadores where cod_contrato = '" & vCodigo _
                      & "' and cedula = '" & .Item(i).Text & "'"
               Call ConectionExecute(strSQL)
               
                Call Bitacora("Borra", "Pagador Id.:" & .Item(i).Text & " de Contrato No.:" & vCodigo)
               
            End If
          Next i
          End With
       
       Case 3 'Cargos Contractuales
    
          With lswC.ListItems
          For i = 1 To .Count
            If .Item(i).Checked Then
               strSQL = "delete CxC_Contratos_Cargos where cod_contrato = '" & vCodigo _
                      & "' and cod_cargo = '" & .Item(i).Text & "'"
               Call ConectionExecute(strSQL)
  
               Call Bitacora("Borra", "Cargo Suscripción Cod:" & .Item(i).Text & " Cnt: " & vCodigo)
            End If
          Next i
          End With
  
    End Select

End Select 'Toolbar

'Refresca la información
tcMain.Item(0).Selected = True


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_Contrato"
  gBusquedas.Orden = "cod_Contrato"
  gBusquedas.Consulta = "select cod_contrato as 'Contrato',Descripcion from CxC_Contratos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()

If Trim(txtCodigo) <> "" Then
   Call sbConsulta(Trim(txtCodigo))
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   tcMain.Item(0).Selected = True
   txtNotas.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select cod_contrato as 'Contrato',Descripcion from CxC_Contratos"
  gBusquedas.Filtro = ""
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
