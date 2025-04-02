VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCajas_Recaudores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cajas: Recaudadores"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8970
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3732
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8652
      _Version        =   1310723
      _ExtentX        =   15261
      _ExtentY        =   6583
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
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "txtDescripcion"
      Item(0).Control(1)=   "Label1(15)"
      Item(0).Control(2)=   "Label1(1)"
      Item(0).Control(3)=   "GroupBox1"
      Item(0).Control(4)=   "txtNotas"
      Item(1).Caption =   "Contactos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Item(2).Caption =   "Conceptos ¦ Cajas"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "tlbServicios"
      Item(2).Control(1)=   "lswServicios"
      Item(2).Control(2)=   "lswCajas"
      Item(2).Control(3)=   "lblServicio"
      Item(2).Control(4)=   "Label1(4)"
      Begin MSComctlLib.Toolbar tlbServicios 
         Height          =   264
         Left            =   -66160
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   372
         _ExtentX        =   661
         _ExtentY        =   476
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Servicios"
               Object.ToolTipText     =   "Servicios"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswServicios 
         Height          =   2772
         Left            =   -69880
         TabIndex        =   8
         Top             =   876
         Visible         =   0   'False
         Width           =   4092
         _ExtentX        =   7223
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ListView lswCajas 
         Height          =   2772
         Left            =   -65680
         TabIndex        =   9
         Top             =   876
         Visible         =   0   'False
         Width           =   4212
         _ExtentX        =   7435
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3132
         Left            =   -69040
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   7092
         _Version        =   524288
         _ExtentX        =   12509
         _ExtentY        =   5524
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   6
         SpreadDesigner  =   "frmCajas_Recaudores.frx":0000
         VisibleRows     =   1
         VScrollSpecialType=   2
         Appearance      =   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   1800
         TabIndex        =   13
         Top             =   480
         Width           =   6732
         _Version        =   1310723
         _ExtentX        =   11874
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   912
         Left            =   1800
         TabIndex        =   15
         Top             =   840
         Width           =   6732
         _Version        =   1310723
         _ExtentX        =   11874
         _ExtentY        =   1609
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         DragMode        =   1  'Automatic
         Height          =   2892
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   8292
         _Version        =   1310723
         _ExtentX        =   14626
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Configuración Contable"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.FlatEdit txtCuenta 
            Height          =   312
            Left            =   1560
            TabIndex        =   18
            Top             =   480
            Width           =   1932
            _Version        =   1310723
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaComision 
            Height          =   312
            Left            =   1560
            TabIndex        =   19
            Top             =   840
            Width           =   1932
            _Version        =   1310723
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaIV 
            Height          =   312
            Left            =   1560
            TabIndex        =   20
            Top             =   1200
            Width           =   1932
            _Version        =   1310723
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   21
            Top             =   480
            Width           =   4812
            _Version        =   1310723
            _ExtentX        =   8488
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaComisionDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   22
            Top             =   840
            Width           =   4812
            _Version        =   1310723
            _ExtentX        =   8488
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaIVDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   23
            Top             =   1200
            Width           =   4812
            _Version        =   1310723
            _ExtentX        =   8488
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
         Begin VB.Label Label1 
            Caption         =   "Cuenta Principal"
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
            Left            =   0
            TabIndex        =   26
            Top             =   480
            Width           =   1692
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta I.V.A"
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
            Index           =   11
            Left            =   0
            TabIndex        =   25
            Top             =   1200
            Width           =   1692
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Comisión"
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
            Index           =   10
            Left            =   0
            TabIndex        =   24
            Top             =   840
            Width           =   1692
         End
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   2052
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
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
         Index           =   15
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Asignacion de servicios a cajas..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   4
         Left            =   -69880
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.Label lblServicio 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   -65680
         TabIndex        =   10
         Top             =   516
         Visible         =   0   'False
         Width           =   2820
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   4860
      Width           =   8964
      _ExtentX        =   15822
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario de Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8970
      _ExtentX        =   15822
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
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   7080
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Recaudores.frx":0705
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   252
      Left            =   3648
      TabIndex        =   2
      Top             =   480
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activo?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3048
      TabIndex        =   3
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
      Left            =   1488
      TabIndex        =   4
      Top             =   480
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Recaudador"
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
      Height          =   216
      Index           =   0
      Left            =   72
      TabIndex        =   5
      Top             =   480
      Width           =   1236
   End
End
Attribute VB_Name = "frmCajas_Recaudores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vPaso As Boolean



Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_recaudador from cajas_recaudador"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_recaudador > '" & txtCodigo & "' order by cod_recaudador asc"
    Else
       strSQL = strSQL & " where cod_recaudador < '" & txtCodigo & "' order by cod_recaudador desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!COD_RECAUDADOR)
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
vModulo = 5

End Sub

Private Sub Form_Load()
On Error GoTo vError

 vModulo = 5
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 

 vEdita = False
 
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpia
 
 Call Formularios(Me)
 Call RefrescaTags(Me)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub lswCajas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CAJAS_SERVICIOS_ASIGNADOS(cod_recaudador,cod_servicio,cod_Caja,registro_Fecha,registro_usuario)" _
          & " values('" & vCodigo & "','" & lblServicio.Tag & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"

   Call Bitacora("Registra", "Asignación en Caja: " & Item.Tag & " (Serv.:" & lblServicio.Tag & " - Rec.:" & vCodigo & ")")
Else
   strSQL = "delete CAJAS_SERVICIOS_ASIGNADOS where cod_recaudador = '" & vCodigo & "' and cod_servicio = '" _
          & lblServicio.Tag & "' and cod_caja = '" & Item.Tag & "'"
   Call Bitacora("Elimina", "Asignación en Caja: " & Item.Tag & " (Serv.:" & lblServicio.Tag & " - Rec.:" & vCodigo & ")")
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswServicios_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If lswServicios.ListItems.Count = 0 Or vPaso Then Exit Sub

lblServicio.Caption = "Servicio..: " & lswServicios.SelectedItem
lblServicio.Tag = lswServicios.SelectedItem.Tag

strSQL = "select C.cod_caja,C.descripcion,X.cod_caja as 'Asignado'" _
        & " from cajas_definicion C left join cajas_servicios_asignados X on C.cod_caja = X.cod_caja" _
        & " and X.cod_recaudador = '" & vCodigo & "' and X.cod_servicio = '" & lblServicio.Tag & "'"

vPaso = True
lswCajas.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswCajas.ListItems.Add(, , Trim(rs!Descripcion))
      itmX.Tag = Trim(rs!cod_caja)
  
  If Not IsNull(rs!Asignado) Then
     itmX.Checked = True
  End If
      
  rs.MoveNext
Loop
rs.Close

vPaso = False


End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vCodigo = "" Then Exit Sub

Select Case Item.Index
   Case 0 'Nada
   
   Case 1 'Contactos
        
          strSQL = "Select linea,identificacion,nombre,tel_trabajo,tel_celular,email" _
                 & " from cajas_recaudador_contactos" _
                 & " where cod_recaudador = '" & vCodigo & "' order by Linea"
          Call sbCargaGrid(vGrid, 6, strSQL)
   
   Case 2 'Servicios
            
            vPaso = True
            
            lswServicios.ListItems.Clear
            lswCajas.ListItems.Clear
            lblServicio.Caption = "Servicio ..:"
            lblServicio.Tag = ""
                        
            strSQL = "select cod_servicio , Descripcion from cajas_servicios" _
                   & " where cod_recaudador = '" & vCodigo & "' order by cod_servicio"
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
              Set itmX = lswServicios.ListItems.Add(, , Trim(rs!Descripcion))
                  itmX.Tag = Trim(rs!COD_SERVICIO)
              rs.MoveNext
            Loop
            rs.Close
            
            vPaso = False
             
End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpia
      txtCodigo.Text = ""
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpia
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_recaudador,descripcion from cajas_recaudador"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo.Text = gBusquedas.Resultado
       txtDescripcion.SetFocus

    Case "REPORTES"

    Case "AYUDA"

End Select


End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Recaudador no es válido ..."
If Not fxgCntCuentaValida(txtCuenta.Text) Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable Prinicipal no es válida.."
If Not fxgCntCuentaValida(txtCuentaIV.Text) Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable para Impuesto de Ventas no es válida.."
If Not fxgCntCuentaValida(txtCuentaComision.Text) Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable para Comisiones no es válida.."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update cajas_recaudador set descripcion = '" & Trim(txtDescripcion.Text) & "'" _
         & ",notas = '" & UCase(txtNotas) & "',activo = " & chkActivo.Value _
         & ",cod_cuenta = '" & fxgCntCuentaFormato(False, txtCuenta.Text) _
         & "',cod_cuenta_Comision = '" & fxgCntCuentaFormato(False, txtCuentaComision.Text) _
         & "',cod_cuenta_IV = '" & fxgCntCuentaFormato(False, txtCuentaIV.Text) _
         & "' where cod_recaudador = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Recaudador: " & vCodigo)

Else
  vCodigo = txtCodigo.Text

   strSQL = "insert cajas_recaudador(cod_recaudador,descripcion,notas,activo,cod_cuenta,cod_cuenta_IV,cod_cuenta_Comision" _
          & ", registro_fecha,registro_usuario)" _
          & " values('" & vCodigo & "','" & Trim(txtDescripcion.Text) & "','" & txtNotas.Text & "'," & chkActivo.Value & ",'" _
          & fxgCntCuentaFormato(False, txtCuenta.Text) & "','" & fxgCntCuentaFormato(False, txtCuentaIV.Text) _
          & "','" & fxgCntCuentaFormato(False, txtCuentaComision.Text) & "', dbo.MyGetdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Recaudador: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub sbLimpia()

vCodigo = ""

txtDescripcion.Text = ""
txtNotas.Text = ""
chkActivo.Value = vbChecked

txtCuenta.Text = ""
txtCuentaDesc.Text = ""
txtCuentaComision.Text = ""
txtCuentaComisionDesc.Text = ""
txtCuentaIV.Text = ""
txtCuentaIVDesc.Text = ""

StatusBarX.Panels.Item(1).Text = ""
StatusBarX.Panels.Item(2).Text = ""

tcMain.Item(0).Selected = True
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False

End Sub

Private Sub tlbServicios_ButtonClick(ByVal Button As MSComctlLib.Button)
GLOBALES.gTag = vCodigo
Call sbFormsCall("frmCajas_Servicios", 1, , , False, Me)
'Call SSTab_Click(2)

End Sub


Private Sub txtCodigo_Change()
Call sbLimpia
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_recaudador"
   gBusquedas.Orden = "cod_recaudador"
   gBusquedas.Consulta = "select cod_recaudador,descripcion from cajas_recaudador"
   frmBusquedas.Show vbModal
   txtCodigo.SetFocus
   txtCodigo = gBusquedas.Resultado
   txtDescripcion.SetFocus
End If

End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo.Text) <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta.Text = gCuenta
    txtCuentaDesc.Text = ""
End If

End Sub

Private Sub txtCuenta_LostFocus()
    txtCuenta.Text = fxgCntCuentaFormato(False, txtCuenta.Text)
    txtCuentaDesc.Text = fxgCntCuentaDesc(txtCuenta.Text)
    txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text)
End Sub


Private Sub txtCuentaComision_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaComisionDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaComision.Text = gCuenta
    txtCuentaComisionDesc.Text = ""
End If

End Sub


Private Sub txtCuentaComision_LostFocus()
    txtCuentaComision.Text = fxgCntCuentaFormato(False, txtCuentaComision.Text)
    txtCuentaComisionDesc.Text = fxgCntCuentaDesc(txtCuentaComision.Text)
    txtCuentaComision.Text = fxgCntCuentaFormato(True, txtCuentaComision.Text)
End Sub

Private Sub txtCuentaComisionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIV.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaComisionDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtCuentaComision.Text = fxgCntCuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaComision.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtCuenta.Text = fxgCntCuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtCuentaIV_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIVDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaIV.Text = gCuenta
    txtCuentaIVDesc.Text = ""
End If

End Sub

Private Sub txtCuentaIV_LostFocus()
    txtCuentaIV.Text = fxgCntCuentaFormato(False, txtCuentaIV.Text)
    txtCuentaIVDesc.Text = fxgCntCuentaDesc(txtCuentaIV.Text)
    txtCuentaIV.Text = fxgCntCuentaFormato(True, txtCuentaIV.Text)
End Sub

Private Sub txtCuentaIVDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaIVDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtCuentaIV.Text = fxgCntCuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_recaudador,descripcion from cajas_recaudador"
   frmBusquedas.Show vbModal
   txtCodigo.SetFocus
   txtCodigo = gBusquedas.Resultado
   Call txtCodigo_LostFocus
End If


End Sub

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select R.*,isnull(Cta.Descripcion,'') as 'CtaDesc',isnull(CtaIv.Descripcion,'') as 'CtaDescIv', isnull(CtaCom.Descripcion,'') as 'CtaDescCom'" _
       & " from cajas_recaudador R left join CntX_Cuentas Cta on R.cod_Cuenta = Cta.Cod_Cuenta and Cta.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " left join CntX_Cuentas CtaIv on R.cod_Cuenta_Iv = CtaIv.Cod_Cuenta and CtaIv.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " left join CntX_Cuentas CtaCom on R.cod_Cuenta_Comision = CtaCom.Cod_Cuenta and CtaCom.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " where R.cod_recaudador = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

Call sbLimpia


If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  txtCodigo = rs!COD_RECAUDADOR
  vCodigo = Trim(rs!COD_RECAUDADOR)

  chkActivo.Value = rs!activo

  txtDescripcion = rs!Descripcion
  txtNotas = rs!Notas
  
  txtCuenta.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
  txtCuentaDesc.Text = rs!CtaDesc
  
  txtCuentaIV.Text = fxgCntCuentaFormato(True, rs!Cod_Cuenta_IV)
  txtCuentaIVDesc.Text = rs!CtaDescIV
  
  txtCuentaComision.Text = fxgCntCuentaFormato(True, rs!Cod_Cuenta_Comision)
  txtCuentaComisionDesc.Text = rs!CtaDescCom
  
  StatusBarX.Panels.Item(1).Text = rs!Registro_Usuario & ""
  StatusBarX.Panels.Item(2).Text = rs!Registro_Fecha & ""

  tcMain.Item(0).Selected = True
  tcMain.Item(1).Enabled = True
  tcMain.Item(2).Enabled = True
  
End If
rs.Close

Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete cajas_recaudadores where cod_recaudador = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Recaudador : " & vCodigo)
  Call sbLimpia
  
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Function fxGuardarContacto() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarContacto = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1


If vGrid.Text = "" Then 'Insertar
    strSQL = "select isnull(max(Linea),0) + 1 as 'Linea' from cajas_recaudador_contactos" _
           & " where cod_recaudador = '" & vCodigo & "'"
    Call OpenRecordSet(rs, strSQL)
      vGrid.Text = CStr(rs!Linea)
    rs.Close
    
    strSQL = "insert into cajas_recaudador_contactos(cod_recaudador,linea,identificacion,nombre,tel_trabajo,tel_celular,email)" _
           & " values('" & vCodigo & "'," & vGrid.Text & ",'"
    vGrid.col = 2
    strSQL = strSQL & Trim(vGrid.Text) & "',"
    vGrid.col = 3
    strSQL = strSQL & "'" & Trim(vGrid.Text) & "',"
    vGrid.col = 4
    strSQL = strSQL & "'" & Trim(vGrid.Text) & "',"
    vGrid.col = 5
    strSQL = strSQL & " '" & Trim(vGrid.Text) & "',"
    vGrid.col = 6
    strSQL = strSQL & "'" & Trim(vGrid.Text) & "')"
    Call ConectionExecute(strSQL)
    
    vGrid.col = 2
    Call Bitacora("Registra", "Contato ..: " & vGrid.Text & " ... Recaudador ..:" & vCodigo)

Else 'Actualizar
    
    vGrid.col = 2
    strSQL = "update cajas_recaudador_contactos set identificacion= " & Trim(vGrid.Text) & ","
    vGrid.col = 3
      strSQL = strSQL & "nombre = '" & Trim(vGrid.Text) & "',"
    vGrid.col = 4
    strSQL = strSQL & "tel_trabajo = '" & Trim(vGrid.Text) & "',"
    vGrid.col = 5
    strSQL = strSQL & "tel_celular = '" & Trim(vGrid.Text) & "',"
    vGrid.col = 6
    strSQL = strSQL & "email = '" & Trim(vGrid.Text) & "'"
    vGrid.col = 1
    strSQL = strSQL & " where linea =  " & vGrid.Text & " and cod_recaudador = '" & vCodigo & "'"
            
    vGrid.col = 2
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Contato ..: " & vGrid.Text & " ... Recaudador ..:" & vCodigo)

End If

fxGuardarContacto = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarContacto
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If



If KeyCode = vbKeyDelete Then
    If MsgBox("¿Desea eliminar este registro?", vbYesNo Or vbQuestion, "") = vbYes Then
        vGrid.col = 1
        If vGrid.Text = "" Then Exit Sub
        
        strSQL = "Delete cajas_recaudador_contactos where cod_recaudador = '" & vCodigo & "' and linea = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        vGrid.col = 2
        Call Bitacora("Elimina", "Contato ..: " & vGrid.Text & " ... Recaudador ..:" & vCodigo)
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.Row = vGrid.ActiveRow
        vGrid.MaxRows = vGrid.MaxRows - 1
    End If
End If

End Sub
