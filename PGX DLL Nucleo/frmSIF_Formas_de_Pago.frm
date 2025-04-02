VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmSIF_Formas_de_Pago 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formas de pago "
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   10610
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
      Item(0).Caption =   "Detalle"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "GroupBox1"
      Item(0).Control(2)=   "GroupBox2"
      Item(1).Caption =   "Cuentas Bancarias Autorizadas"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "scTitulo"
      Item(1).Control(1)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5295
         Left            =   -70000
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1441793
         _ExtentX        =   16325
         _ExtentY        =   9340
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   3015
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   9255
         _Version        =   1441793
         _ExtentX        =   16325
         _ExtentY        =   5318
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.CheckBox chkMaximoAplica 
            Height          =   255
            Left            =   4200
            TabIndex        =   27
            Top             =   1680
            Width           =   3855
            _Version        =   1441793
            _ExtentX        =   6800
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Máximo permitido por transacción"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Top             =   360
            Width           =   7455
            _Version        =   1441793
            _ExtentX        =   13144
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.ComboBox cboTipo 
            Height          =   315
            Left            =   1680
            TabIndex        =   20
            Top             =   840
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
         Begin XtremeSuiteControls.CheckBox chkEfectivo 
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   1320
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Efectivo ?"
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
         Begin XtremeSuiteControls.CheckBox chkSaldosFavor 
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   2040
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Control de Saldos a Favor ?"
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
         Begin XtremeSuiteControls.CheckBox chkAplicaDep 
            Height          =   255
            Left            =   1680
            TabIndex        =   23
            Top             =   2400
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Control de Depósitos Bancarios"
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
         Begin XtremeSuiteControls.FlatEdit txtMaximoMonto 
            Height          =   315
            Left            =   1680
            TabIndex        =   26
            Top             =   1680
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   25
            Top             =   360
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Descripción"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   24
            Top             =   840
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Formato"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1575
         Left            =   0
         TabIndex        =   12
         Top             =   3360
         Width           =   9255
         _Version        =   1441793
         _ExtentX        =   16325
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Origen de Recursos (Solicitar información) "
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
         Begin XtremeSuiteControls.CheckBox chkOrigen 
            Height          =   255
            Left            =   1680
            TabIndex        =   13
            Top             =   360
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10186
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplica Validación de Origen de Recursos?"
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
         End
         Begin XtremeSuiteControls.CheckBox chkOrigenAplDiario 
            Height          =   255
            Left            =   1680
            TabIndex        =   14
            Top             =   720
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Diario a partir de: "
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtOrigenMontoDiario 
            Height          =   315
            Left            =   4200
            TabIndex        =   15
            Top             =   720
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.CheckBox chkOrigenAplMensual 
            Height          =   255
            Left            =   1680
            TabIndex        =   16
            Top             =   1080
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Mensual a partir de: "
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtOrigenMontoMensual 
            Height          =   315
            Left            =   4200
            TabIndex        =   17
            Top             =   1080
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4260
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1095
         Left            =   0
         TabIndex        =   5
         Top             =   4920
         Width           =   9255
         _Version        =   1441793
         _ExtentX        =   16325
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Cuenta Contable:"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   315
            Left            =   3600
            TabIndex        =   6
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   -70000
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1441793
         _ExtentX        =   16319
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Indicar las cuentas bancarias asociadas a la forma de pago (Depósitos) para registros directos en Cajas/Tesoreria "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
      Height          =   252
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   432
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin XtremeSuiteControls.CheckBox chkActiva 
      Height          =   252
      Left            =   7560
      TabIndex        =   4
      Top             =   600
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activa ?  "
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   372
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Forma de Pago"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSIF_Formas_de_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim vEdita  As Boolean
Dim vCodigo As String, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem



Private Function fxExisteForma(vCodigo As String) As Boolean

strSQL = "select isnull(count(*),0) as Existe" _
       & " from sif_formas_pago where cod_forma_pago =  '" & vCodigo & "' "
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  fxExisteForma = False
Else
  fxExisteForma = True
End If
rs.Close
End Function


Private Sub cboTipo_Click()

If vPaso Then Exit Sub


tcMain.Item(1).Enabled = False

chkSaldosFavor.Enabled = True
chkSaldosFavor.Value = vbUnchecked
chkEfectivo.Enabled = True
chkEfectivo.Value = vbUnchecked
chkAplicaDep.Enabled = True
chkAplicaDep.Value = vbUnchecked

Select Case Mid(cboTipo.Text, 1, 2)
    Case "SA" 'Saldo a Favor
        chkSaldosFavor.Value = vbUnchecked
        chkSaldosFavor.Enabled = False
    
        chkEfectivo.Enabled = False
        chkEfectivo.Value = vbUnchecked
    
        chkAplicaDep.Enabled = False
        chkAplicaDep.Value = vbUnchecked
    
    Case "FO" 'Fondos (Planes de Ahorros)
        chkSaldosFavor.Value = vbUnchecked
        chkSaldosFavor.Enabled = False
    
        chkEfectivo.Enabled = False
        chkEfectivo.Value = vbUnchecked
    
        chkAplicaDep.Enabled = False
        chkAplicaDep.Value = vbUnchecked
    
    Case "DE" 'Depositos
        
       tcMain.Item(1).Enabled = True
    Case Else
    
End Select

End Sub


Private Sub FlatScrollBar2_Change()

On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 cod_forma_pago from sif_formas_pago"
    
    If FlatScrollBar2.Value = 1 Then
       strSQL = strSQL & " where cod_forma_pago > '" & txtCodigo.Text & "' order by cod_forma_pago asc"
    Else
       strSQL = strSQL & " where cod_forma_pago < '" & txtCodigo.Text & "' order by cod_forma_pago desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_forma_pago
      Call sbConsulta(txtCodigo.Text)
      
    End If
End If

vScroll = False
FlatScrollBar2.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub


Private Sub Form_Load()
vModulo = 10
 
vEdita = True


cboTipo.Clear
cboTipo.AddItem "EFECTIVO"
cboTipo.AddItem "CHEQUE"
cboTipo.AddItem "TARJETA"
cboTipo.AddItem "DEPOSITO"
cboTipo.AddItem "DOCUMENTO"
cboTipo.AddItem "FONDOS"
cboTipo.AddItem "SALDO A FAVOR"


With lsw.ColumnHeaders
    .Clear
    .Add , , "( Id )", 800
    .Add , , "Cuenta", 2300, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Divisa", 700, vbCenter
    .Add , , "Entidad", 2200
End With


Call sbToolBarIconos(tlb, False)
Call sbToolBar(tlb, "nuevo")

vScroll = False
FlatScrollBar2.Value = 0
vScroll = True

Call sbLimpiaDatos

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpiaDatos()

vCodigo = ""
lsw.ListItems.Clear

tcMain.Item(0).Selected = True

txtCodigo.Text = ""
txtDescripcion.Text = ""
txtCuentaDesc = ""
txtCuentaCod.Text = ""

chkActiva.Value = vbChecked
chkEfectivo.Value = vbUnchecked
chkSaldosFavor.Value = vbUnchecked

vPaso = True
    cboTipo.Text = "EFECTIVO"
vPaso = False


chkMaximoAplica.Value = xtpUnchecked
txtMaximoMonto.Text = Format(0, "Standard")
  
chkOrigen.Value = xtpUnchecked
chkOrigenAplDiario.Value = xtpUnchecked
txtOrigenMontoDiario.Text = Format(0, "Standard")
chkOrigenAplMensual.Value = xtpUnchecked
txtOrigenMontoMensual.Text = Format(0, "Standard")

Call cboTipo_Click

End Sub




Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert SIF_FORMAS_PAGO_BANCOS_ASG(id_banco,cod_forma_pago,registro_fecha,registro_usuario)" _
          & " values(" & Item.Tag & ",'" & vCodigo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete SIF_FORMAS_PAGO_BANCOS_ASG where id_Banco = " & Item.Tag & " and cod_forma_pago = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaCod = gCuenta
    txtCuentaDesc = fxgCntCuentaDesc(gCuenta)
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus
End Sub

Private Sub txtCuentaCod_LostFocus()
txtCuentaCod.Text = fxgCntCuentaFormato(False, txtCuentaCod.Text)
txtCuentaDesc.Text = fxgCntCuentaDesc(txtCuentaCod.Text)
txtCuentaCod.Text = fxgCntCuentaFormato(True, txtCuentaCod.Text)
End Sub


Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" And vEdita = True Then Call sbConsulta(txtCodigo.Text)
End Sub



Private Sub sbLswLlena(pCodigo As String)

On Error GoTo vError

vPaso = True

lsw.ListItems.Clear


strSQL = "select Ban.ID_BANCO, Ban.DESCRIPCION, Ban.Cta, isnull(Fp.Id_Banco,0) as 'Idx'" _
       & ", Ban.Cod_Divisa, isnull(Eb.DESCRIPCION,'') as 'Entidad_Desc'" _
       & " from TES_BANCOS Ban" _
       & "      left join SIF_FORMAS_PAGO_BANCOS_ASG Fp" _
       & "      on Ban.ID_BANCO = Fp.id_banco and Fp.cod_forma_pago = '" & pCodigo & "'" _
       & "      left join TES_BANCOS_GRUPOS Eb on Ban.Cod_Grupo = Eb.Cod_Grupo" _
       & " where Ban.ESTADO = 'A'" _
       & "  order by Fp.id_Banco desc,Ban.ID_BANCO asc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , "( " & rs!id_banco & " )")
      itmX.Tag = rs!id_banco
      itmX.SubItems(1) = Trim(rs!Cta)
      itmX.SubItems(2) = rs!Descripcion
      itmX.SubItems(3) = rs!Cod_Divisa
      itmX.SubItems(4) = rs!Entidad_Desc
      
      If rs!IdX > 0 Then
          itmX.Checked = vbChecked
          itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(pCodigo As String)
On Error GoTo vError

Me.MousePointer = vbHourglass

 strSQL = "select * " _
        & " from vSys_Formas_Pago" _
        & " where cod_forma_pago = '" & pCodigo & "'"
 

Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  txtCodigo.Text = rs!cod_forma_pago
  vCodigo = rs!cod_forma_pago
  
  'Asigna el Tipo de Formato de Documento (Refresca Chk's)
  cboTipo.Text = rs!Tipo_Desc
  
  txtDescripcion.Text = rs!Descripcion
  chkActiva.Value = rs!Activa
  chkEfectivo.Value = rs!Efectivo
  chkSaldosFavor.Value = rs!aplica_saldos_favor
  chkAplicaDep.Value = rs!APLICA_PARA_DEPOSITO
  
  txtCuentaDesc.Text = rs!Cuenta_Desc
  txtCuentaCod.Text = rs!Cuenta_Mask
    
  chkMaximoAplica.Value = rs!MAXIMO_APL
  txtMaximoMonto.Text = Format(rs!MAXIMO_MONTO, "Standard")
    
  chkOrigen.Value = rs!OR_APLICA
  chkOrigenAplDiario.Value = rs!OR_DIARIO_APL
  txtOrigenMontoDiario.Text = Format(rs!OR_DIARIO_MONTO, "Standard")
  chkOrigenAplMensual.Value = rs!OR_MENSUAL_APL
  txtOrigenMontoMensual.Text = Format(rs!OR_MENSUAL_MONTO, "Standard")
    
  Call sbLswLlena(rs!cod_forma_pago)

Else
  MsgBox "No se encontró registro verifique...", vbInformation
  txtCodigo.Text = ""
  txtCodigo.SetFocus
  Call sbLimpiaDatos
End If

Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        Call sbLimpiaDatos
        vEdita = False
        txtCodigo.SetFocus
       Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     Call sbToolBar(tlb, "activo")
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaDatos
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_forma_pago,descripcion from sif_formas_pago "
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

End Select


End Sub


Private Sub sbGuardar()
Dim vTipo As String

On Error GoTo vError

Select Case Mid(cboTipo.Text, 1, 2)
  Case "EF" 'Efectivo
    vTipo = "E"
  Case "DE" 'Deposito
    vTipo = "B"
  Case "TA" 'Tarjeta
    vTipo = "T"
  Case "CH" 'Cheque
    vTipo = "C"
  Case "DO" 'Documento
    vTipo = "D"
  Case "SA" 'Saldo a Favor
    vTipo = "S"
  Case "FO" 'Fondos
    vTipo = "F"

End Select

    
'                   MAXIMO_APL SmallInt
'               , MAXIMO_MONTO       DEC(18,2)
'               , OR_APLICA          SMALLINT
'               , OR_DIARIO_APL      SMALLINT
'               , OR_DIARIO_MONTO    DEC(18,2)
'               , OR_MENSUAL_APL     SMALLINT
'               , OR_MENSUAL_MONTO   DEC(18,2)
    

If fxExisteForma(txtCodigo.Text) Then
  strSQL = "update SIF_FORMAS_PAGO set descripcion = '" & UCase(Trim(txtDescripcion)) & "'" _
         & ", activa = " & chkActiva.Value & ", efectivo = " & chkEfectivo.Value & ", aplica_saldos_favor = " & chkSaldosFavor.Value & "" _
         & ", cod_cuenta = '" & fxgCntCuentaFormato(False, txtCuentaCod) & "', tipo = '" & vTipo _
         & "', aplica_para_deposito = " & chkAplicaDep.Value & ", MAXIMO_APL = " & chkMaximoAplica.Value & ", MAXIMO_MONTO = " & CCur(txtMaximoMonto.Text) _
         & ", OR_APLICA = " & chkOrigen.Value & ", OR_DIARIO_APL = " & chkOrigenAplDiario.Value & ", OR_DIARIO_MONTO = " & CCur(txtOrigenMontoDiario.Text) _
         & ", OR_MENSUAL_APL = " & chkOrigenAplMensual.Value & ", OR_MENSUAL_MONTO = " & CCur(txtOrigenMontoMensual.Text) _
         & "  where cod_forma_pago = '" & vCodigo & "' "
         
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Forma Pago: " & vCodigo)

Else
  vCodigo = txtCodigo.Text

   strSQL = "insert into SIF_FORMAS_PAGO(cod_forma_pago, descripcion, activa, efectivo, aplica_saldos_favor, cod_cuenta, tipo, aplica_para_deposito" _
          & ", MAXIMO_APL, MAXIMO_MONTO, OR_APLICA, OR_DIARIO_APL, OR_DIARIO_MONTO, OR_MENSUAL_APL, OR_MENSUAL_MONTO, REGISTRO_USUARIO, REGISTRO_FECHA)" _
          & " values('" & vCodigo & "','" & UCase(Trim(txtDescripcion)) & "', " & chkActiva.Value & ", " & chkEfectivo.Value _
          & ", " & chkSaldosFavor.Value & " ,'" & fxgCntCuentaFormato(False, txtCuentaCod) & "', '" & vTipo _
          & "'," & chkAplicaDep.Value & ", " & chkMaximoAplica.Value & ", " & CCur(txtMaximoMonto.Text) & ", " & chkOrigen.Value _
          & ", " & chkOrigenAplDiario.Value & ", " & CCur(txtOrigenMontoDiario.Text) & ", " & chkOrigenAplMensual.Value & ", " & CCur(txtOrigenMontoMensual.Text) _
          & ", '" & glogon.Usuario & " ',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Forma Pago: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Function fxValida()

fxValida = True

If Trim(txtCodigo) = "" Then fxValida = False
If Trim(txtDescripcion) = "" Then fxValida = False

If Not IsNumeric(txtMaximoMonto.Text) Then
    txtMaximoMonto.Text = 0
End If

If Not IsNumeric(txtOrigenMontoDiario.Text) Then
    txtOrigenMontoDiario.Text = 0
End If

If Not IsNumeric(txtOrigenMontoMensual.Text) Then
    txtOrigenMontoMensual.Text = 0
End If



End Function



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tcMain.Item(0).Selected = True
    txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_forma_pago"
   gBusquedas.Orden = "cod_forma_pago"
   gBusquedas.Consulta = "select cod_forma_pago,descripcion from sif_formas_pago"
   frmBusquedas.Show vbModal
   txtCodigo.SetFocus
   txtCodigo.Text = gBusquedas.Resultado
   txtDescripcion.SetFocus
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboTipo.SetFocus
End Sub


Private Sub txtMaximoMonto_GotFocus()
On Error GoTo vError
 txtMaximoMonto.Text = CCur(txtMaximoMonto.Text)
vError:
End Sub

Private Sub txtMaximoMonto_LostFocus()
On Error GoTo vError
 txtMaximoMonto.Text = Format(CCur(txtMaximoMonto.Text), "Standard")
vError:
End Sub


Private Sub txtOrigenMontoDiario_GotFocus()
On Error GoTo vError
 txtOrigenMontoDiario.Text = CCur(txtOrigenMontoDiario.Text)
vError:
End Sub

Private Sub txtOrigenMontoDiario_LostFocus()
On Error GoTo vError
 txtOrigenMontoDiario.Text = Format(CCur(txtOrigenMontoDiario.Text), "Standard")
vError:
End Sub


Private Sub txtOrigenMontoMensual_GotFocus()
On Error GoTo vError
 txtOrigenMontoMensual.Text = CCur(txtOrigenMontoMensual.Text)
vError:
End Sub

Private Sub txtOrigenMontoMensual_LostFocus()
On Error GoTo vError
 txtOrigenMontoMensual.Text = Format(CCur(txtOrigenMontoMensual.Text), "Standard")
vError:
End Sub
