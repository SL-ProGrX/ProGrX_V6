VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmTES_Accesos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Accesos"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   Icon            =   "frmTES_Accesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   10695
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   360
      Top             =   360
   End
   Begin XtremeSuiteControls.PushButton cmdAsigna 
      Height          =   252
      Left            =   9600
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1092
      _Version        =   1441792
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Accesos"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      CausesValidation=   0   'False
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   10695
      _Version        =   1441792
      _ExtentX        =   18865
      _ExtentY        =   11880
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
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   4
      Item(0).Caption =   "Cuentas"
      Item(0).Tooltip =   "Cuentas Bancarias"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "cbo"
      Item(0).Control(1)=   "lswUsers"
      Item(0).Control(2)=   "Label1(13)"
      Item(0).Control(3)=   "Label1(0)"
      Item(0).Control(4)=   "Label1(1)"
      Item(0).Control(5)=   "cboBanco"
      Item(0).Control(6)=   "chkTodos(0)"
      Item(1).Caption =   "Usuarios"
      Item(1).Tooltip =   "Acceso a Cuentas por Usuarios"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "txtUsuario"
      Item(1).Control(1)=   "txtNombre"
      Item(1).Control(2)=   "FlatScrollBar"
      Item(1).Control(3)=   "Label1(2)"
      Item(1).Control(4)=   "Label1(3)"
      Item(1).Control(5)=   "lswBancos"
      Item(1).Control(6)=   "Label1(14)"
      Item(1).Control(7)=   "cboBancoX"
      Item(1).Control(8)=   "chkTodos(1)"
      Item(2).Caption =   "Accesos"
      Item(2).Tooltip =   "Detalle de Accesos"
      Item(2).ControlCount=   9
      Item(2).Control(0)=   "txtAsgNombre"
      Item(2).Control(1)=   "txtAsgUsuario"
      Item(2).Control(2)=   "txtAsgBanco"
      Item(2).Control(3)=   "FlatScrollBarX2"
      Item(2).Control(4)=   "Label1(4)"
      Item(2).Control(5)=   "Label1(5)"
      Item(2).Control(6)=   "imgReporte"
      Item(2).Control(7)=   "lswAsgBancos"
      Item(2).Control(8)=   "tcMainAux"
      Item(3).Caption =   "Copia"
      Item(3).Tooltip =   "Copia de Accesos"
      Item(3).ControlCount=   13
      Item(3).Control(0)=   "txtcdUsuario"
      Item(3).Control(1)=   "txtcdNombre"
      Item(3).Control(2)=   "txtcoUsuario"
      Item(3).Control(3)=   "txtcoNombre"
      Item(3).Control(4)=   "FlatScrollBarC"
      Item(3).Control(5)=   "FlatScrollBarD"
      Item(3).Control(6)=   "Label1(10)"
      Item(3).Control(7)=   "Label1(7)"
      Item(3).Control(8)=   "Label1(6)"
      Item(3).Control(9)=   "Label1(9)"
      Item(3).Control(10)=   "Label1(8)"
      Item(3).Control(11)=   "GroupBox1"
      Item(3).Control(12)=   "GroupBox2"
      Begin XtremeSuiteControls.ListView lswBancos 
         Height          =   5292
         Left            =   -68200
         TabIndex        =   35
         Top             =   1320
         Visible         =   0   'False
         Width           =   7332
         _Version        =   1441792
         _ExtentX        =   12933
         _ExtentY        =   9334
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
      Begin XtremeSuiteControls.ListView lswUsers 
         Height          =   5292
         Left            =   2520
         TabIndex        =   34
         Top             =   1320
         Width           =   6372
         _Version        =   1441792
         _ExtentX        =   11239
         _ExtentY        =   9334
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
      Begin XtremeSuiteControls.ListView lswAsgBancos 
         Height          =   1572
         Left            =   -67840
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   7332
         _Version        =   1441792
         _ExtentX        =   12933
         _ExtentY        =   2773
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1092
         Left            =   -69640
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441792
         _ExtentX        =   16954
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Copiar accesos:"
         ForeColor       =   4210752
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton tlbAux 
            Height          =   372
            Left            =   7800
            TabIndex        =   29
            Top             =   360
            Width           =   1692
            _Version        =   1441792
            _ExtentX        =   2984
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Copia Accesos"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin VB.Image Image1 
            Height          =   372
            Index           =   0
            Left            =   360
            Picture         =   "frmTES_Accesos.frx":6852
            Stretch         =   -1  'True
            Top             =   360
            Width           =   372
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Advertencia: Todos los datos del usuario destino serán eliminados y se incluiran el esquema del usuario origen."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   492
            Index           =   11
            Left            =   960
            TabIndex        =   20
            Top             =   360
            Width           =   6012
         End
      End
      Begin XtremeSuiteControls.TabControl tcMainAux 
         Height          =   3492
         Left            =   -69160
         TabIndex        =   8
         Top             =   3000
         Visible         =   0   'False
         Width           =   8652
         _Version        =   1441792
         _ExtentX        =   15261
         _ExtentY        =   6159
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
         AllowReorder    =   -1  'True
         Appearance      =   4
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   4
         Item(0).Caption =   "Documentos"
         Item(0).Tooltip =   "Acciones sobre Tipos de Documentos"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "vGrid"
         Item(1).Caption =   "Conceptos"
         Item(1).Tooltip =   "Conceptos permitidos?"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswConceptos"
         Item(2).Caption =   "Unidades"
         Item(2).Tooltip =   "Unidades permitidas?"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "lswUnidades"
         Item(3).Caption =   "Firmas"
         Item(3).Tooltip =   "Autorizaciones de Firma Digital"
         Item(3).ControlCount=   7
         Item(3).Control(0)=   "txtRngFirmasHasta"
         Item(3).Control(1)=   "txtRngFirmasDesde"
         Item(3).Control(2)=   "chkFirmaRango"
         Item(3).Control(3)=   "chkUserFirma"
         Item(3).Control(4)=   "Label2(2)"
         Item(3).Control(5)=   "Label2(3)"
         Item(3).Control(6)=   "btnFirmas"
         Begin XtremeSuiteControls.ListView lswUnidades 
            Height          =   3090
            Left            =   -70000
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   8655
            _Version        =   1441792
            _ExtentX        =   15266
            _ExtentY        =   5450
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
            ShowBorder      =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswConceptos 
            Height          =   3090
            Left            =   -70000
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   8655
            _Version        =   1441792
            _ExtentX        =   15266
            _ExtentY        =   5450
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
            ShowBorder      =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkUserFirma 
            Height          =   372
            Left            =   -69400
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   7572
            _Version        =   1441792
            _ExtentX        =   13356
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Usuario Autorizado a Utilizar Firmas Electrónicas / en Cheques"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
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
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   2775
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   8055
            _Version        =   524288
            _ExtentX        =   14203
            _ExtentY        =   4890
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
            MaxCols         =   493
            ScrollBars      =   2
            SpreadDesigner  =   "frmTES_Accesos.frx":71B6
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton btnFirmas 
            Height          =   495
            Left            =   -63640
            TabIndex        =   39
            Top             =   2400
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441792
            _ExtentX        =   3413
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Actualiza Firmas"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.CheckBox chkFirmaRango 
            Height          =   372
            Left            =   -69040
            TabIndex        =   47
            Top             =   840
            Visible         =   0   'False
            Width           =   7452
            _Version        =   1441792
            _ExtentX        =   13144
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Utiliza un rango Específicio de Firmas (Diferente al del autorizado por Cuenta Bancaria)"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
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
         Begin XtremeSuiteControls.FlatEdit txtRngFirmasDesde 
            Height          =   312
            Left            =   -67360
            TabIndex        =   48
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1440
            Visible         =   0   'False
            Width           =   2772
            _Version        =   1441792
            _ExtentX        =   4890
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRngFirmasHasta 
            Height          =   312
            Left            =   -67360
            TabIndex        =   49
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1800
            Visible         =   0   'False
            Width           =   2772
            _Version        =   1441792
            _ExtentX        =   4890
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
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
            Left            =   -68200
            TabIndex        =   11
            Top             =   1440
            Visible         =   0   'False
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
            Left            =   -68200
            TabIndex        =   10
            Top             =   1800
            Visible         =   0   'False
            Width           =   732
         End
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarX2 
         Height          =   252
         Left            =   -60484
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarC 
         Height          =   252
         Left            =   -60760
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarD 
         Height          =   252
         Left            =   -60760
         TabIndex        =   13
         Top             =   2880
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1092
         Left            =   -69640
         TabIndex        =   21
         Top             =   5160
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441792
         _ExtentX        =   16954
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Elimina accesos:"
         ForeColor       =   4210752
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton tlbElimina 
            Height          =   372
            Left            =   7800
            TabIndex        =   28
            Top             =   360
            Width           =   1692
            _Version        =   1441792
            _ExtentX        =   2984
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Elimina Accesos"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   336
            Left            =   7680
            TabIndex        =   22
            Top             =   2760
            Width           =   1692
            _ExtentX        =   2990
            _ExtentY        =   556
            ButtonWidth     =   2709
            ButtonHeight    =   550
            Style           =   1
            TextAlignment   =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Copiar Accesos"
                  Key             =   "Copiar"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin VB.Image Image1 
            Height          =   372
            Index           =   1
            Left            =   360
            Picture         =   "frmTES_Accesos.frx":78CB
            Stretch         =   -1  'True
            Top             =   360
            Width           =   372
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Elimina todos los permisos asignados a usuarios que se encuentren su cuenta inactiva."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   492
            Index           =   15
            Left            =   960
            TabIndex        =   23
            Top             =   360
            Width           =   6012
         End
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   252
         Left            =   -60724
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   2520
         TabIndex        =   31
         Top             =   960
         Width           =   6372
         _Version        =   1441792
         _ExtentX        =   11245
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   2520
         TabIndex        =   32
         Top             =   600
         Width           =   6372
         _Version        =   1441792
         _ExtentX        =   11245
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
      Begin XtremeSuiteControls.ComboBox cboBancoX 
         Height          =   312
         Left            =   -68200
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   7332
         _Version        =   1441792
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtAsgNombre 
         Height          =   312
         Left            =   -66160
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1441792
         _ExtentX        =   9970
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsgUsuario 
         Height          =   312
         Left            =   -67840
         TabIndex        =   40
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441792
         _ExtentX        =   2984
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtcoUsuario 
         Height          =   312
         Left            =   -68200
         TabIndex        =   42
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441792
         _ExtentX        =   2984
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtcoNombre 
         Height          =   312
         Left            =   -66520
         TabIndex        =   43
         Top             =   1920
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1441792
         _ExtentX        =   9970
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtcdUsuario 
         Height          =   312
         Left            =   -68200
         TabIndex        =   44
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441792
         _ExtentX        =   2984
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtcdNombre 
         Height          =   312
         Left            =   -66520
         TabIndex        =   45
         Top             =   2880
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1441792
         _ExtentX        =   9970
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsgBanco 
         Height          =   312
         Left            =   -69160
         TabIndex        =   50
         Top             =   2640
         Visible         =   0   'False
         Width           =   8652
         _Version        =   1441792
         _ExtentX        =   15261
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   252
         Index           =   0
         Left            =   9000
         TabIndex        =   51
         Top             =   1320
         Width           =   1212
         _Version        =   1441792
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   -68200
         TabIndex        =   52
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441792
         _ExtentX        =   2984
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   -66520
         TabIndex        =   53
         Top             =   600
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1441792
         _ExtentX        =   9970
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   252
         Index           =   1
         Left            =   -60760
         TabIndex        =   54
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441792
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas"
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
         Index           =   14
         Left            =   -69400
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad"
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
         Index           =   3
         Left            =   -69400
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   -69400
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Origen:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   8
         Left            =   -69640
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Destino:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   9
         Left            =   -69640
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   -68920
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   -68920
         TabIndex        =   15
         Top             =   2880
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Copiar el Esquema de Accesos a Bancos y Documentos de un usuario a otro"
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
         Index           =   10
         Left            =   -70000
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   10695
      End
      Begin VB.Image imgReporte 
         Height          =   252
         Left            =   -59920
         Picture         =   "frmTES_Accesos.frx":8278
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas"
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
         Index           =   5
         Left            =   -69160
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Index           =   4
         Left            =   -69160
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios"
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
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas Bancarias"
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
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad"
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
         Index           =   13
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Width           =   1812
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control de Accesos a Cuentas Bancarias"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   12
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6612
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmTES_Accesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vScrollX2 As Boolean, vPaso As Boolean
Dim vScrollC As Boolean, vScrollD As Boolean

Private Sub sbLlenaLswBancos(vUsuario As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

vPaso = True

lswBancos.ListItems.Clear
chkTodos(0).Value = vbUnchecked

strSQL = "select B.id_banco,B.descripcion,B.cta,A.nombre" _
       & " from Tes_Bancos B left join tes_Banco_Asg A on B.id_banco = A.id_banco" _
       & " and A.nombre = '" & vUsuario & "'"
       
If cboBancoX.Text <> "TODOS" Then
   strSQL = strSQL & " Where B.cod_Grupo = '" & SIFGlobal.fxCodText(cboBancoX.Text) & "'"
End If
       
strSQL = strSQL & " order by A.nombre desc"

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lswBancos.ListItems.Add(, , rs!ID_BANCO)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Cta
     itmX.Checked = IIf(IsNull(rs!Nombre), False, True)
     itmX.ForeColor = IIf(IsNull(rs!Nombre), vbBlack, vbBlue)
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub
vError:

End Sub



Private Sub sbLlenaLswBancosAsg(vUsuario As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

vPaso = True

lswAsgBancos.ListItems.Clear
txtAsgBanco.Tag = ""
txtAsgBanco.Text = ""
vGrid.MaxCols = 6
vGrid.MaxRows = 0

tcMainAux.Item(0).Selected = True

strSQL = "select B.id_banco,B.descripcion,B.cta,A.nombre" _
       & " from Tes_Bancos B inner join tes_Banco_Asg A on B.id_banco = A.id_banco" _
       & " and A.nombre = '" & vUsuario & "' order by A.nombre desc"
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lswAsgBancos.ListItems.Add(, , rs!ID_BANCO)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Cta
 rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub
vError:

End Sub

Private Sub sbLlenaLswUsuarios(vBanco As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

vPaso = True

lswUsers.ListItems.Clear
chkTodos(1).Value = vbUnchecked

strSQL = "select U.nombre,U.descripcion,U.estado,A.id_banco" _
       & " from Usuarios U left join tes_Banco_Asg A on U.nombre = A.nombre" _
       & " and A.id_banco = " & cbo.ItemData(cbo.ListIndex) & " where U.estado = 'A' order by A.id_banco desc"
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lswUsers.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = IIf((rs!Estado = "A"), "Activo", "Inactivo")
     itmX.Checked = IIf(IsNull(rs!ID_BANCO), False, True)
     itmX.ForeColor = IIf(IsNull(rs!ID_BANCO), vbBlack, vbBlue)
    
 rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub
vError:

End Sub



Private Sub btnFirmas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vPaso Then Exit Sub

If Not IsNumeric(txtAsgBanco.Tag) Then Exit Sub

strSQL = "select count(*) as Existe from TES_BANCO_FIRMASAUT where id_banco = " & txtAsgBanco.Tag _
       & " and usuario = '" & txtAsgUsuario.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then
   strSQL = "insert TES_BANCO_FIRMASAUT(usuario,id_banco,UTILIZA_FIRMAS_AUTORIZA,APLICA_RANGO_AUTORIZACION" _
          & ",FIRMAS_AUTORIZA_INICIO,FIRMAS_AUTORIZA_CORTE) values('" & txtAsgUsuario _
          & "'," & txtAsgBanco.Tag & "," & chkUserFirma.Value & "," & chkFirmaRango.Value _
          & "," & CCur(txtRngFirmasDesde.Text) & "," & CCur(txtRngFirmasHasta.Text) & ")"
Else
  If chkUserFirma.Value = vbChecked Then
       strSQL = "update TES_BANCO_FIRMASAUT set UTILIZA_FIRMAS_AUTORIZA = " & chkUserFirma.Value _
              & ",APLICA_RANGO_AUTORIZACION = " & chkFirmaRango.Value _
              & ",FIRMAS_AUTORIZA_INICIO = " & CCur(txtRngFirmasDesde.Text) _
              & ",FIRMAS_AUTORIZA_CORTE = " & CCur(txtRngFirmasHasta.Text) _
              & " where usuario = '" & txtAsgUsuario & "' and id_banco = " & txtAsgBanco.Tag
  Else
       strSQL = "delete TES_BANCO_FIRMASAUT where usuario = '" & txtAsgUsuario & "' and id_banco = " & txtAsgBanco.Tag
  End If
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cbo_Click()
If vPaso Or cbo.ListCount = 0 Then Exit Sub
Call sbLlenaLswUsuarios(cbo.ItemData(cbo.ListIndex))
End Sub

Private Sub cboBanco_Click()
Dim strSQL As String

If vPaso Then Exit Sub

 strSQL = "select id_banco as 'IdX',rtrim(descripcion) as 'ItmX' from Tes_Bancos where estado = 'A'"
 If cboBanco.Text <> "TODOS" Then
    strSQL = strSQL & " and cod_grupo ='" & cboBanco.ItemData(cboBanco.ListIndex) & "'"
 End If
 
 vPaso = True
     Call sbCbo_Llena_New(cbo, strSQL, False, True)
 vPaso = False
 
 Call cbo_Click

End Sub

Private Sub cboBancoX_Click()

If vPaso Then Exit Sub
Call sbLlenaLswBancos(txtUsuario.Text)

End Sub

Private Sub chkFirmaRango_Click()
If chkUserFirma.Value = vbChecked Then
    txtRngFirmasDesde.Enabled = True
    txtRngFirmasHasta.Enabled = True
Else
    txtRngFirmasDesde.Enabled = False
    txtRngFirmasHasta.Enabled = False
End If

End Sub

Private Sub chkTodos_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

'
' Esta opcion se encuentra deshabilitada / No activar hasta revisar el proceso de borrado.
'

On Error GoTo vError


'Deshabilita
Exit Sub

Me.MousePointer = vbHourglass


Select Case Index
  Case 0  'Cuentas
      strSQL = "delete tes_Banco_Asg where id_banco = " & cbo.ItemData(cbo.ListIndex)
      Call ConectionExecute(strSQL)
  
      If chkTodos(Index).Value = vbChecked Then
        strSQL = "insert into tes_Banco_Asg(id_banco,nombre) " _
               & "(select " & cbo.ItemData(cbo.ListIndex) & ",nombre from usuarios where estado = 'A')"
        Call ConectionExecute(strSQL)
      End If
      
      For i = 1 To lswUsers.ListItems.Count
         lswUsers.ListItems.Item(i).Checked = chkTodos(Index).Value
      Next i
  
  
  Case 1  'Usuarios
      strSQL = "delete tes_Banco_Asg where nombre = '" & txtUsuario & "'"
      Call ConectionExecute(strSQL)
      
      If chkTodos(Index).Value = vbChecked Then
        strSQL = "insert into tes_Banco_Asg(id_banco,nombre) " _
               & "(select id_banco,'" & txtUsuario & "' from Tes_Bancos)"
        Call ConectionExecute(strSQL)
      End If
      
      For i = 1 To lswBancos.ListItems.Count
         lswBancos.ListItems.Item(i).Checked = chkTodos(Index).Value
      Next i
    
End Select

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkUserFirma_Click()

If chkUserFirma.Value = vbChecked Then
   chkFirmaRango.Enabled = True
Else
   chkFirmaRango.Enabled = False
End If
Call chkFirmaRango_Click
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select nombre,descripcion from usuarios"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where nombre > '" & txtUsuario & "' and estado = 'A' order by nombre asc"
    Else
       strSQL = strSQL & " where nombre < '" & txtUsuario & "' and estado = 'A' order by nombre desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbLlenaLswBancos(rs!Nombre)
      txtUsuario = rs!Nombre
      txtNombre = rs!Descripcion
    Else
      lswBancos.ListItems.Clear
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub


Private Sub FlatScrollBarX2_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScrollX2 Then
    strSQL = "select nombre,descripcion from usuarios"
    
    If FlatScrollBarX2.Value = 1 Then
       strSQL = strSQL & " where nombre > '" & txtAsgUsuario & "' and estado = 'A' order by nombre asc"
    Else
       strSQL = strSQL & " where nombre < '" & txtAsgUsuario & "' and estado = 'A' order by nombre desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbLlenaLswBancosAsg(rs!Nombre)
      txtAsgUsuario = rs!Nombre
      txtAsgNombre = rs!Descripcion
    Else
      lswAsgBancos.ListItems.Clear
    End If
    rs.Close
End If

vScrollX2 = False
FlatScrollBarX2.Value = 0
vScrollX2 = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarC_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScrollC Then
    strSQL = "select nombre,descripcion from usuarios"
    
    If FlatScrollBarC.Value = 1 Then
       strSQL = strSQL & " where nombre > '" & txtcoUsuario & "' and Estado = 'A' order by nombre asc"
    Else
       strSQL = strSQL & " where nombre < '" & txtcoUsuario & "' and Estado = 'A' order by nombre desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      txtcoUsuario = rs!Nombre
      txtcoNombre = rs!Descripcion
    End If
    rs.Close
End If

vScrollC = False
FlatScrollBarC.Value = 0
vScrollC = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub FlatScrollBarD_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScrollD Then
    strSQL = "select nombre,descripcion from usuarios"
    
    If FlatScrollBarD.Value = 1 Then
       strSQL = strSQL & " where nombre > '" & txtcdUsuario & "' and Estado = 'A' order by nombre asc"
    Else
       strSQL = strSQL & " where nombre < '" & txtcdUsuario & "' and Estado = 'A' order by nombre desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      txtcdUsuario = rs!Nombre
      txtcdNombre = rs!Descripcion
    End If
    rs.Close
End If

vScrollD = False
FlatScrollBarD.Value = 0
vScrollD = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
 
vModulo = 9
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

 tcMain.Item(0).Selected = True
 
 With lswUsers.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1830
    .Add , , "Nombre", 3500
    .Add , , "Estado", 1000, vbCenter
 End With
 
 
 With lswBancos.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "Descripción", 3800
    .Add , , "Cuenta", 2700, vbCenter
 End With
 
 With lswAsgBancos.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "Descripción", 3800
    .Add , , "Cuenta", 2700, vbCenter
 End With
 lswAsgBancos.Checkboxes = False
 
 With lswConceptos.ColumnHeaders
    .Clear
    .Add , , "Código", 1400
    .Add , , "Descripción", 6900
 End With
 
 With lswUnidades.ColumnHeaders
    .Clear
    .Add , , "Código", 1400
    .Add , , "Descripción", 6900
 End With
 
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 vScrollX2 = False
 FlatScrollBarX2.Value = 0
 vScrollX2 = True
 
 
 vScrollC = False
 FlatScrollBarC.Value = 0
 vScrollC = True
 
 vScrollD = False
 FlatScrollBarD.Value = 0
 vScrollD = True
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 
 'Asigna Seguridad
 lswBancos.Enabled = cmdAsigna.Enabled
 lswUsers.Enabled = cmdAsigna.Enabled
 vGrid.Enabled = cmdAsigna.Enabled
 lswConceptos.Enabled = cmdAsigna.Enabled
 lswUnidades.Enabled = cmdAsigna.Enabled
 tlbAux.Enabled = cmdAsigna.Enabled
 btnFirmas.Enabled = cmdAsigna.Enabled
 chkUserFirma.Enabled = cmdAsigna.Enabled
 
 
End Sub


Private Sub imgReporte_Click()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("Banking_ListadoAccesos.rpt")
    .SelectionFormula = "{TES_DOCUMENTOS_ASG.NOMBRE} = '" & txtAsgUsuario & "'"
    .PrintReport
End With


Me.MousePointer = vbDefault

End Sub



Private Sub sbCargaGridLocal(pGrid As Object, pGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

pGrid.MaxCols = pGridMaxCol
pGrid.MaxRows = 1

pGrid.Row = pGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  pGrid.Row = pGrid.MaxRows
  
  For i = 1 To pGrid.MaxCols
    pGrid.col = i
    Select Case i
     Case 1
        pGrid.Text = CStr(rs!Descripcion)
        pGrid.CellTag = CStr(rs!Tipo)
     Case 2
        pGrid.Value = rs!solicita
     Case 3
        pGrid.Value = rs!autoriza
     Case 4
        pGrid.Value = rs!Genera
     Case 5
        pGrid.Value = rs!Asientos
     Case 6
        pGrid.Value = rs!anula
    End Select
  
  Next i
  
  pGrid.MaxRows = pGrid.MaxRows + 1
  
  rs.MoveNext
Loop

rs.Close

pGrid.MaxRows = pGrid.MaxRows - 1

Me.MousePointer = vbDefault

End Sub




Private Sub lswAsgBancos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswAsgBancos.SortKey = ColumnHeader.Index - 1
  If lswAsgBancos.SortOrder = 0 Then lswAsgBancos.SortOrder = 1 Else lswAsgBancos.SortOrder = 0
  lswAsgBancos.Sorted = True
End Sub


Private Sub lswAsgBancos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

'Limpiar Componentes
If vPaso Then Exit Sub
If lswAsgBancos.ListItems.Count = 0 Then Exit Sub
If txtAsgUsuario = "" Then Exit Sub

tcMainAux.Item(0).Selected = True

txtAsgBanco.Tag = Item.Text
txtAsgBanco.Text = Trim(Item.SubItems(1)) & " ¦ Cuenta: " & Trim(Item.SubItems(2))

strSQL = "select T.Tipo,T.descripcion,isnull(A.Solicita,0) as Solicita,isnull(A.Autoriza,0) as Autoriza" _
       & ",isnull(A.Genera,0) as Genera,isnull(A.asientos,0) as Asientos,isnull(A.ANULA,0) as Anula" _
       & " from tes_tipos_doc T left join tes_documentos_asg A on T.tipo = A.tipo" _
       & " and A.id_banco = " & txtAsgBanco.Tag & " and A.nombre = '" & txtAsgUsuario.Text & "'" _
       & " Where T.tipo in(select Tipo from tes_banco_docs where id_banco = " & txtAsgBanco.Tag & ")"
vPaso = True
Call sbCargaGridLocal(vGrid, 6, strSQL)
vPaso = False

End Sub

Private Sub lswBancos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswBancos.SortKey = ColumnHeader.Index - 1
  If lswBancos.SortOrder = 0 Then lswBancos.SortOrder = 1 Else lswBancos.SortOrder = 0
  lswBancos.Sorted = True
End Sub

Private Sub lswBancos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Byte

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then

   strSQL = "insert tes_Banco_Asg(id_banco,nombre) values(" & Item.Text _
          & ",'" & txtUsuario & "')"
   Call ConectionExecute(strSQL)
   
Else
   
   strSQL = "Al quitar acceso a este Nivel, automáticamente todos los permisos relacionados como : " & vbCrLf _
          & "     - Documentos Asignados" & vbCrLf _
          & "     - Conceptos Asignados" & vbCrLf _
          & "     - Unidades Asignadas" & vbCrLf _
          & "     - Firmas Autorizadas" & vbCrLf _
          & " de este Usuario seran removidos, desea continuar ?"

   i = MsgBox(strSQL, vbYesNo)
   If i = vbYes Then
     Call sbEliminaAccesoBanco(txtUsuario, Item.Text)
   Else
    Item.Checked = vbChecked
   End If

End If


Exit Sub

vError:

 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 If Item.Checked Then
    Item.Checked = vbUnchecked
 Else
    Item.Checked = vbChecked
 End If
 

End Sub



Private Sub lswConceptos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswConceptos.SortKey = ColumnHeader.Index - 1
  If lswConceptos.SortOrder = 0 Then lswConceptos.SortOrder = 1 Else lswConceptos.SortOrder = 0
  lswConceptos.Sorted = True
End Sub

Private Sub lswConceptos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub


If Item.Checked Then
   strSQL = "insert tes_conceptos_asg(nombre,cod_concepto,id_banco) values('" & txtAsgUsuario _
          & "','" & Item.Text & "'," & txtAsgBanco.Tag & ")"
Else
   strSQL = "delete tes_conceptos_asg where nombre = '" & txtAsgUsuario & "' and cod_concepto = '" _
          & Item.Text & "' and id_banco = " & txtAsgBanco.Tag
End If
Call ConectionExecute(strSQL)

End Sub


Private Sub lswUnidades_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswUnidades.SortKey = ColumnHeader.Index - 1
  If lswUnidades.SortOrder = 0 Then lswUnidades.SortOrder = 1 Else lswUnidades.SortOrder = 0
  lswUnidades.Sorted = True
End Sub

Private Sub lswUnidades_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert tes_unidad_asg(nombre,cod_unidad,id_banco) values('" & txtAsgUsuario _
          & "','" & Item.Text & "'," & txtAsgBanco.Tag & ")"
Else
   strSQL = "delete tes_unidad_asg where nombre = '" & txtAsgUsuario & "' and cod_unidad = '" _
          & Item.Text & "' and id_banco = " & txtAsgBanco.Tag
End If
Call ConectionExecute(strSQL)

End Sub



Private Sub sbEliminaAccesoBanco(xUsuario As String, xBanco As Integer)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

   strSQL = "delete TES_DOCUMENTOS_ASG where id_banco = " & xBanco _
          & " and nombre = '" & xUsuario & "'"
   
   strSQL = strSQL & Space(10) & "delete TES_CONCEPTOS_ASG where id_banco = " & xBanco _
          & " and nombre = '" & xUsuario & "'"
   
   strSQL = strSQL & Space(10) & "delete TES_UNIDAD_ASG where id_banco = " & xBanco _
          & " and nombre = '" & xUsuario & "'"
   
   strSQL = strSQL & Space(10) & "delete TES_BANCO_FIRMASAUT where id_banco = " & xBanco _
          & " and usuario = '" & xUsuario & "'"
  
   
   strSQL = strSQL & Space(10) & "delete tes_Banco_Asg where id_banco = " & xBanco _
          & " and nombre = '" & xUsuario & "'"
   
   Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswUsers_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswUsers.SortKey = ColumnHeader.Index - 1
  If lswUsers.SortOrder = 0 Then lswUsers.SortOrder = 1 Else lswUsers.SortOrder = 0
  lswUsers.Sorted = True
End Sub

Private Sub lswUsers_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Byte

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert tes_Banco_Asg(id_banco,nombre) values(" & cbo.ItemData(cbo.ListIndex) _
          & ",'" & Item.Text & "')"
   Call ConectionExecute(strSQL)

Else
   strSQL = "Al quitar acceso a este Nivel, automáticamente todos los permisos relacionados como : " & vbCrLf _
          & "     - Documentos Asignados" & vbCrLf _
          & "     - Conceptos Asignados" & vbCrLf _
          & "     - Unidades Asignadas" & vbCrLf _
          & "     - Firmas Autorizadas" & vbCrLf _
          & " de este Usuario seran removidos, desea continuar ?"

   i = MsgBox(strSQL, vbYesNo)
   If i = vbYes Then
     Call sbEliminaAccesoBanco(Item.Text, cbo.ItemData(cbo.ListIndex))
   Else
     Item.Checked = vbChecked
   End If
   
         
End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 If Item.Checked Then
    Item.Checked = vbUnchecked
 Else
    Item.Checked = vbChecked
 End If
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Select Case Item.Index
  Case 0
    If cbo.ListCount > 0 Then
        Call sbLlenaLswUsuarios(cbo.ItemData(cbo.ListIndex))
    End If
    
  Case 1
    Call sbLlenaLswBancos(txtUsuario)
  
  Case 2 'Accesos + Detalles
     tcMainAux.Item(0).Selected = True
     vGrid.MaxRows = False
  Case 3 'Copia
    txtcoUsuario.SetFocus
End Select

vError:
End Sub

Private Sub tcMainAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError


'Inicializa
vGrid.MaxRows = 0
lswConceptos.ListItems.Clear
lswUnidades.ListItems.Clear
chkUserFirma.Value = vbUnchecked
chkFirmaRango.Value = vbUnchecked
txtRngFirmasDesde.Text = "0.00"
txtRngFirmasHasta.Text = "0.00"

If txtAsgBanco.Tag = "" Then Exit Sub
 
vPaso = True
 
Select Case Item.Index
  Case 1 'Conceptos
       strSQL = "select C.cod_concepto,C.descripcion,A.id_Banco" _
              & " from tes_conceptos C left join tes_conceptos_asg A on C.cod_concepto = A.cod_concepto" _
              & " and A.id_banco = " & txtAsgBanco.Tag & " and A.nombre = '" & txtAsgUsuario _
              & "' Order by A.id_Banco desc"
       
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
        Set itmX = lswConceptos.ListItems.Add(, , rs!cod_Concepto)
            itmX.SubItems(1) = rs!Descripcion
        If Not IsNull(rs!ID_BANCO) Then
            itmX.Checked = True
            itmX.ForeColor = vbBlue
        End If
        rs.MoveNext
       Loop
       rs.Close
  Case 2 'Unidades
  
       strSQL = "select U.cod_unidad,U.descripcion,A.id_Banco" _
              & " from CntX_Unidades U left join tes_unidad_asg A on U.cod_unidad = A.cod_unidad" _
              & " and A.id_banco = " & txtAsgBanco.Tag & " and A.nombre = '" & txtAsgUsuario _
              & "' where U.cod_Contabilidad = " & GLOBALES.gEnlace & " Order by A.id_Banco desc"
       
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
        Set itmX = lswUnidades.ListItems.Add(, , rs!Cod_Unidad)
            itmX.SubItems(1) = rs!Descripcion
        If Not IsNull(rs!ID_BANCO) Then
            itmX.Checked = True
            itmX.ForeColor = vbBlue
        End If
        rs.MoveNext
       Loop
       rs.Close
  
  Case 3 'firmas
       
       strSQL = "select * from TES_BANCO_FIRMASAUT" _
              & " where id_banco = " & txtAsgBanco.Tag & " and usuario = '" & txtAsgUsuario & "'"

       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          chkUserFirma.Value = rs!UTILIZA_FIRMAS_AUTORIZA
          chkFirmaRango.Value = rs!APLICA_RANGO_AUTORIZACION
          txtRngFirmasDesde.Text = Format(rs!firmas_autoriza_inicio, "Standard")
          txtRngFirmasHasta.Text = Format(rs!firmas_autoriza_corte, "Standard")
       End If
       rs.Close
       
       Call chkUserFirma_Click
End Select


vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

 vPaso = True
 
     strSQL = "select rtrim(cod_grupo) as  'IdX', rtrim(Descripcion) as 'ItmX' from TES_BANCOS_GRUPOS" _
           & " where Activo = 1"
    Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
    Call sbCbo_Llena_New(cboBancoX, strSQL, True, True)
 
 vPaso = False
 
 Call cboBanco_Click

End Sub

Private Sub tlbAux_Click()
Dim i As Byte, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

'Borra datos Actuales
strSQL = "delete tes_documentos_asg where nombre = '" & txtcdUsuario & "'"

strSQL = strSQL & Space(10) & "delete tes_unidad_Asg where nombre = '" & txtcdUsuario & "'"

strSQL = strSQL & Space(10) & "delete tes_conceptos_Asg where nombre = '" & txtcdUsuario & "'"

strSQL = strSQL & Space(10) & "delete tes_banco_Asg where nombre = '" & txtcdUsuario & "'"

strSQL = strSQL & Space(10) & "delete TES_BANCO_FIRMASAUT where usuario = '" & txtcdUsuario & "'"


'Copiar Datos
strSQL = strSQL & Space(10) & "insert into tes_banco_asg(id_banco,nombre) (select id_banco,'" & txtcdUsuario _
       & "' from tes_banco_asg where nombre = '" & txtcoUsuario & "')"

strSQL = strSQL & Space(10) & "insert into TES_BANCO_FIRMASAUT(id_banco,usuario,UTILIZA_FIRMAS_AUTORIZA,APLICA_RANGO_AUTORIZACION" _
       & ",FIRMAS_AUTORIZA_INICIO,FIRMAS_AUTORIZA_CORTE) (select id_banco,'" & txtcdUsuario _
       & "',UTILIZA_FIRMAS_AUTORIZA,APLICA_RANGO_AUTORIZACION" _
       & ",FIRMAS_AUTORIZA_INICIO,FIRMAS_AUTORIZA_CORTE from TES_BANCO_FIRMASAUT where usuario = '" & txtcoUsuario & "')"

strSQL = strSQL & Space(10) & "insert into tes_documentos_asg(nombre,tipo,id_banco,solicita,autoriza,genera,asientos,anula)" _
       & "(select '" & txtcdUsuario & "',tipo,id_banco,solicita,autoriza,genera,asientos,anula" _
       & " from tes_documentos_asg where nombre = '" & txtcoUsuario & "')"

strSQL = strSQL & Space(10) & "insert into tes_conceptos_asg(nombre,id_banco,cod_concepto) (select '" & txtcdUsuario _
       & "',id_banco,cod_concepto from tes_conceptos_asg where nombre = '" & txtcoUsuario & "')"

strSQL = strSQL & Space(10) & "insert into tes_unidad_asg(nombre,id_banco,cod_unidad) (select '" & txtcdUsuario _
       & "',id_banco,cod_unidad from tes_unidad_asg where nombre = '" & txtcoUsuario & "')"

'Aplica el Lote Completo
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Call Bitacora("Aplica", "Copia de Accesos a Cuentas del usuario " & txtcoUsuario.Text & " a " & txtcdUsuario.Text)

MsgBox "Accesos a Operaciones sobre cuentas Bancarios copiado satisfactoriamente de " & txtcoUsuario & " a " & txtcdUsuario, vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbElimina_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

   strSQL = "delete TES_DOCUMENTOS_ASG where nombre in (select nombre from usuarios where estado = 'I')"
   
   strSQL = strSQL & Space(10) & "delete TES_CONCEPTOS_ASG where nombre in (select nombre from usuarios where estado = 'I')"
   
   strSQL = strSQL & Space(10) & "delete TES_UNIDAD_ASG where nombre in (select nombre from usuarios where estado = 'I')"
   
   strSQL = strSQL & Space(10) & "delete TES_BANCO_FIRMASAUT where usuario in (select nombre from usuarios where estado = 'I')"
   
   strSQL = strSQL & Space(10) & "delete tes_Banco_Asg where nombre in (select nombre from usuarios where estado = 'I')"
      
   'Aplica el Lote
   Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Revisión de actualización terminada satisfactoriamente, se eliminaron todos los permisos de usuarios inactivos.!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtAsgUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "nombre"
 gBusquedas.Orden = "nombre"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = "  and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtAsgUsuario.Text = gBusquedas.Resultado
 txtAsgNombre.Text = gBusquedas.Resultado2
 Call sbLlenaLswBancosAsg(txtAsgUsuario.Text)
End If
End Sub


Private Sub txtAsgNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "Descripcion"
 gBusquedas.Orden = "Descripcion"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = "  and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtAsgUsuario.Text = gBusquedas.Resultado
 txtAsgNombre.Text = gBusquedas.Resultado2
 Call sbLlenaLswBancosAsg(txtAsgUsuario.Text)
End If

End Sub


Private Sub txtcoUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "nombre"
 gBusquedas.Orden = "nombre"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = " and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtcoUsuario = gBusquedas.Resultado
 txtcoNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtcoNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "Descripcion"
 gBusquedas.Orden = "Descripcion"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = " and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtcoUsuario = gBusquedas.Resultado
 txtcoNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtcdUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "nombre"
 gBusquedas.Orden = "nombre"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = "  and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtcdUsuario = gBusquedas.Resultado
 txtcdNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtcdNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "Descripcion"
 gBusquedas.Orden = "Descripcion"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = "  and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtcdUsuario = gBusquedas.Resultado
 txtcdNombre = gBusquedas.Resultado2
End If

End Sub






Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "Descripcion"
 gBusquedas.Orden = "Descripcion"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = "  and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtUsuario = gBusquedas.Resultado
 txtNombre = gBusquedas.Resultado2
 Call sbLlenaLswBancos(txtUsuario)
End If

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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkUserFirma.SetFocus
End Sub

Private Sub txtRngFirmasHasta_LostFocus()
On Error GoTo vError
  txtRngFirmasHasta.Text = Format(CCur(txtRngFirmasHasta.Text), "Standard")
vError:
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "nombre"
 gBusquedas.Orden = "nombre"
 gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
 gBusquedas.Filtro = " and Estado = 'A'"
 frmBusquedas.Show vbModal
 txtUsuario = gBusquedas.Resultado
 txtNombre = gBusquedas.Resultado2
 Call sbLlenaLswBancos(txtUsuario)
End If
End Sub



Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtAsgBanco.Tag = "" Or vPaso Then Exit Sub

If col > 1 Then
  vGrid.Row = Row
  vGrid.col = 1
  strSQL = "select isnull(count(*),0) as Existe from tes_documentos_asg" _
         & " where nombre = '" & txtAsgUsuario & "' and id_banco = " & txtAsgBanco.Tag _
         & " and Tipo = '" & vGrid.CellTag & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
    strSQL = "insert tes_documentos_asg(nombre,id_banco,tipo,solicita,autoriza,genera,asientos,anula) values('" _
           & txtAsgUsuario & "'," & txtAsgBanco.Tag & ",'" & vGrid.CellTag & "',0,0,0,0,0)"
    Call ConectionExecute(strSQL)
  End If
  rs.Close
    
  vGrid.col = col
  vGrid.Row = Row
  Select Case col
    Case 2 'Solicita
        strSQL = "update tes_documentos_asg set solicita = " & vGrid.Value
    Case 3 'Autoriza
        strSQL = "update tes_documentos_asg set Autoriza = " & vGrid.Value
    Case 4 'Genera
        strSQL = "update tes_documentos_asg set Genera = " & vGrid.Value
    Case 5 'Asientos
        strSQL = "update tes_documentos_asg set Asientos = " & vGrid.Value
    Case 6 'Anulaciones
        strSQL = "update tes_documentos_asg set Anula = " & vGrid.Value
  End Select
  vGrid.col = 1
  strSQL = strSQL & " where nombre = '" & txtAsgUsuario & "' and id_banco = " & txtAsgBanco.Tag _
         & "  and Tipo = '" & vGrid.CellTag & "'"
  Call ConectionExecute(strSQL)
End If


Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

