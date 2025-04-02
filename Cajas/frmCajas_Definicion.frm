VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCajas_Definicion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de cajas"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   10605
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   480
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activa?"
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6855
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   10575
      _Version        =   1441793
      _ExtentX        =   18653
      _ExtentY        =   12091
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
      ItemCount       =   8
      Item(0).Caption =   "Definición"
      Item(0).ControlCount=   15
      Item(0).Control(0)=   "Label1(5)"
      Item(0).Control(1)=   "Label1(9)"
      Item(0).Control(2)=   "Label1(10)"
      Item(0).Control(3)=   "Label1(11)"
      Item(0).Control(4)=   "Label1(12)"
      Item(0).Control(5)=   "Label1(15)"
      Item(0).Control(6)=   "txtDescripcion"
      Item(0).Control(7)=   "chkAperturaCompartida"
      Item(0).Control(8)=   "cboPeriodoCierre"
      Item(0).Control(9)=   "cboTipoCierre"
      Item(0).Control(10)=   "txtNotas"
      Item(0).Control(11)=   "txtContrasenaVence"
      Item(0).Control(12)=   "dtpFechaApertura"
      Item(0).Control(13)=   "GroupBox1"
      Item(0).Control(14)=   "GroupBox2"
      Item(1).Caption =   "Saldos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "vGridRangosDivisa"
      Item(1).Control(1)=   "Label2(4)"
      Item(2).Caption =   "Servicios"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "ArbolRecaudador"
      Item(2).Control(1)=   "Label2(6)"
      Item(3).Caption =   "Auxiliares"
      Item(3).ControlCount=   10
      Item(3).Control(0)=   "chkLimita_Patrimonio"
      Item(3).Control(1)=   "chkLimita_Consulta"
      Item(3).Control(2)=   "chkLimita_CxC"
      Item(3).Control(3)=   "chkLimita_Fondos"
      Item(3).Control(4)=   "optLimita(2)"
      Item(3).Control(5)=   "optLimita(1)"
      Item(3).Control(6)=   "chkLimita_Creditos"
      Item(3).Control(7)=   "optLimita(0)"
      Item(3).Control(8)=   "lswAuxiliares"
      Item(3).Control(9)=   "Label2(3)"
      Item(4).Caption =   "Formas de Pago"
      Item(4).ControlCount=   2
      Item(4).Control(0)=   "lswFormasPago"
      Item(4).Control(1)=   "Label2(0)"
      Item(5).Caption =   "Documentos"
      Item(5).ControlCount=   2
      Item(5).Control(0)=   "lswDocumentos"
      Item(5).Control(1)=   "Label2(2)"
      Item(6).Caption =   "Usuarios"
      Item(6).ControlCount=   2
      Item(6).Control(0)=   "vGrid"
      Item(6).Control(1)=   "btnUsuario(1)"
      Item(7).Caption =   "Copia"
      Item(7).ControlCount=   2
      Item(7).Control(0)=   "ShortcutCaption2"
      Item(7).Control(1)=   "GroupBox3"
      Begin XtremeSuiteControls.ListView lswAuxiliares 
         Height          =   5775
         Left            =   -66640
         TabIndex        =   40
         Top             =   1080
         Visible         =   0   'False
         Width           =   6735
         _Version        =   1441793
         _ExtentX        =   11880
         _ExtentY        =   10186
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswFormasPago 
         Height          =   6255
         Left            =   -67480
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   11033
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswDocumentos 
         Height          =   6255
         Left            =   -67480
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   11033
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   2292
         Left            =   -69760
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   9972
         _Version        =   1441793
         _ExtentX        =   17590
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Información de la Caja destino.: "
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.FlatEdit txtCajaDestino 
            Height          =   330
            Left            =   1428
            TabIndex        =   51
            Top             =   480
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
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
         Begin XtremeSuiteControls.FlatEdit txtCajaDestinoNombre 
            Height          =   330
            Left            =   3120
            TabIndex        =   52
            Top             =   480
            Width           =   6732
            _Version        =   1441793
            _ExtentX        =   11874
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCopiar 
            Height          =   492
            Index           =   0
            Left            =   7200
            TabIndex        =   53
            Top             =   1320
            Width           =   1344
            _Version        =   1441793
            _ExtentX        =   2371
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Copiar"
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
            Picture         =   "frmCajas_Definicion.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnCopiar 
            Height          =   492
            Index           =   1
            Left            =   8520
            TabIndex        =   54
            Top             =   1320
            Width           =   1344
            _Version        =   1441793
            _ExtentX        =   2371
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Cerrar"
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
            Picture         =   "frmCajas_Definicion.frx":06F0
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Caja"
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
            Index           =   3
            Left            =   840
            TabIndex        =   50
            Top             =   480
            Width           =   552
         End
      End
      Begin VB.OptionButton optLimita 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Créditos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   -66640
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   600
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton optLimita 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fondos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   -64360
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.OptionButton optLimita 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   -62200
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   975
         Left            =   360
         TabIndex        =   26
         Top             =   5880
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Cuenta Contable"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaDev 
            Height          =   330
            Left            =   1680
            TabIndex        =   27
            Top             =   600
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDescCuentaDev 
            Height          =   330
            Left            =   3480
            TabIndex        =   28
            Top             =   600
            Width           =   6015
            _Version        =   1441793
            _ExtentX        =   10610
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkCtaCaja 
            Height          =   255
            Left            =   1680
            TabIndex        =   56
            Top             =   240
            Width           =   6855
            _Version        =   1441793
            _ExtentX        =   12091
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Utiliza Cuenta de la Caja en lugar de la Forma de Pago en Efectivo"
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
            UseVisualStyle  =   -1  'True
            TextAlignment   =   4
            Appearance      =   16
            Value           =   1
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   330
         Left            =   2040
         TabIndex        =   14
         Top             =   960
         Width           =   7815
         _Version        =   1441793
         _ExtentX        =   13785
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkAperturaCompartida 
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   6975
         _Version        =   1441793
         _ExtentX        =   12303
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Utiliza Apertura/Cierre de caja compartido entre los usuarios asignados?  "
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
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboPeriodoCierre 
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.ComboBox cboTipoCierre 
         Height          =   330
         Left            =   7200
         TabIndex        =   17
         Top             =   1800
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   675
         Left            =   2040
         TabIndex        =   18
         Top             =   2280
         Width           =   7815
         _Version        =   1441793
         _ExtentX        =   13785
         _ExtentY        =   1191
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtContrasenaVence 
         Height          =   315
         Left            =   8880
         TabIndex        =   19
         Top             =   1440
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaApertura 
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2775
         Left            =   360
         TabIndex        =   21
         Top             =   3120
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   4895
         _StockProps     =   79
         Caption         =   "Definición de la Oficina de la Caja y Autorización de Cobros Judiciales"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.CheckBox chkMovCbrJud 
            Height          =   252
            Left            =   1680
            TabIndex        =   22
            Top             =   480
            Width           =   6372
            _Version        =   1441793
            _ExtentX        =   11239
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Esta Caja realiza movimientos a operaciones en Cobro Judicial?"
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
            UseVisualStyle  =   -1  'True
            TextAlignment   =   4
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkUtilizaUsuario 
            Height          =   255
            Left            =   1680
            TabIndex        =   23
            Top             =   1920
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11239
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Utiliza la Oficina asignada al Usuario?"
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
            UseVisualStyle  =   -1  'True
            TextAlignment   =   4
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.ComboBox cboOficina 
            Height          =   330
            Left            =   1680
            TabIndex        =   24
            Top             =   2280
            Width           =   7815
            _Version        =   1441793
            _ExtentX        =   13785
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
         Begin XtremeSuiteControls.CheckBox chkPermiteRC 
            Height          =   255
            Left            =   1680
            TabIndex        =   55
            Top             =   840
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11239
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar Retiros de Efectivo?"
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
            UseVisualStyle  =   -1  'True
            TextAlignment   =   4
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkBoveda 
            Height          =   255
            Left            =   1680
            TabIndex        =   57
            Top             =   1560
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11239
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Rol de Boveda (Aprovisionamientos - Reintegros) ?"
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
            UseVisualStyle  =   -1  'True
            TextAlignment   =   4
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkTrasladoEfectivo 
            Height          =   255
            Left            =   1680
            TabIndex        =   58
            Top             =   1200
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11239
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite Traslado de Efectivo entre Cajas ?"
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
            UseVisualStyle  =   -1  'True
            TextAlignment   =   4
            Appearance      =   16
            Value           =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Oficina"
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
            Left            =   720
            TabIndex        =   25
            Top             =   2280
            Width           =   975
         End
      End
      Begin FPSpreadADO.fpSpread vGridRangosDivisa 
         Height          =   5655
         Left            =   -69280
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   9375
         _Version        =   524288
         _ExtentX        =   16536
         _ExtentY        =   9975
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmCajas_Definicion.frx":0E06
         VisibleRows     =   1
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.TreeView ArbolRecaudador 
         Height          =   5985
         Left            =   -68200
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   10557
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         ImageList       =   "imgExplorer"
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
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5895
         Left            =   -70120
         TabIndex        =   39
         Top             =   900
         Visible         =   0   'False
         Width           =   10335
         _Version        =   524288
         _ExtentX        =   18230
         _ExtentY        =   10398
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
         MaxCols         =   5
         SpreadDesigner  =   "frmCajas_Definicion.frx":1575
         VisibleRows     =   1
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.CheckBox chkLimita_Creditos 
         Height          =   252
         Left            =   -69640
         TabIndex        =   43
         Top             =   1560
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Limita Líneas de Créditos"
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
      Begin XtremeSuiteControls.CheckBox chkLimita_Fondos 
         Height          =   252
         Left            =   -69640
         TabIndex        =   44
         Top             =   2040
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Limita Fondos (Planes)"
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
      Begin XtremeSuiteControls.CheckBox chkLimita_CxC 
         Height          =   492
         Left            =   -69640
         TabIndex        =   45
         Top             =   2520
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Limita Conceptos de Cuentas por Cobrar"
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
      Begin XtremeSuiteControls.CheckBox chkLimita_Patrimonio 
         Height          =   252
         Left            =   -69640
         TabIndex        =   46
         Top             =   3240
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Limita Rubros de Patrimonio"
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
      Begin XtremeSuiteControls.CheckBox chkLimita_Consulta 
         Height          =   252
         Left            =   -69640
         TabIndex        =   47
         Top             =   3720
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Filtros en Consultas "
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
      Begin XtremeSuiteControls.PushButton btnUsuario 
         Height          =   420
         Index           =   1
         Left            =   -63520
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         _Version        =   1441793
         _ExtentX        =   5953
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Asignar Usuarios a esta Caja"
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
         Picture         =   "frmCajas_Definicion.frx":1C1D
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   495
         Left            =   -70000
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   873
         _StockProps     =   14
         Caption         =   "Copiar la caja actual en una nueva:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
         Alignment       =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione los Tipos de Documentos asignados a esta caja "
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
         Height          =   1056
         Index           =   2
         Left            =   -69760
         TabIndex        =   38
         Top             =   660
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione las formas de pago que esta caja puede registrar"
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
         Height          =   1056
         Index           =   0
         Left            =   -69760
         TabIndex        =   37
         Top             =   660
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione las líneas, planes y conceptos de los diferentes auxiliares que desea filtrar para esta caja.:"
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
         Height          =   936
         Index           =   3
         Left            =   -69880
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Indique los servicios que esta caja tiene autorizados!"
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
         Height          =   1296
         Index           =   6
         Left            =   -69760
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Indicar los rangos de montos en documentos y efectivo que estan autorizados en Cajas por cierres mínimos y máximos por divisa"
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
         Height          =   456
         Index           =   4
         Left            =   -69880
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   7920
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
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha apertura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Cierre"
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
         Left            =   5520
         TabIndex        =   9
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo Cierre"
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
         Index           =   10
         Left            =   360
         TabIndex        =   8
         Top             =   1830
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tiempo de Renovación de Contraseña de Cajas [En días]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   4215
         TabIndex        =   7
         Top             =   1440
         Width           =   4575
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
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9240
      Top             =   600
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3315
      TabIndex        =   0
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgExplorer"
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
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   9600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Definicion.frx":233D
            Key             =   "imgDocu"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Definicion.frx":3217
            Key             =   "imgFormu"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Definicion.frx":9A79
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8085
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario de Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Fecha de Registro"
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   450
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   794
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
   Begin VB.Image imgCopia 
      Height          =   240
      Left            =   3960
      Picture         =   "frmCajas_Definicion.frx":9B75
      ToolTipText     =   "Copiar Configuración de Cajas"
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   255
      TabIndex        =   1
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "frmCajas_Definicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vPaso As Boolean

Dim vScroll As Boolean
Dim vDocumento As String
Dim vRecaudador As String, vServicio As String, iNodo As Long, vNode As Node
Dim vInicial As Boolean

Private Sub ArbolBancos_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim vCodigo As String
Dim strSQL As String


'If Node.Text = "Todos" Then
' Call sbAsignaTodos
' Exit Sub
'End If
'
vCodigo = fxIndiceCodigo(Node.Key)


On Error GoTo vError


If Node.Checked Then
   strSQL = "insert into cajas_bancos(cod_bancos,cod_caja,registro_usuario,registro_fecha,activo) values('" & vCodigo _
            & "','" & txtCodigo & "','" & glogon.Usuario & "',dbo.MyGetdate(),1)"
   Node.ForeColor = vbBlue
Else
   strSQL = "Delete cajas_bancos where cod_bancos ='" & vCodigo _
          & "' and cod_caja = '" & txtCodigo.Text & "'"
   Node.ForeColor = vbRed
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description)
      

End Sub



Private Sub cmdNuevo_Click()
End Sub

'Private Sub sbCargaNodos()
'Dim strSQL As String, rs As New ADODB.Recordset
'With ArbolCajas
'  .Nodes.Clear
'  'Crear Root
'  Set xNode = vArbol.Nodes.Add(, , "US", "Todas")
'  xNode.Bold = True
'
'  strSQL = "select cod_caja,descripcion from cajas_definicion" _
'           & " where cod_caja not in (select caja from cajas_usuarios)order by cod_caja"
'
'
'  Call OpenRecordSet(rs, strSQL)
'  Do While Not rs.EOF
'   Call sbCreaNodos("US", rs!descripcion, "", True, "0x0" & rs!cod_caja & "P")
'    .Nodes(.Nodes.Count).Expanded = True
'    rs.MoveNext
'  Loop
'  rs.Close
'
'   xNode.Expanded = True
'
'End With
'
'
'Me.MousePointer = vbDefault
'
'End Sub

Private Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, varbol As TreeView, Optional xkey As String = "N")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = varbol.Nodes.Add(vPadre, tvwChild)
    nodX.Image = vImagen
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & varbol.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
Set vNode = nodX
End Sub
Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function

Private Sub ArbolDocumentos_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Image = "imgDocu" Then
    iNodo = Node.Index
    vDocumento = fxIndiceMultiple(Node.Key, "D")
    vModulo = fxIndiceCodigo(Node.Parent.Key)
     'Call sbCargaDatosLsw
End If
End Sub


Private Sub ArbolRecaudador_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim vCodigo As String
Dim strSQL As String


On Error GoTo vError


If Node.Image = "imgDocu" Then
    vServicio = fxIndiceMultiple(Node.Key, "D")
    vRecaudador = fxIndiceCodigo(Node.Parent.Key)
   
    If Node.Checked Then
       strSQL = "insert into cajas_servicios_asignados(cod_recaudador,cod_caja,cod_servicio,registro_usuario,registro_fecha) values('" & vRecaudador _
              & "','" & txtCodigo & "','" & vServicio & "','" & glogon.Usuario & "',dbo.MyGetdate())"
       Node.ForeColor = vbBlue
    Else
       strSQL = "Delete cajas_servicios_asignados where cod_caja ='" & txtCodigo _
              & "' and cod_servicio = '" & vServicio & "' and cod_recaudador = '" & vRecaudador & "'"
       Node.ForeColor = vbRed
    End If
    Call ConectionExecute(strSQL)

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description)
      
      

End Sub


Private Sub btnCopiar_Click(Index As Integer)
Dim strSQL As String

If Index = 1 Then
  tcMain.Item(0).Selected = True
  Exit Sub
End If

Me.MousePointer = vbHourglass

On Error GoTo vError

'Inserta linea
If Trim(txtCajaDestino.Text) <> "" And Trim(txtCajaDestinoNombre.Text) <> "" Then
      strSQL = "exec spCaja_Copia '" & txtCodigo.Text & "','" & txtCajaDestino.Text & "','" & glogon.Usuario _
             & "','" & txtCajaDestinoNombre.Text & "'"
      Call ConectionExecute(strSQL)

      Call Bitacora("Aplica", "Copia Caja.:" & txtCodigo.Text & " a " & txtCajaDestino.Text)

End If

Me.MousePointer = vbDefault
MsgBox "Copia Realizada Satisfactoriamente...", vbInformation


Call sbConsulta(txtCajaDestino.Text)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnUsuario_Click(Index As Integer)
Call sbFormsCall("frmCajas_Usuarios", 1, , , False, Me)

Dim strSQL As String

strSQL = "select rtrim(usuario) as 'Usuario',registro_fecha, rtrim(registro_usuario) as 'registro_usuario',salida_fecha,salida_usuario from cajas_usuarios_h " _
    & "where cod_caja = '" & vCodigo & "' order by usuario"
Call sbCargaGrid(vGrid, 5, strSQL)
End Sub

Private Sub cboPeriodoCierre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboTipoCierre.SetFocus
End Sub

Private Sub cboTipoCierre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub chkUtilizaUsuario_Click()
If chkUtilizaUsuario.Value = 1 Then
   cboOficina.Enabled = False
Else
   cboOficina.Enabled = True
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

tcMain.Item(0).Selected = True


If vScroll Then
' If txtCodigo = "" Then txtCodigo = 0
    strSQL = "select Top 1 cod_caja from cajas_definicion"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_caja > '" & txtCodigo.Text & "' order by cod_caja asc"
    Else
       strSQL = strSQL & " where cod_caja < '" & txtCodigo.Text & "' order by cod_caja desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!Cod_Caja)
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
 
 vEdita = True

tcMain.Item(7).Visible = False


cboPeriodoCierre.Clear
cboPeriodoCierre.AddItem "Abierto"
cboPeriodoCierre.ItemData(cboPeriodoCierre.ListCount - 1) = "A"
cboPeriodoCierre.AddItem "Diario"
cboPeriodoCierre.ItemData(cboPeriodoCierre.ListCount - 1) = "D"
cboPeriodoCierre.AddItem "Semanal"
cboPeriodoCierre.ItemData(cboPeriodoCierre.ListCount - 1) = "S"

cboTipoCierre.Clear
cboTipoCierre.AddItem "Saldos Abierto"
cboTipoCierre.ItemData(cboTipoCierre.ListCount - 1) = "A"
cboTipoCierre.AddItem "Cierre Ciego"
cboTipoCierre.ItemData(cboTipoCierre.ListCount - 1) = "C"


 With lswAuxiliares.ColumnHeaders
    .Clear
    .Add , , "Código", 1800, vbCenter
    .Add , , "Descripción", 4500
 End With
 
 
 With lswFormasPago.ColumnHeaders
    .Clear
    .Add , , "Código", 1800, vbCenter
    .Add , , "Descripción", 4500
 End With
 
 With lswDocumentos.ColumnHeaders
    .Clear
    .Add , , "Código", 1800, vbCenter
    .Add , , "Descripción", 4500
 End With
 
 
 
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")

 
 vScroll = False
     FlatScrollBar.Value = 0
 vScroll = True
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub imgCopia_Click()

tcMain.Item(7).Selected = True

End Sub



Private Sub lswAuxiliares_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, vTipo As String, vAsigna As String

On Error GoTo vError

If vPaso Then Exit Sub


Select Case True
   Case optLimita.Item(0).Value 'Creditos
        vTipo = "CRD"
           
   Case optLimita.Item(1).Value 'Fondos
        vTipo = "FND"
   
   Case optLimita.Item(2).Value 'CxC
        vTipo = "CXC"
End Select

If Item.Checked Then

   vAsigna = "REGISTRA"
   strSQL = "insert into CAJAS_AUXILIARES_ASG(tipo,cod_auxiliar,cod_caja,registro_fecha,registro_usuario)" _
           & "values('" & vTipo & "','" & Item.Text & "', '" & vCodigo & "', dbo.MyGetdate(),'" & glogon.Usuario & "')"
   
Else
   vAsigna = "ELIMINA"
   strSQL = "Delete CAJAS_AUXILIARES_ASG where tipo = '" & vTipo & "' and cod_auxiliar = '" & Item.Text _
          & "' and cod_caja = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

strSQL = "CONCEPTO: " & Item.Text & " -> AUXILIAR: " & vTipo & " -> CAJA: " & vCodigo
Call Bitacora(vAsigna, strSQL)


Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswDocumentos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert into CAJAS_DOCUMENTOS(TIPO_DOCUMENTO,cod_caja,registro_fecha,registro_usuario)" _
           & "values('" & Item.Text & "', '" & vCodigo & "', dbo.MyGetdate(),'" & glogon.Usuario & "')"
   
Else
   strSQL = "Delete CAJAS_DOCUMENTOS where TIPO_DOCUMENTO = '" & Item.Text _
          & "' and cod_caja = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswFormasPago_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert into CAJAS_FORMAS_PAGO(cod_forma_pago,cod_caja,registro_fecha,registro_usuario)" _
           & "values('" & Item.Text & "', '" & vCodigo & "', dbo.MyGetdate(),'" & glogon.Usuario & "')"
   
Else
   strSQL = "Delete CAJAS_FORMAS_PAGO where cod_forma_pago ='" & Item.Text _
          & "' and cod_caja = '" & vCodigo & "'"
       
End If
Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargaAuxiliares()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

If vCodigo = "" Then Exit Sub

Select Case True
   Case optLimita.Item(0).Value 'Creditos
        
        strSQL = "select A.CODIGO AS 'CODIGO',A.descripcion,C.COD_AUXILIAR as 'Asignado'" _
               & " from CATALOGO A" _
               & " left join CAJAS_AUXILIARES_ASG C on A.CODIGO = C.COD_AUXILIAR AND C.TIPO = 'CRD'" _
               & " and C.cod_caja = '" & vCodigo & "' order by C.COD_AUXILIAR desc, A.CODIGO"
   
   Case optLimita.Item(1).Value 'Fondos
        
        strSQL = "select A.COD_PLAN AS 'CODIGO',A.descripcion,C.COD_AUXILIAR as 'Asignado'" _
               & " from FND_PLANES A" _
               & " left join CAJAS_AUXILIARES_ASG C on A.COD_PLAN = C.COD_AUXILIAR AND C.TIPO = 'FND'" _
               & " and C.cod_caja = '" & vCodigo & "' order by C.COD_AUXILIAR desc, A.COD_PLAN"
   
   Case optLimita.Item(2).Value 'CxC
   
        strSQL = "select A.COD_CONCEPTO AS 'CODIGO',A.descripcion,C.COD_AUXILIAR as 'Asignado'" _
               & " from CXC_CONCEPTOS A" _
               & " left join CAJAS_AUXILIARES_ASG C on A.COD_CONCEPTO = C.COD_AUXILIAR AND C.TIPO = 'CXC'" _
               & " and C.cod_caja = '" & vCodigo & "' order by C.COD_AUXILIAR desc, A.COD_CONCEPTO"
         
End Select


vPaso = True
lswAuxiliares.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswAuxiliares.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion

  If Not IsNull(rs!asignado) Then
     itmX.Checked = True
     itmX.ForeColor = vbBlue
     itmX.ListSubItems(1).ForeColor = vbBlue
  End If

  rs.MoveNext
Loop
rs.Close
vPaso = False


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub optLimita_Click(Index As Integer)
Call sbCargaAuxiliares
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

If vCodigo = "" Then Exit Sub

Select Case Item.Index

   Case 1 'Rangos Efectivo/Documentos
        strSQL = "Select D.Cod_Divisa,D.descripcion,isnull(P.efectivo_maximo,0)as Efec_Max," _
               & "isnull(P.Efectivo_minimo,0) as Efec_Min ,isnull(P.Documentos_Maximo, 0) As Doc_Max," _
               & "isnull(P.Documentos_minimo,0) as Doc_Min" _
               & " from cntx_divisas D left join cajas_politicas_saldos P on D.Cod_divisa = P.cod_divisa and" _
               & " cod_caja = '" & vCodigo & "' where D.cod_Contabilidad = " & GLOBALES.gEnlace
        Call sbCargaGrid(vGridRangosDivisa, 6, strSQL)
        vGridRangosDivisa.MaxRows = vGridRangosDivisa.MaxRows - 1
        
   Case 2 'Servicios Asociados
        Call sbCargaNodosRecaudadorServicios
        
   Case 3 'Carga Auxiliares
        Call sbCargaAuxiliares

   Case 4 'Formas de Pago
        strSQL = "select F.COD_FORMA_PAGO,F.DESCRIPCION,C.cod_caja as 'Asignado'" _
               & " from SIF_FORMAS_PAGO F" _
               & " left join CAJAS_FORMAS_PAGO C on F.COD_FORMA_PAGO = C.COD_FORMA_PAGO " _
               & " AND  F.Activa = 1 and   C.cod_caja =  '" & vCodigo & "' order by C.cod_Caja desc,F.cod_Forma_Pago"
        vPaso = True
        lswFormasPago.ListItems.Clear
        
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Set itmX = lswFormasPago.ListItems.Add(, , rs!Cod_Forma_Pago)
              itmX.SubItems(1) = rs!Descripcion
          
          If Not IsNull(rs!asignado) Then
             itmX.Checked = True
             itmX.ForeColor = vbBlue
             itmX.ListSubItems(1).ForeColor = vbBlue
          End If
          
          rs.MoveNext
        Loop
        rs.Close
        vPaso = False
   
   Case 5 'Documentos
        strSQL = "select D.TIPO_DOCUMENTO,D.DESCRIPCION,C.cod_caja as 'Asignado'" _
               & " from SIF_DOCUMENTOS D" _
               & " left join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
               & " AND   D.Activo = 1 and C.cod_caja =  '" & txtCodigo.Text & "' order by C.cod_Caja desc,D.TIPO_DOCUMENTO"
        vPaso = True
        lswDocumentos.ListItems.Clear
        
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Set itmX = lswDocumentos.ListItems.Add(, , rs!TIPO_DOCUMENTO)
              itmX.SubItems(1) = rs!Descripcion
          
          If Not IsNull(rs!asignado) Then
             itmX.Checked = True
             itmX.ForeColor = vbBlue
             itmX.ListSubItems(1).ForeColor = vbBlue
          End If
          
          rs.MoveNext
        Loop
        rs.Close
        vPaso = False
        
   Case 6 'Usuarios
        strSQL = "select rtrim(usuario) as 'Usuario',registro_fecha, rtrim(registro_usuario) as 'registro_usuario',salida_fecha,salida_usuario from cajas_usuarios_h " _
            & "where cod_caja = '" & vCodigo & "' order by usuario"
        Call sbCargaGrid(vGrid, 5, strSQL)
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
 TimerX.Interval = 0
 TimerX.Enabled = False
 
 Call sbCargaOficinas
 
 Call sbLimpia
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpia
      Call sbToolBar(tlb, "edicion")
      txtCodigo.Text = ""
      txtCodigo.SetFocus
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      Call sbToolBar(tlb, "edicion")
      txtDescripcion.SetFocus
    
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
       gBusquedas.Consulta = "select cod_Caja,descripcion from cajas_definicion"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

    Case "REPORTES"

    Case "AYUDA"

End Select

End Sub




Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    txtCodigo.Text = ""
    gBusquedas.Consulta = "Select cod_caja,descripcion from cajas_definicion"
    gBusquedas.Columna = "cod_caja"
    gBusquedas.Orden = "cod_caja"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtCodigo.Text = Trim(gBusquedas.Resultado)
    txtDescripcion = Trim(gBusquedas.Resultado2)
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" Then Call sbConsulta(txtCodigo.Text)
End Sub

Private Sub txtContrasenaVence_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPeriodoCierre.SetFocus
End Sub


Private Sub txtCuentaDev_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaDev = gCuenta
    txtDescCuentaDev = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescCuentaDev.SetFocus
End Sub

Private Sub txtCuentaDev_LostFocus()
txtCuentaDev = fxgCntCuentaFormato(False, txtCuentaDev)
txtDescCuentaDev = fxgCntCuentaDesc(txtCuentaDev)
txtCuentaDev = fxgCntCuentaFormato(True, txtCuentaDev)

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpFechaApertura.SetFocus
End Sub

Private Sub sbLimpia()

vCodigo = ""
txtCodigo.Text = ""
txtDescripcion.Text = ""
txtContrasenaVence.Text = ""
txtNotas.Text = ""

txtContrasenaVence.Text = 60
chkActivo.Value = vbChecked
dtpFechaApertura.Value = Format(fxFechaServidor, "dd/mm/yyyy")

chkAperturaCompartida.Value = vbUnchecked
chkMovCbrJud.Value = vbUnchecked


chkBoveda.Value = vbUnchecked
chkTrasladoEfectivo.Value = vbUnchecked
chkCtaCaja.Value = vbUnchecked


chkLimita_Creditos.Value = vbUnchecked
chkLimita_Fondos.Value = vbUnchecked
chkLimita_CxC.Value = vbUnchecked
chkLimita_Patrimonio.Value = vbUnchecked
chkLimita_Consulta.Value = vbUnchecked

Call sbCboAsignaDato(cboPeriodoCierre, "Diario", True, "D")
Call sbCboAsignaDato(cboTipoCierre, "Cierre Ciego", True, "C")

tcMain.Item(0).Selected = True
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False
tcMain.Item(4).Enabled = False
tcMain.Item(5).Enabled = False
tcMain.Item(6).Enabled = False

StatusBarX.Panels.Item(1).Text = ""
StatusBarX.Panels.Item(2).Text = ""


End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update cajas_definicion set descripcion = '" & UCase(Trim(txtDescripcion)) & "'" _
         & ", notas = '" & UCase(txtNotas) & "',activa = " & chkActivo.Value _
         & ", apertura_fecha = '" & Format(dtpFechaApertura, "yyyymmdd") & "',Apertura_Compartida = " & chkAperturaCompartida.Value _
         & ", cierre_periocidad = '" & cboPeriodoCierre.ItemData(cboPeriodoCierre.ListIndex) & "'" _
         & ", cierre_tipo ='" & cboTipoCierre.ItemData(cboTipoCierre.ListIndex) & "',periocidad_contrasena = " & txtContrasenaVence & " " _
         & ", oficina_utiliza_usuario = " & chkUtilizaUsuario.Value & ",cod_oficina = '" & cboOficina.ItemData(cboOficina.ListIndex) & "' " _
         & ", cod_cuenta_dev = '" & fxgCntCuentaFormato(False, txtCuentaDev.Text) & "', PERMITE_MOV_CBRJUD = " & chkMovCbrJud.Value _
         & ", Limita_Consulta = " & chkLimita_Consulta.Value & ", Limita_Creditos = " & chkLimita_Creditos.Value _
         & ", Limita_Fondos = " & chkLimita_Fondos.Value & ", Limita_Patrimonio = " & chkLimita_Patrimonio.Value _
         & ", Limita_CxC = " & chkLimita_CxC.Value & ", PERMITE_RC = " & chkPermiteRC.Value _
         & ", PERMITE_TRASLADOS_EF = " & chkTrasladoEfectivo.Value & ", ROL_BOVEDA = " & chkBoveda.Value & ", UTILIZA_CTA_CAJA_EF = " & chkCtaCaja.Value _
         & " where cod_caja = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Definición Cajas: " & vCodigo)

Else
  vCodigo = txtCodigo.Text
   
   strSQL = "insert into cajas_definicion(cod_caja,descripcion,notas,activa,apertura_fecha,apertura_compartida" _
           & ",cierre_periocidad,cierre_tipo,periocidad_contrasena,oficina_utiliza_usuario,cod_oficina" _
           & ",cod_cuenta_dev,PERMITE_MOV_CBRJUD, Limita_Consulta, Limita_Creditos, Limita_Fondos, Limita_CxC" _
           & ",Limita_Patrimonio, PERMITE_RC, PERMITE_TRASLADOS_EF , ROL_BOVEDA, UTILIZA_CTA_CAJA_EF, REGISTRO_FECHA,REGISTRO_USUARIO)" _
           & " values('" & vCodigo & "','" & UCase(txtDescripcion) & "','" & UCase(txtNotas) & "'," & chkActivo.Value & "," _
           & "'" & Format(dtpFechaApertura, "yyyy/mm/dd") & "'," & chkAperturaCompartida.Value & ",'" & cboPeriodoCierre.ItemData(cboPeriodoCierre.ListIndex) & "'," _
           & "'" & cboTipoCierre.ItemData(cboTipoCierre.ListIndex) & "'," & txtContrasenaVence & "," & chkUtilizaUsuario & "," _
           & "'" & cboOficina.ItemData(cboOficina.ListIndex) & "','" & fxgCntCuentaFormato(False, txtCuentaDev.Text) _
           & "', " & chkMovCbrJud.Value & "," & chkLimita_Consulta.Value & "," & chkLimita_Creditos.Value & "," & chkLimita_Fondos.Value _
           & ", " & chkLimita_CxC.Value & "," & chkLimita_Patrimonio & ", " & chkPermiteRC.Value _
           & ", " & chkTrasladoEfectivo.Value & ", " & chkBoveda.Value & ", " & chkCtaCaja.Value _
           & ", dbo.MyGetdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Definición Cajas: " & vCodigo)

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
  strSQL = "delete cajas_definicion where cod_caja = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Caja : " & vCodigo)
  Call sbLimpia
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui
Call sbSIFCleanTxtInject(txtNotas)
Call sbSIFCleanTxtInject(txtDescripcion)
Call sbSIFCleanTxtInject(txtCodigo)

If txtDescripcion.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Recaudador no es válido ..."
If Not fxgCntCuentaValida(txtCuentaDev.Text) Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable Prinicipal no es válida.."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function



Private Sub sbConsulta(pCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.*,rtrim(O.Cod_Oficina) as 'Cod_Oficina', rtrim(O.Descripcion) as 'OficinaDesc'" _
       & ", isnull(Cta.Descripcion,'') as 'CuentaDesc'" _
       & " from cajas_definicion C inner join Sif_Oficinas O on C.cod_Oficina = O.cod_Oficina" _
       & " left join CntX_Cuentas Cta on C.Cod_Cuenta_Dev = Cta.Cod_Cuenta and Cta.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " where C.cod_caja = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

tcMain.Item(0).Selected = True

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  txtCodigo.Text = rs!Cod_Caja
  vCodigo = rs!Cod_Caja

  txtDescripcion.Text = rs!Descripcion
  txtDescripcion.SetFocus
  
  txtNotas.Text = rs!NOTAS
 
  chkActivo.Value = rs!Activa
  chkAperturaCompartida.Value = rs!Apertura_Compartida
  chkMovCbrJud.Value = rs!PERMITE_MOV_CBRJUD
  chkPermiteRC.Value = rs!PERMITE_RC
  dtpFechaApertura = Format(rs!Apertura_Fecha, "dd/mm/yyyy")
  
  chkBoveda.Value = rs!ROL_BOVEDA
  chkTrasladoEfectivo.Value = rs!PERMITE_TRASLADOS_EF
  chkCtaCaja.Value = rs!UTILIZA_CTA_CAJA_EF

 
  Select Case rs!Cierre_Periocidad
    Case "D"
       Call sbCboAsignaDato(cboPeriodoCierre, "Diario", True, "D")
    Case "A"
       Call sbCboAsignaDato(cboPeriodoCierre, "Abierto", True, "A")
    Case "S"
       Call sbCboAsignaDato(cboPeriodoCierre, "Semanal", True, "S")
  End Select
  
  If rs!Cierre_Tipo = "C" Then
       Call sbCboAsignaDato(cboTipoCierre, "Cierre Ciego", True, "C")
  Else
       Call sbCboAsignaDato(cboTipoCierre, "Saldos Abiertos", True, "A")
  End If
  
  txtContrasenaVence = rs!Periocidad_Contrasena
  
  txtCuentaDev.Text = fxgCntCuentaFormato(True, rs!Cod_Cuenta_Dev)
  txtDescCuentaDev.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, rs!Cod_Cuenta_Dev))
  
  chkUtilizaUsuario = rs!oficina_utiliza_usuario
  Call sbCboAsignaDato(cboOficina, rs!OficinaDesc, True, rs!COD_OFICINA)

  chkLimita_Consulta.Value = rs!Limita_Consulta
  chkLimita_Creditos.Value = rs!Limita_Creditos
  chkLimita_CxC.Value = rs!Limita_CxC
  chkLimita_Fondos.Value = rs!Limita_Fondos
  chkLimita_Patrimonio.Value = rs!Limita_Patrimonio

    tcMain.Item(1).Enabled = True
    tcMain.Item(2).Enabled = True
    tcMain.Item(3).Enabled = True
    tcMain.Item(4).Enabled = True
    tcMain.Item(5).Enabled = True
    tcMain.Item(6).Enabled = True
    
    StatusBarX.Panels.Item(1).Text = rs!Registro_Usuario & ""
    StatusBarX.Panels.Item(2).Text = rs!Registro_Fecha & ""

End If

rs.Close

Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxIndiceMultiple(xkey As String, vTipo As String) As String
Dim i As Long, strResultado As String, blnPaso As Boolean

xkey = fxIndiceCodigo(xkey)

blnPaso = True

If xkey = "" Then
  fxIndiceMultiple = ""
  Exit Function
End If

If vTipo = "D" Then ' Tipo
  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xkey, i, 1)
    Else
     blnPaso = False
    End If
    i = i + 1
  Loop
  
Else 'Numero

  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) = "-" Then blnPaso = False
    i = i + 1
  Loop
  strResultado = Mid(xkey, i, 50) '50 es un default ningun asiento es tan largo

End If

fxIndiceMultiple = strResultado

End Function


Private Function fxConceptoAsignado(vConcepto As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(cod_concepto),0) as cantidad from " _
        & "sif_conceptos_documento where tipo_documento ='" & vDocumento & "' and modulo = " & vModulo & " and " _
        & " cod_concepto = '" & vConcepto & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Cantidad = 0 Then
  fxConceptoAsignado = False
Else
  fxConceptoAsignado = True
End If
rs.Close
End Function



Private Sub sbCargaNodosRecaudadorServicios()
Dim strSQL As String, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
Dim xNode As Node, lng As Long
Dim i As Long


Me.MousePointer = vbHourglass

i = 0

With ArbolRecaudador
    .Nodes.Clear
    'Crear Root
    Set xNode = .Nodes.Add(, , "US", "Principal")
    xNode.Bold = True
    
    strSQL = "select cod_recaudador,descripcion from cajas_recaudador order by cod_recaudador"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
        Call sbCreaNodos("US", rs!Descripcion, "imgFormu", True, ArbolRecaudador, "0x0" & rs!COD_RECAUDADOR & "R")
        
        strSQL = "select C.cod_servicio,C.descripcion,X.cod_servicio as 'Asignado'" _
               & " from cajas_Servicios C left join cajas_servicios_asignados X on C.cod_servicio = X.cod_servicio" _
               & " and X.cod_recaudador = '" & rs!COD_RECAUDADOR & "' and X.Cod_caja = '" & vCodigo _
               & "' where C.cod_recaudador = '" & rs!COD_RECAUDADOR & "' order by   X.cod_servicio desc, C.cod_servicio"

        rs2.Open strSQL, glogon.Conection, adOpenStatic
        i = i + 1
        
        Do While Not rs2.EOF
            Call sbCreaNodos("0x0" & rs!COD_RECAUDADOR & "R", rs2!Descripcion, "imgDocu", True, ArbolRecaudador, "0x0" & rs2!COD_SERVICIO & "-" & i & "D")
            If Not IsNull(rs2!asignado) Then
                .Nodes.Item(.Nodes.Count).Checked = True
                .Nodes.Item(.Nodes.Count).ForeColor = vbBlue
            Else
                .Nodes.Item(.Nodes.Count).Checked = False
                .Nodes.Item(.Nodes.Count).ForeColor = vbRed
            End If
            .Nodes(.Nodes.Count).Expanded = True
            rs2.MoveNext
        Loop
        rs2.Close
        rs.MoveNext
        .Nodes(.Nodes.Count).Expanded = True
    Loop
    rs.Close
    
    
    xNode.Expanded = True
    
End With
Me.MousePointer = vbDefault


End Sub




Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGridRangosDivisa.Row = vGridRangosDivisa.ActiveRow
vGridRangosDivisa.col = 1
If vGridRangosDivisa.Text = "" Then vGridRangosDivisa.Text = 0

strSQL = "select isnull(count(*),0) as Existe from cajas_politicas_saldos  " _
       & " where cod_divisa ='" & vGridRangosDivisa.Text & "' and cod_caja = '" & txtCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
    If Trim(vGridRangosDivisa.Text) = "" Then Exit Function
    strSQL = "insert into cajas_politicas_saldos(cod_caja,cod_divisa,efectivo_maximo,efectivo_minimo,documentos_maximo,documentos_minimo," _
            & "registro_usuario,registro_fecha)" _
            & " values('" & txtCodigo & "','" & UCase(vGridRangosDivisa.Text) & "',"
    vGridRangosDivisa.col = 3
    strSQL = strSQL & "" & CCur(vGridRangosDivisa.Text) & ","
    vGridRangosDivisa.col = 4
    strSQL = strSQL & CCur(vGridRangosDivisa.Text) & ","
    vGridRangosDivisa.col = 5
    strSQL = strSQL & CCur(vGridRangosDivisa.Text) & ","
    vGridRangosDivisa.col = 6
    strSQL = strSQL & CCur(vGridRangosDivisa.Text) & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
    
    Call ConectionExecute(strSQL)
    
    vGridRangosDivisa.col = 1
    Call Bitacora("Registra", "Mantenimieto Rangos Divisas: " & vGridRangosDivisa.Text & " " & txtCodigo)

Else 'Actualizar
    
    vGridRangosDivisa.col = 3
    strSQL = "update cajas_politicas_saldos set efectivo_maximo = " & CCur(vGridRangosDivisa.Text) & ",efectivo_minimo = " & ""
    vGridRangosDivisa.col = 4
    strSQL = strSQL & CCur(vGridRangosDivisa.Text) & ",documentos_maximo = " & ""
    vGridRangosDivisa.col = 5
    strSQL = strSQL & CCur(vGridRangosDivisa.Text) & ",documentos_minimo = " & ""
    vGridRangosDivisa.col = 6
    strSQL = strSQL & CCur(vGridRangosDivisa.Text) & "  where cod_divisa =  '"
    vGridRangosDivisa.col = 1
    strSQL = strSQL & UCase(vGridRangosDivisa.Text) & " ' and cod_caja = '" & txtCodigo.Text & "'"
       
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Mantenimiento Rangos Divisas : " & vGridRangosDivisa.Text & " " & txtCodigo & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'Resume
End Function



Private Sub vGridRangosDivisa_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGridRangosDivisa.ActiveCol = vGridRangosDivisa.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
End If


Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaOficinas()
Dim strSQL As String

strSQL = "select rtrim(cod_oficina) as 'IdX',  rtrim(Descripcion) as 'Itmx'" _
       & " from sif_oficinas where estado = 1 order by cod_oficina"
Call sbCbo_Llena_New(cboOficina, strSQL, False, True)

End Sub

