VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_TraspasoTesoreria 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traslado de Solicitudes formalizadas a Bancos"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   HelpContextID   =   3030
   Icon            =   "CR_TraspasoTesoreria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   84
      Top             =   7920
      Visible         =   0   'False
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20558
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.GroupBox fraFiltros 
      Height          =   2775
      Left            =   2520
      TabIndex        =   65
      Top             =   3360
      Visible         =   0   'False
      Width           =   7935
      _Version        =   1572864
      _ExtentX        =   13991
      _ExtentY        =   4890
      _StockProps     =   79
      Caption         =   "Filtros adicionales"
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   1800
         TabIndex        =   66
         Top             =   600
         Width           =   6012
         _Version        =   1572864
         _ExtentX        =   10610
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
         Appearance      =   17
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboUsuarios 
         Height          =   312
         Left            =   1800
         TabIndex        =   67
         Top             =   1080
         Width           =   6012
         _Version        =   1572864
         _ExtentX        =   10610
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
         Appearance      =   17
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboSistema 
         Height          =   312
         Left            =   1800
         TabIndex        =   68
         Top             =   1560
         Width           =   6012
         _Version        =   1572864
         _ExtentX        =   10610
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
         Appearance      =   17
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   420
         Index           =   0
         Left            =   6000
         TabIndex        =   69
         Top             =   2160
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":000C
      End
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   420
         Index           =   1
         Left            =   7320
         TabIndex        =   70
         Top             =   2160
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   741
         _StockProps     =   79
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":070C
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   11
         Left            =   360
         TabIndex        =   73
         Top             =   600
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cuenta Bancaria:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   12
         Left            =   360
         TabIndex        =   72
         Top             =   1080
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Usuario:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   13
         Left            =   360
         TabIndex        =   71
         Top             =   1560
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Sistema:"
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
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20553
      _ExtentY        =   11451
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
      ItemCount       =   7
      SelectedItem    =   2
      Item(0).Caption =   "Remesa"
      Item(0).ControlCount=   24
      Item(0).Control(0)=   "Label8(0)"
      Item(0).Control(1)=   "Label8(1)"
      Item(0).Control(2)=   "Label8(2)"
      Item(0).Control(3)=   "Label8(3)"
      Item(0).Control(4)=   "Label8(4)"
      Item(0).Control(5)=   "Label8(5)"
      Item(0).Control(6)=   "Label8(6)"
      Item(0).Control(7)=   "txtRemesa"
      Item(0).Control(8)=   "txtFecha"
      Item(0).Control(9)=   "txtEstado"
      Item(0).Control(10)=   "txtUsuario"
      Item(0).Control(11)=   "txtNotas"
      Item(0).Control(12)=   "dtpInicio"
      Item(0).Control(13)=   "dtpCorte"
      Item(0).Control(14)=   "btnBarra(0)"
      Item(0).Control(15)=   "btnBarra(1)"
      Item(0).Control(16)=   "btnBarra(2)"
      Item(0).Control(17)=   "lswRemesas"
      Item(0).Control(18)=   "fraReporte"
      Item(0).Control(19)=   "btnBarra(9)"
      Item(0).Control(20)=   "txtRemesa_Casos"
      Item(0).Control(21)=   "txtRemesa_Monto"
      Item(0).Control(22)=   "Label8(22)"
      Item(0).Control(23)=   "Label8(23)"
      Item(1).Caption =   "Cargar"
      Item(1).ControlCount=   12
      Item(1).Control(0)=   "Label8(9)"
      Item(1).Control(1)=   "Label8(10)"
      Item(1).Control(2)=   "cboCarga"
      Item(1).Control(3)=   "cboOficina"
      Item(1).Control(4)=   "chkFiltros"
      Item(1).Control(5)=   "chkCarga"
      Item(1).Control(6)=   "lswCarga"
      Item(1).Control(7)=   "txtCargaTotal"
      Item(1).Control(8)=   "btnBarra(3)"
      Item(1).Control(9)=   "btnBarra(4)"
      Item(1).Control(10)=   "btnBarra(5)"
      Item(1).Control(11)=   "Label8(18)"
      Item(2).Caption =   "Trasladar"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "Label8(14)"
      Item(2).Control(1)=   "cboTraslado"
      Item(2).Control(2)=   "lswTraslado"
      Item(2).Control(3)=   "txtPagoTotal"
      Item(2).Control(4)=   "btnBarra(6)"
      Item(2).Control(5)=   "btnBarra(7)"
      Item(2).Control(6)=   "Label8(19)"
      Item(2).Control(7)=   "ShortcutCaption1(1)"
      Item(3).Caption =   "Informes"
      Item(3).ControlCount=   9
      Item(3).Control(0)=   "opt(0)"
      Item(3).Control(1)=   "txtRepRemesas"
      Item(3).Control(2)=   "lblRemesa"
      Item(3).Control(3)=   "opt(1)"
      Item(3).Control(4)=   "chkRemesaInd"
      Item(3).Control(5)=   "lswRep"
      Item(3).Control(6)=   "btnBarra(8)"
      Item(3).Control(7)=   "ShortcutCaption1(0)"
      Item(3).Control(8)=   "ShortcutCaption1(2)"
      Item(4).Caption =   "Reactivación"
      Item(4).ControlCount=   6
      Item(4).Control(0)=   "Label3(0)"
      Item(4).Control(1)=   "cmdReactivar"
      Item(4).Control(2)=   "Label8(15)"
      Item(4).Control(3)=   "Label8(16)"
      Item(4).Control(4)=   "txtOperacion"
      Item(4).Control(5)=   "txtDetalle"
      Item(5).Caption =   "Cambio de Concepto"
      Item(5).ControlCount=   4
      Item(5).Control(0)=   "txtCamConOP"
      Item(5).Control(1)=   "vGrid"
      Item(5).Control(2)=   "Label3(3)"
      Item(5).Control(3)=   "Label8(17)"
      Item(6).Caption =   "Consulta"
      Item(6).ControlCount=   5
      Item(6).Control(0)=   "txtConsulta_Operacion"
      Item(6).Control(1)=   "Label8(20)"
      Item(6).Control(2)=   "Label3(1)"
      Item(6).Control(3)=   "txtConsulta_Remesa"
      Item(6).Control(4)=   "Label8(21)"
      Begin XtremeSuiteControls.ListView lswRemesas 
         Height          =   3132
         Left            =   -68440
         TabIndex        =   18
         Top             =   3240
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1572864
         _ExtentX        =   17801
         _ExtentY        =   5524
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswCarga 
         Height          =   3972
         Left            =   -70000
         TabIndex        =   35
         Top             =   2040
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1572864
         _ExtentX        =   20553
         _ExtentY        =   7006
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
      Begin XtremeSuiteControls.ListView lswTraslado 
         Height          =   3975
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   11415
         _Version        =   1572864
         _ExtentX        =   20135
         _ExtentY        =   7011
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRep 
         Height          =   3612
         Left            =   -70000
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1572864
         _ExtentX        =   20553
         _ExtentY        =   6371
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   0
         Left            =   -69520
         TabIndex        =   59
         Top             =   5160
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(Pendientes) Detalle de Remesa"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCamConOP 
         Height          =   432
         Left            =   -68320
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1572864
         _ExtentX        =   4254
         _ExtentY        =   762
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.GroupBox fraReporte 
         Height          =   2052
         Left            =   -65680
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   7452
         _Version        =   1572864
         _ExtentX        =   13144
         _ExtentY        =   3619
         _StockProps     =   79
         Caption         =   "Informes"
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
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   0
            Left            =   1920
            TabIndex        =   28
            Top             =   1200
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Pendientes"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkRepFechas 
            Height          =   255
            Left            =   4680
            TabIndex        =   25
            Top             =   360
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todas"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.DateTimePicker dtpRepCorte 
            Height          =   315
            Left            =   3240
            TabIndex        =   23
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   556
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtpRepInicio 
            Height          =   315
            Left            =   1920
            TabIndex        =   22
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   556
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.ComboBox cboRepOficina 
            Height          =   312
            Left            =   1920
            TabIndex        =   24
            Top             =   720
            Width           =   4932
            _Version        =   1572864
            _ExtentX        =   8705
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   0
            Left            =   5760
            TabIndex        =   26
            Top             =   1200
            Width           =   612
            _Version        =   1572864
            _ExtentX        =   1080
            _ExtentY        =   741
            _StockProps     =   79
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "CR_TraspasoTesoreria.frx":0E0C
         End
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   1
            Left            =   6360
            TabIndex        =   27
            Top             =   1200
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   741
            _StockProps     =   79
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "CR_TraspasoTesoreria.frx":1513
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   1
            Left            =   3600
            TabIndex        =   29
            Top             =   1200
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Trasladadas"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin VB.Image imgRepRefresca 
            Height          =   240
            Left            =   6600
            Picture         =   "CR_TraspasoTesoreria.frx":1B51
            ToolTipText     =   "Actualizar Oficinas"
            Top             =   360
            Width           =   240
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   372
            Index           =   8
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Oficina/Agencia:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   372
            Index           =   7
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   1212
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Corte:"
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
         End
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   0
         Left            =   -65680
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Nueva"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":2241
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa 
         Height          =   432
         Left            =   -68440
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1572864
         _ExtentX        =   4254
         _ExtentY        =   762
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   312
         Left            =   -68440
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1572864
         _ExtentX        =   4254
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
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   -64840
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   -64840
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   -68440
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1572864
         _ExtentX        =   17801
         _ExtentY        =   1397
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
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -68440
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -67240
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   1
         Left            =   -63880
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   741
         _StockProps     =   79
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":2941
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   2
         Left            =   -63400
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   741
         _StockProps     =   79
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":2EE5
      End
      Begin XtremeSuiteControls.ComboBox cboCarga 
         Height          =   312
         Left            =   -67600
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1572864
         _ExtentX        =   13573
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
         Appearance      =   17
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   312
         Left            =   -67600
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1572864
         _ExtentX        =   13573
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
         Appearance      =   17
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkFiltros 
         Height          =   372
         Left            =   -67600
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Filtros adicionales?"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkCarga 
         Height          =   252
         Left            =   -69880
         TabIndex        =   36
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboTraslado 
         Height          =   312
         Left            =   2160
         TabIndex        =   38
         Top             =   600
         Width           =   7692
         _Version        =   1572864
         _ExtentX        =   13573
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
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4572
         Left            =   -68800
         TabIndex        =   41
         Top             =   1800
         Visible         =   0   'False
         Width           =   10332
         _Version        =   524288
         _ExtentX        =   18225
         _ExtentY        =   8065
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   494
         ScrollBars      =   2
         SpreadDesigner  =   "CR_TraspasoTesoreria.frx":35EC
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   432
         Left            =   -68440
         TabIndex        =   46
         Top             =   1080
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1572864
         _ExtentX        =   4254
         _ExtentY        =   762
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   3
         Left            =   -63880
         TabIndex        =   50
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":3C40
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   4
         Left            =   -62560
         TabIndex        =   51
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Cargar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":4340
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   5
         Left            =   -61240
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":4A48
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   6
         Left            =   6960
         TabIndex        =   53
         Top             =   960
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":5154
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   7
         Left            =   8280
         TabIndex        =   54
         Top             =   960
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Traslado"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":5854
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   8
         Left            =   -60760
         TabIndex        =   55
         Top             =   5640
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Informe"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":6125
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdReactivar 
         Height          =   420
         Left            =   -62800
         TabIndex        =   56
         Top             =   5760
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "ReActivar Desemsolsos de la Operación"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":682C
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   312
         Left            =   -59200
         TabIndex        =   57
         Top             =   4590
         Visible         =   0   'False
         Width           =   852
         _Version        =   1572864
         _ExtentX        =   1503
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
         Text            =   "15"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkRemesaInd 
         Height          =   372
         Left            =   -60640
         TabIndex        =   58
         Top             =   5040
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Indicar Remesa"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   1
         Left            =   -69520
         TabIndex        =   60
         Top             =   5520
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(Traslado) Detalle Agrupado de Remesa"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCargaTotal 
         Height          =   312
         Left            =   -60760
         TabIndex        =   62
         Top             =   6120
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1572864
         _ExtentX        =   4254
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
      End
      Begin XtremeSuiteControls.FlatEdit txtPagoTotal 
         Height          =   312
         Left            =   9120
         TabIndex        =   64
         Top             =   6000
         Width           =   2412
         _Version        =   1572864
         _ExtentX        =   4254
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
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   3912
         Left            =   -68440
         TabIndex        =   47
         Top             =   1560
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1572864
         _ExtentX        =   17378
         _ExtentY        =   6900
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtConsulta_Operacion 
         Height          =   432
         Left            =   -68320
         TabIndex        =   74
         Top             =   960
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1572864
         _ExtentX        =   4254
         _ExtentY        =   762
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtConsulta_Remesa 
         Height          =   3912
         Left            =   -68320
         TabIndex        =   77
         Top             =   1560
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1572864
         _ExtentX        =   17378
         _ExtentY        =   6900
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   9
         Left            =   -64360
         TabIndex        =   79
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   741
         _StockProps     =   79
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "CR_TraspasoTesoreria.frx":6F2C
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa_Monto 
         Height          =   312
         Left            =   -61480
         TabIndex        =   81
         Top             =   1680
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa_Casos 
         Height          =   312
         Left            =   -61480
         TabIndex        =   80
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   2
         Left            =   -64600
         TabIndex        =   88
         Top             =   4560
         Visible         =   0   'False
         Width           =   11655
         _Version        =   1572864
         _ExtentX        =   20558
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Remesas - visualizar últimas"
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
      Begin XtremeShortcutBar.ShortcutCaption lblRemesa 
         Height          =   375
         Left            =   -70000
         TabIndex        =   87
         Top             =   4560
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9551
         _ExtentY        =   661
         _StockProps     =   14
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
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   86
         Top             =   1440
         Width           =   11415
         _Version        =   1572864
         _ExtentX        =   20135
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Operaciones Pendientes a Trasladar"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   85
         Top             =   480
         Visible         =   0   'False
         Width           =   11655
         _Version        =   1572864
         _ExtentX        =   20558
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   23
         Left            =   -62320
         TabIndex        =   83
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Casos:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   22
         Left            =   -62320
         TabIndex        =   82
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Monto:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   21
         Left            =   -69640
         TabIndex        =   78
         Top             =   1440
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Remesa:"
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
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Consulta la Remesa en donde se registró la Operación"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   372
         Index           =   1
         Left            =   -69880
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   11412
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   20
         Left            =   -69640
         TabIndex        =   75
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "No. Operación:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   19
         Left            =   7320
         TabIndex        =   63
         Top             =   6000
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Total:"
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
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   18
         Left            =   -62560
         TabIndex        =   61
         Top             =   6120
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Total:"
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
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   17
         Left            =   -69640
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "No. Operación:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   16
         Left            =   -69760
         TabIndex        =   45
         Top             =   1440
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Detalle:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   15
         Left            =   -69760
         TabIndex        =   44
         Top             =   1080
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "No. Operación:"
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
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cambio de Conceptos de desembolsos a terceros; de operaciones pendientes de girar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   372
         Index           =   3
         Left            =   -69880
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   11412
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   $"CR_TraspasoTesoreria.frx":765D
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   372
         Index           =   0
         Left            =   -69880
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   11412
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   14
         Left            =   840
         TabIndex        =   37
         Top             =   600
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Remesa Id:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   10
         Left            =   -69400
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Oficina/Agencia:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   9
         Left            =   -69400
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Remesa Id:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   6
         Left            =   -68440
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Últimas Remesas Registradas:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   5
         Left            =   -69400
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Notas:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   4
         Left            =   -65680
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Usuario:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   3
         Left            =   -69400
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Registro:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   2
         Left            =   -65680
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Estado:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   1
         Left            =   -69400
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Corte:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   0
         Left            =   -69400
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Remesa Id:"
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
      End
   End
   Begin XtremeSuiteControls.Label lblTituloMain 
      Height          =   615
      Left            =   2280
      TabIndex        =   89
      Top             =   360
      Width           =   8775
      _Version        =   1572864
      _ExtentX        =   15478
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Traslado de Operaciones a Bancos"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12372
   End
End
Attribute VB_Name = "frmCR_TraspasoTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Dim itmX As ListViewItem, vPaso As Boolean

Dim mRequiereAutorizacion As Boolean
Dim vDuplicado As Boolean
Dim strLista  As String

Private Sub btnBarra_Click(Index As Integer)
Dim i As Integer

On Error GoTo vError

Select Case Index
  Case 0 'NUEVO"
     
    Call sbLimpia
    
  Case 9 'GUARDAR
    If txtRemesa.Text = "" Then
     
            strSQL = "select isnull(max(cod_remesa),0) + 1 as Ultimo from CRD_REMESAS_TES"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert CRD_REMESAS_TES(cod_remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas) values(" & rs!Ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!Ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de CRD Traslado a Tesoreria : " & txtRemesa)
    
    Else
        If txtEstado.Text = "Abierta" Then
                    
            strSQL = "update CRD_REMESAS_TES set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where cod_remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
             
            Call Bitacora("Modifica", "Remesa de CRD Traslado a Tesoreria : " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
  Case 1 'BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Abierta" Then
            strSQL = "delete CRD_REMESAS_TES_detalle where Cod_Remesa = " & txtRemesa
            
            strSQL = strSQL & Space(10) & "delete CRD_REMESAS_TES where Cod_Remesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            
            Call Bitacora("Elimina", "Remesa de CRD Traslado a Tesoreria : " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case 2 'REPORTES"
     fraReporte.Visible = Not fraReporte.Visible


  '---------Carga
  Case 3 'Carga: Buscar
    If cboCarga.ListCount = 0 Then Exit Sub
    Call sbCargaBuscar
  
  Case 4 'Carga: Cargar
    If lswCarga.ListItems.Count = 0 Then Exit Sub
    Call sbCarga
  
  Case 5 'Carga: Cerrar Remesa
    Call sbCerrar

  '---------Traslado
  Case 6 'Traslado: Buscar
    If cboTraslado.ListCount = 0 Then Exit Sub
    Call sbTrasladoBuscar
  
  Case 7 'Traslado: Traslado
    If cboTraslado.ListCount = 0 Then Exit Sub
    If lswTraslado.ListItems.Count = 0 Then
            Call sbTrasladoBuscar
    End If
    Call sbTraslado
  
  '---------Reportes
  Case 8
    Call sbInforme_Remesa

End Select


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnFiltros_Click(Index As Integer)
Select Case Index
  Case 0 'Buscar
    Call sbCargaBuscar
    fraFiltros.Visible = False
    chkFiltros.Value = vbUnchecked
  
  Case 1 'Refrescar
    Call sbFiltros
    
End Select

End Sub

Private Sub btnReporte_Click(Index As Integer)

Select Case Index
    Case 0 'Reporte
        Select Case True
          Case optReporte.Item(0).Value
            Call sbReportePendientes
          Case optReporte.Item(1).Value
            Call sbReporteEnviadas
        End Select
    Case 1 'Cerrar
      fraReporte.Visible = False
End Select
End Sub

Private Sub cboCarga_Click()
Dim vFechaInicio As Date, vFechaCorte As Date

If cboCarga.ListCount = 0 Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear

strSQL = "select fecha_inicio,fecha_corte from CRD_REMESAS_TES where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!fecha_inicio
  vFechaCorte = rs!fecha_corte
rs.Close


'Carga Oficinas
strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas  where cod_oficina in(" _
       & " select R.cod_oficina_R" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.tesoreria is null and R.estado in('A','C') and id_solicitud not in(select id_solicitud from CRD_REMESAS_TES_DETALLE)" _
       & " group by R.cod_oficina_R)" _
       & " order by cod_oficina"
Call sbCbo_Llena_New(cboOficina, strSQL, True, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(pRemesa As Long)

Call sbLimpia
  
strSQL = "select T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
       & " from CRD_REMESAS_TES T left join vCrd_Remesa_Tes_Rsm D on T.cod_Remesa = D.cod_Remesa" _
       & " where T.Cod_Remesa = " & pRemesa

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa.Text = CStr(rs!cod_Remesa)
  txtUsuario.Text = rs!Usuario
  txtFecha.Text = rs!fecha
  
  Select Case rs!ESTADO
    Case "A"
      txtEstado = "Abierta"
    Case "C"
      txtEstado = "Cerrada"
    Case "T"
      txtEstado = "Trasladada"
  End Select
  
  dtpInicio.Value = rs!fecha_inicio
  dtpCorte.Value = rs!fecha_corte
  
  txtNotas.Text = rs!notas
  txtRemesa_Casos.Text = Format(rs!Casos, "###,##0")
  txtRemesa_Monto.Text = Format(rs!Monto, "Standard")
  
End If
rs.Close

End Sub



Private Sub sbFiltros()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select fecha_inicio,fecha_corte from CRD_REMESAS_TES where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!fecha_inicio
  vFechaCorte = rs!fecha_corte
rs.Close


'Cargado de Bancos
strSQL = " select B.id_Banco as 'Idx',B.descripcion as 'Itmx'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & "  inner join Tes_Bancos B on R.cod_Banco = B.id_Banco" _
       & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.tesoreria is null and R.estado in('A','C') and R.id_solicitud not in(select id_solicitud from CRD_REMESAS_TES_DETALLE)" _
       & " group by B.id_Banco,B.descripcion"
Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
        

'Cargado de Usuarios
strSQL = " select R.UserFor as Itmx" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.tesoreria is null and R.estado in('A','C') and R.id_solicitud not in(select id_solicitud from CRD_REMESAS_TES_DETALLE)" _
       & " group by R.UserFor"
Call sbCbo_Llena_New(cboUsuarios, strSQL, True, False)



'Cargado de Sistemas
strSQL = " select ISNULL(R.COD_APP,'" & App.ProductName & "')  as Itmx" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.tesoreria is null and R.estado in('A','C') and R.id_solicitud not in(select id_solicitud from CRD_REMESAS_TES_DETALLE)" _
       & " group by ISNULL(R.COD_APP,'" & App.ProductName & "')"
       
Call sbCbo_Llena_New(cboSistema, strSQL, True, False)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkFiltros_Click()
If chkFiltros.Value = vbChecked Then
   fraFiltros.Visible = True
   Call sbFiltros
Else
   fraFiltros.Visible = False
End If
End Sub

Private Sub chkRepFechas_Click()
If chkRepFechas.Value = vbChecked Then
  dtpRepInicio.Enabled = False
Else
  dtpRepInicio.Enabled = True
End If

dtpRepCorte.Enabled = dtpRepInicio.Enabled

End Sub

Private Sub sbInforme_Remesa()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim vTipoUser As String, xRemesa As String

On Error GoTo vError

If lblRemesa.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


vSubTitulo = ""
vFiltro = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Crédito"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Traslado a Tesoreria")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If

 Select Case True
  Case opt.Item(0).Value 'Pendiente Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_SGTRemesaTESDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Traslado Detalle Agrupado Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_SGTRemesaTESDetalleAgrp.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA TRASLADO A TESORERIA : CREDITOS'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = "{CRD_REMESAS_TES.COD_REMESA} = " & lblRemesa.Tag
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub imgRepRefresca_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

 
If chkRepFechas.Value = vbChecked Then
  vFechaInicio = "1900/01/01"
  vFechaCorte = fxFechaServidor
Else
  vFechaInicio = dtpRepInicio.Value
  vFechaCorte = dtpRepCorte.Value
End If


'Carga Oficinas
strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas  where cod_oficina in(" _
       & " select R.cod_oficina_R" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.tesoreria is null and R.estado in('A','C') and id_solicitud not in(select id_solicitud from CRD_REMESAS_TES_DETALLE)" _
       & " group by R.cod_oficina_R)" _
       & " order by cod_oficina"
Call sbCbo_Llena_New(cboRepOficina, strSQL, True, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswCarga_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCarga.SortKey = ColumnHeader.Index - 1
  If lswCarga.SortOrder = 0 Then lswCarga.SortOrder = 1 Else lswCarga.SortOrder = 0
  lswCarga.Sorted = True
End Sub

Private Sub lswCarga_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curTotal As Currency

If Trim(txtCargaTotal.Text) = "" Then txtCargaTotal.Text = 0

curTotal = CCur(txtCargaTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(8))
Else
   curTotal = curTotal - CCur(Item.SubItems(8))
End If

txtCargaTotal.Text = Format(curTotal, "Standard")
End Sub


Private Sub lswRemesas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRemesas.SortKey = ColumnHeader.Index - 1
  If lswRemesas.SortOrder = 0 Then lswRemesas.SortOrder = 1 Else lswRemesas.SortOrder = 0
  lswRemesas.Sorted = True
End Sub

Private Sub lswRemesas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    If lswRemesas.ListItems.Count <= 0 Then Exit Sub
    Call sbConsulta(Item.Text)
End Sub

Private Sub lswRep_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRep.SortKey = ColumnHeader.Index - 1
  If lswRep.SortOrder = 0 Then lswRep.SortOrder = 1 Else lswRep.SortOrder = 0
  lswRep.Sorted = True
End Sub

Private Sub lswRep_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lswRep.ListItems.Count <= 0 Then Exit Sub

lblRemesa.Caption = Item.Text & " ¦ " & Item.SubItems(1) _
            & " ¦ " & Item.SubItems(2)
lblRemesa.Tag = Item.Text
End Sub


Private Sub cboTraslado_Click()
    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
End Sub

Private Sub chkCarga_Click()
Dim i As Integer, curTotal As Currency


For i = 1 To lswCarga.ListItems.Count
  lswCarga.ListItems.Item(i).Checked = chkCarga.Value
  
   If chkCarga.Value = vbChecked Then
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(8))
   End If
  
Next i

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub



Private Sub sbReporteRemesas(pRemesa As Long)
Dim vSubTitulo As String, vFiltro As String
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Crédito > Seguimiento Tramites"
 
 .Connect = glogon.ConectRPT
  
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(4) = "fxFiltro='" & vFiltro & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("AfiComisionRemesas.rpt")
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbLimpia()

Me.MousePointer = vbHourglass

fraFiltros.Visible = False

Select Case tcMain.Selected.Index
  Case 0 'Remesas
     txtEstado.Text = ""
     txtFecha.Text = ""
     txtUsuario.Text = ""
     txtRemesa.Text = ""
     
     txtRemesa_Casos.Text = ""
     txtRemesa_Monto.Text = ""
     
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    
    dtpRepInicio.Value = dtpInicio.Value
    dtpRepCorte.Value = dtpInicio.Value
    
    fraReporte.Visible = False
    
    txtNotas.Text = ""
     
     strSQL = "select TOP 50 T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
            & " from CRD_REMESAS_TES T left join vCrd_Remesa_Tes_Rsm D on T.cod_Remesa = D.cod_Remesa" _
            & " order by T.fecha desc"
     
     
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!cod_Remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                
                Select Case rs!ESTADO
                  Case "A"
                     itmX.SubItems(3) = "Abierta"
                  Case "C"
                     itmX.SubItems(3) = "Cerrada"
                  Case "T"
                     itmX.SubItems(3) = "Trasladada"
                End Select
                
                itmX.SubItems(4) = Format(rs!fecha_inicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!fecha_corte, "dd/mm/yyyy")
                itmX.SubItems(6) = Format(rs!Casos, "###,###0")
                itmX.SubItems(7) = Format(rs!Monto, "Standard")
                itmX.SubItems(8) = rs!notas
                
                
       End With
       rs.MoveNext
     Loop
     rs.Close
     
  Case 1 'Carga
    'Solo busca las remesas que se encuentran Abiertas
    
    vPaso = True
    
    cboCarga.Clear

    lswCarga.ListItems.Clear
    chkCarga.Value = vbUnchecked
        
    strSQL = "select * from CRD_REMESAS_TES where estado = 'A' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!cod_Remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!fecha_inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
      
      cboCarga.ItemData(cboCarga.ListCount - 1) = CStr(rs!cod_Remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!cod_Remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!fecha_inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
    End If
    

    vPaso = False
    Call cboCarga_Click
    Call chkFiltros_Click
   
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
        
    strSQL = "select * from CRD_REMESAS_TES where estado = 'C' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboTraslado.AddItem (Format(rs!cod_Remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!fecha_inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
      cboTraslado.ItemData(cboTraslado.ListCount - 1) = CStr(rs!cod_Remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboTraslado.Text = (Format(rs!cod_Remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!fecha_inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboTraslado_Click

  
  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
            & " from CRD_REMESAS_TES T left join vCrd_Remesa_Tes_Rsm D on T.cod_Remesa = D.cod_Remesa" _
            & " order by T.fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!cod_Remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                
                Select Case rs!ESTADO
                  Case "A"
                     itmX.SubItems(3) = "Abierta"
                  Case "C"
                     itmX.SubItems(3) = "Cerrada"
                  Case "T"
                     itmX.SubItems(3) = "Trasladada"
                End Select
                
      
                itmX.SubItems(4) = Format(rs!fecha_inicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!fecha_corte, "dd/mm/yyyy")
                itmX.SubItems(6) = Format(rs!Casos, "###,###0")
                itmX.SubItems(7) = Format(rs!Monto, "Standard")
                itmX.SubItems(8) = rs!notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close

    
  Case 4 'Re-Activaciones
    txtOperacion.Tag = 0
    txtOperacion = ""
    txtDetalle = ""
 
   
  Case 5 'Cambio de Concepto
     txtCamConOP.Text = ""
     vGrid.MaxRows = 0
 
 End Select


Me.MousePointer = vbDefault

End Sub




Private Sub lswTraslado_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswTraslado.SortKey = ColumnHeader.Index - 1
  If lswTraslado.SortOrder = 0 Then lswTraslado.SortOrder = 1 Else lswTraslado.SortOrder = 0
  lswTraslado.Sorted = True
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 Call sbLimpia
End Sub


Private Sub sbCargaBuscar()
Dim rs2 As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency
Dim bSueprvisar As Boolean

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0

bSueprvisar = True

strSQL = "select fecha_inicio,fecha_corte from CRD_REMESAS_TES where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!fecha_inicio
  vFechaCorte = rs!fecha_corte
rs.Close

If cboBanco.Text <> "TODOS" And chkFiltros.Value = vbChecked Then
    bSueprvisar = fxSupervisaBanco
End If

If bSueprvisar Then
    strSQL = "select R.id_solicitud,R.codigo,S.cedula,S.nombre,R.montoapr,R.monto_girado,R.TES_SUPERVISION_FECHA," _
           & "dbo.fxTesSupervisa(S.cedula,S.nombre,R.monto_girado,0,'C') as 'Duplicado'" _
           & ", isnull(vD.Numero,0) as 'DESEM_NUM', isnull(vD.MONTO,0) as 'DESEM_MONTO'" _
           & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
           & "   inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
           & "    left join CRD_REMESAS_TES_DETALLE Td on R.id_solicitud = Td.id_Solicitud" _
           & "    left join vCrdOperacion_DesembolsosGiro vD on R.id_solicitud = vD.id_Solicitud" _
           & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
           & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
           & " and R.tesoreria is null and R.estado in('A','C') and Td.id_solicitud is null" _
           & " and (R.Emitir in('CK','TE') or isnull(vD.MONTO,0) > 0)"
Else
    strSQL = "select R.id_solicitud,R.codigo,S.cedula,S.nombre,R.montoapr,R.monto_girado,R.TES_SUPERVISION_FECHA,0 as 'Duplicado'" _
           & ", isnull(vD.Numero,0) as 'DESEM_NUM', isnull(vD.MONTO,0) as 'DESEM_MONTO'" _
           & "  from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
           & "    inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
           & "    left join CRD_REMESAS_TES_DETALLE Td on R.id_solicitud = Td.id_Solicitud" _
           & "    left join vCrdOperacion_DesembolsosGiro vD on R.id_solicitud = vD.id_Solicitud" _
           & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
           & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
           & " and R.tesoreria is null and R.estado in('A','C') and Td.id_solicitud is null" _
           & " and (R.Emitir in('CK','TE') or isnull(vD.MONTO,0) > 0)"
           
End If
If cboOficina.Text <> "TODOS" Then
   strSQL = strSQL & " and R.cod_Oficina_R = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
End If

If mRequiereAutorizacion Then
   strSQL = strSQL & " and (R.ANALISTAS_REVISION = 1 or R.AUTORIZA_TRANSFERENCIA = 1)"
End If


If chkFiltros.Value = vbChecked Then
    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And R.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And R.UserFor like '" & cboUsuarios.Text & "%'"
    End If

    If cboSistema.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(R.cod_app,'" & App.ProductName & "') like '" & cboSistema.Text & "%'"
    End If
End If

strSQL = strSQL & " order by id_solicitud"

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
prgBar.Visible = True
vDuplicado = False
strLista = ""

With lswCarga
 .ListItems.Clear
 Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!Id_solicitud)
       If rs!Duplicado = 1 And IsNull(rs!TES_SUPERVISION_FECHA) Then
          itmX.ForeColor = vbRed
          vDuplicado = True
         strLista = strLista & rs!Id_solicitud & " " & rs!codigo & " " & rs!Cedula & " " & Format(rs!Monto_girado, "Standard") & vbCrLf
       Else
          itmX.ForeColor = vbBlack
       End If
       
       itmX.SubItems(1) = rs!codigo
       itmX.SubItems(2) = rs!Cedula
       itmX.SubItems(3) = rs!Nombre
       itmX.SubItems(4) = Format(rs!montoapr, "Standard")
       itmX.SubItems(5) = Format(rs!Monto_girado, "Standard")
       itmX.SubItems(6) = rs!DESEM_NUM
       itmX.SubItems(7) = Format(rs!DESEM_MONTO, "Standard")
       itmX.SubItems(8) = Format(rs!Monto_girado + rs!DESEM_MONTO, "Standard")
       itmX.SubItems(9) = IIf(vDuplicado = True, rs!Duplicado, 0)
   
       itmX.Checked = chkCarga.Value
         
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(8))
       End If
        
        rs.MoveNext
        
        prgBar.Value = prgBar.Value + 1
 Loop
End With

rs.Close

prgBar.Visible = False

txtCargaTotal.Text = Format(curTotal, "Standard")

If vDuplicado = True Then
   MsgBox "Estas operaciones necesitan autorización para ser trasladadas ya que cuentan" _
          & "con una transacción por un monto igual en Tesorería " & vbCrLf & vbCrLf & strLista, vbCritical
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear
 
End Sub


Private Sub sbTrasladoBuscar()
Dim rs2 As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from CRD_REMESAS_TES where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!fecha_inicio
  vFechaCorte = rs!fecha_corte
rs.Close


strSQL = "select R.id_solicitud,R.codigo,S.cedula,S.nombre,R.montoapr,R.monto_girado" _
       & ", isnull(D.Numero,0) as 'Desembolsos_Numero', isnull(D.Monto,0) as 'Desembolsos'" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & "  left join vCrdOperacion_DesembolsosGiro D on R.id_Solicitud = D.id_Solicitud" _
       & " where R.estadosol='F' and R.fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.estado in('A','C') and R.id_solicitud in(select id_solicitud from CRD_REMESAS_TES_DETALLE" _
       & " where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) & ")" _
       & " order by R.id_solicitud"

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
prgBar.Visible = True

With lswTraslado
 .ListItems.Clear
 Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!Id_solicitud)
       itmX.SubItems(1) = rs!codigo
       itmX.SubItems(2) = rs!Cedula
       itmX.SubItems(3) = rs!Nombre
       itmX.SubItems(4) = Format(rs!montoapr, "Standard")
       itmX.SubItems(5) = Format(rs!Monto_girado, "Standard")
       
       itmX.SubItems(6) = rs!Desembolsos_Numero
       itmX.SubItems(7) = Format(rs!Desembolsos, "Standard")
       itmX.SubItems(8) = Format(rs!Monto_girado + rs!Desembolsos, "Standard")
   
       curTotal = curTotal + CCur(itmX.SubItems(8))
       
       rs.MoveNext
       prgBar.Value = prgBar.Value + 1
 Loop

End With

rs.Close

prgBar.Visible = False

txtPagoTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswTraslado.ListItems.Clear

End Sub








Private Sub sbCerrar()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CRD_REMESAS_TES" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
strSQL = "update CRD_REMESAS_TES set estado = 'C'" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("Aplica", "Cierra Remesa Traslado a Tesoreria : " & cboCarga.ItemData(cboCarga.ListIndex))


MsgBox "Remesa Cerrada Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCarga()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CRD_REMESAS_TES" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

'Calcula los casos a procesar
vCasos = 1
For i = 1 To lswCarga.ListItems.Count
 If lswCarga.ListItems.Item(i).Checked Then
    vCasos = vCasos + 1
 End If
Next i

prgBar.Max = vCasos
prgBar.Value = 1
prgBar.Visible = True


With lswCarga.ListItems

For i = 1 To .Count
 If .Item(i).Checked And .Item(i).SubItems(9) = 0 Then
 
     strSQL = "insert CRD_REMESAS_TES_DETALLE(cod_remesa,id_solicitud,monto,desembolsos) values(" _
            & cboCarga.ItemData(cboCarga.ListIndex) & "," & .Item(i).Text & "," & CCur(.Item(i).SubItems(5)) _
            & "," & CCur(.Item(i).SubItems(7)) & ")"
     Call ConectionExecute(strSQL)
   
    prgBar.Value = prgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Carga Remesa Traslado a Tesoreria : " & cboCarga.ItemData(cboCarga.ListIndex))
End If

End With

prgBar.Visible = False

Me.MousePointer = vbDefault

MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation
Call sbCargaBuscar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear
 

End Sub




Private Sub txtCamConOP_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

vGrid.MaxRows = 0

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  strSQL = "select D.id_desembolso,R.id_solicitud,R.codigo,D.monto,D.concepto" _
         & " from reg_creditos R inner join desembolsos D on R.id_solicitud = D.id_solicitud" _
         & " Where D.retener = 0 And R.tesoreria Is Null and R.estadosol = 'F' and R.id_solicitud = " & txtCamConOP.Text

 Call sbCargaGrid(vGrid, 5, strSQL, True)
 
 vGrid.MaxRows = vGrid.MaxRows - 1
 
 If vGrid.MaxRows = 0 Then
    MsgBox "La operación se encuentra en un estado que no es posible el cambio de conceptos, verifique...", vbExclamation
 End If
 
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub txtConsulta_Operacion_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

Me.MousePointer = vbHourglass


If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then

 txtConsulta_Remesa = ""
 txtConsulta_Operacion.Tag = 0
 strSQL = "select R.id_solicitud,R.codigo,R.cedula,R.monto_girado,C.descripcion,S.nombre" _
        & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
        & " inner join Catalogo C on R.codigo = C.codigo" _
        & " where R.id_solicitud = " & txtConsulta_Operacion & " and R.estado in('A','C')"
 Call OpenRecordSet(rs, strSQL)
 
 If Not rs.EOF And Not rs.BOF Then
   
   txtConsulta_Remesa.Tag = rs!codigo
   txtConsulta_Remesa = txtConsulta_Remesa & "Línea         : " & rs!codigo & vbCrLf
   txtConsulta_Remesa = txtConsulta_Remesa & "Descripción   : " & rs!Descripcion & vbCrLf
   txtConsulta_Remesa = txtConsulta_Remesa & "Identificación: " & rs!Cedula & vbCrLf
   txtConsulta_Remesa = txtConsulta_Remesa & "Nombre        : " & rs!Nombre & vbCrLf
   txtConsulta_Remesa = txtConsulta_Remesa & "Monto a Girar : " & Format(rs!Monto_girado, "Standard") & vbCrLf & vbCrLf & vbCrLf
   rs.Close
    
   'Remesa
   strSQL = "select Td.*, T.ESTADO, T.USUARIO , T.FECHA_INICIO , T.FECHA_CORTE, T.FECHA" _
          & "  from CRD_REMESAS_TES T inner join CRD_REMESAS_TES_DETALLE Td on T.COD_REMESA = Td.COD_REMESA" _
          & "  Where Td.Id_solicitud = " & txtConsulta_Operacion.Text
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
      txtConsulta_Remesa = txtConsulta_Remesa & ":::. REMESA DE PAGO .::" & vbCrLf & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Remesa Id      :" & rs!cod_Remesa & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Estado         :"
      Select Case rs!ESTADO
        Case "A"
            txtConsulta_Remesa = txtConsulta_Remesa & "Abierta" & vbCrLf
        Case "C"
            txtConsulta_Remesa = txtConsulta_Remesa & "Cerrada" & vbCrLf
        Case "T"
            txtConsulta_Remesa = txtConsulta_Remesa & "Trasladada" & vbCrLf
      End Select
      txtConsulta_Remesa = txtConsulta_Remesa & "Fecha Creación :" & rs!fecha & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Usuario        :" & rs!Usuario & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Monto          :" & Format(rs!Monto, "Standard") & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Desembolsos Add:" & Format(rs!Desembolsos, "Standard") & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Tesoreria Id   :" & rs!Nsolicitud & vbCrLf & vbCrLf & vbCrLf
   Else
      txtConsulta_Remesa = txtConsulta_Remesa & ">>> REMESA DE PAGO: NO SE LOCALIZO NINGUNA <<<" & vbCrLf & vbCrLf
   End If
   rs.Close
   
   'Bancos
   strSQL = "select T.NSOLICITUD,T.id_banco,T.ndocumento,T.BENEFICIARIO, T.FECHA_SOLICITUD, T.FECHA_EMISION , T.ESTADO " _
          & ", '[' + B.CTA  + '] ' + B.DESCRIPCION   as 'Cuenta_Desc' , Bg.DESCRIPCION as 'Banco_Desc', Td.DESCRIPCION as 'Tipo_Desc'" _
          & "  from Tes_Transacciones T inner join TES_BANCOS B on T.ID_BANCO = B.ID_BANCO" _
          & "      inner join TES_BANCOS_GRUPOS Bg on B.COD_GRUPO = Bg.COD_GRUPO" _
          & "      inner join TES_TIPOS_DOC Td on T.tipo = Td.TIPO" _
          & " where T.op = " & txtConsulta_Operacion & " and T.estado in('I','T','P', 'E')"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
      txtConsulta_Remesa = txtConsulta_Remesa & ":::. BANCOS .::" & vbCrLf & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Solicitud    :" & rs!Nsolicitud & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Documento    :" & rs!ndocumento & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Beneficiario :" & rs!Beneficiario & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Banco        :" & rs!BANCO_DESC & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Cuenta       :" & rs!Cuenta_Desc & vbCrLf
      txtConsulta_Remesa = txtConsulta_Remesa & "Tipo         :" & rs!Tipo_Desc & vbCrLf
   End If
   rs.Close

   
 
 Else
   
   MsgBox "La Operacion Digitada no existe...", vbExclamation
 
 End If

End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If

End Sub


Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Integer, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vCodApp As String _
                              , vOficina As String, vUnidad As String, vToken As String, vRemesaTipo As String, vRemesa As Integer _
                              , Optional vRef_01 As String = "", Optional vRef_02 As String = "" _
                              , Optional vRef_03 As String = "") As Long   'Regresa el NSOLICITUD
                              
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long
 
strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza" _
       & ",fecha_autorizacion,Cod_App,Ref_01,Ref_02,Ref_03,ID_TOKEN,REMESA_TIPO,REMESA_ID)" _
       & " values('GEN','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "','" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CC','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate(),'" & vCodApp & "','" & vRef_01 & "','" & vRef_02 & "','" & vRef_03 _
                   & "','" & vToken & "','" & vRemesaTipo & "' ," & vRemesa & ")"
Else
   strSQL = strSQL & ",'N',null,null,'" & vCodApp & "','" & vRef_01 & "','" & vRef_02 & "','" & vRef_03 & "','" & vToken _
                   & "','" & vRemesaTipo & "' ," & vRemesa & ")"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
rsX.Open strSQL, glogon.Conection, adOpenStatic
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

rsX.Open strSQL, glogon.Conection, adOpenStatic
If Trim(rsX!codigo) = Trim(vCodigo) Then lngSol = rsX!Nsolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "' and op = " & vOP
  rsX.CursorLocation = adUseServer
  rsX.Open strSQL, glogon.Conection, adOpenStatic
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidad & "')"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCtaBanco(pBanco As Integer) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & pBanco
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function


Private Function fxCtaPuente(pCodigo As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select ctaPuente from catalogo where codigo ='" & pCodigo & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
     fxCtaPuente = ""
Else
     fxCtaPuente = rsX!ctapuente
End If

rsX.Close

End Function


Private Sub sbCreaDesembolsos(vReferencia As Long, vOP As Long, vFecha As Date, vTipo As String, vBanco As Integer _
                             , vCod_App As String, vOficina As String, vUnidad As String, vToken As String _
                             , vRemesaTipo As String, vRemesa As Integer)
Dim rsTemp As New ADODB.Recordset, lngSolicitud As Long

strSQL = "select * from desembolsos where retener = 0 and id_solicitud = " & vOP


With rsTemp
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
 Do While Not .EOF
     lngSolicitud = fxMaestroTesoreria(vTipo, vBanco, !Monto, !id_desembolso _
                   , !Concepto, !Id_solicitud, !Id_solicitud, vReferencia, !codigo, "0" _
                   , vFecha, vCod_App, vOficina, vUnidad, vToken, vRemesaTipo, vRemesa _
                   , CStr(vOP), CStr(!id_desembolso))
     
     Call sbCreaDetalle(lngSolicitud, fxCtaBanco(vBanco), !Monto, "H", 1, vUnidad)
     Call sbCreaDetalle(lngSolicitud, !cuenta_conta, !Monto, "D", 2, vUnidad)
     
     strSQL = "update desembolsos set tdocumento = '" & vTipo & "',cod_banco = " & vBanco & ",nsolicitud = " & lngSolicitud _
            & " where id_desembolso = " & !id_desembolso
     Call ConectionExecute(strSQL)
  .MoveNext
 Loop
 .Close
End With

End Sub

Private Sub sbTraslado()
Dim lngSolicitud As Long, vFecha As Date
Dim vTipo As String, vBanco As Integer
Dim vToken As String

Me.MousePointer = vbHourglass

On Error GoTo vError

vFecha = fxFechaServidor
strSQL = "select top 1 id_token from tes_tokens where estado = 'A' order by registro_fecha "
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  vToken = rs!ID_TOKEN
Else
  vToken = fxCreaToquen
End If
rs.Close




'strSQL = "select cod_parametro,valor from crd_parametros where cod_parametro in('11','12')"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
'  Select Case Trim(rs!cod_parametro)
'    Case "11"
'       vBanco = Trim(rs!valor)
'    Case "12"
'       vTipo = Trim(rs!valor)
'  End Select
'  rs.MoveNext
'Loop
'rs.Close


'strSQL = "select R.id_solicitud,R.codigo,S.cedula,S.nombre,R.emitir,R.cod_banco,R.monto_girado,R.cta_banco,isnull(R.cod_app,'S.I.F.') as 'Cod_App'" _
'       & ",Ofi.cod_Oficina,Ofi.Cod_Unidad" _
'       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
'       & " inner join SIF_Oficinas Ofi on R.cod_oficina_r = Ofi.cod_Oficina" _
'       & " where R.estado in('A','C') and R.estadosol = 'F' and R.tesoreria is null" _
'       & " and R.id_solicitud in(select id_solicitud from CRD_REMESAS_TES_DETALLE where cod_remesa = " _
'       & cboTraslado.ItemData(cboTraslado.ListIndex) & ")"

strSQL = "select id_solicitud,codigo" _
       & " from reg_creditos" _
       & " where estado in('A','C') and estadosol = 'F' and tesoreria is null" _
       & " and id_solicitud in(select id_solicitud from CRD_REMESAS_TES_DETALLE where cod_remesa = " _
       & cboTraslado.ItemData(cboTraslado.ListIndex) & ")"

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
prgBar.Visible = True


Do While Not rs.EOF

    'Nuevo Proceso (Integrado)
    strSQL = "exec spCrdCreditoEnviaTesoreria_Todo " & rs!Id_solicitud & ",'" & vToken & "'," & cboTraslado.ItemData(cboTraslado.ListIndex) & ",'CRD'"
    Call ConectionExecute(strSQL)
 
    'Actualiza Bitacora
    Call Bitacora("Registra", "Traspaso a Tesoreria de la Operacion y Desembol OP:" & rs!Id_solicitud)
     
    'Tags de Seguimiento
    Call sbCrdOperacionTags(rs!Id_solicitud, rs!codigo, "S04", "", "Remesa de Traslado No..:" & cboTraslado.ItemData(cboTraslado.ListIndex))

'' 'Graba y Devuelve el registro Maestro en Tesoreria
''
'' If rs!Monto_girado > 0 And (rs!Emitir = "CK" Or rs!Emitir = "TE") Then
''    lngSolicitud = fxMaestroTesoreria(rs!Emitir, rs!cod_banco, rs!Monto_girado, Trim(rs!CEDULA) _
''                   , rs!Nombre, rs!Id_solicitud, rs!Id_solicitud, 0 _
''                   , rs!Codigo, rs!CTA_BANCO, vFecha, rs!Cod_App, rs!cod_oficina, rs!cod_unidad _
''                   , vToken, "CRD", cboTraslado.ItemData(cboTraslado.ListIndex), rs!Id_solicitud, rs!Codigo, rs!cod_oficina)
''
''    'Mata el Pasivo de la Nota de Debito de la Formalizacion contra Tes_Bancos
''    Call sbCreaDetalle(lngSolicitud, fxCtaBanco(rs!cod_banco), rs!Monto_girado, "H", 1, rs!cod_unidad)
''    Call sbCreaDetalle(lngSolicitud, fxCtaPuente(rs!Codigo), rs!Monto_girado, "D", 2, rs!cod_unidad)
''
''
'' Else 'Monto a Girar > 0
''
''   lngSolicitud = 0
''
'' End If
''
''  'Crea los Documentos de Desembolsos, para Ambos Procesos el Procedimiento es el mismo
''  Call sbCreaDesembolsos(lngSolicitud, rs!Id_solicitud, vFecha, vTipo, vBanco, rs!Cod_App, rs!cod_oficina, rs!cod_unidad _
''                        , vToken, "CRD", cboTraslado.ItemData(cboTraslado.ListIndex))
''
''
''
'' 'Actualiza Campo Tesoreria
''  strSQL = "update reg_creditos set tesoreria = dbo.MyGetdate(),ID_TOKEN = '" & vToken & "'" _
''         & " where id_solicitud = " & rs!Id_solicitud
''  Call ConectionExecute(strSQL)
''
''
'' 'Actualiza  Remesa Detalle
''  strSQL = "update CRD_REMESAS_TES_DETALLE SET id_banco = " & rs!cod_banco & ",nsolicitud = " & lngSolicitud _
''         & "  Where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
''         & "  and id_solicitud = " & rs!Id_solicitud
''  Call ConectionExecute(strSQL)

  
 If prgBar.Max > prgBar.Value Then prgBar.Value = prgBar.Value + 1
 rs.MoveNext
 
Loop
rs.Close

'Actualiza y Carga Remesa
strSQL = "update CRD_REMESAS_TES SET Estado = 'T'" _
       & "  Where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call ConectionExecute(strSQL)

Call sbLimpia


Me.MousePointer = vbDefault

prgBar.Visible = False

MsgBox "Operaciones Enviadas a Tesoreria Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 If rs.State = 1 Then rs.Close

End Sub


Private Sub cmdReactivar_Click()

On Error GoTo vError

If txtOperacion.Tag = 1 Then
  
  strSQL = "update reg_creditos set tesoreria = null where id_solicitud = " & txtOperacion
  Call ConectionExecute(strSQL)
  
  
  Call Bitacora("Aplica", "ReActivacion Traslado Tes. Op:" & txtOperacion)
  
  'Tags de Seguimiento
  Call sbCrdOperacionTags(txtOperacion.Text, txtDetalle.Tag, "S04", "", ">>> Re.Activación del Desembolso <<<")
  
  txtOperacion = 0
  txtOperacion.Tag = 0
  txtDetalle = ""

  MsgBox "Operación ReActivada Satisfactoriamente...", vbInformation

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbReportePendientes()
Dim strTitulo As String
Dim strRuta As String, strInicio As String, strFinal As String


On Error GoTo vError

Me.MousePointer = vbHourglass

strTitulo = "Operaciones pendientes de Traslado a Tesorería"


strRuta = SIFGlobal.fxPathReportes("Credito_SGTEnvioTesoreria.rpt")
strInicio = "Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")"
strFinal = "Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     
     .Connect = glogon.ConectRPT
     
     .WindowTitle = "Solicitudes a trasladar a Tesorería"
     
    .ReportFileName = strRuta
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy hh:mm:ss") & "'"
    .Formulas(3) = "Usuario ='" & glogon.Usuario & "'"
    .Formulas(4) = "Titulo='" & strTitulo & "'"
    
    strSQL = "{REG_CREDITOS.ESTADOSOL} = 'F'"
    If chkRepFechas.Value = vbUnchecked Then
      strSQL = strSQL & " and {REG_CREDITOS.FECHAFORP} >= Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ")" _
             & " and {REG_CREDITOS.FECHAFORP} <= Date(" & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
        .Formulas(5) = "de='" & Format(dtpRepInicio.Value, "dd/mm/yyyy") & "'"
        .Formulas(6) = "a='" & Format(dtpRepCorte.Value, "dd/mm/yyyy") & "'"
    Else
        .Formulas(5) = "de=' --- '"
        .Formulas(6) = "a=' --- '"
    End If
    
    
    If cboRepOficina.Text <> "TODOS" Then
       strSQL = strSQL & " AND {REG_CREDITOS.COD_OFICINA_R} = '" & SIFGlobal.fxCodText(cboRepOficina.Text) & "'"
    End If
    
    
    strSQL = strSQL & " and ISNULL({REG_CREDITOS.TESORERIA}) AND {REG_CREDITOS.ESTADO}='A'"
    
    .SelectionFormula = strSQL
    
    .SubreportToChange = "subCkDesembolsos"
    .SelectionFormula = "{DESEMBOLSOS.ID_SOLICITUD} = {?Pm-REG_CREDITOS.ID_SOLICITUD} AND {DESEMBOLSOS.RETENER} = 0"
    
    .PrintReport
    

End With

 Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporteEnviadas()

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "OPERACIONES ENVIADAS A TESORERIA"

 .Connect = glogon.ConectRPT

.ReportFileName = SIFGlobal.fxPathReportes("Credito_SGTEnvioTesoreriaRec.rpt")
.Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
.Formulas(2) = "fxFecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
.Formulas(3) = "fxTitulo='Desembolsos Solicitados en Tesorería'"
.Formulas(4) = "fxUsuario='" & glogon.Usuario & "'"
.Formulas(5) = "fxSubTitulo='INICIO : " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " CORTE : " & Format(dtpRepCorte.Value, "dd/mm/yyyy") & "'"

strSQL = "{TES_TRANSACCIONES.FECHA_SOLICITUD} in date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ") to date(" _
    & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ") and {TES_TRANSACCIONES.MODULO} ='CC'"

.SelectionFormula = strSQL
.Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

On Error GoTo vError

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

 tcMain.Item(0).Selected = True
 
 With lswRemesas.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Estado", 1400
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Casos", 1000, vbRightJustify
    .Add , , "Monto", 2400, vbRightJustify
    .Add , , "Notas", 3400
 End With
 
 With lswRep.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Estado", 1400, vbCenter
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Casos", 1000, vbRightJustify
    .Add , , "Monto", 2400, vbRightJustify
    .Add , , "Notas", 3400
 End With
  
 
 With lswCarga.ColumnHeaders
    .Clear
    .Add , , "No.Operación", 1400
    .Add , , "Línea", 1000, vbCenter
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3400
    .Add , , "Aprobado", 1800, vbRightJustify
    .Add , , "A Girar", 1800, vbRightJustify
    .Add , , "Desembolsos", 1200, vbCenter
    .Add , , "Otros Giros", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Duplicado?", 1200, vbCenter
 End With
 
 
 With lswTraslado.ColumnHeaders
    .Clear
    .Add , , "No.Operación", 1400
    .Add , , "Línea", 1000, vbCenter
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3400
    .Add , , "Aprobado", 1800, vbRightJustify
    .Add , , "A Girar", 1800, vbRightJustify
    .Add , , "Desembolsos", 1200, vbCenter
    .Add , , "Otros Giros", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Duplicado?", 1200, vbCenter
 End With
 
 
 Call Formularios(Me)
 
 btnBarra(9).Tag = btnBarra(0).Tag
 
 Call RefrescaTags(Me)
 
 Call sbLimpia
 Call sbRequiereAutorizacion
 
strSQL = "select rtrim(cod_oficina) as 'Idx', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbCbo_Llena_New(cboRepOficina, strSQL, True, False)
 
Exit Sub

vError:

 
End Sub


Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rsTmp As New ADODB.Recordset


If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 txtDetalle = ""
 txtOperacion.Tag = 0
 strSQL = "select R.id_solicitud,R.codigo,R.cedula,R.monto_girado,C.descripcion,S.nombre, R.Emitir" _
        & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
        & " inner join Catalogo C on R.codigo = C.codigo" _
        & " where R.id_solicitud = " & txtOperacion & " and R.estado = 'A'"
 Call OpenRecordSet(rs, strSQL)
 If Not rs.EOF And Not rs.BOF Then
   txtDetalle.Tag = rs!codigo
   txtDetalle = txtDetalle & "Línea         : " & rs!codigo & vbCrLf
   txtDetalle = txtDetalle & "Descripción   : " & rs!Descripcion & vbCrLf
   txtDetalle = txtDetalle & "Identificación: " & rs!Cedula & vbCrLf
   txtDetalle = txtDetalle & "Nombre        : " & rs!Nombre & vbCrLf
   txtDetalle = txtDetalle & "Monto a Girar : " & Format(rs!Monto_girado, "Standard") & vbCrLf
   
   Select Case rs!Emitir
        Case "TE", "CK"
                'Verifica que no existan documentos emitidos con anterioridad
                strSQL = "select NSOLICITUD,id_banco,tipo,ndocumento from Tes_Transacciones" _
                       & " where op = " & txtOperacion & " and estado in('I','T','P')"
                rsTmp.Open strSQL, glogon.Conection, adOpenStatic
                If rsTmp.EOF And rsTmp.BOF Then
                   txtOperacion.Tag = 1
                   'Verificar si tiene desembolsos asociados en Tesoreria
                   rsTmp.Close
                   strSQL = "select NSOLICITUD,id_banco,tipo,ndocumento from Tes_Transacciones" _
                          & " where op = " & txtOperacion & " and estado in('I','T','P')"
                   rsTmp.Open strSQL, glogon.Conection, adOpenStatic
                   If Not rsTmp.EOF And Not rsTmp.BOF Then
                      txtOperacion.Tag = 0
                      txtDetalle = txtDetalle & "EXISTEN DESEMBOLSOS ASOCIADOS EN TESORERIA" & vbCrLf
                   End If
                
                Else 'Mov del Deudor Directamente
                   txtOperacion.Tag = 0
                   txtDetalle = txtDetalle & " / EXISTE UN DOCUMENTO O SOLICITUD DE EMISION EN TESORERIA / " & rs!Nombre & vbCrLf
                   txtDetalle = txtDetalle & "Solicitud :" & rsTmp!Nsolicitud & vbCrLf
                   txtDetalle = txtDetalle & "Documento :" & rsTmp!ndocumento & vbCrLf
                   txtDetalle = txtDetalle & "Tipo/Banco:" & rsTmp!Tipo & "/" & rsTmp!id_banco & vbCrLf
                End If
                
                rsTmp.Close
        Case "RC" 'Retiro en Cajas
                   txtOperacion.Tag = 0
                   txtDetalle = txtDetalle & " ** EL TIPO DE EMISION -Retiro de Efectivo en Cajas - NO PERMITE REACTIVACION"
        Case "CP" 'Pago a Proveedor
                   txtOperacion.Tag = 0
                   txtDetalle = txtDetalle & " ** EL TIPO DE EMISION -Pago a Proveedor - NO PERMITE REACTIVACION"
        Case Else
                   txtOperacion.Tag = 0
                   txtDetalle = txtDetalle & " ** EL TIPO DE EMISION ACTUAL NO PERMITE REACTIVACION"
        
    End Select 'Emitir
   
   'Segunda verificacion con el nuevo esquema
   
   If CLng(txtOperacion.Tag) = 1 Then
    'Verificar AQUI; pero como deseable porque el codigo nuevo es compatible con este codigo de verificacion
   
   
   
   End If
   
 
 Else
   
   MsgBox "La Operacion Digitada no existe...", vbExclamation
 
 End If
 rs.Close

End If

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vOP As Long, vCodigo As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
   'Actualiza Desembolso
   vGrid.Row = vGrid.ActiveRow
   
   
   vGrid.Col = 2
   vOP = vGrid.Text
   vGrid.Col = 3
   vCodigo = vGrid.Text
   
   vGrid.Col = 5
   strSQL = "update desembolsos set concepto = '" & vGrid.Text & "' where id_desembolso = "
   vGrid.Col = 1
   strSQL = strSQL & vGrid.Text
   Call ConectionExecute(strSQL)

   Call sbBitacoraCredito("15", "Cambio de Concepto de Desembolso de CRD ID:" & vGrid.Text, "C", vOP, vCodigo)
   
   MsgBox "Cambio de Concepto Realizado Satisfactoriamente...", vbInformation

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbRequiereAutorizacion()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    strSQL = "select isnull(valor,'') from CRD_PARAMETROS where cod_parametro = '27'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs.Fields(0) = "S" Then
            mRequiereAutorizacion = True
        Else
            mRequiereAutorizacion = False
        End If
    Else
        mRequiereAutorizacion = False
    End If
    rs.Close
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxSupervisaBanco() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = " Select isnull(SUPERVISION,0) as 'SUPERVISION' from tes_bancos where id_banco =  " & cboBanco.ItemData(cboBanco.ListIndex)
Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs!SUPERVISION > 0 Then
           fxSupervisaBanco = True
        Else
           fxSupervisaBanco = False
        End If
    Else
      fxSupervisaBanco = False
    End If
rs.Close
End Function


Private Function fxCreaToquen() As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim strToken As String

strToken = Format(fxFechaServidor, "yyyy.mm.dd")


strSQL = "select  isnull(COUNT(id_token),0)+ 1 as 'consec'  from tes_tokens where id_token like('" & strToken & "%')"
Call OpenRecordSet(rs, strSQL)

strToken = strToken & "." & rs!Consec

rs.Close

strSQL = "insert tes_tokens(id_token,registro_fecha,registro_usuario,estado)" _
      & "values('" & strToken & "',dbo.MyGetdate(),'" & glogon.Usuario & "','A') "
Call ConectionExecute(strSQL)

fxCreaToquen = strToken

End Function

