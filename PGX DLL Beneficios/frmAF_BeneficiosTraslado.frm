VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmAF_BeneficiosTraslado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Beneficios a Tesorería"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   12075
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   150
      Left            =   0
      TabIndex        =   67
      Top             =   7800
      Visible         =   0   'False
      Width           =   12135
      _Version        =   1572864
      _ExtentX        =   21405
      _ExtentY        =   265
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.GroupBox fraFiltros 
      Height          =   2295
      Left            =   1800
      TabIndex        =   60
      Top             =   3720
      Width           =   7935
      _Version        =   1572864
      _ExtentX        =   13991
      _ExtentY        =   4043
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
         TabIndex        =   61
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboUsuarios 
         Height          =   312
         Left            =   1800
         TabIndex        =   62
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   420
         Index           =   0
         Left            =   6000
         TabIndex        =   63
         Top             =   1680
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   420
         Index           =   1
         Left            =   7320
         TabIndex        =   64
         Top             =   1680
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":0700
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   12
         Left            =   360
         TabIndex        =   66
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
         Index           =   11
         Left            =   360
         TabIndex        =   65
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
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   12135
      _Version        =   1572864
      _ExtentX        =   21405
      _ExtentY        =   11456
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
      ItemCount       =   4
      SelectedItem    =   3
      Item(0).Caption =   "Remesa"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "Label8(1)"
      Item(0).Control(1)=   "Label8(2)"
      Item(0).Control(2)=   "Label8(3)"
      Item(0).Control(3)=   "Label8(4)"
      Item(0).Control(4)=   "Label8(5)"
      Item(0).Control(5)=   "Label8(6)"
      Item(0).Control(6)=   "txtRemesa"
      Item(0).Control(7)=   "txtFecha"
      Item(0).Control(8)=   "txtEstado"
      Item(0).Control(9)=   "txtUsuario"
      Item(0).Control(10)=   "txtNotas"
      Item(0).Control(11)=   "dtpInicio"
      Item(0).Control(12)=   "dtpCorte"
      Item(0).Control(13)=   "btnBarra(0)"
      Item(0).Control(14)=   "btnBarra(1)"
      Item(0).Control(15)=   "btnBarra(2)"
      Item(0).Control(16)=   "lswRemesas"
      Item(0).Control(17)=   "fraReporte"
      Item(0).Control(18)=   "Labe9(0)"
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
      Item(2).Control(1)=   "Label2(16)"
      Item(2).Control(2)=   "cboTraslado"
      Item(2).Control(3)=   "lswTraslado"
      Item(2).Control(4)=   "txtPagoTotal"
      Item(2).Control(5)=   "btnBarra(6)"
      Item(2).Control(6)=   "btnBarra(7)"
      Item(2).Control(7)=   "Label8(19)"
      Item(3).Caption =   "Informes"
      Item(3).ControlCount=   9
      Item(3).Control(0)=   "opt(0)"
      Item(3).Control(1)=   "txtRepRemesas"
      Item(3).Control(2)=   "Label16(2)"
      Item(3).Control(3)=   "lblRemesa"
      Item(3).Control(4)=   "opt(1)"
      Item(3).Control(5)=   "Label16(4)"
      Item(3).Control(6)=   "chkRemesaInd"
      Item(3).Control(7)=   "lswRep"
      Item(3).Control(8)=   "btnBarra(8)"
      Begin XtremeSuiteControls.ListView lswRemesas 
         Height          =   3132
         Left            =   -68440
         TabIndex        =   2
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
         TabIndex        =   3
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
         Height          =   4092
         Left            =   -69880
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   11412
         _Version        =   1572864
         _ExtentX        =   20129
         _ExtentY        =   7218
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
         Left            =   0
         TabIndex        =   5
         Top             =   840
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
         Left            =   480
         TabIndex        =   6
         Top             =   5160
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
      Begin XtremeSuiteControls.GroupBox fraReporte 
         Height          =   2055
         Left            =   -65800
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   7455
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   0
            Left            =   1920
            TabIndex        =   8
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
            Left            =   4560
            TabIndex        =   9
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.DateTimePicker dtpRepCorte 
            Height          =   315
            Left            =   3120
            TabIndex        =   10
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
            Left            =   1800
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   720
            Width           =   4932
            _Version        =   1572864
            _ExtentX        =   8705
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
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   0
            Left            =   5760
            TabIndex        =   13
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
            Picture         =   "frmAF_BeneficiosTraslado.frx":0E00
         End
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   1
            Left            =   6360
            TabIndex        =   14
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
            Picture         =   "frmAF_BeneficiosTraslado.frx":1507
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   1
            Left            =   3600
            TabIndex        =   15
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
            Picture         =   "frmAF_BeneficiosTraslado.frx":1B45
            ToolTipText     =   "Actualizar Oficinas"
            Top             =   360
            Width           =   240
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   372
            Index           =   8
            Left            =   360
            TabIndex        =   17
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
            TabIndex        =   16
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
         TabIndex        =   18
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
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_BeneficiosTraslado.frx":2235
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa 
         Height          =   432
         Left            =   -68440
         TabIndex        =   19
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   312
         Left            =   -68440
         TabIndex        =   20
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   -64840
         TabIndex        =   21
         Top             =   1320
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   -64840
         TabIndex        =   22
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   -68440
         TabIndex        =   23
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -68440
         TabIndex        =   24
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
         TabIndex        =   25
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
         Left            =   -64360
         TabIndex        =   26
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":2867
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   2
         Left            =   -63880
         TabIndex        =   27
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":2E0B
      End
      Begin XtremeSuiteControls.ComboBox cboCarga 
         Height          =   312
         Left            =   -67600
         TabIndex        =   28
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   312
         Left            =   -67600
         TabIndex        =   29
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkFiltros 
         Height          =   372
         Left            =   -67600
         TabIndex        =   30
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
         TabIndex        =   31
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
         Left            =   -67840
         TabIndex        =   32
         Top             =   600
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   3
         Left            =   -63880
         TabIndex        =   33
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":3512
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   4
         Left            =   -62560
         TabIndex        =   34
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":3C12
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   5
         Left            =   -61240
         TabIndex        =   35
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":432B
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   6
         Left            =   -63040
         TabIndex        =   36
         Top             =   960
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":4A37
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   7
         Left            =   -61720
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":5137
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   8
         Left            =   9240
         TabIndex        =   38
         Top             =   5640
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
         Picture         =   "frmAF_BeneficiosTraslado.frx":5A08
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   312
         Left            =   10800
         TabIndex        =   39
         Top             =   4560
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
         Left            =   9360
         TabIndex        =   40
         Top             =   5040
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
         Left            =   480
         TabIndex        =   41
         Top             =   5520
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
         TabIndex        =   42
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
         Left            =   -60880
         TabIndex        =   43
         Top             =   6000
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   19
         Left            =   -62680
         TabIndex        =   59
         Top             =   6000
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
         Height          =   252
         Index           =   18
         Left            =   -62560
         TabIndex        =   58
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
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remesas - visualizar últimas"
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
         Height          =   300
         Index           =   4
         Left            =   5400
         TabIndex        =   57
         Top             =   4560
         Width           =   5412
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   300
         Left            =   120
         TabIndex        =   56
         Top             =   4560
         Width           =   5292
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
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
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   55
         Top             =   600
         Width           =   11652
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lista de Operaciones Pendientes de Traslado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   16
         Left            =   -69880
         TabIndex        =   54
         Top             =   1560
         Visible         =   0   'False
         Width           =   11412
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   14
         Left            =   -69160
         TabIndex        =   53
         Top             =   600
         Visible         =   0   'False
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
      Begin XtremeSuiteControls.Label Labe9 
         Height          =   372
         Index           =   0
         Left            =   -69400
         TabIndex        =   44
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Traspaso de Beneficios a Bancos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12372
   End
End
Attribute VB_Name = "frmAF_BeneficiosTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Dim itmX As ListViewItem, vPaso As Boolean

Dim mRequiereAutorizacion As Boolean
Dim vCedula As String, vNombre As String, vTipo As String
Dim vDuplicado As Boolean
Dim strLista  As String


Private Sub btnBarra_Click(Index As Integer)
Dim i As Integer

On Error GoTo vError

Select Case Index
  Case 0 'NUEVO"
     
    If txtRemesa.Text = "" Then
     
            strSQL = "select isnull(max(cod_remesa),0) + 1 as Ultimo from AFI_BENEFICIOS_REMESAS"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert AFI_BENEFICIOS_REMESAS(cod_remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas) values(" & rs!Ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa.Text = rs!Ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de Beneficios Traslado a Tesoreria: " & txtRemesa)
    
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update AFI_BENEFICIOS_REMESAS set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where cod_remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
             
            Call Bitacora("Modifica", "Remesa de Beneficios Traslado a Tesoreria: " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
    
  Case 1 'BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Remesa Abierta" Then
            strSQL = "update  afi_bene_otorga set COD_REMESA = NULL where Cod_Remesa = " & txtRemesa.Text
            
            strSQL = strSQL & Space(10) & "delete AFI_BENEFICIOS_REMESAS where Cod_Remesa = " & txtRemesa.Text
            Call ConectionExecute(strSQL)
            
            
            Call Bitacora("Elimina", "Remesa de Beneficios Traslado a Tesoreria: " & txtRemesa)
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
            Call sbReporte("S")
          Case optReporte.Item(1).Value
            Call sbReporte("E")
        End Select
    Case 1 'Cerrar
      fraReporte.Visible = False
End Select
End Sub

Private Sub sbReporte(strEstado As String)
Dim strTitulo As String
Dim strRuta As String, strInicio As String, strFinal As String


On Error GoTo vError

Me.MousePointer = vbHourglass

strTitulo = "Beneficios  pendientes de Traslado a Tesorería"


strRuta = SIFGlobal.fxPathReportes("Beneficios_Pendientes.rpt")
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
     
     .WindowTitle = "Beneficios a trasladar a Tesorería"
     
    .ReportFileName = strRuta
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Titulo='" & strTitulo & "'"
    
    strSQL = "{AFI_BENE_PAGO.ESTADO} = '" & strEstado & "'"
    If chkRepFechas.Value = vbUnchecked Then
      strSQL = strSQL & " and {AFI_BENE_OTORGA.REGISTRA_FECHA} >= Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ")"
            
        .Formulas(4) = "de='" & Format(dtpRepInicio.Value, "dd/mm/yyyy") & "'"
        .Formulas(5) = "a='" & Format(dtpRepCorte.Value, "dd/mm/yyyy") & "'"
    Else
        .Formulas(4) = "de=' --- '"
        .Formulas(5) = "a=' --- '"
    End If
    
    
    If cboRepOficina.Text <> "TODOS" Then
       strSQL = strSQL & " AND {AFI_BENE_OTORGA.COD_OFICINA} = '" & SIFGlobal.fxCodText(cboRepOficina.Text) & "'"
    End If
    
    If strEstado = "S" Then
      strSQL = strSQL & " and ISNULL({AFI_BENE_PAGO.TESORERIA}))"
    Else
      strSQL = strSQL & " and {AFI_BENE_PAGO.TESORERIA} > 0)"
    End If
    
    .SelectionFormula = strSQL
    .PrintReport
    

End With

 Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboCarga_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
If cboCarga.Text = "" Then Exit Sub
strSQL = "select fecha_inicio,fecha_corte from AFI_BENEFICIOS_REMESAS where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close


'Carga Oficinas
strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas  where cod_oficina in(" _
       & " select cod_oficina" _
       & " from afi_bene_otorga" _
       & " where Estado = 'S' and registra_fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and (cod_remesa is null or cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " ) group by cod_oficina)"
       
If fxSIFParametros("16") = "S" Then
   strSQL = strSQL & " and Analista_Revision = 'S'"
End If

strSQL = strSQL & " order by cod_oficina"

Call sbCbo_Llena_New(cboOficina, strSQL, True, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbConsulta(pRemesa As Long)

Call sbLimpia
  
strSQL = "select * from AFI_BENEFICIOS_REMESAS where Cod_Remesa = " & pRemesa
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa = CStr(rs!cod_remesa)
  txtUsuario = rs!Usuario
  txtFecha = rs!fecha
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Remesa Abierta"
    Case "C"
      txtEstado = "Remesa Cerrada"
    Case "T"
      txtEstado = "Remesa Trasladada"
  End Select
  
  dtpInicio.Value = rs!Fecha_Inicio
  dtpCorte.Value = rs!Fecha_Corte
  
  txtNotas.Text = rs!notas
  
End If
rs.Close

End Sub



Private Sub sbFiltros()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select fecha_inicio,fecha_corte from AFI_BENEFICIOS_REMESAS where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close


'Cargado de Bancos

strSQL = "Select B.cod_Banco as 'Idx',TB.descripcion as 'Itmx'" _
         & " from afi_bene_pago B" _
         & " inner join tes_bancos TB on B.cod_banco = TB.id_banco " _
         & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio" _
         & " and B.consec = O.consec and  registra_fecha between '" & Format(vFechaInicio, "yyyymmdd") & "  00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
         & " where B.ESTADO = 'S' and B.tesoreria is null group by B.cod_Banco,TB.descripcion"
Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
        

'Cargado de Usuarios
strSQL = "Select O.Registra_User as 'IdX',O.Registra_User as 'Itmx'" _
         & " from afi_bene_pago B" _
         & " inner join tes_bancos TB on B.cod_banco = TB.id_banco " _
         & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio" _
         & " and B.consec = O.consec and  registra_fecha between '" & Format(vFechaInicio, "yyyymmdd") & "' and '" & Format(vFechaCorte, "yyyymmdd") & "'" _
         & " where B.ESTADO = 'S' and B.tesoreria is null"
Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)


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
 .WindowTitle = "Módulo de Beneficios"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Traslado a Tesoreria")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If
     
 Select Case True
  Case opt.Item(0).Value 'Pendiente  Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Beneficios_RemesaGrDet.rpt")
     strSQL = "{AFI_BENEFICIOS_REMESAS.COD_REMESA} = " & lblRemesa.Tag & "" _
           & " and {AFI_BENEFICIOS_REMESAS.Estado} = 'C'  "
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Traslado  Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Beneficios_RemesaTraslado.rpt")
     strSQL = "{AFI_BENEFICIOS_REMESAS.COD_REMESA} = " & lblRemesa.Tag & "" _
           & " and {AFI_BENEFICIOS_REMESAS.Estado} = 'T' "
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA TRASLADO A TESORERIA : BENEFICIOS'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = strSQL
 
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
       & " and R.tesoreria is null and R.estado in('A','C') and id_solicitud not in(select id_solicitud from AFI_BENEFICIOS_REMESAS_DETALLE)" _
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
   curTotal = curTotal + CCur(Item.SubItems(3))
Else
   curTotal = curTotal - CCur(Item.SubItems(3))
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
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(3))
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
     txtEstado = ""
     txtFecha = ""
     txtUsuario = ""
     txtRemesa = ""
     
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    
    dtpRepInicio.Value = dtpInicio.Value
    dtpRepCorte.Value = dtpInicio.Value
    
    fraReporte.Visible = False
    
    txtNotas.Text = ""
     
     strSQL = "select TOP 50 * from AFI_BENEFICIOS_REMESAS order by fecha desc"
     
     
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!cod_remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                
                Select Case rs!Estado
                  Case "A"
                     itmX.SubItems(3) = "Remesa Abierta"
                  Case "C"
                     itmX.SubItems(3) = "Remesa Cerrada"
                  Case "T"
                     itmX.SubItems(3) = "Remesa Trasladada"
                End Select
                
                itmX.SubItems(4) = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!Fecha_Corte, "dd/mm/yyyy")
                itmX.SubItems(6) = rs!notas
                
                
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
        
    strSQL = "select * from AFI_BENEFICIOS_REMESAS where estado = 'A' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
      
      cboCarga.ItemData(cboCarga.ListCount - 1) = CStr(rs!cod_remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click
    Call chkFiltros_Click
   
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
        
    strSQL = "select * from AFI_BENEFICIOS_REMESAS where estado = 'C' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboTraslado.AddItem (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
      cboTraslado.ItemData(cboTraslado.ListCount - 1) = CStr(rs!cod_remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboTraslado.Text = (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboTraslado_Click

  
  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " * from AFI_BENEFICIOS_REMESAS order by fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!cod_remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                itmX.SubItems(3) = rs!Fecha_Inicio
                itmX.SubItems(4) = rs!Fecha_Corte
                itmX.SubItems(5) = rs!notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close

    
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

Dim strNombre As String, vTipo As String
Dim curTotal As Currency
Dim bSueprvisar As Boolean
Dim strLista As String

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from AFI_BENEFICIOS_REMESAS where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close

bSueprvisar = True

If cboBanco.Text <> "TODOS" And chkFiltros.Value = vbChecked Then
    bSueprvisar = fxSupervisaBanco
End If

If bSueprvisar Then
    strSQL = "Select B.*,S.Nombre,E.Descripcion as 'EstadoPersona',Ban.Descripcion as 'BancoDesc',B.TES_SUPERVISION_FECHA," _
             & "dbo.fxTesSupervisa(B.cedula,S.nombre,B.monto,0,'C') as 'Duplicado'" _
             & " from afi_bene_pago B inner join socios S on B.cedula = S.cedula" _
             & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio and B.consec = O.consec" _
             & " inner join Afi_Estados_Persona E on S.EstadoActual = E.Cod_Estado" _
             & " inner join Tes_Bancos Ban on B.cod_Banco = Ban.id_Banco" _
             & " where O.cod_remesa is null " _
             & "   and O.registra_fecha between '" & Format(vFechaInicio, "yyyy/mm/dd 00:00:00") & "' and '" & Format(vFechaCorte, "yyyy/mm/dd 23:59:59") & "'" _
             & "   and B.ESTADO = 'S' and B.tesoreria is null"
Else
    strSQL = "Select B.*,S.Nombre,E.Descripcion as 'EstadoPersona',Ban.Descripcion as 'BancoDesc'" _
             & " from afi_bene_pago B inner join socios S on B.cedula = S.cedula" _
             & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio and B.consec = O.consec" _
             & " inner join Afi_Estados_Persona E on S.EstadoActual = E.Cod_Estado" _
             & " inner join Tes_Bancos Ban on B.cod_Banco = Ban.id_Banco" _
             & " where O.cod_remesa is null " _
             & "   and O.registra_fecha between '" & Format(vFechaInicio, "yyyy/mm/dd 00:00:00") & "' and '" & Format(vFechaCorte, "yyyy/mm/dd 23:59:59") & "'" _
             & "   and B.ESTADO = 'S' and B.tesoreria is null"
             
    
End If

If fxSIFParametros("16") = "S" Then
   strSQL = strSQL & " and O.Analista_Revision = 'S'"
End If

If cboOficina.Text <> "TODOS" And cboOficina.ListCount > 0 Then
   strSQL = strSQL & " and O.cod_Oficina = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
End If

If chkFiltros.Value = vbChecked Then
    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And B.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And O.Registra_User like '%" & cboUsuarios.Text & "%'"
    End If

End If
        
        
strSQL = strSQL & " order by B.CEDULA"

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswCarga
 .ListItems.Clear
 Do While Not rs.EOF
   
    Select Case rs!Tipo
       Case "S"
         vTipo = "Socio"
         vCedula = rs!Cedula
       Case "B"
         vTipo = "Beneficiario"
    End Select
    'carga la lista
    
    Set itmX = lswCarga.ListItems.Add(, , rs!consec)
        itmX.SubItems(1) = rs!Cod_Beneficio
        itmX.SubItems(2) = rs!EstadoPersona
        
        If rs!Duplicado = 1 And IsNull(rs!TES_SUPERVISION_FECHA) Then
           itmX.ForeColor = vbRed
           vDuplicado = True
           strLista = strLista & rs!consec & " " & rs!Cod_Beneficio & " " & rs!Cedula & " " & Format(rs!MONTO, "Standard") & vbCrLf
        Else
           itmX.ForeColor = vbBlack
        End If
       
        vNombre = rs!Nombre
        'en caso de que sea un beneficiario
'        If Trim(vNombre) = "" Or UCase(vNombre) = "ASEASECCSS" Then
'           vNombre = fxBeneficiario(rs!Cedula, rs!consec)
'        End If
        
        itmX.SubItems(3) = Format(rs!MONTO, "Standard")
        itmX.SubItems(4) = rs!Cedula
        itmX.SubItems(5) = rs!Nombre
        itmX.SubItems(6) = fxTipoDocumento(rs!Tipo_Emision)
        itmX.SubItems(7) = rs!cta_bancaria & ""
        
        itmX.SubItems(8) = rs!BancoDesc
        itmX.SubItems(9) = rs!cod_banco
        itmX.SubItems(10) = IIf(vDuplicado, rs!Duplicado, 0)
    rs.MoveNext
    
   
       itmX.Checked = chkCarga.Value
         
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(3))
       End If
        
     '   rs.MoveNext
        
        PrgBar.Value = PrgBar.Value + 1
 Loop
End With

rs.Close

If vDuplicado Then
      MsgBox "Estos Beneficios  necesitan autorización para ser trasladados ya que cuentan" _
          & "con una transacción por un monto igual en Tesorería " & vbCrLf & vbCrLf & strLista, vbCritical
End If


PrgBar.Visible = False

txtCargaTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub








Private Sub sbCerrar()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from AFI_BENEFICIOS_REMESAS" _
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
strSQL = "update AFI_BENEFICIOS_REMESAS set estado = 'C'" _
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
strSQL = "select count(*) as Existe from AFI_BENEFICIOS_REMESAS" _
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

PrgBar.Max = vCasos
PrgBar.Value = 1
PrgBar.Visible = True


With lswCarga.ListItems

For i = 1 To .Count
 If .Item(i).Checked And .Item(i).SubItems(10) = 0 Then
       strSQL = "update afi_bene_otorga set cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
              & " where consec = " & .Item(i).Text _
              & " and cod_beneficio = '" & Trim(.Item(i).SubItems(1)) & "'"
       Call ConectionExecute(strSQL)
    
    PrgBar.Value = PrgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Carga Remesa Traslado a Tesoreria: " & cboCarga.ItemData(cboCarga.ListIndex))
End If

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault

MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation
Call sbCargaBuscar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear
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
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
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
     lngSolicitud = fxMaestroTesoreria(vTipo, vBanco, !MONTO, !id_desembolso _
                   , !Concepto, !Id_solicitud, !Id_solicitud, vReferencia, !Codigo, "0" _
                   , vFecha, vCod_App, vOficina, vUnidad, vToken, vRemesaTipo, vRemesa _
                   , CStr(vOP), CStr(!id_desembolso))
     
     Call sbCreaDetalle(lngSolicitud, fxCtaBanco(vBanco), !MONTO, "H", 1, vUnidad)
     Call sbCreaDetalle(lngSolicitud, !cuenta_conta, !MONTO, "D", 2, vUnidad)
     
     strSQL = "update desembolsos set tdocumento = '" & vTipo & "',cod_banco = " & vBanco & ",nsolicitud = " & lngSolicitud _
            & " where id_desembolso = " & !id_desembolso
     Call ConectionExecute(strSQL)
  .MoveNext
 Loop
 .Close
End With

End Sub


Private Sub sbTrasladoBuscar()
Dim rs2 As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear
curTotal = 0
If cboTraslado.Text = "" Then Exit Sub

strSQL = "select fecha_inicio,fecha_corte from AFI_BENEFICIOS_REMESAS where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close



strSQL = "Select B.*,S.Nombre,E.Descripcion as 'EstadoPersona',Ban.Descripcion as 'BancoDesc'" _
         & " from afi_bene_pago B inner join socios S on B.cedula = S.cedula" _
         & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio and B.consec = O.consec" _
         & " inner join Afi_Estados_Persona E on S.EstadoActual = E.Cod_Estado" _
         & " inner join Tes_Bancos Ban on B.cod_Banco = Ban.id_Banco" _
         & " Where O.cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
         & "   and O.registra_fecha between '" & Format(vFechaInicio, "yyyy/mm/dd 00:00:00") & "' and '" & Format(vFechaCorte, "yyyy/mm/dd 23:59:59") & "'" _
         & "   and O.ESTADO in('S','A') and B.tesoreria is null"
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswTraslado
 .ListItems.Clear
 Do While Not rs.EOF
    Set itmX = lswTraslado.ListItems.Add(, , rs!consec)
        itmX.SubItems(1) = rs!Cod_Beneficio
        itmX.SubItems(2) = rs!EstadoPersona
        
        vNombre = rs!Nombre
        'en caso de que sea un beneficiario
'        If Trim(vNombre) = "" Or UCase(vNombre) = "ASEASECCSS" Then
'           vNombre = fxBeneficiario(rs!Cedula, rs!consec)
'        End If
        
        itmX.SubItems(3) = Format(rs!MONTO, "Standard")
        itmX.SubItems(4) = rs!Cedula
        itmX.SubItems(5) = rs!Nombre
        itmX.SubItems(6) = fxTipoDocumento(rs!Tipo_Emision)
        itmX.SubItems(7) = rs!cta_bancaria & ""
        
        itmX.SubItems(8) = rs!BancoDesc
        itmX.SubItems(9) = rs!cod_banco
    
        curTotal = curTotal + CCur(itmX.SubItems(3))
       rs.MoveNext
       PrgBar.Value = PrgBar.Value + 1
 Loop

End With

rs.Close

PrgBar.Visible = False

txtPagoTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswTraslado.ListItems.Clear

End Sub


Private Sub sbTraslado()
Dim iContador As Integer, vTesoreria As Long, vMonto As Currency, vIDBanco As Long
Dim vNombre As String, vConsec As Long, vTipo As String
Dim vBeneficio As String, vCedula As String, vEmite As String
Dim vCtaBanco As String, vCta As String, vBanco As String
Dim vDetalle As String, strSQL As String, vDetalle2 As String
Dim vCtaBene As String, rs As New ADODB.Recordset
Dim vRemesa As Integer, i As Integer
Dim vToken As String, vFecha As Date


Me.MousePointer = vbHourglass

On Error GoTo vError


strSQL = "select top 1 id_token from tes_tokens where estado = 'A' order by registro_fecha "
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  vToken = rs!ID_TOKEN
Else
  vToken = fxTesToken
End If
rs.Close


vFecha = fxFechaServidor
vRemesa = cboTraslado.ItemData(cboTraslado.ListIndex)

For i = 1 To lswTraslado.ListItems.Count
  iContador = iContador + 1
Next i

If iContador > 0 Then

    PrgBar.Max = iContador
    PrgBar.Value = 1
    
   
    
   For i = 1 To lswTraslado.ListItems.Count
   
        Select Case UCase(lswTraslado.ListItems.Item(i).SubItems(2))
          Case "SOCIO"
            vTipo = "S"
          Case "BENEFICIARIO"
            vTipo = "B"
        End Select
          
          vConsec = lswTraslado.ListItems.Item(i).Text
          vBeneficio = lswTraslado.ListItems.Item(i).SubItems(1)
          vMonto = CCur(lswTraslado.ListItems.Item(i).SubItems(3))
          vCedula = lswTraslado.ListItems.Item(i).SubItems(4)
          vNombre = Trim(lswTraslado.ListItems.Item(i).SubItems(5))
          vEmite = fxTipoDocumento(lswTraslado.ListItems.Item(i).SubItems(6))
          vIDBanco = lswTraslado.ListItems.Item(i).SubItems(9)
          vBanco = lswTraslado.ListItems.Item(i).SubItems(8)
          vCtaBanco = fxgCtaBanco(vIDBanco)
          vCta = lswTraslado.ListItems.Item(i).SubItems(7)
          
          
          strSQL = "select descripcion,cod_cuenta  from afi_beneficios where cod_beneficio = '" & vBeneficio & "'"
          Call OpenRecordSet(rs, strSQL)
                vCtaBene = IIf(Not IsNull(rs!cod_cuenta), rs!cod_cuenta, "0")
                vDetalle = vBeneficio
                vDetalle2 = rs!Descripcion
          rs.Close
          

          vTesoreria = fxgTesoreriaMaestro(vEmite, vIDBanco, vMonto, vCedula, vNombre, _
                            vConsec, vDetalle, 0, vDetalle2, vCta, vFecha, GLOBALES.gOficinaUnidad, , , , , "BENE", vToken, "BEN", vRemesa)
          
         
          'Actualiza el estado en tabla afi_bene_otorga
          strSQL = "Update afi_bene_otorga set estado = 'E',autoriza_user = '" & glogon.Usuario & "'," _
                  & "autoriza_fecha = dbo.MyGetdate(),cod_remesa = " & vRemesa & "  where cedula = '" & vCedula & "'" _
                  & " and cod_beneficio = '" & vBeneficio & "' and consec = '" & vConsec & "'"
                  
          Call ConectionExecute(strSQL)
          
          'Actualiza estado en afi_bene_pago
          strSQL = "Update afi_bene_pago set estado = 'E',tesoreria = " & vTesoreria & "," _
                  & "envio_user = '" & glogon.Usuario & "',envio_fecha = dbo.MyGetdate()" _
                  & ",ID_TOKEN = '" & vToken & "'" _
                  & " where cedula = '" & vCedula & "'" _
                  & " and cod_beneficio = '" & vBeneficio & "' and consec = '" & vConsec & "'"
                  
          Call ConectionExecute(strSQL)
         
          'Detalle de tesoreria
        
          Call sbgTesoreriaDetalle(vTesoreria, vCtaBanco, vMonto, "H", 1)
          Call sbgTesoreriaDetalle(vTesoreria, vCtaBene, vMonto, "D", 2)
          
    If PrgBar.Max < iContador Then PrgBar.Value = PrgBar.Value + 1
   
   Next i
  
  
End If

PrgBar.Value = 0


Me.MousePointer = vbDefault

'Actualiza y Carga Remesa
strSQL = "update AFI_BENEFICIOS_REMESAS SET Estado = 'T'" _
       & "  Where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call ConectionExecute(strSQL)

Call sbLimpia


Me.MousePointer = vbDefault

PrgBar.Visible = False

MsgBox "Operaciones Enviadas a Tesoreria Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
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

vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

 tcMain.Item(0).Selected = True
 
 With lswRemesas.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Estado", 1400
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Notas", 3400
 End With
 
 With lswRep.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Notas", 3400
 End With
  
 
 With lswCarga.ColumnHeaders
    .Clear
    .Add , , "Id", 500
    .Add , , "Beneficio", 2500
    .Add , , "Tipo", 600, vbCenter
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3000
    .Add , , "Emite", 1800, vbCenter
    .Add , , "Cuenta", 2100
    .Add , , "Banco", 2500
    .Add , , "Banco Id", 1800, vbCenter
    .Add , , "Duplicado?", 1500, vbCenter
 End With
 
 
 With lswTraslado.ColumnHeaders
    .Clear
    .Add , , "Id", 500
    .Add , , "Beneficio", 2500
    .Add , , "Tipo", 600, vbCenter
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3000
    .Add , , "Emite", 1800, vbCenter
    .Add , , "Cuenta", 2100
    .Add , , "Banco", 2500
    .Add , , "Banco Id", 1800, vbCenter
 End With
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia
 Call sbRequiereAutorizacion
 
strSQL = "select rtrim(cod_oficina) as 'Idx', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbCbo_Llena_New(cboRepOficina, strSQL, True, False)
 
 
 
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

strToken = strToken & "." & rs!consec

rs.Close

strSQL = "insert tes_tokens(id_token,registro_fecha,registro_usuario,estado)" _
      & "values('" & strToken & "',dbo.MyGetdate(),'" & glogon.Usuario & "','A') "
Call ConectionExecute(strSQL)

fxCreaToquen = strToken

End Function





