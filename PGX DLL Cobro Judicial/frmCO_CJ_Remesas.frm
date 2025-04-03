VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCO_CJ_Remesas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros: Remesas de Gastos Desembolsables"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11655
      _Version        =   1310723
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
      ItemCount       =   4
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
      Item(1).ControlCount=   11
      Item(1).Control(0)=   "Label8(9)"
      Item(1).Control(1)=   "Label8(10)"
      Item(1).Control(2)=   "cboCarga"
      Item(1).Control(3)=   "cboOficina"
      Item(1).Control(4)=   "chkCarga"
      Item(1).Control(5)=   "lswCarga"
      Item(1).Control(6)=   "txtCargaTotal"
      Item(1).Control(7)=   "btnBarra(3)"
      Item(1).Control(8)=   "btnBarra(4)"
      Item(1).Control(9)=   "btnBarra(5)"
      Item(1).Control(10)=   "Label8(18)"
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
      Item(3).ControlCount=   7
      Item(3).Control(0)=   "txtRepRemesas"
      Item(3).Control(1)=   "lblRemesa"
      Item(3).Control(2)=   "chkRemesaInd"
      Item(3).Control(3)=   "lswRep"
      Item(3).Control(4)=   "btnBarra(8)"
      Item(3).Control(5)=   "ShortcutCaption1(0)"
      Item(3).Control(6)=   "ShortcutCaption1(2)"
      Begin XtremeSuiteControls.ListView lswRemesas 
         Height          =   3132
         Left            =   1560
         TabIndex        =   2
         Top             =   3240
         Width           =   10092
         _Version        =   1310723
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
         _Version        =   1310723
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
         Left            =   -69880
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   11415
         _Version        =   1310723
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
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1310723
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
      Begin XtremeSuiteControls.GroupBox fraReporte 
         Height          =   2052
         Left            =   4320
         TabIndex        =   6
         Top             =   960
         Width           =   7452
         _Version        =   1310723
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
            TabIndex        =   7
            Top             =   1200
            Width           =   1572
            _Version        =   1310723
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
            TabIndex        =   8
            Top             =   360
            Width           =   1215
            _Version        =   1310723
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
            TabIndex        =   9
            Top             =   360
            Width           =   1335
            _Version        =   1310723
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
            TabIndex        =   10
            Top             =   360
            Width           =   1335
            _Version        =   1310723
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
            TabIndex        =   11
            Top             =   720
            Width           =   4932
            _Version        =   1310723
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
            TabIndex        =   12
            Top             =   1200
            Width           =   612
            _Version        =   1310723
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
            Picture         =   "frmCO_CJ_Remesas.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   1
            Left            =   6360
            TabIndex        =   13
            Top             =   1200
            Width           =   492
            _Version        =   1310723
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
            Picture         =   "frmCO_CJ_Remesas.frx":0707
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   1
            Left            =   3600
            TabIndex        =   14
            Top             =   1200
            Width           =   1692
            _Version        =   1310723
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
         Begin XtremeSuiteControls.Label Label8 
            Height          =   372
            Index           =   7
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Width           =   1212
            _Version        =   1310723
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
            Index           =   8
            Left            =   360
            TabIndex        =   15
            Top             =   720
            Width           =   1452
            _Version        =   1310723
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
         Begin VB.Image imgRepRefresca 
            Height          =   240
            Left            =   6600
            Picture         =   "frmCO_CJ_Remesas.frx":0D45
            ToolTipText     =   "Actualizar Oficinas"
            Top             =   360
            Width           =   240
         End
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   0
         Left            =   4320
         TabIndex        =   17
         Top             =   480
         Width           =   1332
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":1435
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa 
         Height          =   432
         Left            =   1560
         TabIndex        =   18
         Top             =   480
         Width           =   2412
         _Version        =   1310723
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
         Left            =   1560
         TabIndex        =   19
         Top             =   1680
         Width           =   2412
         _Version        =   1310723
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
         Left            =   5160
         TabIndex        =   20
         Top             =   1320
         Width           =   2052
         _Version        =   1310723
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
         Left            =   5160
         TabIndex        =   21
         Top             =   1680
         Width           =   2052
         _Version        =   1310723
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
         Left            =   1560
         TabIndex        =   22
         Top             =   2040
         Width           =   10092
         _Version        =   1310723
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
         Left            =   1560
         TabIndex        =   23
         Top             =   1320
         Width           =   1212
         _Version        =   1310723
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
         Left            =   2760
         TabIndex        =   24
         Top             =   1320
         Width           =   1212
         _Version        =   1310723
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
         Left            =   6120
         TabIndex        =   25
         Top             =   480
         Width           =   492
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":1B35
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   2
         Left            =   6600
         TabIndex        =   26
         Top             =   480
         Width           =   492
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":20D9
      End
      Begin XtremeSuiteControls.ComboBox cboCarga 
         Height          =   312
         Left            =   -67600
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1310723
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
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1310723
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
      Begin XtremeSuiteControls.CheckBox chkCarga 
         Height          =   252
         Left            =   -69880
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
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
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1310723
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
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   3
         Left            =   -63880
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":27E0
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   4
         Left            =   -62560
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":2EE0
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   5
         Left            =   -61240
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":35E8
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   6
         Left            =   -63040
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":3CF4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   7
         Left            =   -61720
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":43F4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   8
         Left            =   -60760
         TabIndex        =   36
         Top             =   5640
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":4CC5
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   312
         Left            =   -59200
         TabIndex        =   37
         Top             =   4590
         Visible         =   0   'False
         Width           =   852
         _Version        =   1310723
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
         TabIndex        =   38
         Top             =   5040
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310723
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
      Begin XtremeSuiteControls.FlatEdit txtCargaTotal 
         Height          =   312
         Left            =   -60760
         TabIndex        =   39
         Top             =   6120
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1310723
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
         TabIndex        =   40
         Top             =   6000
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1310723
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
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   9
         Left            =   5640
         TabIndex        =   41
         Top             =   480
         Width           =   492
         _Version        =   1310723
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
         Picture         =   "frmCO_CJ_Remesas.frx":53CC
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa_Monto 
         Height          =   312
         Left            =   8520
         TabIndex        =   42
         Top             =   1680
         Width           =   2052
         _Version        =   1310723
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
         Left            =   8520
         TabIndex        =   43
         Top             =   1320
         Width           =   2052
         _Version        =   1310723
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   0
         Left            =   600
         TabIndex        =   61
         Top             =   480
         Width           =   1212
         _Version        =   1310723
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   1
         Left            =   600
         TabIndex        =   60
         Top             =   1320
         Width           =   1212
         _Version        =   1310723
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
         Index           =   2
         Left            =   4320
         TabIndex        =   59
         Top             =   1320
         Width           =   1212
         _Version        =   1310723
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
         Index           =   3
         Left            =   600
         TabIndex        =   58
         Top             =   1680
         Width           =   1212
         _Version        =   1310723
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
         Index           =   4
         Left            =   4320
         TabIndex        =   57
         Top             =   1680
         Width           =   1212
         _Version        =   1310723
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
         Index           =   5
         Left            =   600
         TabIndex        =   56
         Top             =   2040
         Width           =   1212
         _Version        =   1310723
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
         Index           =   6
         Left            =   1560
         TabIndex        =   55
         Top             =   2880
         Width           =   2892
         _Version        =   1310723
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
         Index           =   9
         Left            =   -69400
         TabIndex        =   54
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310723
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
         Index           =   10
         Left            =   -69400
         TabIndex        =   53
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310723
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
         Height          =   252
         Index           =   14
         Left            =   -69160
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310723
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
         Height          =   252
         Index           =   18
         Left            =   -62560
         TabIndex        =   51
         Top             =   6120
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310723
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
         Index           =   19
         Left            =   -62680
         TabIndex        =   50
         Top             =   6000
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310723
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
         Index           =   22
         Left            =   7680
         TabIndex        =   49
         Top             =   1680
         Width           =   1212
         _Version        =   1310723
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
         Index           =   23
         Left            =   7680
         TabIndex        =   48
         Top             =   1320
         Width           =   1212
         _Version        =   1310723
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   11655
         _Version        =   1310723
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   46
         Top             =   1440
         Visible         =   0   'False
         Width           =   11415
         _Version        =   1310723
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
      Begin XtremeShortcutBar.ShortcutCaption lblRemesa 
         Height          =   375
         Left            =   -70000
         TabIndex        =   45
         Top             =   4560
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1310723
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
         Index           =   2
         Left            =   -64600
         TabIndex        =   44
         Top             =   4560
         Visible         =   0   'False
         Width           =   11655
         _Version        =   1310723
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
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   62
      Top             =   7920
      Visible         =   0   'False
      Width           =   11655
      _Version        =   1310723
      _ExtentX        =   20558
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.Label lblTituloMain 
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   8775
      _Version        =   1310723
      _ExtentX        =   15478
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Traslado de Gestiones de Cobro Judicial para Desembolso en Bancos"
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
      Width           =   15732
   End
End
Attribute VB_Name = "frmCO_CJ_Remesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Dim itmX As ListViewItem, vPaso As Boolean
Dim mRequiereAutorizacion As Boolean
Dim mUnidad As String, mConcepto As String

Private Sub btnBarra_Click(Index As Integer)
Dim i As Integer

On Error GoTo vError

Select Case Index
  Case 0 'NUEVO"
     
    Call sbLimpia
    
  Case 9 'GUARDAR
    If txtRemesa.Text = "" Then
     
            strSQL = "select isnull(max(cod_remesa),0) + 1 as Ultimo from CBR_CJ_REMESAS"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert CBR_CJ_REMESAS(cod_remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas) values(" & rs!Ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!Ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de CJ Traslado a Tesorería:  " & txtRemesa)
    
    Else
        If txtEstado.Text = "Abierta" Then
                    
            strSQL = "update CBR_CJ_REMESAS set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where cod_remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
             
            Call Bitacora("Modifica", "Remesa de CJ Traslado a Tesorería:  " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
  Case 1 'BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Abierta" Then
            strSQL = "Update CBR_CJ_TRAMITE_GASTOS set cod_Remesa = null where Cod_Remesa = " & txtRemesa
            
            strSQL = strSQL & Space(10) & "delete CBR_CJ_REMESAS where Cod_Remesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            
            Call Bitacora("Elimina", "Remesa de CJ Traslado a Tesorería:  " & txtRemesa)
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


On Error GoTo vError

lswCarga.ListItems.Clear
If cboCarga.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select fecha_inicio,fecha_corte from CBR_CJ_REMESAS where cod_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!fecha_inicio
  vFechaCorte = rs!fecha_corte
rs.Close


'Carga Oficinas
strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as Itmx" _
       & " from SIF_Oficinas  where cod_oficina in(" _
       & "Select R.COD_OFICINA_F" _
       & " from CBR_CJ_TRAMITE_GASTOS G " _
       & " inner join CBR_CJ_TIPOS_GASTOS T on G.TIPO_GASTO = T.TIPO_GASTO" _
       & " inner join CBR_CJ_TRAMITE X on G.COD_TRAMITE = X.COD_TRAMITE" _
       & " inner join REG_CREDITOS R on X.ID_SOLICITUD = R.ID_SOLICITUD " _
       & " Where G.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and cod_Remesa is null) " _
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
       & " from CBR_CJ_REMESAS T left join vCBR_CJ_REMESAS_Rsm D on T.cod_Remesa = D.cod_Remesa" _
       & " where T.Cod_Remesa = " & pRemesa

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa.Text = CStr(rs!cod_Remesa)
  txtUsuario.Text = rs!Usuario
  txtFecha.Text = rs!fecha
  
  Select Case rs!Estado
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
 .WindowTitle = "Reportes del Módulo de Cobro Judicial"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Traslado a Tesoreria")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If
'
' Select Case True
'  Case opt.Item(0).Value 'Pendiente Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_RemesaTESDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
'  Case opt.Item(1).Value 'Traslado Detalle Agrupado Remesa
'     .ReportFileName = SIFGlobal.fxPathReportes("CxC_RemesaTESDetalleAgrp.rpt")
'     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
' End Select
'
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA TRASLADO A TESORERIA : Cbr J.'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = "{CBR_CJ_REMESAS.cod_Remesa} = " & lblRemesa.Tag
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgCancelar_Click()
fraReporte.Visible = False
End Sub

Private Sub imgReporte_Click()

Select Case fraReporte.Caption
  Case "Pendientes"
    Call sbReportePendientes
  Case "Trasladadas"
    Call sbReporteEnviadas
End Select

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
strSQL = "select rtrim(cod_oficina) + ' - ' + rtrim(descripcion) as Itmx" _
       & " from SIF_Oficinas  where cod_oficina in(" _
       & "Select R.COD_OFICINA_F" _
       & " from CBR_CJ_TRAMITE_GASTOS G " _
       & " inner join CBR_CJ_TIPOS_GASTOS T on G.TIPO_GASTO = T.TIPO_GASTO" _
       & " inner join CBR_CJ_TRAMITE X on G.COD_TRAMITE = X.COD_TRAMITE" _
       & " inner join REG_CREDITOS R on X.ID_SOLICITUD = R.ID_SOLICITUD " _
       & " Where G.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and cod_Remesa is null) " _
       & " order by cod_oficina"

Call sbLlenaCbo(cboRepOficina, strSQL, True, False)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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
  
   If lswCarga.ListItems.Item(i).Checked Then
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(4))
   End If
  
Next i

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub

Private Sub lswCarga_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCarga.SortKey = ColumnHeader.Index - 1
  If lswCarga.SortOrder = 0 Then lswCarga.SortOrder = 1 Else lswCarga.SortOrder = 0
  lswCarga.Sorted = True
End Sub


Private Sub lswCarga_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curTotal As Currency


If vPaso Then Exit Sub
If Trim(txtCargaTotal.Text) = "" Then txtCargaTotal.Text = 0

curTotal = CCur(txtCargaTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(4))
Else
   curTotal = curTotal - CCur(Item.SubItems(4))
End If

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub




Private Sub sbReporteRemesas()
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
 
 .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_ListadoRemesa.rpt")
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

Select Case tcMain.Selected.Index
  Case 0 'Remesas
     txtEstado.Text = ""
     txtFecha.Text = ""
     txtUsuario.Text = ""
     txtRemesa.Text = ""
     
     txtRemesa_Casos.Text = ""
     txtRemesa_Monto.Text = ""
     
    fraReporte.Visible = False
    
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    
    dtpRepInicio.Value = dtpInicio.Value
    dtpRepCorte.Value = dtpInicio.Value
    
    txtNotas.Text = ""
        
        strSQL = "select TOP 50 T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
          & " from CBR_CJ_REMESAS T left join vCBR_CJ_REMESAS_Rsm D on T.cod_Remesa = D.cod_Remesa" _
          & " order by T.fecha desc"
     
     
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!cod_Remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                
                Select Case rs!Estado
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
        
    strSQL = "select * from CBR_CJ_REMESAS where estado = 'A' order by fecha desc"
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
   
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
        
    strSQL = "select * from CBR_CJ_REMESAS where estado = 'C' order by fecha desc"
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
            & " from CBR_CJ_REMESAS T left join vCBR_CJ_REMESAS_Rsm D on T.cod_Remesa = D.cod_Remesa" _
            & " order by T.fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!cod_Remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                
                Select Case rs!Estado
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

Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from CBR_CJ_REMESAS where cod_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!fecha_inicio
  vFechaCorte = rs!fecha_corte
rs.Close


If cboOficina.Text = "TODOS" Then



    strSQL = "Select G.NUM_LINEA,G.COD_TRAMITE,g.BENEFICIARIO,G.MONTO,T.DESCRIPCION as 'Descripcion'" _
             & " from CBR_CJ_TRAMITE_GASTOS G " _
             & " inner join CBR_CJ_TIPOS_GASTOS T on G.TIPO_GASTO = T.TIPO_GASTO" _
             & " Where G.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
             & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and cod_Remesa is null "
Else

    strSQL = "Select G.NUM_LINEA,G.COD_TRAMITE,g.BENEFICIARIO,G.MONTO,T.DESCRIPCION as 'Descripcion'" _
             & " from CBR_CJ_TRAMITE_GASTOS G " _
             & " inner join CBR_CJ_TIPOS_GASTOS T on G.TIPO_GASTO = T.TIPO_GASTO" _
             & " inner join CBR_CJ_TRAMITE X on G.COD_TRAMITE = X.COD_TRAMITE" _
             & " inner join REG_CREDITOS R on X.ID_SOLICITUD = R.ID_SOLICITUD and" _
             & " R.COD_OFICINA_F = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'" _
             & " Where G.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
             & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and cod_Remesa is null "
             
End If

strSQL = strSQL & " and T.Aplica_Desembolso = 1 and G.TESORERIA_FECHA is null order by G.cod_tramite,g.Num_linea"

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

vPaso = True

With lswCarga
 .ListItems.Clear
 Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!COD_TRAMITE)
       itmX.SubItems(1) = rs!NUM_LINEA
       itmX.SubItems(2) = rs!Beneficiario
       itmX.SubItems(3) = rs!Descripcion
       itmX.SubItems(4) = Format(rs!Monto, "Standard")
       
       itmX.Checked = vbChecked
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(4))
       End If
        
        rs.MoveNext
        
        PrgBar.Value = PrgBar.Value + 1
 Loop
End With

rs.Close

vPaso = False

PrgBar.Visible = False

txtCargaTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear


End Sub


Private Sub sbTrasladoBuscar()
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from CBR_CJ_REMESAS where cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!fecha_inicio
  vFechaCorte = rs!fecha_corte
rs.Close

strSQL = "Select G.NUM_LINEA, G.COD_TRAMITE,g.BENEFICIARIO,G.MONTO,T.DESCRIPCION as 'Descripcion'" _
        & " from CBR_CJ_TRAMITE_GASTOS G " _
        & " inner join CBR_CJ_TIPOS_GASTOS T on G.TIPO_GASTO = T.TIPO_GASTO" _
        & " Where G.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
        & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
        & " and G.Tesoreria_Fecha is null" _
        & " order by G.cod_tramite"

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswTraslado
 .ListItems.Clear
 Do While Not rs.EOF
Set itmX = .ListItems.Add(, , rs!COD_TRAMITE)

       itmX.SubItems(1) = rs!NUM_LINEA

       itmX.SubItems(2) = rs!Beneficiario
       itmX.SubItems(3) = rs!Descripcion
       itmX.SubItems(4) = Format(rs!Monto, "Standard")
       
       itmX.Checked = vbChecked
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(4))
       End If
       
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



Private Sub sbCerrar()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CBR_CJ_REMESAS" _
       & " where cod_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
strSQL = "update CBR_CJ_REMESAS set estado = 'C'" _
       & " where cod_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("Aplica", "Cierra Remesa CJ Traslado a Tesoreria: " & cboCarga.ItemData(cboCarga.ListIndex))


MsgBox "Remesa Cerrada Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCarga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CBR_CJ_REMESAS" _
       & " where cod_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close


Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear



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
 If .Item(i).Checked Then
 
     strSQL = "update CBR_CJ_TRAMITE_GASTOS set cod_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
            & " where cod_tramite = " & .Item(i).Text & "  And num_linea = " & .Item(i).SubItems(1) & ""
     Call ConectionExecute(strSQL)
   
    PrgBar.Value = PrgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Carga Remesa CJ Traslado a Tesoreria: " & cboCarga.ItemData(cboCarga.ListIndex))
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
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String) As Long                                 'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza,fecha_autorizacion)" _
       & " values('" & mConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "','" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CC','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = strSQL & ",'N',null,null)"
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
         & "'"
  rsX.CursorLocation = adUseServer
  rsX.Open strSQL, glogon.Conection, adOpenStatic
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)


strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidad & "')"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCtaBanco(pBanco As Integer) As String
Dim rsX As New ADODB.Recordset

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & pBanco
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function


Private Sub sbTraslado()
Dim vToken As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vToken = ""
    
'Nuevo Proceso (Integrado)
strSQL = "exec spCbr_CJ_Traslado_Bancos " & cboTraslado.ItemData(cboTraslado.ListIndex) & ", '" & vToken & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Remesa CJ Traslado a Tesoreria: " & cboTraslado.ItemData(cboTraslado.ListIndex))
 
Call sbLimpia


Me.MousePointer = vbDefault

MsgBox "Operaciones Enviadas a Tesoreria Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbReportePendientes()

Dim strTitulo As String
Dim strRuta As String, strInicio As String, strFinal As String
Dim strFiltro As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strTitulo = "Honorarios pendientes de Traslado a Tesorería"


strRuta = SIFGlobal.fxPathReportes("Cobro_Judicial_GastoPenEnviar.rpt")
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
     
     .WindowTitle = "Honorarios a trasladar a Tesorería"
     
    .ReportFileName = strRuta
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Titulo='" & strTitulo & "'"
    
  
    If chkRepFechas.Value = vbUnchecked Then
        strSQL = "  cdate({CBR_CJ_TRAMITE_GASTOS.Registro_Fecha}) in Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd")
        strSQL = strSQL & ") to Date (" & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
        strFiltro = "Desde " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
    Else
         strFiltro = "Todas las fechas "
    End If
    
    
    
    If cboRepOficina.Text <> "TODOS" Then
       strSQL = strSQL & " AND {REG_CREDITOS.COD_OFICINA_F} = '" & cboRepOficina.ItemData(cboRepOficina.ListIndex) & "'"
       
       strFiltro = strFiltro & " /OFICINA " & cboRepOficina.ItemData(cboRepOficina.ListIndex)
    Else
       strFiltro = strFiltro & "Todas las Oficinas"
    End If
    
    If strSQL = "" Then
      strSQL = "ISNULL({CBR_CJ_TRAMITE_GASTOS.TESORERIA_NUMERO})"
    Else
      strSQL = strSQL & " AND ISNULL({CBR_CJ_TRAMITE_GASTOS.TESORERIA_NUMERO})"
    End If
    .Formulas(4) = "Filtro='" & strFiltro & "'"
    
    
    .SelectionFormula = strSQL
    
    '.SubreportToChange = "subCkDesembolsos"
    '.SelectionFormula = "{DESEMBOLSOS.Operacion} = {?Pm-CxC_Cuentas.Operacion}"
    
    .PrintReport


End With

 Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporteEnviadas()

Dim strFiltro As String

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "HONORARIOS ENVIADOS A TESORERIA"

 .Connect = glogon.ConectRPT

 .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_GastoTrasladadas.rpt")
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "Titulo='Desembolsos Solicitados en Tesorería'"
 .Formulas(4) = "Usuario='" & glogon.Usuario & "'"
 strFiltro = "INICIO : " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " CORTE : " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
 
 strSQL = "{CBR_CJ_TRAMITE_GASTOS.tesoreria_fecha} in date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ") to date(" _
       & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
    
 If cboRepOficina.Text <> "TODOS" Then
    strSQL = strSQL & " AND {REG_CREDITOS.COD_OFICINA_F} = '" & cboRepOficina.ItemData(cboRepOficina.ListIndex) & "'"
    strFiltro = strFiltro & " /OFICINA " & cboRepOficina.ItemData(cboRepOficina.ListIndex)
 Else
    strFiltro = strFiltro & "Todas las Oficinas"
 End If

 .Formulas(5) = "filtro='" & strFiltro & "'"

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

vModulo = 6

On Error GoTo vError

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
    .Add , , "No. Trámite", 1400
    .Add , , "Línea", 1000
    .Add , , "Beneficiario", 2500
    .Add , , "Concepto", 2500
    .Add , , "Monto", 1800, vbRightJustify
 End With
 
 
 With lswTraslado.ColumnHeaders
    .Clear
    .Add , , "No. Trámite", 1400
    .Add , , "Línea", 1000
    .Add , , "Beneficiario", 2500
    .Add , , "Concepto", 2500
    .Add , , "Monto", 1800, vbRightJustify
 End With
 
 
 Call Formularios(Me)
 
 btnBarra(9).Tag = btnBarra(0).Tag
 
 Call RefrescaTags(Me)
 
 Call sbLimpia
' Call sbRequiereAutorizacion
 
strSQL = "select rtrim(cod_oficina) as 'Idx', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbCbo_Llena_New(cboRepOficina, strSQL, True, False)
 
Exit Sub

vError:


 
 
End Sub

Private Function fxTraeCuentaGasto(vGasto As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_cuenta from cbr_cj_tipos_gastos where tipo_gasto ='" & vGasto & "' "
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
  fxTraeCuentaGasto = rs!cod_cuenta
Else
  fxTraeCuentaGasto = "0"
End If
rs.Close

End Function

Private Function fxTraeCuentaBanco(vBanco As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select ctaconta from tes_bancos where id_banco = " & vBanco & " "
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
  fxTraeCuentaBanco = rs!ctaConta
Else
  fxTraeCuentaBanco = "0"
End If
rs.Close

End Function

