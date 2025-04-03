VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmVivRemesasTesoreria 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10920
      Top             =   360
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      _Version        =   1441793
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
      SelectedItem    =   2
      Item(0).Caption =   "Remesa"
      Item(0).ControlCount=   23
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
      Item(0).Control(18)=   "btnBarra(9)"
      Item(0).Control(19)=   "txtRemesa_Casos"
      Item(0).Control(20)=   "txtRemesa_Monto"
      Item(0).Control(21)=   "Label8(22)"
      Item(0).Control(22)=   "Label8(23)"
      Item(1).Caption =   "Cargar"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "Label8(9)"
      Item(1).Control(1)=   "cboCarga"
      Item(1).Control(2)=   "chkCarga"
      Item(1).Control(3)=   "lswCarga"
      Item(1).Control(4)=   "txtCargaTotal"
      Item(1).Control(5)=   "btnBarra(3)"
      Item(1).Control(6)=   "btnBarra(4)"
      Item(1).Control(7)=   "btnBarra(5)"
      Item(1).Control(8)=   "Label8(18)"
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
         TabIndex        =   4
         Top             =   3240
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
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
         Height          =   4332
         Left            =   -70000
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1441793
         _ExtentX        =   20553
         _ExtentY        =   7641
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
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   11412
         _Version        =   1441793
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
         Left            =   -70000
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1441793
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
         TabIndex        =   5
         Top             =   5160
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtRemesa 
         Height          =   432
         Left            =   -68440
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441793
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
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441793
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
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   -64840
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   2052
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   -68440
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
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
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ComboBox cboCarga 
         Height          =   312
         Left            =   -67600
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1441793
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
      Begin XtremeSuiteControls.CheckBox chkCarga 
         Height          =   252
         Left            =   -69880
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         TabIndex        =   15
         Top             =   600
         Width           =   7692
         _Version        =   1441793
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
         Index           =   8
         Left            =   -60760
         TabIndex        =   16
         Top             =   5640
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
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
         Appearance      =   16
         Picture         =   "frmVivRemesasTesoreria.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   312
         Left            =   -59200
         TabIndex        =   17
         Top             =   4560
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441793
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
         TabIndex        =   18
         Top             =   5040
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
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
         TabIndex        =   19
         Top             =   5520
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1441793
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
         TabIndex        =   20
         Top             =   6120
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441793
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
         TabIndex        =   21
         Top             =   6000
         Width           =   2412
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtRemesa_Monto 
         Height          =   312
         Left            =   -61480
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   2052
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa_Casos 
         Height          =   312
         Left            =   -61480
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   0
         Left            =   -65920
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":07BC
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   1
         Left            =   -64120
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":0EBC
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   2
         Left            =   -63640
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":1460
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   9
         Left            =   -64600
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":1B67
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   3
         Left            =   -64480
         TabIndex        =   45
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":2298
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   4
         Left            =   -63160
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":2998
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   5
         Left            =   -61840
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":30A0
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   6
         Left            =   6600
         TabIndex        =   48
         Top             =   960
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":37AC
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   7
         Left            =   7920
         TabIndex        =   49
         Top             =   960
         Width           =   1575
         _Version        =   1441793
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
         Picture         =   "frmVivRemesasTesoreria.frx":3EAC
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   11415
         _Version        =   1441793
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   0
         Left            =   -69400
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         Left            =   -69400
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         Left            =   -65680
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         Left            =   -69400
         TabIndex        =   36
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         Left            =   -65680
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         Left            =   -69400
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         Left            =   -68440
         TabIndex        =   33
         Top             =   2880
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
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
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
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
         Height          =   252
         Index           =   14
         Left            =   840
         TabIndex        =   31
         Top             =   600
         Width           =   1572
         _Version        =   1441793
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
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
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
         Height          =   300
         Index           =   2
         Left            =   -70000
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   11652
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
         Left            =   -69880
         TabIndex        =   29
         Top             =   4560
         Visible         =   0   'False
         Width           =   5292
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remesas - visualizar últimas"
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
         Height          =   300
         Index           =   4
         Left            =   -64600
         TabIndex        =   28
         Top             =   4560
         Visible         =   0   'False
         Width           =   5412
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   18
         Left            =   -62560
         TabIndex        =   27
         Top             =   6120
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
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
         Left            =   7320
         TabIndex        =   26
         Top             =   6000
         Width           =   1572
         _Version        =   1441793
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
         Left            =   -62320
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
         Left            =   -62320
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
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
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   50
      Top             =   7920
      Visible         =   0   'False
      Width           =   11655
      _Version        =   1441793
      _ExtentX        =   20558
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Traspaso de Desembolsos de Hipotecarios a Bancos"
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
      Height          =   492
      Left            =   1680
      TabIndex        =   40
      Top             =   360
      Width           =   9732
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12372
   End
End
Attribute VB_Name = "frmVivRemesasTesoreria"
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
     
            strSQL = "select isnull(max(Remesa),0) + 1 as Ultimo from viviendaRemesasTesoreria"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert viviendaRemesasTesoreria(Remesa,RegistroUsuario,RegistroFecha,Estado,FechaInicio,FechaCorte,notas) values(" & rs!Ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!Ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de CRD Hipotecario Traslado a Tesoreria : " & txtRemesa)
    
    Else
        If txtEstado.Text = "Abierta" Then
                    
            strSQL = "update viviendaRemesasTesoreria set RegistroUsuario = '" & glogon.Usuario & "',FechaInicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',FechaCorte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where Remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
             
            Call Bitacora("Modifica", "Remesa de CRD Hipotecario Traslado a Tesoreria : " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
  Case 1 'BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Abierta" Then
            strSQL = "delete viviendaRemesasTesoreria_detalle where Remesa = " & txtRemesa
            
            strSQL = strSQL & Space(10) & "update ViviendaDesembolsos set TesoreriaRemesa = Null" _
                        & ", TesoreriaSolicitud = Null, TesoreriaFecha = Nul, TesoreriaUsuario = Null" _
                        & "  where TesoreriaRemesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            
            Call Bitacora("Elimina", "Remesa de CRD Hipotecario Traslado a Tesoreria : " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case 2 'REPORTES"
'     fraReporte.Visible = Not fraReporte.Visible


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
    Call sbTraslado
  
  '---------Reportes
  Case 8
    Call sbInforme_Remesa

End Select


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 Resume

End Sub


Private Sub cboCarga_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date

If cboCarga.ListCount = 0 Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear

strSQL = "select FechaInicio,FechaCorte from viviendaRemesasTesoreria where Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FechaInicio
  vFechaCorte = rs!FechaCorte
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(pRemesa As Long)

Call sbLimpia

strSQL = "select T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
       & " from viviendaRemesasTesoreria T left join vCrd_Hipotecario_Remesa_Tes_Rsm D on T.Remesa = D.Remesa" _
       & " where T.Remesa = " & pRemesa

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa.Text = CStr(rs!Remesa)
  txtUsuario.Text = rs!RegistroUsuario
  txtFecha.Text = rs!RegistroFecha
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Abierta"
    Case "C"
      txtEstado = "Cerrada"
    Case "T"
      txtEstado = "Trasladada"
  End Select
  
  dtpInicio.Value = rs!FechaInicio
  dtpCorte.Value = rs!FechaCorte
  
  txtNotas.Text = rs!notas
  txtRemesa_Casos.Text = Format(rs!Casos, "###,##0")
  txtRemesa_Monto.Text = Format(rs!Monto, "Standard")
  
End If
rs.Close

End Sub


Private Sub sbInforme_Remesa()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String, xRemesa As String

On Error GoTo vError

If lblRemesa.Tag = "" Then Exit Sub

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
 .WindowTitle = "Reportes del Módulo de Crédito Hipotecario"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Tesorería")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If



 Select Case True
  Case opt.Item(0).Value 'Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_Hipotecario_RemesaTESDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Detalle Agrupado Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_Hipotecario_RemesaTESDetalleAgrp.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
 End Select

 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE PAGO: DESEMBOLSOS VIVIENDA'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"
 .SelectionFormula = "{viviendaRemesasTesoreria.Remesa} = " & lblRemesa.Tag
' .PrintReport
 .Action = 1

End With

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
   curTotal = curTotal + CCur(Item.SubItems(5))
Else
   curTotal = curTotal - CCur(Item.SubItems(5))
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
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(5))
   End If
  
Next i

txtCargaTotal.Text = Format(curTotal, "Standard")

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
     
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    
   
    txtNotas.Text = ""
     
     strSQL = "select TOP 50 T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
            & " from viviendaRemesasTesoreria T left join vCrd_Hipotecario_Remesa_Tes_Rsm D on T.Remesa = D.Remesa" _
            & " order by T.RegistroFecha desc"
     
     
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!Remesa)
                itmX.SubItems(1) = rs!RegistroUsuario
                itmX.SubItems(2) = rs!RegistroFecha
                
                Select Case rs!Estado
                  Case "A", "X"
                     itmX.SubItems(3) = "Abierta"
                  Case "C"
                     itmX.SubItems(3) = "Cerrada"
                  Case "T", "P"
                     itmX.SubItems(3) = "Trasladada"
                End Select
                
                itmX.SubItems(4) = Format(rs!FechaInicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!FechaCorte, "dd/mm/yyyy")
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
        
    strSQL = "select * from viviendaRemesasTesoreria where estado in('A','X') order by RegistroFecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!Remesa, "0000") & "..." & Trim(rs!RegistroUsuario) & "..." _
            & rs!RegistroFecha & " I:" & Format(rs!FechaInicio, "dd/mm/yyyy") & " C:" & Format(rs!FechaCorte, "dd/mm/yyyy"))
      
      cboCarga.ItemData(cboCarga.ListCount - 1) = CStr(rs!Remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!Remesa, "0000") & "..." & Trim(rs!RegistroUsuario) & "..." _
            & rs!RegistroFecha & " I:" & Format(rs!FechaInicio, "dd/mm/yyyy") & " C:" & Format(rs!FechaCorte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click
       
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
        
    strSQL = "select * from viviendaRemesasTesoreria where estado = 'C' order by RegistroFecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboTraslado.AddItem (Format(rs!Remesa, "0000") & "..." & Trim(rs!RegistroUsuario) & "..." _
            & rs!RegistroFecha & " I:" & Format(rs!FechaInicio, "dd/mm/yyyy") & " C:" & Format(rs!FechaCorte, "dd/mm/yyyy"))
      cboTraslado.ItemData(cboTraslado.ListCount - 1) = CStr(rs!Remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboTraslado.Text = (Format(rs!Remesa, "0000") & "..." & Trim(rs!RegistroUsuario) & "..." _
            & rs!RegistroFecha & " I:" & Format(rs!FechaInicio, "dd/mm/yyyy") & " C:" & Format(rs!FechaCorte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboTraslado_Click

  
  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
            & " from viviendaRemesasTesoreria T left join vCrd_Hipotecario_Remesa_Tes_Rsm D on T.Remesa = D.Remesa" _
            & " order by T.RegistroFecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!Remesa)
                itmX.SubItems(1) = rs!RegistroUsuario
                itmX.SubItems(2) = rs!RegistroFecha
                
                Select Case rs!Estado
                  Case "A", "X"
                     itmX.SubItems(3) = "Abierta"
                  Case "C"
                     itmX.SubItems(3) = "Cerrada"
                  Case "T", "P"
                     itmX.SubItems(3) = "Trasladada"
                End Select
                
      
                itmX.SubItems(4) = Format(rs!FechaInicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!FechaCorte, "dd/mm/yyyy")
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




Private Sub sbCerrar()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from viviendaRemesasTesoreria" _
       & " where Remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado in('A','X')"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
strSQL = "update viviendaRemesasTesoreria set estado = 'C'" _
       & " where Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("Aplica", "Cierra Remesa Crd Hipotecario Traslado a Tesoreria : " & cboCarga.ItemData(cboCarga.ListIndex))


MsgBox "Remesa Cerrada Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbLimpia

End Sub

Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If

End Sub



Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

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
    .Add , , "Id", 1100
    .Add , , "No.Operación", 1400, vbCenter
    .Add , , "Línea", 1000, vbCenter
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3400
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Beneficiario", 3400
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2400, vbCenter
    .Add , , "Duplicado?", 1400, vbCenter
 End With
 
 With lswTraslado.ColumnHeaders
    .Clear
    .Add , , "Id", 1100
    .Add , , "No.Operación", 1400, vbCenter
    .Add , , "Línea", 1000, vbCenter
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3400
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Beneficiario", 3400
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2400, vbCenter
    
 End With
     
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 btnBarra(9).Tag = btnBarra(0).Tag
 
 tcMain.Item(0).Selected = True
 
 Call sbRequiereAutorizacion
 
 
Exit Sub

vError:

 
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




Private Function fxTesToken() As String
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

fxTesToken = strToken

End Function










'Private Sub sbReporte()
'Dim vSubTitulo As String, vFiltro As String
'Dim strSQL As String
'
'On Error GoTo vError
'
'Me.MousePointer = vbHourglass
'
'vSubTitulo = ""
'vFiltro = ""
'strSQL = ""
'
'
'With frmContenedor.Crt
' .Reset
' .WindowShowGroupTree = True
' .WindowShowPrintSetupBtn = True
' .WindowShowRefreshBtn = True
' .WindowShowSearchBtn = True
' .WindowState = crptMaximized
' .WindowTitle = "Reportes del Módulo de Credito Hipotecario"
'
' .Connect = glogon.ConectRPT
'
'
' .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
' .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
' .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
' .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
' .Formulas(4) = "fxFiltro='" & vFiltro & "'"
'
' .ReportFileName = App.Path & SIFGlobal.fxPathReportes("Credito_Hipotecario_Remesas.rpt")
' .PrintReport
'
'End With
'
'Me.MousePointer = vbDefault
'Exit Sub
'
'vError:
' Me.MousePointer = vbDefault
' MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'
'End Sub

Private Sub sbCargaBuscar()
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0



strSQL = "select FechaInicio,FechaCorte from viviendaRemesasTesoreria where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FechaInicio
  vFechaCorte = rs!FechaCorte
rs.Close


strSQL = "select D.CodigoDesembolso,D.NumeroOperacion,D.Beneficiario,D.Monto,D.RegistroFecha,D.RegistroUsuario" _
       & ",S.cedula,S.nombre,R.codigo,D.TES_SUPERVISION_FECHA  " _
       & ",dbo.fxTesSupervisa(D.Identificacion,D.Beneficiario,D.Monto,0,'V') as 'Duplicado'" _
       & " From ViviendaDesembolsos D inner join Reg_Creditos R on D.numeroOperacion = R.id_solicitud" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where D.TesoreriaRemesa is null" _
       & " and D.RegistroFecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") _
       & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswCarga.ListItems.Add(, , rs!CodigoDesembolso)
     If rs!duplicado = 1 And IsNull(rs!TES_SUPERVISION_FECHA) Then
          itmX.ForeColor = vbRed
          vDuplicado = True
         strLista = strLista & rs!CodigoDesembolso & " " & rs!cedula & " " & Format(rs!Monto, "Standard") & vbCrLf
       Else
          itmX.ForeColor = vbBlack
       End If

     itmX.SubItems(1) = rs!NumeroOperacion
     itmX.SubItems(2) = rs!codigo
     itmX.SubItems(3) = rs!cedula
     itmX.SubItems(4) = rs!Nombre
     itmX.SubItems(5) = Format(rs!Monto, "Standard")
     itmX.SubItems(6) = rs!Beneficiario
     itmX.SubItems(7) = rs!RegistroFecha
     itmX.SubItems(8) = rs!RegistroUsuario
     
     If rs!duplicado = 1 And IsNull(rs!TES_SUPERVISION_FECHA) Then
        itmX.ForeColor = vbRed
        vDuplicado = True
        strLista = strLista & rs!NumeroOperacion & " " & rs!codigo & " " & rs!cedula & " " & Format(rs!Monto, "Standard") & vbCrLf
     Else
        itmX.ForeColor = vbBlack
     End If
     itmX.SubItems(9) = IIf(vDuplicado = True, rs!duplicado, 0)

     itmX.Checked = chkCarga.Value

     If itmX.Checked Then
        curTotal = curTotal + CCur(itmX.SubItems(5))
     End If

 rs.MoveNext

 PrgBar.Value = PrgBar.Value + 1

Loop
rs.Close

PrgBar.Visible = False

txtCargaTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

If vDuplicado = True Then
   MsgBox "Estas operaiones necesitan autorización para ser trasladadas ya que cuentan" _
          & "con una transacción por un monto igual en Tesorería " & vbCrLf & vbCrLf & strLista, vbCritical
End If


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
strSQL = "select count(*) as Existe from viviendaRemesasTesoreria" _
       & " where remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado in('A','X') "
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
 If .Item(i).Checked And .Item(i).SubItems(9) = 0 Then

     strSQL = "update viviendaDesembolsos set TesoreriaRemesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
            & " where CodigoDesembolso = " & .Item(i).Text
     Call ConectionExecute(strSQL)

    PrgBar.Value = PrgBar.Value + 1
  End If
Next i

If vCasos > 0 Then
    'Actualiza el Estado de la Remesa como cerrada
    strSQL = "update viviendaRemesasTesoreria set estado = 'X'" _
           & " where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
    Call ConectionExecute(strSQL)
    Call Bitacora("Genera", "Remesa de Desembolsos de Vivienda: " & cboCarga.ItemData(cboCarga.ListIndex))
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


Private Sub sbTrasladoBuscar()
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear
curTotal = 0

strSQL = "select D.CodigoDesembolso,D.NumeroOperacion,D.Beneficiario,D.Monto,D.RegistroFecha,D.RegistroUsuario" _
       & ",S.cedula,S.nombre,R.codigo " _
       & " From ViviendaDesembolsos D inner join Reg_Creditos R on D.numeroOperacion = R.id_solicitud" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where D.TesoreriaFecha is null and D.TesoreriaRemesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswTraslado.ListItems.Add(, , rs!CodigoDesembolso)
     itmX.SubItems(1) = rs!NumeroOperacion
     itmX.SubItems(2) = rs!codigo
     itmX.SubItems(3) = rs!cedula
     itmX.SubItems(4) = rs!Nombre
     itmX.SubItems(5) = Format(rs!Monto, "Standard")
     itmX.SubItems(6) = rs!Beneficiario
     itmX.SubItems(7) = rs!RegistroFecha
     itmX.SubItems(8) = rs!RegistroUsuario

     curTotal = curTotal + CCur(itmX.SubItems(5))

 rs.MoveNext
 PrgBar.Value = PrgBar.Value + 1
Loop
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
Dim i As Integer, vCasos As Integer
Dim vFecha As Date
Dim strToken As String

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from ViviendaRemesasTesoreria" _
       & " where Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
       & " and estado in('C') "
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra en procesada...", vbExclamation
    Exit Sub
 End If
rs.Close


strSQL = "select top 1 id_token from tes_tokens where estado = 'A' order by registro_fecha "
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  strToken = rs!id_token
Else
  strToken = fxTesToken
End If
rs.Close


'Actualiza el Estado de la Remesa como Cola de Pago / Al finalizar Revisa si ya fue Totalmente Pagada
strSQL = "update ViviendaRemesasTesoreria set estado = 'P'" _
       & " where Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call ConectionExecute(strSQL)

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

PrgBar.Max = lswTraslado.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswTraslado.ListItems

For i = 1 To .Count
'spAFIComisionPago

     strSQL = "exec spCRDViviendaDesembolsoPago " & cboTraslado.ItemData(cboTraslado.ListIndex) & "," & .Item(i).Text _
            & ",'" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & strToken & "'"
     Call ConectionExecute(strSQL)


     Call Bitacora("Aplica", "Desembolso de Vivienda a Tesoreria Remesa:" & cboTraslado.ItemData(cboTraslado.ListIndex) _
                    & " IdDesem:" & .Item(i).Text)

    PrgBar.Value = PrgBar.Value + 1
Next i

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub


vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswTraslado.ListItems.Clear

End Sub



