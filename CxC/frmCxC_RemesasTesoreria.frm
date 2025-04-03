VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxC_RemesasTesoreria 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CxC: Remesa de desembolsos (Tesorería)"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   11910
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10680
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   8010
      Visible         =   0   'False
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
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
      ItemCount       =   5
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
      Item(2).ControlCount=   9
      Item(2).Control(0)=   "Label8(14)"
      Item(2).Control(1)=   "Label2(16)"
      Item(2).Control(2)=   "cboTraslado"
      Item(2).Control(3)=   "lswTraslado"
      Item(2).Control(4)=   "txtPagoTotal"
      Item(2).Control(5)=   "btnBarra(6)"
      Item(2).Control(6)=   "btnBarra(7)"
      Item(2).Control(7)=   "Label8(19)"
      Item(2).Control(8)=   "chkTrasladoAgrupar"
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
      Item(4).Caption =   "Reactivación"
      Item(4).ControlCount=   6
      Item(4).Control(0)=   "Label3(0)"
      Item(4).Control(1)=   "cmdReactivar"
      Item(4).Control(2)=   "Label8(15)"
      Item(4).Control(3)=   "Label8(16)"
      Item(4).Control(4)=   "txtOperacion"
      Item(4).Control(5)=   "txtDetalle"
      Begin XtremeSuiteControls.ListView lswRep 
         Height          =   3612
         Left            =   -70000
         TabIndex        =   3
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
      Begin XtremeSuiteControls.ListView lswTraslado 
         Height          =   4092
         Left            =   -69880
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.ListView lswCarga 
         Height          =   3972
         Left            =   -70000
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ListView lswRemesas 
         Height          =   3132
         Left            =   1560
         TabIndex        =   6
         Top             =   3240
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
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   0
         Left            =   -69520
         TabIndex        =   7
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox fraReporte 
         Height          =   2055
         Left            =   4200
         TabIndex        =   8
         Top             =   960
         Width           =   7455
         _Version        =   1441793
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
            TabIndex        =   9
            Top             =   1200
            Width           =   1572
            _Version        =   1441793
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
            Left            =   4800
            TabIndex        =   10
            Top             =   360
            Width           =   1215
            _Version        =   1441793
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
            Left            =   3240
            TabIndex        =   11
            Top             =   360
            Width           =   1335
            _Version        =   1441793
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
            TabIndex        =   12
            Top             =   360
            Width           =   1335
            _Version        =   1441793
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
            TabIndex        =   13
            Top             =   720
            Width           =   4932
            _Version        =   1441793
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
            TabIndex        =   14
            Top             =   1200
            Width           =   612
            _Version        =   1441793
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
            Picture         =   "frmCxC_RemesasTesoreria.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   1
            Left            =   6360
            TabIndex        =   15
            Top             =   1200
            Width           =   492
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
            Picture         =   "frmCxC_RemesasTesoreria.frx":0707
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   1
            Left            =   3600
            TabIndex        =   16
            Top             =   1200
            Width           =   1692
            _Version        =   1441793
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
            TabIndex        =   18
            Top             =   360
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
            Index           =   8
            Left            =   360
            TabIndex        =   17
            Top             =   720
            Width           =   1452
            _Version        =   1441793
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
            Picture         =   "frmCxC_RemesasTesoreria.frx":0D45
            ToolTipText     =   "Actualizar Oficinas"
            Top             =   360
            Width           =   240
         End
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   0
         Left            =   4320
         TabIndex        =   19
         Top             =   480
         Width           =   1332
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":1435
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa 
         Height          =   432
         Left            =   1560
         TabIndex        =   20
         Top             =   480
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   312
         Left            =   1560
         TabIndex        =   21
         Top             =   1680
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   5160
         TabIndex        =   22
         Top             =   1320
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   5160
         TabIndex        =   23
         Top             =   1680
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1560
         TabIndex        =   24
         Top             =   2040
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1560
         TabIndex        =   25
         Top             =   1320
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
         Left            =   2760
         TabIndex        =   26
         Top             =   1320
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
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   1
         Left            =   6120
         TabIndex        =   27
         Top             =   480
         Width           =   492
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":1A67
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   2
         Left            =   6600
         TabIndex        =   28
         Top             =   480
         Width           =   492
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":200B
      End
      Begin XtremeSuiteControls.ComboBox cboCarga 
         Height          =   312
         Left            =   -67600
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1441793
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
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1441793
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
      Begin XtremeSuiteControls.CheckBox chkCarga 
         Height          =   252
         Left            =   -69880
         TabIndex        =   31
         Top             =   1680
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
         Left            =   -67840
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   432
         Left            =   -68440
         TabIndex        =   33
         Top             =   1080
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   3
         Left            =   -63880
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":2712
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   4
         Left            =   -62560
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":2E12
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   5
         Left            =   -61240
         TabIndex        =   36
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":352B
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   6
         Left            =   -63040
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":3C37
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   7
         Left            =   -61720
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":4337
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   8
         Left            =   -60760
         TabIndex        =   39
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCxC_RemesasTesoreria.frx":4C08
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdReactivar 
         Height          =   420
         Left            =   -63280
         TabIndex        =   40
         Top             =   5760
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8276
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":530F
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   312
         Left            =   -59200
         TabIndex        =   41
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkRemesaInd 
         Height          =   372
         Left            =   -60640
         TabIndex        =   42
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
         TabIndex        =   43
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtCargaTotal 
         Height          =   315
         Left            =   -60760
         TabIndex        =   44
         Top             =   6120
         Visible         =   0   'False
         Width           =   2415
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPagoTotal 
         Height          =   312
         Left            =   -60880
         TabIndex        =   45
         Top             =   6000
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   3912
         Left            =   -68440
         TabIndex        =   46
         Top             =   1560
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   9
         Left            =   5640
         TabIndex        =   47
         Top             =   480
         Width           =   492
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
         Picture         =   "frmCxC_RemesasTesoreria.frx":5A0F
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa_Monto 
         Height          =   312
         Left            =   8520
         TabIndex        =   48
         Top             =   1680
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa_Casos 
         Height          =   312
         Left            =   8520
         TabIndex        =   49
         Top             =   1320
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTrasladoAgrupar 
         Height          =   372
         Left            =   -65200
         TabIndex        =   71
         Top             =   960
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Agrupar por Beneficiario   "
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   0
         Left            =   600
         TabIndex        =   70
         Top             =   480
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
         Left            =   600
         TabIndex        =   69
         Top             =   1320
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
         Left            =   4320
         TabIndex        =   68
         Top             =   1320
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
         Left            =   600
         TabIndex        =   67
         Top             =   1680
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
         Left            =   4320
         TabIndex        =   66
         Top             =   1680
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
         Left            =   600
         TabIndex        =   65
         Top             =   2040
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
         Left            =   1560
         TabIndex        =   64
         Top             =   2880
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
         TabIndex        =   63
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
         Height          =   372
         Index           =   10
         Left            =   -69400
         TabIndex        =   62
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
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
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lista de Operaciones Pendientes a Trasladar"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   60
         Top             =   1560
         Visible         =   0   'False
         Width           =   11412
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
         Left            =   -70000
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   11652
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   58
         Top             =   4560
         Visible         =   0   'False
         Width           =   5292
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
         Left            =   -64600
         TabIndex        =   57
         Top             =   4560
         Visible         =   0   'False
         Width           =   5412
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   $"frmCxC_RemesasTesoreria.frx":6140
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Index           =   0
         Left            =   -69880
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   11412
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   15
         Left            =   -69760
         TabIndex        =   55
         Top             =   1080
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
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
         TabIndex        =   54
         Top             =   1440
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
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
         Height          =   252
         Index           =   18
         Left            =   -62560
         TabIndex        =   53
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
         Left            =   -62680
         TabIndex        =   52
         Top             =   6000
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
         Height          =   372
         Index           =   22
         Left            =   7680
         TabIndex        =   51
         Top             =   1680
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
         Left            =   7680
         TabIndex        =   50
         Top             =   1320
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Generar Solicitudes para Desembolsos en Bancos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Top             =   0
      Width           =   13092
   End
End
Attribute VB_Name = "frmCxC_RemesasTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset

Dim itmX As ListViewItem, vPaso As Boolean

Dim mRequiereAutorizacion As Boolean

Dim vDuplicado As Boolean, strLista  As String
Dim mUnidad As String, mConcepto As String

Private Sub btnBarra_Click(Index As Integer)
Dim i As Integer

On Error GoTo vError

Select Case Index
  Case 0 'NUEVO"
     
    Call sbLimpia
    
  Case 9 'GUARDAR
    If txtRemesa.Text = "" Then

            strSQL = "select isnull(max(Tesoreria_Remesa),0) + 1 as Ultimo from CxC_REMESAS_TES"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert CxC_REMESAS_TES(Tesoreria_Remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas) values(" & rs!ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)

                txtRemesa = rs!ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de CxC Traslado a Tesoreria : " & txtRemesa)

    Else
        If txtEstado.Text = "Abierta" Then

            strSQL = "update CxC_REMESAS_TES set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where Tesoreria_Remesa = " & txtRemesa
             Call ConectionExecute(strSQL)

            Call Bitacora("Modifica", "Remesa de CxC Traslado a Tesoreria : " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If

    Call sbLimpia
    
  Case 1 'BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Abierta" Then
            strSQL = "update CxC_Cuentas set Tesoreria_Remesa = Null where Tesoreria_Remesa = " & txtRemesa.Text
            Call ConectionExecute(strSQL)

            strSQL = "delete CxC_REMESAS_TES where Tesoreria_Remesa = " & txtRemesa.Text
            Call ConectionExecute(strSQL)


            Call Bitacora("Elimina", "Remesa de CxC Traslado a Tesoreria : " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case 2 'REPORTES"
     fraReporte.Left = 4200
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

strSQL = "select fecha_inicio,fecha_corte from CxC_REMESAS_TES where Tesoreria_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL, 0)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close


'Carga Oficinas
strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as Itmx" _
       & " from SIF_Oficinas  where cod_oficina in(" _
       & " select R.Cod_Oficina" _
       & " from CxC_Cuentas R inner join CxC_Conceptos C on R.cod_concepto = C.cod_concepto " _
       & " where R.Autoriza_Estado='F' and R.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.Tesoreria_Estado is null and R.estado in('A','C') and Operacion not in(select Operacion from CxC_Cuentas)" _
       & " group by R.Cod_Oficina)" _
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
       & " from CxC_REMESAS_TES T left join vCxC_Remesa_Tes_Rsm D on T.TESORERIA_REMESA = D.TESORERIA_REMESA" _
       & " where T.TESORERIA_REMESA = " & pRemesa

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa.Text = CStr(rs!TESORERIA_REMESA)
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
  
  dtpInicio.Value = rs!Fecha_Inicio
  dtpCorte.Value = rs!Fecha_Corte
  
  txtNotas.Text = rs!Notas
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
 .WindowTitle = "Reportes del Módulo de Cuentas por Cobrar"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Traslado a Tesoreria")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If

 Select Case True
  Case opt.Item(0).Value 'Pendiente Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("CxC_RemesaTESDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Traslado Detalle Agrupado Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("CxC_RemesaTESDetalleAgrp.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA TRASLADO A TESORERIA : CxC'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = "{CxC_REMESAS_TES.Tesoreria_Remesa} = " & lblRemesa.Tag
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 31
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
strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as Itmx" _
       & " from SIF_Oficinas  where cod_oficina in(" _
       & " select R.Cod_Oficina" _
       & " from CxC_Cuentas R inner join CxC_Conceptos C on R.cod_concepto = C.cod_concepto " _
       & " where R.Autoriza_Estado='F' and R.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and R.TESORERIA_FECHA is null and R.estado in('A','C') and Operacion not in(select Operacion from CxC_Cuentas)" _
       & " group by R.Cod_Oficina)" _
       & " order by cod_oficina"
Call sbCbo_Llena_New(cboRepOficina, strSQL, True, True)

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

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 Call sbLimpia
End Sub


Private Sub lswTraslado_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswTraslado.SortKey = ColumnHeader.Index - 1
  If lswTraslado.SortOrder = 0 Then lswTraslado.SortOrder = 1 Else lswTraslado.SortOrder = 0
  lswTraslado.Sorted = True
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


Private Sub SSTab_Click(PreviousTab As Integer)
 Call sbLimpia
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
 
 .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionRemesas.rpt")
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbLimpia()

On Error GoTo vError

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
    
     dtpRepInicio.Value = dtpInicio.Value
     dtpRepCorte.Value = dtpInicio.Value
    
     fraReporte.Visible = False
    
     txtNotas.Text = ""
     
     strSQL = "select TOP 50 T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
            & " from CxC_REMESAS_TES T left join vCxC_Remesa_Tes_Rsm D on T.TESORERIA_REMESA = D.TESORERIA_REMESA" _
            & " order by T.fecha desc"
     
     
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!TESORERIA_REMESA)
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
                
                itmX.SubItems(4) = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!Fecha_Corte, "dd/mm/yyyy")
                itmX.SubItems(6) = Format(rs!Casos, "###,###0")
                itmX.SubItems(7) = Format(rs!Monto, "Standard")
                itmX.SubItems(8) = rs!Notas
                
                
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
        
    strSQL = "select * from CxC_REMESAS_TES where estado = 'A' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
      cboCarga.ItemData(cboCarga.ListCount - 1) = CStr(rs!TESORERIA_REMESA)

      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click
   
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
        
    strSQL = "select * from CxC_REMESAS_TES where estado = 'C' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboTraslado.AddItem (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
      cboTraslado.ItemData(cboTraslado.ListCount - 1) = CStr(rs!TESORERIA_REMESA)
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboTraslado.Text = (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboTraslado_Click

  
  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " T.*, isnull(D.Casos,0) as 'Casos', isnull(D.Monto,0) as 'Monto' " _
            & " from CxC_REMESAS_TES T left join vCxC_Remesa_Tes_Rsm D on T.TESORERIA_REMESA = D.TESORERIA_REMESA" _
            & " order by T.fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!TESORERIA_REMESA)
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
                
      
                itmX.SubItems(4) = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!Fecha_Corte, "dd/mm/yyyy")
                itmX.SubItems(6) = Format(rs!Casos, "###,###0")
                itmX.SubItems(7) = Format(rs!Monto, "Standard")
                itmX.SubItems(8) = rs!Notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close


    
  Case 4 'Re-Activaciones
    txtOperacion.Tag = 0
    txtOperacion = ""
    txtDetalle = ""
 

 End Select


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub tlbTraslado_ButtonClick(ByVal Button As MSComctlLib.Button)

If cboTraslado.ListCount = 0 Then Exit Sub

Select Case Button.Key
  Case "buscar"
    Call sbTrasladoBuscar
  
  Case "traslado"
    Call sbTraslado

End Select

End Sub


Private Sub sbCargaBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from CxC_REMESAS_TES where Tesoreria_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close


strSQL = "select * from vCxC_Cuentas_Desembolsos_Pendientes" _
       & " where Activa_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'"

If cboOficina.Text <> "TODOS" Then
   strSQL = strSQL & " and Cod_Oficina = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
End If

strSQL = strSQL & " order by Operacion"

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswCarga
 .ListItems.Clear
 Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!Operacion)
       itmX.SubItems(1) = rs!cod_Concepto
       itmX.SubItems(2) = rs!Cedula
       itmX.SubItems(3) = rs!Nombre
       itmX.SubItems(4) = Format(rs!Monto, "Standard")
       itmX.SubItems(5) = Format(rs!DESEMBOLSO_PENDIENTE, "Standard")
       itmX.SubItems(6) = "0.00"
       itmX.SubItems(7) = "0.00"
       itmX.SubItems(8) = Format(rs!DESEMBOLSO_PENDIENTE, "Standard")
       itmX.Checked = chkCarga.Value
         
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(8))
       End If
        
        rs.MoveNext
        
        PrgBar.Value = PrgBar.Value + 1
 Loop
End With

rs.Close

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
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from CxC_REMESAS_TES where Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close

'
'strSQL = "select R.Operacion,R.cod_concepto,S.cedula,S.nombre,R.Monto,R.Desembolso_Monto" _
'       & " from CxC_Cuentas R inner join CxC_Personas S on R.cedula = S.cedula" _
'       & " inner join CxC_Conceptos C on R.cod_concepto = C.cod_concepto " _
'       & " where R.Autoriza_Estado='A' and R.Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
'       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
'       & " and R.estado in('A','C') and R.Tesoreria_Fecha is null" _
'       & " and R.Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
'       & " order by R.Operacion"

strSQL = "select * from vCxC_Cuentas_Desembolsos_Cargados" _
       & " where Registro_Fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
       & " and Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
       & " order by Operacion"
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswTraslado
 .ListItems.Clear
 Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!Operacion)
       itmX.SubItems(1) = rs!cod_Concepto
       itmX.SubItems(2) = rs!Cedula
       itmX.SubItems(3) = rs!Nombre
       itmX.SubItems(4) = Format(rs!Monto, "Standard")
       itmX.SubItems(5) = Format(rs!Desembolso_Monto, "Standard")
       
       itmX.SubItems(6) = "0.00"
       itmX.SubItems(7) = "0.00"
       itmX.SubItems(8) = Format(rs!Desembolso_Monto, "Standard")

  
       curTotal = curTotal + CCur(itmX.SubItems(8))
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CxC_REMESAS_TES" _
       & " where Tesoreria_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
strSQL = "update CxC_REMESAS_TES set estado = 'C'" _
       & " where Tesoreria_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CxC_REMESAS_TES" _
       & " where Tesoreria_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
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

strSQL = ""
For i = 1 To .Count
 If .Item(i).Checked Then
 
'     strSQL = "update CxC_Cuentas set Tesoreria_Remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
'            & " where Operacion = " & .Item(i).Text
     strSQL = strSQL & Space(10) & "exec spCxC_Cuenta_Desembolso_Carga " & .Item(i).Text & "," & cboCarga.ItemData(cboCarga.ListIndex) _
            & ",'" & glogon.Usuario & "'"
     
     If Len(strSQL) > 20000 Then
         Call ConectionExecute(strSQL)
         strSQL = ""
     End If
     
    PrgBar.Value = PrgBar.Value + 1
  End If
Next i
 
'Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If
 
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Carga Remesa Traslado a Tesoreria : " & cboCarga.ItemData(cboCarga.ListIndex))
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



Private Sub tlbCarga_ButtonClick(ByVal Button As MSComctlLib.Button)
If cboCarga.ListCount = 0 Then Exit Sub

Select Case Button.Key
  Case "buscar"
    Call sbCargaBuscar
  
  Case "cargar"
    If lswCarga.ListItems.Count = 0 Then Exit Sub
    Call sbCarga
  
  Case "cerrar"
    Call sbCerrar
End Select

End Sub



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
'
'dtpRepInicio.Value = fxFechaServidor
'dtpRepCorte.Value = dtpRepInicio.Value

Call sbLimpia


End Sub

Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If

End Sub


Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Integer, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String) As Long                                 'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,tipo_Beneficiario,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza,fecha_autorizacion)" _
       & " values('" & mConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "',5,'" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CxC','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = strSQL & ",'N',null,null)"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL, 0)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL, 0)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "'"
  rsX.CursorLocation = adUseServer
  Call OpenRecordSet(rsX, strSQL, 0)
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
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function


Private Function fxCtaPuente(pCodigo As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select Cod_Cuenta_Salida from CxC_Conceptos where cod_concepto ='" & pCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
     fxCtaPuente = ""
Else
     fxCtaPuente = rsX!cod_Cuenta_Salida
End If

rsX.Close

End Function


Private Sub sbCreaDesembolsos(vReferencia As Long, vOP As Long, vFecha As Date, vTipo As String, vBanco As Integer)
Dim rsTemp As New ADODB.Recordset, lngSolicitud As Long
'
'strSQL = "select * from desembolsos where retener = 0 and Operacion = " & vOP
'
'With rsTemp
' .CursorLocation = adUseServer
' .Open strSQL, glogon.Conection, adOpenStatic
' Do While Not .EOF
'     lngSolicitud = fxMaestroTesoreria(vTipo, vBanco, !Monto, !id_desembolso _
'                   , !Concepto, !Operacion, !Operacion, vReferencia, !cod_concepto, "0", vFecha)
'     Call sbCreaDetalle(lngSolicitud, fxCtaBanco(vBanco), !Monto, "H", 1)
'     Call sbCreaDetalle(lngSolicitud, !cuenta_conta, !Monto, "D", 2)
'
'     strSQL = "update desembolsos set tdocumento = '" & vTipo & "',Emitir_Tipo_Banco = " & vBanco & ",nsolicitud = " & lngSolicitud _
'            & " where id_desembolso = " & !id_desembolso
'     Call ConectionExecute(strSQL)
'  .MoveNext
' Loop
' .Close
'End With

End Sub

Private Sub sbTraslado()
Dim rsTmp As New ADODB.Recordset
Dim lngSolicitud As Long, vFecha As Date, vLinea As Integer
Dim vTipo As String, vBanco As Integer

Me.MousePointer = vbHourglass

On Error GoTo vError

vFecha = fxFechaServidor
mUnidad = fxCxC_Parametro("05")
mConcepto = fxCxC_Parametro("06")

If chkTrasladoAgrupar.Value = vbChecked Then
'    strSQL = "select R.cod_concepto,S.cedula,S.nombre,R.Emitir_Tipo,R.Emitir_Banco,sum(R.Desembolso_Monto) as 'Desembolso_Monto',R.Emitir_Cuenta" _
'           & ",Ofi.cod_unidad,Ofi.cod_centro_Costo,R.cod_Contrato,C.cod_Cuenta_Salida as 'ConceptoCta',B.CTACONTA as 'BancoCta'" _
'           & ",RTRIM(MIN(R.Num_Documento)) + '-' + rtrim(MAX(R.Num_Documento)) as 'Num_documento',rtrim(CONVERT(varchar(30),min(R.Operacion))) + '-' + rtrim(CONVERT(varchar(30),max(R.Operacion)))  as 'Operacion'" _
'           & " from CxC_Cuentas R inner join CxC_Personas S on R.cedula = S.cedula" _
'           & " inner join CxC_Conceptos C on R.cod_concepto = C.cod_Concepto" _
'           & " inner join Tes_Bancos B on R.Emitir_Banco = B.id_Banco" _
'           & " inner join SIF_Oficinas Ofi on R.cod_Oficina = Ofi.cod_Oficina" _
'           & " where R.estado in('A','C') and R.Autoriza_Estado = 'A' and R.tesoreria_Fecha is null" _
'           & " and R.Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
'           & " group by S.cedula,S.nombre,R.Emitir_Tipo,R.Emitir_Banco,R.Emitir_Cuenta" _
'           & ",Ofi.cod_unidad,Ofi.cod_centro_Costo,R.cod_concepto,R.cod_Contrato,C.cod_Cuenta_Salida,B.CTACONTA"
           
    strSQL = "select cedula, nombre, Emitir_Tipo, Emitir_Banco,sum(Desembolso_Monto) as 'Desembolso_Monto',Emitir_Cuenta" _
           & ",cod_unidad, cod_centro_Costo, BancoCta, 'Agrupado' as 'Cod_Concepto'" _
           & ",RTRIM(MIN(Num_Documento)) + '..' + rtrim(MAX(Num_Documento)) as 'Num_documento',rtrim(CONVERT(varchar(30),min(Operacion))) + '..' + rtrim(CONVERT(varchar(30),max(Operacion)))  as 'Operacion'" _
           & " from vCxC_Cuentas_Desembolsos_Traslado_Pendiente" _
           & " Where Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
           & " group by cedula,nombre,Emitir_Tipo,Emitir_Banco,Emitir_Cuenta" _
           & ",cod_unidad,cod_centro_Costo,BancoCta"
           
           

Else
'    strSQL = "select R.Operacion,R.cod_concepto,S.cedula,S.nombre,R.Emitir_Tipo,R.Emitir_Banco,R.Desembolso_Monto,R.Emitir_Cuenta" _
'           & ",Ofi.cod_unidad,Ofi.cod_centro_Costo,R.num_documento,R.cod_Contrato,C.cod_Cuenta_Salida as 'ConceptoCta',B.CTACONTA as 'BancoCta'" _
'           & " from CxC_Cuentas R inner join CxC_Personas S on R.cedula = S.cedula" _
'           & " inner join CxC_Conceptos C on R.cod_concepto = C.cod_Concepto" _
'           & " inner join Tes_Bancos B on R.Emitir_Banco = B.id_Banco" _
'           & " inner join SIF_Oficinas Ofi on R.cod_Oficina = Ofi.cod_Oficina" _
'           & " where R.estado in('A','C') and R.Autoriza_Estado = 'A' and R.tesoreria_Fecha is null" _

'           & " and R.Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
    
    strSQL = "select * from vCxC_Cuentas_Desembolsos_Traslado_Pendiente" _
           & " Where Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
            
           
End If

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True


Do While Not rs.EOF

 'Graba y Devuelve el registro Maestro en Tesoreria ('Or rs!Emitir_Tipo = "ND")
 
 If rs!Desembolso_Monto > 0 And (rs!Emitir_Tipo = "CK" Or rs!Emitir_Tipo = "TE") Then
    lngSolicitud = fxMaestroTesoreria(rs!Emitir_Tipo, rs!Emitir_Banco, rs!Desembolso_Monto, Trim(rs!Cedula) _
                   , rs!Nombre, 0, "Ops:" & rs!Operacion & " Cp:" & rs!cod_Concepto, 0 _
                   , ("Docs:" & Trim(rs!Num_Documento)), rs!Emitir_Cuenta, vFecha, rs!Cod_Unidad)
                   
    'Mata el Pasivo de la Nota de Debito de la Formalizacion contra Tes_Bancos
    Call sbCreaDetalle(lngSolicitud, rs!BancoCta, rs!Desembolso_Monto, "H", 1, rs!Cod_Unidad)
    
    If chkTrasladoAgrupar.Value = vbChecked Then
        strSQL = "select cod_concepto, cedula, nombre, Emitir_Tipo, Emitir_Banco,sum(Desembolso_Monto) as 'Desembolso_Monto',Emitir_Cuenta" _
               & ",cod_unidad, cod_centro_Costo, cod_Contrato, ConceptoCta, BancoCta" _
               & ",RTRIM(MIN(Num_Documento)) + '..' + rtrim(MAX(Num_Documento)) as 'Num_documento',rtrim(CONVERT(varchar(30),min(Operacion))) + '..' + rtrim(CONVERT(varchar(30),max(Operacion)))  as 'Operacion'" _
               & " from vCxC_Cuentas_Desembolsos_Traslado_Pendiente" _
               & " Where Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
               & "   and Cedula = '" & rs!Cedula & "'" _
               & " group by cedula,nombre,Emitir_Tipo,Emitir_Banco,Emitir_Cuenta" _
               & ",cod_unidad,cod_centro_Costo,cod_concepto,cod_Contrato,ConceptoCta,BancoCta"
       Call OpenRecordSet(rsTmp, strSQL)
       vLinea = 2
       Do While Not rsTmp.EOF
           Call sbCreaDetalle(lngSolicitud, rsTmp!ConceptoCta, rsTmp!Desembolso_Monto, "D", vLinea, rsTmp!Cod_Unidad)
           vLinea = vLinea + 1
           rsTmp.MoveNext
       Loop
       rsTmp.Close
    
    Else
        Call sbCreaDetalle(lngSolicitud, rs!ConceptoCta, rs!Desembolso_Monto, "D", 2, rs!Cod_Unidad)
    End If

 Else 'Monto a Girar > 0
   
   lngSolicitud = 0
 
 End If
  
 'Actualiza Campo Tesoreria
 If chkTrasladoAgrupar.Value = vbChecked Then
'    strSQL = "update CxC_Cuentas set tesoreria_Fecha = dbo.MyGetdate(), Tesoreria_Usuario = '" & glogon.Usuario _
'           & "',Tesoreria_Solicitud = " & lngSolicitud & ",Tesoreria_Estado = 'G'" _
'           & " where cedula = '" & rs!Cedula & "' and Emitir_Tipo = '" & rs!Emitir_Tipo & "' and Emitir_Banco = " & rs!Emitir_Banco _
'           & " and cod_concepto = '" & rs!Cod_Concepto & "' and cod_contrato = '" & rs!cod_contrato _
'           & "' and estado in('A','C') and Autoriza_Estado = 'A' and tesoreria_Fecha is null" _
'           & " and Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
    
    strSQL = "exec spCxC_Cuenta_Desembolso_TesoreriaId_Agrupado '" & rs!Cedula & "'," & cboTraslado.ItemData(cboTraslado.ListIndex) _
           & "," & lngSolicitud & ",'" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
    
    'Actualiza Bitacora
    Call Bitacora("Registra", "Traspaso a Tesoreria para Desembolso Cliente :" & rs!Cedula & "-" & rs!cod_Concepto)
 
 Else
'    strSQL = "update CxC_Cuentas set tesoreria_Fecha = dbo.MyGetdate(), Tesoreria_Usuario = '" & glogon.Usuario _
'           & "',Tesoreria_Solicitud = " & lngSolicitud & ",Tesoreria_Estado = 'G' where Operacion = " & rs!Operacion
           
    strSQL = "exec spCxC_Cuenta_Desembolso_TesoreriaId " & rs!Operacion & "," & cboTraslado.ItemData(cboTraslado.ListIndex) _
           & "," & lngSolicitud & ",'" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
    
    'Actualiza Bitacora
    Call Bitacora("Registra", "Traspaso a Tesoreria de la Operacion y Desembol OP:" & rs!Operacion)
 End If
 
 If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
 rs.MoveNext
 
Loop
rs.Close




'Actualiza y Carga Remesa
strSQL = "update CxC_REMESAS_TES SET Estado = 'T'" _
       & "  Where Tesoreria_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call ConectionExecute(strSQL)

Call sbLimpia


Me.MousePointer = vbDefault

PrgBar.Visible = False

MsgBox "Operaciones Enviadas a Tesoreria Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 Resume

End Sub


Private Sub cmdReactivar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtOperacion.Tag = 1 Then
  
  strSQL = "update CxC_Cuentas set tesoreria_Fecha = null where Operacion = " & txtOperacion
  Call ConectionExecute(strSQL)
  
  
  Call Bitacora("Aplica", "ReActivacion Traslado Tes. Op:" & txtOperacion)
  
'  'Tags de Seguimiento
'  Call sbCrdOperacionTags(txtOperacion.Text, txtDetalle.Tag, "S04", "", ">>> Re.Activación del Desembolso <<<")
  
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
Dim strSQL As String, rs As New ADODB.Recordset

Dim strTitulo As String
Dim strRuta As String, strInicio As String, strFinal As String


On Error GoTo vError

Me.MousePointer = vbHourglass

strTitulo = "Operaciones pendientes de Traslado a Tesorería"


strRuta = SIFGlobal.fxPathReportes("CxC_Tesoreria_Envio.rpt")
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
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Titulo='" & strTitulo & "'"
    
    strSQL = "{CxC_Cuentas.Autoriza_Estado} = 'F'"
    If chkRepFechas.Value = vbUnchecked Then
      strSQL = strSQL & " and {CxC_Cuentas.Registro_Fecha} >= Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ")" _
             & " and {CxC_Cuentas.Registro_Fecha} <= Date(" & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
        .Formulas(4) = "de='" & Format(dtpRepInicio.Value, "dd/mm/yyyy") & "'"
        .Formulas(5) = "a='" & Format(dtpRepCorte.Value, "dd/mm/yyyy") & "'"
    Else
        .Formulas(4) = "de=' --- '"
        .Formulas(5) = "a=' --- '"
    End If
    
    
    If cboRepOficina.Text <> "TODOS" Then
       strSQL = strSQL & " AND {CxC_Cuentas.Cod_Oficina} = '" & SIFGlobal.fxCodText(cboRepOficina.Text) & "'"
    End If
    
    
    strSQL = strSQL & " and ISNULL({CxC_Cuentas.TESORERIA}) AND {CxC_Cuentas.ESTADO}='A'"
    
    .SelectionFormula = strSQL
    
    .SubreportToChange = "subCkDesembolsos"
    .SelectionFormula = "{DESEMBOLSOS.Operacion} = {?Pm-CxC_Cuentas.Operacion}"
    
    .PrintReport
    

End With

 Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporteEnviadas()
Dim strSQL As String, rs As New ADODB.Recordset

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

.ReportFileName = SIFGlobal.fxPathReportes("CxC_Tesoreria_Envio_Rec.rpt")
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

vModulo = 31

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

 
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
 
 
 tcMain.Item(0).Selected = True
 
strSQL = "select rtrim(cod_oficina) as 'Idx', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbCbo_Llena_New(cboRepOficina, strSQL, True, False)
 
 Call Formularios(Me)
 
 btnBarra(9).Tag = btnBarra(0).Tag
 
 Call RefrescaTags(Me)
 
 
 
 Call sbLimpia
 
 
Exit Sub

vError:
 
End Sub


Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

Dim rsTmp As New ADODB.Recordset

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 txtDetalle = ""
 txtOperacion.Tag = 0
 strSQL = "select R.Operacion,R.cod_concepto,R.cedula,R.Desembolso_Monto,C.descripcion,S.nombre,R.Tesoreria_Solicitud" _
        & " from CxC_Cuentas R inner join CxC_Personas S on R.cedula = S.cedula" _
        & " inner join CxC_Conceptos C on R.cod_concepto = C.cod_concepto" _
        & " where R.Operacion = " & txtOperacion & " and R.estado = 'A'"
 Call OpenRecordSet(rs, strSQL)
 If Not rs.EOF And Not rs.BOF Then
   txtDetalle.Tag = rs!cod_Concepto
   txtDetalle = txtDetalle & "Línea         : " & rs!cod_Concepto & vbCrLf
   txtDetalle = txtDetalle & "Descripción   : " & rs!DESCRIPCION & vbCrLf
   txtDetalle = txtDetalle & "Cédula        : " & rs!Cedula & vbCrLf
   txtDetalle = txtDetalle & "Nombre        : " & rs!Nombre & vbCrLf
   txtDetalle = txtDetalle & "Monto a Girar : " & Format(rs!Desembolso_Monto, "Standard") & vbCrLf
   
   'Verifica que no existan documentos emitidos con anterioridad
   strSQL = "select NSOLICITUD,id_banco,tipo,ndocumento from Tes_Transacciones" _
          & " where Nsolicitud = " & rs!Tesoreria_Solicitud & " and estado in('I','T','P')"
   Call OpenRecordSet(rsTmp, strSQL, 0)
   If rsTmp.EOF And rsTmp.BOF Then
       txtOperacion.Tag = 1
'      'Verificar si tiene desembolsos asociados en Tesoreria
'      rsTmp.Close
'      strSQL = "select NSOLICITUD,id_banco,tipo,ndocumento from Tes_Transacciones" _
'             & " where op = " & txtOperacion & " and estado in('I','T','P')"
'      Call OpenRecordSet(rsTmp, strSQL, 0)
'      If Not rsTmp.EOF And Not rsTmp.BOF Then
'         txtOperacion.Tag = 0
'         txtDetalle = txtDetalle & "EXISTEN DESEMBOLSOS ASOCIADOS EN TESORERIA" & vbCrLf
'      End If
   
   Else 'Mov del Deudor Directamente
      txtOperacion.Tag = 0
      txtDetalle = txtDetalle & " / EXISTE UN DOCUMENTO O SOLICITUD DE EMISION EN TESORERIA / " & rs!Nombre & vbCrLf
      txtDetalle = txtDetalle & "Solicitud :" & rsTmp!NSolicitud & vbCrLf
      txtDetalle = txtDetalle & "Documento :" & rsTmp!nDocumento & vbCrLf
      txtDetalle = txtDetalle & "Tipo/Banco:" & rsTmp!Tipo & "/" & rsTmp!Id_Banco & vbCrLf
   End If
   
   rsTmp.Close
   
   
   'Segunda verificacion con el nuevo esquema
   
   If CLng(txtOperacion.Tag) = 1 Then
    'Verificar AQUI; pero como deseable porque el cod_concepto nuevo es compatible con este cod_concepto de verificacion
   
   
   
   End If
   
 
 Else
   
   MsgBox "La Operacion Digitada no existe...", vbExclamation
 
 End If
 rs.Close

End If

End Sub


