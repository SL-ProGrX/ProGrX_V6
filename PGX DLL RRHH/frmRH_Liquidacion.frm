VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmRH_Liquidacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Liquidación de la Persona"
   ClientHeight    =   8220
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11295
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7092
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   11412
      _Version        =   1441793
      _ExtentX        =   20129
      _ExtentY        =   12509
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
      Item(0).Caption =   "Liquidación"
      Item(0).ControlCount=   46
      Item(0).Control(0)=   "btnAplicar"
      Item(0).Control(1)=   "cboTipo"
      Item(0).Control(2)=   "txtNotas"
      Item(0).Control(3)=   "txtSalarioPromedio"
      Item(0).Control(4)=   "txtSalarioDiario"
      Item(0).Control(5)=   "txtVacaciones"
      Item(0).Control(6)=   "txtCesantia"
      Item(0).Control(7)=   "txtCesantiaAsoc"
      Item(0).Control(8)=   "txtPreaviso"
      Item(0).Control(9)=   "txtPreavisoDias"
      Item(0).Control(10)=   "txtTotalLiquidar"
      Item(0).Control(11)=   "dtpIngreso"
      Item(0).Control(12)=   "dtpSalida"
      Item(0).Control(13)=   "txtDiasBase"
      Item(0).Control(14)=   "txtVacacionesDias"
      Item(0).Control(15)=   "txtEstadoActual"
      Item(0).Control(16)=   "txtEstadoResultante"
      Item(0).Control(17)=   "txtAguinaldo"
      Item(0).Control(18)=   "txtNomina"
      Item(0).Control(19)=   "Label15(14)"
      Item(0).Control(20)=   "Label15(1)"
      Item(0).Control(21)=   "Label15(13)"
      Item(0).Control(22)=   "Label15(12)"
      Item(0).Control(23)=   "Label15(11)"
      Item(0).Control(24)=   "Label15(10)"
      Item(0).Control(25)=   "Label15(9)"
      Item(0).Control(26)=   "Label15(7)"
      Item(0).Control(27)=   "Label15(8)"
      Item(0).Control(28)=   "Label15(6)"
      Item(0).Control(29)=   "Label15(5)"
      Item(0).Control(30)=   "Label15(3)"
      Item(0).Control(31)=   "Label15(2)"
      Item(0).Control(32)=   "lblUltSalarios"
      Item(0).Control(33)=   "Label15(0)"
      Item(0).Control(34)=   "Label15(4)"
      Item(0).Control(35)=   "txtMeses"
      Item(0).Control(36)=   "Label15(15)"
      Item(0).Control(37)=   "dtpRenuncia"
      Item(0).Control(38)=   "chkConResponsabilidad"
      Item(0).Control(39)=   "txtNominaId"
      Item(0).Control(40)=   "txtNominaCorte"
      Item(0).Control(41)=   "Label15(16)"
      Item(0).Control(42)=   "Label15(17)"
      Item(0).Control(43)=   "Label15(18)"
      Item(0).Control(44)=   "txtCesantiaDias"
      Item(0).Control(45)=   "txtCesantiaAnios"
      Item(1).Caption =   "Aguinaldo"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lswAguinaldo"
      Item(1).Control(1)=   "scAguinaldo"
      Item(2).Caption =   "Salarios"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "ShortcutCaption1"
      Item(2).Control(1)=   "lswSalarios"
      Begin XtremeSuiteControls.ListView lswSalarios 
         Height          =   6252
         Left            =   -69880
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1441793
         _ExtentX        =   19494
         _ExtentY        =   11028
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
      Begin XtremeSuiteControls.ListView lswAguinaldo 
         Height          =   6252
         Left            =   -69880
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1441793
         _ExtentX        =   19494
         _ExtentY        =   11028
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
      Begin XtremeSuiteControls.CheckBox chkConResponsabilidad 
         Height          =   372
         Left            =   6120
         TabIndex        =   48
         Top             =   4080
         Width           =   3132
         _Version        =   1441793
         _ExtentX        =   5524
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Liquidación con Responsabilidad Patronal ?"
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
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   612
         Left            =   7680
         TabIndex        =   7
         Top             =   6360
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmRH_Liquidacion.frx":0000
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   2400
         TabIndex        =   8
         Top             =   4920
         Width           =   6852
         _Version        =   1441793
         _ExtentX        =   12091
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   912
         Left            =   2400
         TabIndex        =   9
         Top             =   5280
         Width           =   6852
         _Version        =   1441793
         _ExtentX        =   12086
         _ExtentY        =   1609
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtSalarioPromedio 
         Height          =   312
         Left            =   7200
         TabIndex        =   10
         Top             =   600
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSalarioDiario 
         Height          =   312
         Left            =   7200
         TabIndex        =   11
         Top             =   960
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtVacaciones 
         Height          =   312
         Left            =   7200
         TabIndex        =   12
         Top             =   1920
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCesantia 
         Height          =   312
         Left            =   7200
         TabIndex        =   13
         Top             =   2280
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCesantiaAsoc 
         Height          =   312
         Left            =   7200
         TabIndex        =   14
         Top             =   2760
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
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
         Text            =   "0.00"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPreaviso 
         Height          =   312
         Left            =   7200
         TabIndex        =   15
         Top             =   3240
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPreavisoDias 
         Height          =   312
         Left            =   9360
         TabIndex        =   16
         Top             =   3240
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   550
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
         Text            =   "0"
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalLiquidar 
         Height          =   312
         Left            =   7200
         TabIndex        =   17
         Top             =   3720
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpIngreso 
         Height          =   312
         Left            =   2400
         TabIndex        =   18
         Top             =   1560
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.DateTimePicker dtpRenuncia 
         Height          =   312
         Left            =   2400
         TabIndex        =   19
         Top             =   600
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtDiasBase 
         Height          =   312
         Left            =   2400
         TabIndex        =   20
         Top             =   4080
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   550
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
         Text            =   "30"
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtVacacionesDias 
         Height          =   312
         Left            =   9360
         TabIndex        =   21
         Top             =   1920
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   550
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
         Text            =   "0"
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstadoActual 
         Height          =   312
         Left            =   2400
         TabIndex        =   22
         Top             =   2040
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      End
      Begin XtremeSuiteControls.FlatEdit txtEstadoResultante 
         Height          =   312
         Left            =   2400
         TabIndex        =   23
         Top             =   2400
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      End
      Begin XtremeSuiteControls.FlatEdit txtAguinaldo 
         Height          =   312
         Left            =   7200
         TabIndex        =   24
         Top             =   1560
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNomina 
         Height          =   312
         Left            =   2400
         TabIndex        =   25
         Top             =   2880
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      End
      Begin XtremeSuiteControls.FlatEdit txtMeses 
         Height          =   312
         Left            =   2400
         TabIndex        =   46
         Top             =   4440
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   550
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
         Text            =   "0"
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNominaId 
         Height          =   312
         Left            =   2400
         TabIndex        =   49
         Top             =   3240
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNominaCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   50
         Top             =   3600
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpSalida 
         Height          =   312
         Left            =   2400
         TabIndex        =   53
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtCesantiaDias 
         Height          =   312
         Left            =   9360
         TabIndex        =   55
         ToolTipText     =   "Dias a Reconocer por Años Laborados"
         Top             =   2280
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   550
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
         Text            =   "0"
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCesantiaAnios 
         Height          =   312
         Left            =   9840
         TabIndex        =   56
         ToolTipText     =   "Años a Reconocer"
         Top             =   2280
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   550
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
         Text            =   "0"
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Salida:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   18
         Left            =   600
         TabIndex        =   54
         Top             =   960
         Width           =   1452
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ult. Nómina: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   17
         Left            =   600
         TabIndex        =   52
         Top             =   3600
         Width           =   1812
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Ultima Nómina Id: "
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
         Index           =   16
         Left            =   600
         TabIndex        =   51
         Top             =   3240
         Width           =   1812
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Meses Base: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   15
         Left            =   600
         TabIndex        =   47
         Top             =   4440
         Width           =   1332
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   492
         Left            =   -69880
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1441793
         _ExtentX        =   19494
         _ExtentY        =   868
         _StockProps     =   14
         Caption         =   "Ultimos Salarios Utilizados para el Promedio"
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
      Begin XtremeShortcutBar.ShortcutCaption scAguinaldo 
         Height          =   492
         Left            =   -69880
         TabIndex        =   43
         Top             =   360
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1441793
         _ExtentX        =   19494
         _ExtentY        =   868
         _StockProps     =   14
         Caption         =   "Meses Incluidos en el Aguinaldo"
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
      Begin VB.Label Label15 
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
         Height          =   372
         Index           =   4
         Left            =   600
         TabIndex        =   41
         Top             =   5280
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
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
         Index           =   0
         Left            =   600
         TabIndex        =   40
         Top             =   4920
         Width           =   1092
      End
      Begin VB.Label lblUltSalarios 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Promedio Ultimos 6 Salarios:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4440
         TabIndex        =   39
         Top             =   600
         Width           =   2412
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salario diario:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   4440
         TabIndex        =   38
         Top             =   960
         Width           =   2412
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vacaciones:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   4440
         TabIndex        =   37
         Top             =   1920
         Width           =   2412
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cesantía:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   5
         Left            =   4440
         TabIndex        =   36
         Top             =   2280
         Width           =   2412
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Cesantía Trasladada Asociación:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   6
         Left            =   4440
         TabIndex        =   35
         Top             =   2640
         Width           =   2412
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Preaviso:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   8
         Left            =   4440
         TabIndex        =   34
         Top             =   3240
         Width           =   2412
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Liquidar:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   7
         Left            =   4440
         TabIndex        =   33
         Top             =   3720
         Width           =   2412
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ingreso:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   9
         Left            =   600
         TabIndex        =   32
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Renuncia:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   10
         Left            =   600
         TabIndex        =   31
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Días base: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   11
         Left            =   600
         TabIndex        =   30
         Top             =   4080
         Width           =   1332
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Inicial: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   12
         Left            =   600
         TabIndex        =   29
         Top             =   2040
         Width           =   1332
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Final: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   13
         Left            =   600
         TabIndex        =   28
         Top             =   2400
         Width           =   1332
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aguinaldo:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   4440
         TabIndex        =   27
         Top             =   1560
         Width           =   2412
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nómina: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   14
         Left            =   600
         TabIndex        =   26
         Top             =   2880
         Width           =   1332
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8640
      Top             =   120
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEmpleadoId 
      Height          =   312
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   5892
      _Version        =   1441793
      _ExtentX        =   10393
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Id. Empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmRH_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean, mHrsBase As Integer

Private Sub btnAplicar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass


'spRH_Liquidacion_Registro(@EmpleadoId varchar(20), @Nomina varchar(10), @NominaId int, @fNomina datetime, @fIngreso datetime, @fRenuncia datetime, @fSalida datetime
'            , @TipoLiq varchar(10), @Notas varchar(500), @ConResponsabilidad smallint, @PreavisoActiva smallint
'            , @EstadoInicial varchar(10), @EstadoResult varchar(10), @DiasBase smallint, @MesesBase smallint
'            , @SalarioPromedio dec(14,2), @SalarioDiario dec(14,2), @Aguinaldo dec(14,2)
'            , @Vacaciones dec(14,2), @VacaDias dec(10,2), @Preaviso dec(14,2), @PreavisoDias smallint
'            , @Cesantia dec(14,2), @CesantiaDias dec(10,2), @CesantiaAnios dec(10,2), @Asociacion dec(14,2)
'            , @TotalLiq dec(14,2), @Usuario varchar(30))

If txtNominaId.Text = "" Then
  txtNominaId.Text = "0"
  txtNominaCorte.Text = Format(dtpIngreso.Value, "yyyy-mm-dd")
End If

strSQL = "exec spRH_Liquidacion_Registro '" & txtEmpleadoId.Text & "','" & txtNomina.Text & "', " & txtNominaId.Text & ", '" & txtNominaCorte.Text _
       & "','" & Format(dtpIngreso.Value, "yyyy-mm-dd") & "','" & Format(dtpRenuncia.Value, "yyyy-mm-dd") & "','" & Format(dtpSalida.Value, "yyyy-mm-dd") _
       & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "','" & txtNotas.Text & "'," & chkConResponsabilidad.Value & ", " & IIf(txtPreavisoDias.Locked, 0, 1) _
       & ", '" & txtEstadoActual.Tag & "','" & txtEstadoResultante.Tag & "'," & txtDiasBase.Text & ", " & txtMeses.Text _
       & ", " & CCur(txtSalarioPromedio.Text) & ", " & CCur(txtSalarioDiario.Text) & ", " & CCur(txtAguinaldo.Text) _
       & ", " & CCur(txtVacaciones.Text) & ", " & CCur(txtVacacionesDias.Text) & ", " & CCur(txtPreaviso.Text) & ", " & CCur(txtPreavisoDias.Text) _
       & ", " & CCur(txtCesantia.Text) & ", " & CCur(txtCesantiaDias.Text) & ", " & CCur(txtCesantiaAnios.Text) & ", " & CCur(txtCesantiaAsoc.Text) _
       & ", " & CCur(txtTotalLiquidar.Text) & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Call sbBoleta_Liquidacion(rs!BOLETA_LIQ)

Call sbLimpia

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub

If cboTipo.ListCount = 0 Then Exit Sub

strSQL = "select T.Estado_Persona, T.Estado_Persona_Resultante, T.Con_Responsabilidad, T.Preaviso_Activa" _
       & ", El.Descripcion as 'EstadoResDesc'" _
       & ", Ea.Descripcion as 'EstadoActDesc'" _
       & " from RH_LIQUIDACION_TIPOS T" _
       & " inner join RH_ESTADOS_TIPOS El on T.Estado_Persona_Resultante = El.Estado_Persona" _
       & " inner join RH_ESTADOS_TIPOS Ea on T.Estado_Persona = Ea.Estado_Persona" _
       & " WHERE T.TIPO_LIQUIDACION = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
 txtEstadoActual.Tag = rs!ESTADO_PERSONA
 txtEstadoActual.Text = rs!EstadoActDesc
 
 txtEstadoResultante.Tag = rs!ESTADO_PERSONA_RESULTANTE
 txtEstadoResultante.Text = rs!EstadoResDesc
 
 txtPreavisoDias.Text = "0"
 
 chkConResponsabilidad.Value = rs!CON_RESPONSABILIDAD
 If rs!PREAVISO_ACTIVA = 1 Then
    txtPreavisoDias.Locked = False
 Else
    txtPreavisoDias.Locked = True
 End If
 
rs.Close

Call sbConsulta(1)

End Sub

Private Sub dtpSalida_Change()
If vPaso Then Exit Sub

Call sbConsulta(1)

End Sub

Private Sub Form_Load()

vModulo = 23

tcMain.Item(0).Selected = True

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpia()

txtEmpleadoId.Text = ""
txtIdentificacion.Text = ""
txtNombre.Text = ""

txtVacaciones.Text = "0"
txtAguinaldo.Text = "0"
txtPreaviso.Text = "0"
txtTotalLiquidar.Text = "0"
txtCesantia.Text = "0"
txtSalarioDiario.Text = "0"
txtCesantiaAsoc.Text = "0"

txtNomina.Text = ""
txtNominaId.Text = ""
txtNominaCorte.Text = "2000-01-01"

Call sbConsulta

End Sub

Private Sub sbInicializa()

On Error GoTo vError

vPaso = True

    dtpIngreso.Value = fxFechaServidor
    dtpSalida.Value = dtpIngreso.Value
  
    strSQL = "select TIPO_LIQUIDACION as Idx, rtrim(Descripcion) as ItmX" _
           & " from RH_LIQUIDACION_TIPOS" _
           & " Where Activo = 1"
    Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

vPaso = False

txtMeses.Text = fxgRH_Parametro("01")
lblUltSalarios.Caption = "Promedio Ultimos " & Trim(txtMeses.Text) & " Salarios:"

txtEmpleadoId.SetFocus


Call cboTipo_Click


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbAguinaldo_List()

On Error GoTo vError

Me.MousePointer = vbHourglass

'strSQL = "exec spRH_Aguinaldos_List '" & txtNomina.Text _
'       & "','" & Format(dtpSalida.Value, "yyyy-mm-dd") & "','D','" & txtEmpleadoId.Text & "'"

strSQL = "exec spRH_Aguinaldos_List '" & txtNomina.Text _
       & "','" & txtNominaCorte.Text & "','D','" & txtEmpleadoId.Text & "'"

With lswAguinaldo.ColumnHeaders
    .Clear
    .Add , , "Corte", 1400
    .Add , , "Qty", 1000, vbCenter
    .Add , , "Salario", 1800, vbRightJustify
    .Add , , "T.Ingresos", 1800, vbRightJustify
    .Add , , "T.Egresos", 1800, vbRightJustify
    .Add , , "Salario Neto", 1800, vbRightJustify
    .Add , , "Salario AP", 1800, vbRightJustify
    
End With

lswAguinaldo.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswAguinaldo.ListItems.Add(, , Format(rs!Corte, "yyyy-mm-dd"))
     itmX.SubItems(1) = rs!Qty
     itmX.SubItems(2) = Format(rs!Salario_Mes, "Standard")
     itmX.SubItems(3) = Format(rs!Total_Ingresos, "Standard")
     itmX.SubItems(4) = Format(rs!Total_Egresos, "Standard")
     itmX.SubItems(5) = Format(rs!Salario_Neto, "Standard")
     itmX.SubItems(6) = Format(rs!Salario_Devengado, "Standard")
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbSalarios_List()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Liquidacion_Salarios_List '" & txtNomina.Text _
       & "','" & txtNominaCorte.Text & "','D','" & txtEmpleadoId.Text & "'"

With lswSalarios.ColumnHeaders
    .Clear
    .Add , , "Corte", 1400
    .Add , , "Salario", 1800, vbRightJustify
    .Add , , "T.Ingresos", 1800, vbRightJustify
    .Add , , "T.Egresos", 1800, vbRightJustify
    .Add , , "Salario Neto", 1800, vbRightJustify
    .Add , , "Salario AP", 1800, vbRightJustify
    
End With

lswSalarios.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswSalarios.ListItems.Add(, , Format(rs!Corte, "yyyy-mm-dd"))
     itmX.SubItems(1) = Format(rs!Salario_Mes, "Standard")
     itmX.SubItems(2) = Format(rs!Total_Ingresos, "Standard")
     itmX.SubItems(3) = Format(rs!Total_Egresos, "Standard")
     itmX.SubItems(4) = Format(rs!Salario_Neto, "Standard")
     itmX.SubItems(5) = Format(rs!Salario_Devengado, "Standard")
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 1 'Aguinaldos
        Call sbAguinaldo_List
        
    Case 2 'Salarios para el Promedio
        Call sbSalarios_List
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub sbBusca()
   gBusquedas.Convertir = "N"
   gBusquedas.Col1Name = "Empleado Id"
   gBusquedas.Col2Name = "Persona Id"
   gBusquedas.Col3Name = "Nombre"
   gBusquedas.Columna = "Empleado_ID"
   gBusquedas.Orden = "Empleado_ID"
   gBusquedas.Consulta = "Select Empleado_ID,Identificacion,Nombre_Completo From Rh_Personas"
   gBusquedas.Filtro = " and ESTADO_PERSONA = 'A'"
   
   frmBusquedas.Show vbModal
   
   txtEmpleadoId.Text = gBusquedas.Resultado
   txtIdentificacion.Text = Trim(gBusquedas.Resultado2)
   txtNombre.Text = gBusquedas.Resultado3
    
   Call sbConsulta
    
End Sub

Public Sub sbConsulta_Externa(pEmpleadoId As String)

txtEmpleadoId.Text = pEmpleadoId
Call sbConsulta

End Sub


Private Sub sbCalculos()

On Error GoTo vError

Dim pSalarioDiario As Currency, pLiquida As Currency

pSalarioDiario = (CCur(txtSalarioPromedio.Text) / mHrsBase) * 8

txtSalarioDiario.Text = Format(pSalarioDiario, "Standard")

txtVacaciones.Text = Format(pSalarioDiario * CInt(txtVacacionesDias.Text), "Standard")

If Not txtPreavisoDias.Locked Then
    txtPreaviso.Text = Format(pSalarioDiario * CInt(txtPreavisoDias.Text), "Standard")
Else
    txtPreaviso.Text = Format(0, "Standard")
End If

pLiquida = CCur(txtAguinaldo.Text) + CCur(txtVacaciones.Text) + CCur(txtPreaviso.Text)

If chkConResponsabilidad.Value = xtpUnchecked Then
    txtCesantia.Text = Format(0, "Standard")
End If

If CCur(txtCesantia.Text) > CCur(txtCesantiaAsoc.Text) Then
    pLiquida = pLiquida + (CCur(txtCesantia.Text) - CCur(txtCesantiaAsoc.Text))
End If

If chkConResponsabilidad.Value = xtpUnchecked Then
    txtCesantia.Text = Format(0, "Standard")
End If

txtTotalLiquidar.Text = Format(pLiquida, "Standard")

Exit Sub

vError:

End Sub


Private Sub sbConsulta(Optional pSegmentos As Integer = 0)
Dim pFecha As Date

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True



If pSegmentos = 0 Then
    strSQL = "select EMPLEADO_ID,IDENTIFICACION,NOMBRE_COMPLETO, Fecha_Ingreso" _
           & " , isnull(SALIDA_FECHA , dbo.Mygetdate() ) as 'Fecha_Salida', COD_NOMINA, dbo.Mygetdate() as 'Fecha'" _
           & " from Rh_Personas" _
           & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
        
        dtpIngreso.Value = rs!FECHA_INGRESO
        dtpIngreso.MaxDate = rs!FECHA_INGRESO
        dtpIngreso.MinDate = rs!FECHA_INGRESO
        
        dtpRenuncia.Value = rs!Fecha_Salida
        
        dtpSalida.Value = rs!Fecha_Salida
        
        pFecha = rs!Fecha
        
        txtEmpleadoId.Text = rs!Empleado_ID
        txtIdentificacion.Text = rs!IDENTIFICACION
        txtNombre.Text = rs!NOMBRE_COMPLETO
        
        txtNomina.Text = rs!COD_NOMINA
        
    End If
    rs.Close
Else
    pFecha = dtpSalida.Value
End If



strSQL = "exec spRH_Liquidacion_Load_Inicial '" & txtNomina.Text & "', '" & txtEmpleadoId.Text _
       & "', '" & Format(pFecha, "yyyy-mm-dd") & "'"
Call OpenRecordSet(rs, strSQL)

mHrsBase = rs!Horas

txtAguinaldo.Text = Format(rs!Aguinaldo, "Standard")
txtSalarioPromedio.Text = Format(rs!Salarios, "Standard")
txtCesantiaAsoc.Text = Format(rs!Asociacion, "Standard")

txtVacacionesDias.Text = Format(rs!Vaca_Dias, "###0")
txtPreavisoDias.Text = "0"

txtCesantia.Text = Format(rs!Cesantia, "Standard")
txtCesantiaDias.Text = Format(rs!CESANTIA_DIAS, "Standard")

txtCesantiaAnios.Text = Format(rs!CESANTIA_ANIOS, "Standard")

If Not IsNull(rs!Nomina_Num) Then
    txtNominaId.Text = rs!Nomina_Num
    txtNominaCorte.Text = Format(rs!Nomina_Corte, "yyyy-mm-dd")
'    dtpSalida.MinDate = rs!Nomina_Corte
End If

rs.Close

Call sbCalculos

Me.MousePointer = vbDefault

vPaso = False
Exit Sub


vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtCesantiaAsoc_KeyUp(KeyCode As Integer, Shift As Integer)
Call sbCalculos
End Sub

Private Sub txtCesantiaAsoc_LostFocus()
txtCesantiaAsoc.Text = Format(CCur(txtCesantiaAsoc.Text), "Standard")
End Sub

Private Sub txtEmpleadoId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub


Private Sub txtPreavisoDias_KeyUp(KeyCode As Integer, Shift As Integer)

Call sbCalculos

End Sub



Private Sub txtVacacionesDias_KeyUp(KeyCode As Integer, Shift As Integer)

Call sbCalculos

End Sub
