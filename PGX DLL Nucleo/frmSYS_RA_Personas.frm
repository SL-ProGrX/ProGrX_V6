VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmSYS_RA_Personas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RA Expedientes: Personas con Accesos Restringidos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6252
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11292
      _Version        =   1310723
      _ExtentX        =   19918
      _ExtentY        =   11028
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
      SelectedItem    =   1
      Item(0).Caption =   "Consultas"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "txtConCedula"
      Item(0).Control(2)=   "txtConNombre"
      Item(0).Control(3)=   "Label1(26)"
      Item(0).Control(4)=   "dtpConInicio"
      Item(0).Control(5)=   "Label1(0)"
      Item(0).Control(6)=   "dtpConCorte"
      Item(0).Control(7)=   "chkConFechas"
      Item(0).Control(8)=   "cboConEstado"
      Item(0).Control(9)=   "Label1(4)"
      Item(0).Control(10)=   "btnBuscar"
      Item(0).Control(11)=   "btnExport"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   17
      Item(1).Control(0)=   "txtNombre"
      Item(1).Control(1)=   "txtCedula"
      Item(1).Control(2)=   "Label1(2)"
      Item(1).Control(3)=   "Label1(1)"
      Item(1).Control(4)=   "txtCodigo"
      Item(1).Control(5)=   "txtNotas"
      Item(1).Control(6)=   "Label1(19)"
      Item(1).Control(7)=   "cboEstado"
      Item(1).Control(8)=   "Label1(8)"
      Item(1).Control(9)=   "Label1(9)"
      Item(1).Control(10)=   "cboTipo"
      Item(1).Control(11)=   "Label1(5)"
      Item(1).Control(12)=   "chkVence"
      Item(1).Control(13)=   "dtpVence"
      Item(1).Control(14)=   "btnGuardar"
      Item(1).Control(15)=   "btnNuevo"
      Item(1).Control(16)=   "btnAutorizacion"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4695
         Left            =   -69880
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1310723
         _ExtentX        =   19500
         _ExtentY        =   8281
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
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkVence 
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   3240
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Vence ?"
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
      End
      Begin XtremeSuiteControls.CheckBox chkConFechas 
         Height          =   252
         Left            =   -62200
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtConCedula 
         Height          =   312
         Left            =   -68560
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
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
      Begin XtremeSuiteControls.FlatEdit txtConNombre 
         Height          =   312
         Left            =   -66880
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1310723
         _ExtentX        =   9758
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   330
         Left            =   -61240
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   582
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
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   330
         Left            =   -60040
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Exportar"
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpConInicio 
         Height          =   312
         Left            =   -64960
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.DateTimePicker dtpConCorte 
         Height          =   312
         Left            =   -63640
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   3120
         TabIndex        =   9
         Top             =   2280
         Width           =   6495
         _Version        =   1310723
         _ExtentX        =   11456
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   4800
         TabIndex        =   10
         Top             =   1800
         Width           =   4815
         _Version        =   1310723
         _ExtentX        =   8488
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   330
         Left            =   3120
         TabIndex        =   11
         Top             =   1800
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2984
         _ExtentY        =   582
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   6360
         TabIndex        =   12
         Top             =   840
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Picture         =   "frmSYS_RA_Personas.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnNuevo 
         Height          =   495
         Left            =   5040
         TabIndex        =   13
         Top             =   840
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmSYS_RA_Personas.frx":0731
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   570
         Left            =   3120
         TabIndex        =   14
         Top             =   840
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   1005
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1095
         Left            =   3120
         TabIndex        =   15
         Top             =   3720
         Width           =   6495
         _Version        =   1310723
         _ExtentX        =   11451
         _ExtentY        =   1926
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   3120
         TabIndex        =   16
         Top             =   2760
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Top             =   3240
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.ComboBox cboConEstado 
         Height          =   312
         Left            =   -68560
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.PushButton btnAutorizacion 
         Height          =   495
         Left            =   7680
         TabIndex        =   30
         Top             =   840
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Autoriza"
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
         Picture         =   "frmSYS_RA_Personas.frx":0D63
         ImageAlignment  =   4
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo"
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
         Left            =   2040
         TabIndex        =   28
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   26
         Left            =   -69880
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   -66520
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Persona Id"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   2040
         TabIndex        =   22
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   21
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   20
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado"
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
         Index           =   4
         Left            =   -69520
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   852
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expedientes Restringidos"
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
      Height          =   495
      Index           =   3
      Left            =   1880
      TabIndex        =   27
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmSYS_RA_Personas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vCodigo As String


Public Sub sbConsulta_Externa(pPersonaId As Long)

Call sbCaso_Load(pPersonaId)

End Sub


Private Sub btnAutorizacion_Click()

If IsNumeric(txtCodigo.Text) Then
  Dim frm As Form


 Call sbFormsCall("frmSYS_RA_Autorizaciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmSYS_RA_Autorizaciones") Then
     Call frm.sbConsulta_Externa(txtCodigo.Text)
     Exit For
   End If
 Next frm
 
End If

End Sub

Private Sub btnExport_Click()
 Call Excel_Exportar_Lsw(lsw)
End Sub


Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub


Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVence.SetFocus
End Sub


Private Sub chkConFechas_Click()
If chkConFechas.Value = vbChecked Then
  dtpConInicio.Enabled = False
Else
  dtpConInicio.Enabled = True
End If
dtpConCorte.Enabled = dtpConInicio.Enabled
End Sub

Private Sub btnBuscar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

txtConCedula.Text = fxSysCleanTxtInject(txtConCedula.Text)
txtConNombre.Text = fxSysCleanTxtInject(txtConNombre.Text)

strSQL = "select *, isnull(Fecha_Vence, '2300/01/01') as 'Vence_Fix' from vSYS_RA_Casos" _
       & " where cedula like '%" & txtConCedula & "%' and nombre like '%" & txtConNombre.Text & "%'" _
       & " and Estado = '" & Mid(cboConEstado.Text, 1, 1) & "'"
       
If chkConFechas.Value = vbUnchecked Then
   strSQL = strSQL & " and isnull(Fecha_Vence, '2300/01/01') between '" & Format(dtpConInicio.Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpConCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Persona_Id)
     itmX.SubItems(1) = rs!EstadoDesc
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = rs!TipoDesc
     itmX.SubItems(5) = Format(rs!registro_Fecha, "yyyy-mm-dd")
     itmX.SubItems(6) = Format(rs!Vence_Fix, "yyyy-mm-dd")
     itmX.SubItems(7) = rs!Notas
     itmX.Tag = rs!Persona_Id
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkVence_Click()
If chkVence.Value = xtpChecked Then
    dtpVence.Enabled = True
Else
    dtpVence.Enabled = False
End If
End Sub

Private Sub btnGuardar_Click()


If Not fxVerifica Then
  Exit Sub
End If

On Error GoTo vError

Dim pVence As String

txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)

If chkVence.Value = xtpChecked Then
    pVence = "'" & Format(dtpVence.Value, "yyyy-mm-dd") & "'"
Else
    pVence = "Null"
End If

If txtCodigo.Text = "" Then
        
   strSQL = "exec spSYS_RA_Persona_Add 0, '" & txtCedula.Text & "', '" & Mid(cboEstado.Text, 1, 1) _
          & "', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', " & pVence & ", '" & txtNotas.Text _
          & "', '" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
    txtCodigo.Text = CStr(rs!Persona_Id)
   rs.Close
   
   Call Bitacora("Registra", "Expediente Restringido: " & txtCodigo.Text & " Cedula = " & txtCedula)

Else
   strSQL = "exec spSYS_RA_Persona_Add " & txtCodigo.Text & ", '" & txtCedula.Text & "', '" & Mid(cboEstado.Text, 1, 1) _
          & "', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', " & pVence & ", '" & txtNotas.Text _
          & "', '" & glogon.Usuario & "'"
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Modifica", "Expediente Restringido: " & txtCodigo.Text & " Cedula = " & txtCedula.Text)
End If

MsgBox "Datos Actualizados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""

If Len(txtNombre.Text) = 0 Then vMensaje = vMensaje & " - No se Indicó ninguna persona!"


If Len(vMensaje) = 0 Then
    fxVerifica = True
Else
    MsgBox vMensaje, vbExclamation
    fxVerifica = False
End If

End Function

Private Sub btnNuevo_Click()

txtCodigo.Text = ""

txtCedula.Text = ""
txtNombre.Text = ""
txtNotas.Text = ""

tcMain.Item(1).Selected = True

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"
cboEstado.Text = "Activo"

dtpVence.MinDate = fxFechaServidor
dtpVence.Value = dtpVence.MinDate

chkVence.Value = xtpUnchecked
Call chkVence_Click

'Cargar Aqui las causas
strSQL = "select rtrim(TIPO_ID) as 'IdX',  rtrim(descripcion) as  'ItmX' from SYS_EXP_TIPOS" _
       & " where Activo = 1 order by TIPO_ID"
Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

txtCedula.SetFocus

End Sub


Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Persona Id", 1800
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Identificación", 1800, vbCenter
    .Add , , "Nombre", 3800
    .Add , , "Tipo", 2800
    .Add , , "Inicio", 1200, vbCenter
    .Add , , "Corte", 1200, vbCenter
    .Add , , "Notas", 3800
End With


tcMain.Item(0).Selected = True

dtpConInicio.Value = fxFechaServidor
dtpConCorte.Value = dtpConInicio.Value

cboConEstado.Clear
cboConEstado.AddItem "Activo"
cboConEstado.AddItem "Inactivo"
cboConEstado.Text = "Activo"

Call chkConFechas_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError

tcMain.Item(1).Selected = True

Call sbCaso_Load(Item.Text)

vError:

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String


Select Case Item.Index
    Case 1
        Call btnNuevo_Click
    Case Else
End Select

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "cedula"
   gBusquedas.Orden = "cedula"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtCedula.Text = gBusquedas.Resultado
   txtNombre.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "nombre"
   gBusquedas.Orden = "nombre"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtCedula.Text = gBusquedas.Resultado
   txtNombre.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub sbCaso_Load(pPersonaId As Long)

On Error GoTo vError

tcMain.Item(1).Selected = True

strSQL = "select *, isnull(Fecha_Vence, '2300/01/01') as 'Fecha_Vence_Id' from vSYS_RA_Casos" _
       & " where Persona_Id = " & pPersonaId
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtCodigo.Text = CStr(rs!Persona_Id)
    
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    
    Call sbCboAsignaDato(cboTipo, rs!TipoDesc, True, rs!Tipo_Id)
 
    If rs!Estado = "A" Then
        cboEstado.Text = "Activo"
    Else
        cboEstado.Text = "Inactivo"
    End If
    
    If IsNull(rs!Fecha_Vence) Then
        chkVence.Value = xtpUnchecked
    Else
        chkVence.Value = xtpChecked
    End If
    
    dtpVence.Value = rs!Fecha_Vence_Id
    txtNotas.Text = rs!Notas
    
End If
rs.Close

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "Persona_Id"
   gBusquedas.Orden = "Persona_Id"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT Persona_Id, Cedula, NOMBRE, Estado" _
             & " FROM vSYS_RA_Casos"
   gBusquedas.Filtro = " and cedula like '" & txtCedula & "%'"
   frmBusquedas.Show vbModal
   
   txtCodigo.Text = gBusquedas.Resultado
   If IsNumeric(txtCodigo.Text) Then
      Call sbCaso_Load(txtCodigo.Text)
   End If
   
End If
vError:

End Sub

Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConNombre.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "CEDULA"
   gBusquedas.Orden = "CEDULA"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtConCedula = gBusquedas.Resultado
   txtConNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtConNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnBuscar.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "NOMBRE"
   gBusquedas.Orden = "NOMBRE"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtConCedula = gBusquedas.Resultado
   txtConNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnGuardar.SetFocus
vError:
End Sub






