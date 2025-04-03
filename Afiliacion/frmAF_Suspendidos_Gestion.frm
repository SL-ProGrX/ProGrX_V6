VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Suspendidos_Gestion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Estado Art. 24"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6252
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11292
      _Version        =   1441793
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
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "Consultas"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "txtConCedula"
      Item(0).Control(2)=   "txtConNombre"
      Item(0).Control(3)=   "Label1(26)"
      Item(0).Control(4)=   "Label1(0)"
      Item(0).Control(5)=   "btnAccion(0)"
      Item(0).Control(6)=   "btnAccion(1)"
      Item(0).Control(7)=   "chkFechas"
      Item(0).Control(8)=   "dtpInicio"
      Item(0).Control(9)=   "dtpCorte"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "txtNombre"
      Item(1).Control(1)=   "txtCedula"
      Item(1).Control(2)=   "Label1(1)"
      Item(1).Control(3)=   "txtNotas"
      Item(1).Control(4)=   "Label1(19)"
      Item(1).Control(5)=   "Label1(8)"
      Item(1).Control(6)=   "btnGuardar"
      Item(1).Control(7)=   "cboEvento"
      Item(1).Control(8)=   "cboTipo"
      Item(2).Caption =   "Mantenimiento"
      Item(2).ControlCount=   11
      Item(2).Control(0)=   "txtArchivo"
      Item(2).Control(1)=   "btnArchivo(0)"
      Item(2).Control(2)=   "btnArchivo(1)"
      Item(2).Control(3)=   "btnArchivo(2)"
      Item(2).Control(4)=   "Label1(2)"
      Item(2).Control(5)=   "Label2(7)"
      Item(2).Control(6)=   "btnAplicar"
      Item(2).Control(7)=   "lswCarga"
      Item(2).Control(8)=   "scMain"
      Item(2).Control(9)=   "cboM_Evento"
      Item(2).Control(10)=   "cboM_Tipo"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4695
         Left            =   -69880
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ListView lswCarga 
         Height          =   3615
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   6376
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
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   252
         Left            =   -62200
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441793
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
         ToolTipText     =   "Presione F4"
         Top             =   600
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
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
         Locked          =   -1  'True
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
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -64960
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -63640
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   -66400
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   4815
         _Version        =   1441793
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
         Left            =   -68080
         TabIndex        =   8
         ToolTipText     =   "Presione F4"
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   -62920
         TabIndex        =   9
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmAF_Suspendidos_Gestion.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1095
         Left            =   -68080
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
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
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   0
         Left            =   -61120
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "frmAF_Suspendidos_Gestion.frx":0731
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   1
         Left            =   -60640
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "frmAF_Suspendidos_Gestion.frx":0E31
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   435
         Left            =   2040
         TabIndex        =   19
         Top             =   600
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12086
         _ExtentY        =   762
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   0
         Left            =   9000
         TabIndex        =   20
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Suspendidos_Gestion.frx":0F9B
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   1
         Left            =   9480
         TabIndex        =   21
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Suspendidos_Gestion.frx":169B
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   2
         Left            =   9960
         TabIndex        =   22
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Suspendidos_Gestion.frx":1DB4
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   495
         Left            =   9480
         TabIndex        =   25
         Top             =   5640
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   868
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
         Picture         =   "frmAF_Suspendidos_Gestion.frx":24CD
      End
      Begin XtremeSuiteControls.ComboBox cboM_Evento 
         Height          =   330
         Left            =   2040
         TabIndex        =   28
         Top             =   1080
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboM_Tipo 
         Height          =   330
         Left            =   3960
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEvento 
         Height          =   330
         Left            =   -68080
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   -66160
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeShortcutBar.ShortcutCaption scMain 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Resultados de la busqueda:"
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Acción"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   23
         Top             =   600
         Width           =   1335
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
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas"
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
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   1212
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
         Left            =   -69760
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
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
         Left            =   -69160
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
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
         Left            =   -69160
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion de Estado Suspendido"
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
      TabIndex        =   16
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
Attribute VB_Name = "frmAF_Suspendidos_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbBuscar()

On Error GoTo vError

If chkFechas.Value = xtpChecked Then
    strSQL = "exec spPAT_AsociadosSinAportes_Bitacora '" & txtConCedula.Text & "', '1900-01-01', '2200-01-01'"
Else
    strSQL = "exec spPAT_AsociadosSinAportes_Bitacora '" & txtConCedula.Text & "', '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
           & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"

End If

Call OpenRecordSet(rs, strSQL)

With lsw.ListItems
    .Clear
    Do While Not rs.EOF
     Set itmX = .Add(, , rs!Id_Bitacora)
         itmX.SubItems(1) = rs!FECHA
         itmX.SubItems(2) = rs!Usuario & ""
         itmX.SubItems(3) = rs!Detalle & ""
         
     rs.MoveNext
    Loop
    rs.Close
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnAccion_Click(Index As Integer)

Select Case Index
    Case 0 'Buscar
        Call sbBuscar
    Case 1 'Exportar
        Call Excel_Exportar_Lsw(lsw)
End Select

End Sub

Private Sub sbLimpia()

lsw.ListItems.Clear

End Sub


Private Sub sbArchivo_Busca()

With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        txtArchivo.Text = .FileName

End With


End Sub


Private Sub sbArchivo_Carga()

Dim pCedula As String, pNombre As String, pMonto As Currency
Dim pLinea As Long

Dim strCadena As String, curMonto As Currency
Dim fn As Long, lCasos As Long
Dim i As Integer, vCampos As Boolean


On Error GoTo vError

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass


Set rs = Excel_Load(txtArchivo.Text, "IMPORT")
    
'Validaciónn del Archivo
vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "CEDULA" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos son Cedula, Nombre ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "NOMBRE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son Cedula, Nombre ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

'FIN: Validación del Archivo




'Sube, Revisa y Carga
With lswCarga.ListItems
    .Clear
    pLinea = 0
    strSQL = ""
    
    Do While Not rs.EOF
      If Trim(rs!Cedula) <> "" Then
        pCedula = rs!Cedula
        pNombre = rs!Nombre
        pLinea = pLinea + 1
        
        If pLinea = 1 Then
            strSQL = strSQL & Space(10) & "exec spPAT_AsociadosSinAportes_Carga_Masiva 'A','AFI-SUSP', '" & glogon.Usuario _
                   & "','" & pCedula & "', '" & pLinea & "','" & pNombre & "' , 1"
        Else
            strSQL = strSQL & Space(10) & "exec spPAT_AsociadosSinAportes_Carga_Masiva 'A','AFI-SUSP', '" & glogon.Usuario _
                   & "','" & pCedula & "', '" & pLinea & "', '" & pNombre & "' , 0"
        End If
        
        If Len(strSQL) > 20000 Then
           Call ConectionExecute(strSQL)
           If glogon.error Then
              Exit Sub
           End If
           strSQL = ""
        End If
        
      End If
      rs.MoveNext
    Loop
    rs.Close

'Procesa Ultimo Bloque

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   If glogon.error Then
      Exit Sub
   End If
   strSQL = ""
End If

'Revisa Lote y lo Carga
strSQL = "exec spPAT_AsociadosSinAportes_Carga_Masiva_Consulta 'A', 'AFI-SUSP', '" & glogon.Usuario _
       & "', " & cboM_Tipo.ItemData(cboM_Tipo.ListIndex)

Call OpenRecordSet(rs, strSQL)
If glogon.error Then
   Exit Sub
End If
    
    Do While Not rs.EOF
           Set itmX = .Add(, , rs!Cedula)
               itmX.SubItems(1) = rs!Nombre
            
      rs.MoveNext
    Loop
    rs.Close


End With 'Lista


Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia


End Sub

Private Sub btnAplicar_Click()


If lsw.ListItems.Count = 0 Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPAT_AsociadosSinAportes_Carga_Masiva_Procesa 'A', 'AFI-SUSP', '" & glogon.Usuario _
       & "', " & cboM_Tipo.ItemData(cboM_Tipo.ListIndex)

Call ConectionExecute(strSQL)

txtArchivo.Text = ""
Call sbLimpia

Me.MousePointer = vbDefault

MsgBox "Información Actualizada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia

End Sub

Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String

Select Case Index
  
  Case 0 'buscar
        txtArchivo.Text = ""
       Call sbArchivo_Busca
  
  Case 1 'cargar
       Call sbArchivo_Carga

  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: IMPORT" & vbCrLf _
              & " 3. Columnas.: CEDULA,  NOMBRE"
     
     MsgBox vMensaje, vbInformation

End Select
End Sub

Private Sub btnGuardar_Click()

On Error GoTo vError

strSQL = "exec spPAT_AsociadosSinAportes_Gestion '" & txtCedula.Text & "', " & cboTipo.ItemData(cboTipo.ListIndex) _
       & ", '" & Mid(txtNotas.Text, 1, 500) & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Se ha cambiado la condición a: " & cboTipo.Text & " a " & txtNombre.Text & ", satisfactoriamente!", vbInformation

txtCedula.Text = ""
txtNombre.Text = ""

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboEvento_Click()

cboTipo.Clear
If cboEvento.Text = "Activar" Then
    cboTipo.AddItem "Activar"
    cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(1)
    cboTipo.Text = "Activar"

Else
    cboTipo.AddItem "Suspender"
    cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(0)
    cboTipo.AddItem "Condición Especial"
    cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(2)
    cboTipo.Text = "Suspender"
End If

End Sub

Private Sub cboM_Evento_Click()
cboM_Tipo.Clear
If cboM_Evento.Text = "Activar" Then
    cboM_Tipo.AddItem "Activar"
    cboM_Tipo.ItemData(cboM_Tipo.ListCount - 1) = CStr(1)
    cboM_Tipo.Text = "Activar"

Else
    cboM_Tipo.AddItem "Suspender"
    cboM_Tipo.ItemData(cboM_Tipo.ListCount - 1) = CStr(0)
    cboM_Tipo.AddItem "Condición Especial"
    cboM_Tipo.ItemData(cboM_Tipo.ListCount - 1) = CStr(2)
    cboM_Tipo.Text = "Suspender"
End If

End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = xtpChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled
End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id Gestión", 1800
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Detalle", 3100
End With



With lswCarga.ColumnHeaders
    .Clear
    .Add , , "Cédula", 1800, vbCenter
    .Add , , "Nombre", 4100
End With

tcMain.Item(0).Selected = True

cboEvento.AddItem "Activar"
cboEvento.AddItem "Inactivar"
cboEvento.Text = "Activar"

Call cboEvento_Click

cboM_Evento.AddItem "Activar"
cboM_Evento.AddItem "Inactivar"
cboM_Evento.Text = "Activar"

Call cboM_Evento_Click


dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -60, dtpCorte.Value)

Call chkFechas_Click

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterna"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "cedula"
    gBusquedas.Orden = "cedula"
    gBusquedas.Filtro = " and estadoactual = 'S'"
    frmBusquedas.Show vbModal
    
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado3
End If

End Sub


Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterna"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "cedula"
    gBusquedas.Orden = "cedula"
    gBusquedas.Filtro = " and estadoactual = 'S'"
    frmBusquedas.Show vbModal
    
    txtConCedula.Text = gBusquedas.Resultado
    txtConNombre.Text = gBusquedas.Resultado3
End If

End Sub
