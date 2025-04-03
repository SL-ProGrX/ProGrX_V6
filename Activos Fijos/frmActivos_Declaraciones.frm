VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmActivos_Declaraciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Declaraciones de Activos Fijos"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   9975
      _Version        =   1572864
      _ExtentX        =   17595
      _ExtentY        =   7858
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
      Item(0).Caption =   "Declaración"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "cbo"
      Item(0).Control(1)=   "dtpInicio"
      Item(0).Control(2)=   "Label6(2)"
      Item(0).Control(3)=   "Label6(1)"
      Item(0).Control(4)=   "Label6(0)"
      Item(0).Control(5)=   "dtpCorte"
      Item(0).Control(6)=   "txtA_Fecha"
      Item(0).Control(7)=   "txtC_Fecha"
      Item(0).Control(8)=   "Label1(0)"
      Item(0).Control(9)=   "Label1(4)"
      Item(0).Control(10)=   "txtA_Usuario"
      Item(0).Control(11)=   "txtC_Usuario"
      Item(0).Control(12)=   "Label1(1)"
      Item(0).Control(13)=   "Label1(2)"
      Item(0).Control(14)=   "txtP_Fecha"
      Item(0).Control(15)=   "Label1(3)"
      Item(0).Control(16)=   "txtP_Usuario"
      Item(0).Control(17)=   "txtNotas"
      Item(1).Caption =   "Historial"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "lbl"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3495
         Left            =   -70000
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
         _ExtentY        =   6165
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   915
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   7695
         _Version        =   1572864
         _ExtentX        =   13573
         _ExtentY        =   1614
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   330
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         Top             =   2160
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   315
         Left            =   3480
         TabIndex        =   22
         Top             =   2160
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtA_Fecha 
         Height          =   315
         Left            =   1920
         TabIndex        =   23
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3000
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Fecha 
         Height          =   315
         Left            =   1920
         TabIndex        =   24
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3360
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtA_Usuario 
         Height          =   315
         Left            =   4200
         TabIndex        =   27
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3000
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Usuario 
         Height          =   315
         Left            =   4200
         TabIndex        =   28
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3360
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtP_Fecha 
         Height          =   315
         Left            =   1920
         TabIndex        =   31
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3720
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtP_Usuario 
         Height          =   315
         Left            =   4200
         TabIndex        =   33
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3720
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesa"
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
         Index           =   3
         Left            =   720
         TabIndex        =   32
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Index           =   2
         Left            =   1920
         TabIndex        =   30
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   29
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Apertura"
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
         Index           =   4
         Left            =   720
         TabIndex        =   26
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cierre"
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
         Index           =   0
         Left            =   720
         TabIndex        =   25
         Top             =   3360
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo de Registro"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fechas"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption lbl 
         Height          =   375
         Left            =   -70000
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Listado de Declaraciones Registradas"
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
         VisualTheme     =   3
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Nuevo"
      Top             =   120
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Declaraciones.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Editar"
      Top             =   120
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Declaraciones.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Eliminar"
      Top             =   120
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Declaraciones.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   "Guardar"
      Top             =   120
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Declaraciones.frx":11D1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Deshacer"
      Top             =   120
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Declaraciones.frx":1902
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   3120
      TabIndex        =   5
      ToolTipText     =   "Reporte"
      Top             =   120
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Declaraciones.frx":2002
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnProcesar 
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Procesar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmActivos_Declaraciones.frx":2709
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cerrar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmActivos_Declaraciones.frx":2E22
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   2040
      TabIndex        =   9
      Top             =   720
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   767
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   435
      Left            =   7320
      TabIndex        =   11
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   720
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   375
      Index           =   7
      Left            =   6000
      TabIndex        =   12
      Top             =   720
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estado"
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
   Begin XtremeSuiteControls.Label Label5 
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   10
      Top             =   720
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Declaración"
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
End
Attribute VB_Name = "frmActivos_Declaraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vEdita As Boolean, vCodigo As Long, vPaso As Boolean

Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)

Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpiaPantalla
        txtCodigo.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
      If txtCodigo.Text = "0" Then
        MsgBox "Seleccione una Declaración para modificacion...", vbInformation
      Else
        vEdita = True
        txtCodigo.SetFocus
        Call sbBarra_Accion("Editar")
      End If
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      End If
    
    Case 5 'REPORTES
   
End Select

End Sub


Private Sub btnCerrar_Click()
On Error GoTo vError

If txtCodigo.Text = "" Then
    MsgBox "No se ha indicado ninguna Declaración de Activos?", vbExclamation
    Exit Sub
End If

If Mid(txtEstado.Text, 1, 1) = "C" Then
    MsgBox "La Declaración ya ha sido cerrada, verifique!", vbExclamation
    Exit Sub
End If

If Mid(txtEstado.Text, 1, 1) <> "A" Then
    MsgBox "La Declaración tiene que estar en Estado de Abierta!", vbExclamation
    Exit Sub
End If

Dim i As Integer

i = MsgBox("Esta seguro que desea Cerrar esta Declaración de Activos?", vbYesNo)
If i = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spActivos_Declara_Main_Cierra " & txtCodigo.Text & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Declaración de Activos: Cerrada Satisfactoriamente!", vbInformation
   Call sbConsulta(txtCodigo.Text)
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnProcesar_Click()
On Error GoTo vError

If txtCodigo.Text = "" Then
    MsgBox "No se ha indicado ninguna Declaración de Activos?", vbExclamation
    Exit Sub
End If

If Mid(txtEstado.Text, 1, 1) = "P" Then
    MsgBox "La Declaración ya ha sido procesada, verifique!", vbExclamation
    Exit Sub
End If

If Mid(txtEstado.Text, 1, 1) <> "C" Then
    MsgBox "La Declaración tiene que Cerrarse antes de procesar!", vbExclamation
    Exit Sub
End If

Dim i As Integer

i = MsgBox("Esta seguro que desea procesar esta Declaración de Activos?", vbYesNo)
If i = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spActivos_Declara_Main_Procesa " & txtCodigo.Text & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Declaración de Activos: Procesada Satisfactoriamente!", vbInformation
   Call sbConsulta(txtCodigo.Text)
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 36
 
 vEdita = True


cbo.Clear
cbo.AddItem "Visible"
cbo.AddItem "Ciego"
cbo.Text = "Visible"


 With lsw.ColumnHeaders
        .Add , , "Declara Id", 1200
        .Add , , "Tipo", 1600, vbCenter
        .Add , , "Estado", 1600, vbCenter
        .Add , , "Inicio", 1800, vbCenter
        .Add , , "Corte", 1800, vbCenter
        .Add , , "Notas", 3600
        .Add , , "R.Fecha", 1800
        .Add , , "R.Usuario", 1800
        .Add , , "C.Fecha", 1800
        .Add , , "C.Usuario", 1800
        .Add , , "P.Fecha", 1800
        .Add , , "P.Usuario", 1800
 End With
 
 Call sbLimpiaPantalla
 Call sbBarra_Accion("Nuevo")

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True

vCodigo = 0
txtCodigo = ""

txtEstado.Text = ""
cbo.Text = "Visible"

txtNotas.Text = ""
dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

txtA_Fecha.Text = ""
txtA_Usuario.Text = ""

txtC_Fecha.Text = ""
txtC_Usuario.Text = ""

txtP_Fecha.Text = ""
txtP_Usuario.Text = ""

End Sub



Private Sub sbConsulta(xCodigo As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vActivos_Declara where ID_DECLARA = " & xCodigo
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbBarra_Accion("Activo")
  vEdita = True
  
  tcMain.Item(0).Selected = True
  
  vCodigo = rs!ID_DECLARA
  txtCodigo.Text = rs!ID_DECLARA
      
  txtEstado.Text = rs!Estado_Desc
  cbo.Text = rs!Tipo_Desc
  
  txtNotas.Text = rs!Notas
  
  dtpInicio.Value = rs!fecha_Inicio
  dtpCorte.Value = rs!fecha_Corte
  
    txtA_Fecha.Text = rs!Registro_fecha
    txtA_Usuario.Text = rs!Registro_Usuario
    
    txtC_Fecha.Text = rs!Cerrado_fecha & ""
    txtC_Usuario.Text = rs!Cerrado_Usuario & ""
    
    txtP_Fecha.Text = rs!Procesado_fecha & ""
    txtP_Usuario.Text = rs!Procesado_Usuario & ""
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

strSQL = "select estado, dbo.fxActivos_PeriodoActual() as 'PeriodoActual'" _
       & " from Activos_periodos where anio = " & Year(dtpInicio.Value) _
       & " and mes = " & Month(dtpInicio.Value)
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
' If rs!Estado <> "P" Then
'      vMensaje = vMensaje & vbCrLf & " - El Periodo del Movimiento ya fue cerrado ..."
' End If
 
 If Year(dtpCorte.Value) <> Year(rs!PeriodoActual) Or Month(dtpCorte.Value) <> Month(rs!PeriodoActual) Then
      vMensaje = vMensaje & vbCrLf & " - La fecha de aplicación del movimiento no corresponde al periodo abierto!"
 End If

End If
rs.Close

If dtpInicio.Value >= dtpCorte.Value Then vMensaje = vMensaje & vbCrLf & " - Verifique el rango de Fechas..."
If txtNotas.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Indique una Nota válida para esta declaración..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbGuardar()
Dim vMovimiento As String

Me.MousePointer = vbHourglass

On Error GoTo vError

If vCodigo = 0 Then
   vMovimiento = "Registra"
Else
   vMovimiento = "Modifica"
End If

strSQL = "exec spActivos_Declara_Main_Add " & vCodigo & ", '" & txtNotas.Text & "', '" & Mid(cbo.Text, 1, 1) _
        & "', '" & Format(dtpInicio.Value, "yyyy/mm/dd") & "', '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
        & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Pass = 0 Then
    Me.MousePointer = vbDefault
    MsgBox rs!Mensaje, vbExclamation
    Exit Sub
End If

Me.MousePointer = vbDefault

vCodigo = rs!ID_DECLARA
 

strSQL = "Declaración de Activo Id: " & vCodigo & ", Inicio: " & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " - " & Format(dtpCorte.Value, "yyyy/mm/dd") & ", Tipo: " & cbo.Text
Call Bitacora("Registra", strSQL)
  
MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Call sbBarra_Accion("Activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

If vCodigo = 0 Then Exit Sub

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "exec spActivos_Declara_Main_Delete " & vCodigo & ", '" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Pass = 1 Then
        Call Bitacora("Elimina", "Declaración de Activo Id: " & vCodigo)
    Else
        MsgBox rs!Mensaje, vbExclamation
    End If
  
  Call sbLimpiaPantalla
  Call sbBarra_Accion("Nuevo")
  
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 Call sbConsulta(Item.Text)
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Select Case Item.Index
 Case 0 'Nada
 Case 1 'Lista
    
    strSQL = "select Top 200 * from vActivos_Declara order by ID_Declara desc"
    Call OpenRecordSet(rs, strSQL, 0)
    lsw.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!ID_DECLARA)
          itmX.SubItems(1) = rs!Tipo_Desc
          itmX.SubItems(2) = rs!Estado_Desc
          itmX.SubItems(3) = Format(rs!fecha_Inicio, "yyyy-mm-dd")
          itmX.SubItems(4) = Format(rs!fecha_Corte, "yyyy-mm-dd")
          itmX.SubItems(5) = rs!Notas
          itmX.SubItems(6) = rs!Registro_fecha
          itmX.SubItems(7) = rs!Registro_Usuario
          
          itmX.SubItems(8) = rs!Cerrado_fecha & ""
          itmX.SubItems(9) = rs!Cerrado_Usuario & ""
          
          itmX.SubItems(10) = rs!Procesado_fecha & ""
          itmX.SubItems(11) = rs!Procesado_Usuario & ""
      rs.MoveNext
    Loop
    rs.Close


End Select

vError:


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  tcMain.Item(0).Selected = True
  cbo.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Id_Declara"
  gBusquedas.Orden = "Id_Declara"
  
  gBusquedas.Col1Name = "Id Declara"
  gBusquedas.Col2Name = "Inicio"
  gBusquedas.Col3Name = "Corte"
  
  gBusquedas.Consulta = "select Id_Declara, Fecha_Inicio, Fecha_Corte, Estado_Desc, Tipo_Desc  from vActivos_Declara"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  If IsNumeric(txtCodigo.Text) Then
    Call sbConsulta(txtCodigo.Text)
  End If
End If

End Sub

Private Sub txtCodigo_LostFocus()
  If IsNumeric(txtCodigo.Text) Then
    Call sbConsulta(txtCodigo.Text)
  End If
End Sub



