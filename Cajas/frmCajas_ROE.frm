VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_ROE 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ROE"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox gbInfo 
      Height          =   5415
      Left            =   2880
      TabIndex        =   72
      Top             =   1320
      Visible         =   0   'False
      Width           =   6975
      _Version        =   1572864
      _ExtentX        =   12303
      _ExtentY        =   9551
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Begin XtremeSuiteControls.FlatEdit txtI_Fecha 
         Height          =   330
         Left            =   3960
         TabIndex        =   85
         Top             =   1680
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtN_Fecha 
         Height          =   330
         Left            =   3960
         TabIndex        =   84
         Top             =   2280
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtR_Usuario 
         Height          =   330
         Left            =   1800
         TabIndex        =   75
         Top             =   840
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtA_Usuario 
         Height          =   330
         Left            =   1800
         TabIndex        =   74
         Top             =   1200
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtA_Fecha 
         Height          =   330
         Left            =   3960
         TabIndex        =   78
         Top             =   1200
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtR_Fecha 
         Height          =   330
         Left            =   3960
         TabIndex        =   79
         Top             =   840
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtN_Usuario 
         Height          =   330
         Left            =   1800
         TabIndex        =   80
         Top             =   2280
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtI_Usuario 
         Height          =   330
         Left            =   1800
         TabIndex        =   81
         Top             =   1680
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.PushButton btnControlCierra 
         Height          =   495
         Left            =   6000
         TabIndex        =   86
         ToolTipText     =   "Cerrar Información"
         Top             =   4800
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   79
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
         Appearance      =   17
         Picture         =   "frmCajas_ROE.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtN_Nota 
         Height          =   2010
         Left            =   1800
         TabIndex        =   87
         Top             =   2640
         Width           =   4815
         _Version        =   1572864
         _ExtentX        =   8493
         _ExtentY        =   3545
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSesionId 
         Height          =   330
         Left            =   4560
         TabIndex        =   90
         Top             =   480
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   34
         Left            =   3600
         TabIndex        =   89
         Top             =   480
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sesión Id:"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Index           =   33
         Left            =   120
         TabIndex        =   88
         Top             =   2640
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Anotaciones de la Anulación:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   83
         Top             =   1680
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Impreso por:"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   82
         Top             =   2280
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Anulado por:"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registrado por:"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   76
         Top             =   1200
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Actualizado por:"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   6975
         _Version        =   1572864
         _ExtentX        =   12303
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Información de Control"
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   12375
      _Version        =   1572864
      _ExtentX        =   21828
      _ExtentY        =   5530
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtA_Nombre 
         Height          =   330
         Left            =   2880
         TabIndex        =   12
         Top             =   720
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
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
      Begin XtremeSuiteControls.FlatEdit txtA_Cedula 
         Height          =   330
         Left            =   6600
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtA_Provincia 
         Height          =   330
         Left            =   2880
         TabIndex        =   14
         Top             =   1920
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtA_Canton 
         Height          =   330
         Left            =   2880
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtA_Distrito 
         Height          =   330
         Left            =   2880
         TabIndex        =   16
         Top             =   2640
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtA_Direccion 
         Height          =   1050
         Left            =   4800
         TabIndex        =   17
         Top             =   1920
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   1852
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtA_Telefono 
         Height          =   330
         Left            =   10320
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.DateTimePicker dtpA_FechaNac 
         Height          =   330
         Left            =   6600
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   330
         Left            =   2880
         TabIndex        =   20
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtA_Pais 
         Height          =   330
         Left            =   2880
         TabIndex        =   64
         Top             =   1560
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeShortcutBar.ShortcutCaption scSesion 
         Height          =   375
         Left            =   9840
         TabIndex        =   70
         Top             =   120
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Sesión Id: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   63
         Top             =   1560
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "País"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Distrito"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cantón"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Provincia"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   8760
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Teléfono"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Nacimiento"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Identificación"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Identificación"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre Completo / Razón Social"
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
      Begin XtremeShortcutBar.ShortcutCaption scParte 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   9855
         _Version        =   1572864
         _ExtentX        =   17383
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Parte A: Datos del Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtROE_Id 
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnImprimir 
      Height          =   495
      Left            =   10320
      TabIndex        =   21
      ToolTipText     =   "Imprimir el ROE"
      Top             =   120
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   79
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
      Appearance      =   17
      Picture         =   "frmCajas_ROE.frx":063E
   End
   Begin XtremeSuiteControls.PushButton btnRefresh 
      Height          =   495
      Left            =   10920
      TabIndex        =   22
      ToolTipText     =   "Refrescar Información"
      Top             =   120
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   79
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
      Appearance      =   17
      Picture         =   "frmCajas_ROE.frx":0D45
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   3135
      Left            =   0
      TabIndex        =   23
      Top             =   3720
      Width           =   12375
      _Version        =   1572864
      _ExtentX        =   21828
      _ExtentY        =   5530
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtD_Nombre 
         Height          =   330
         Left            =   2880
         TabIndex        =   24
         Top             =   720
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
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
      Begin XtremeSuiteControls.FlatEdit txtD_Cedula 
         Height          =   330
         Left            =   6600
         TabIndex        =   25
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtD_Direccion 
         Height          =   1050
         Left            =   4800
         TabIndex        =   26
         Top             =   1920
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   1852
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
      Begin XtremeSuiteControls.FlatEdit txtD_Telefono 
         Height          =   330
         Left            =   10320
         TabIndex        =   27
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.DateTimePicker dtpD_FechaNac 
         Height          =   330
         Left            =   6600
         TabIndex        =   28
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboTipoId_D 
         Height          =   330
         Left            =   2880
         TabIndex        =   29
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboPais 
         Height          =   330
         Left            =   2880
         TabIndex        =   66
         Top             =   1560
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   330
         Left            =   2880
         TabIndex        =   67
         Top             =   1920
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCanton 
         Height          =   330
         Left            =   2880
         TabIndex        =   68
         Top             =   2280
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDistrito 
         Height          =   330
         Left            =   2880
         TabIndex        =   69
         Top             =   2640
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "País"
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
         Left            =   0
         TabIndex        =   38
         Top             =   120
         Width           =   12375
         _Version        =   1572864
         _ExtentX        =   21828
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Parte B: Datos de la Persona Física que realiza la Transacción"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre Completo / Razón Social"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Identificación"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   14
         Left            =   4920
         TabIndex        =   35
         Top             =   1080
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Identificación"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   13
         Left            =   4920
         TabIndex        =   34
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Nacimiento"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   12
         Left            =   8760
         TabIndex        =   33
         Top             =   1440
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Teléfono"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Provincia"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cantón"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Distrito"
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
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   3375
      Left            =   0
      TabIndex        =   39
      Top             =   6720
      Width           =   12375
      _Version        =   1572864
      _ExtentX        =   21828
      _ExtentY        =   5953
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtC_OrigenRecursos 
         Height          =   330
         Left            =   2880
         TabIndex        =   40
         Top             =   1800
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_BeneficiarioDatos 
         Height          =   330
         Left            =   2880
         TabIndex        =   42
         Top             =   2760
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
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
      Begin XtremeSuiteControls.FlatEdit txtC_Observaciones 
         Height          =   570
         Left            =   2880
         TabIndex        =   43
         Top             =   2160
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   1005
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
      Begin XtremeSuiteControls.ComboBox cboTipoTransaccion 
         Height          =   330
         Left            =   2880
         TabIndex        =   45
         Top             =   600
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipoOperacion 
         Height          =   330
         Left            =   6600
         TabIndex        =   52
         Top             =   600
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Comprobantes 
         Height          =   1170
         Left            =   10320
         TabIndex        =   44
         Top             =   600
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   2064
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_MontoDol 
         Height          =   330
         Left            =   2880
         TabIndex        =   41
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_MontoLocal 
         Height          =   330
         Left            =   2880
         TabIndex        =   56
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Hora 
         Height          =   330
         Left            =   6600
         TabIndex        =   58
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Fecha 
         Height          =   330
         Left            =   6600
         TabIndex        =   59
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Origen de los Recursos"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   26
         Left            =   4920
         TabIndex        =   57
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Hora"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto en Dólares"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   23
         Left            =   8760
         TabIndex        =   54
         Top             =   600
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Comprobantes"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   17
         Left            =   4920
         TabIndex        =   53
         Top             =   600
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Operación"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   51
         Top             =   2760
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Datos del Beneficiario"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   50
         Top             =   2280
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaciones"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   20
         Left            =   4920
         TabIndex        =   49
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto Divisa Local"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Transaccion"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   0
         TabIndex        =   46
         Top             =   120
         Width           =   12375
         _Version        =   1572864
         _ExtentX        =   21828
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Parte C: Descripción de la Transacción"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   495
      Left            =   7080
      TabIndex        =   61
      Top             =   120
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.PushButton btnActualizar 
      Height          =   495
      Left            =   9480
      TabIndex        =   62
      ToolTipText     =   "Actualizar ROE"
      Top             =   120
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   79
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
      Appearance      =   17
      Picture         =   "frmCajas_ROE.frx":1445
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   495
      Left            =   11520
      TabIndex        =   71
      ToolTipText     =   "Información de Control"
      Top             =   120
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   79
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
      Appearance      =   17
      Picture         =   "frmCajas_ROE.frx":1B76
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Formulario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
Attribute VB_Name = "frmCajas_ROE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub sbROE_Load()

On Error GoTo vError

If Not IsNumeric(txtROE_Id.Text) Then Exit Sub

strSQL = "select * from vCajas_ROE where ID_ROE = " & txtROE_Id.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  scSesion.Caption = "Sesión Id: " & rs!Id_Sesion & ""

  txtA_Cedula.Text = Trim(rs!Cedula)
  txtA_Nombre.Text = rs!Nombre
  txtA_Pais.Text = rs!ASO_PaisDesc
  txtA_Provincia.Text = rs!ASO_ProvinciaDesc
  txtA_Canton.Text = rs!ASO_CantonDesc
  txtA_Distrito.Text = rs!ASO_DistritoDesc
  txtA_Direccion.Text = rs!ASO_Direccion

 Call sbCboAsignaDato(cboTipoId, rs!ASO_TipoIdDesc, True, rs!ASO_Tipo_Id)
  
  dtpA_FechaNac.Value = rs!ASO_Fecha_Nac
  txtA_Telefono.Text = rs!ASO_Telefono
  
 '---Seccion B: Depositante
  txtD_Cedula.Text = Trim(rs!Identificacion_Depo)
  txtD_Nombre.Text = rs!Nombre_Depo
  
'  txtD_Provincia.Text = rs!Provincia & ""
'  txtD_Canton.Text = rs!Canton & ""
'  txtD_Distrito.Text = rs!Distrito & ""

  txtD_Direccion.Text = rs!Dir_Referencia1 & ""

 vPaso = True
        Call sbCboAsignaDato(cboTipoId_D, rs!DEP_TipoIdDesc, True, rs!Tipo_Id & "")
        Call sbCboAsignaDato(cboPais, rs!DEP_PaisDesc, True, rs!cod_Pais & "")
        Call sbCboAsignaDato(cboProvincia, rs!DEP_ProvinciaDesc, True, rs!cod_Provincia & "")
        Call sbCboAsignaDato(cboCanton, rs!DEP_CantonDesc, True, rs!cod_Canton & "")
        Call sbCboAsignaDato(cboDistrito, rs!DEP_DistritoDesc, True, rs!Cod_Distrito & "")
 vPaso = False
  
  If Not IsNull(rs!Fecha_Nac_Const_Empr) Then
      dtpD_FechaNac.Value = rs!Fecha_Nac_Const_Empr
  Else
      dtpD_FechaNac.Value = Now
  End If
  
  txtD_Telefono.Text = rs!Telefono_Depo & ""
  
  '---Seccion C
  
  txtC_BeneficiarioDatos.Text = rs!Datos_Beneficiario & ""
  txtC_Comprobantes.Text = rs!Num_Doc & ""
  txtC_Fecha.Text = rs!fecha & ""
  txtC_Hora.Text = rs!hora & ""
  txtC_Observaciones.Text = rs!Observacion & ""
  txtC_MontoLocal.Text = Format(rs!Monto_Local, "Standard")
  txtC_MontoDol.Text = Format(rs!Monto_Dol, "Standard")
     
  txtC_OrigenRecursos.Text = rs!Origen_Fondos & ""
  
  Call sbCboAsignaDato(cboTipoTransaccion, rs!TIPO_TRANS & "", False)
  Call sbCboAsignaDato(cboTipoOperacion, rs!TIPO_OPERACION & "", False)
  
  '--Estado
  If rs!Estado = "A" Then
     txtEstado.Text = "Activo"
  Else
     txtEstado.Text = "Inactivo"
  End If
  
  
  txtSesionId.Text = rs!Id_Sesion & ""
  txtR_Usuario.Text = rs!Registro_Usuario & ""
  txtR_Fecha.Text = rs!Registro_Fecha & ""
  
  txtA_Usuario.Text = rs!actualiza_Usuario & ""
  txtA_Fecha.Text = rs!actualiza_fecha & ""
  
  txtI_Usuario.Text = rs!Imprime_Usuario & ""
  txtI_Fecha.Text = rs!Imprime_fecha & ""
  
  txtN_Usuario.Text = rs!USUARIO_ANULACION & ""
  txtN_Fecha.Text = rs!FECHA_ANULACION & ""
  
  txtN_Nota.Text = rs!OBSERV_ANULACION & ""
  
End If
rs.Close

Exit Sub

vError:

End Sub


Private Sub btnActualizar_Click()

On Error GoTo vError

Dim vMensaje As String

vMensaje = ""

If cboTipoId_D.Text = "" Then
    vMensaje = vMensaje & " - No ha indicado el Tipo de Identificación" & vbCrLf
End If
If cboPais.Text = "" Then
    vMensaje = vMensaje & " - No ha indicado el País de Procedencia" & vbCrLf
End If
If cboProvincia.Text = "" Then
    vMensaje = vMensaje & " - No ha indicado la Provincia" & vbCrLf
End If
If cboCanton.Text = "" Then
    vMensaje = vMensaje & " - No ha indicado el Cantón" & vbCrLf
End If
If cboDistrito.Text = "" Then
    vMensaje = vMensaje & " - No ha indicado el distrito" & vbCrLf
End If

If Len(txtD_Direccion.Text) <= 10 Then
    vMensaje = vMensaje & " - No ha indicado la Dirección Exacta" & vbCrLf
End If

If Len(txtD_Telefono.Text) < 8 Then
    vMensaje = vMensaje & " - No ha indicado un número de Teléfono válido" & vbCrLf
End If

If Len(txtC_OrigenRecursos.Text) <= 10 Then
    vMensaje = vMensaje & " - No ha indicado el Origen de los Recursos" & vbCrLf
End If

If Len(txtC_BeneficiarioDatos.Text) <= 10 Then
    vMensaje = vMensaje & " - No ha indicado los datos del Beneficiario" & vbCrLf
End If

If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation
   Exit Sub
End If


Dim pPais As String, pProvincia As String, pCanton As String, pDistrito As String, pTipoId As String, pPaisDesc As String

pTipoId = cboTipoId_D.ItemData(cboTipoId_D.ListIndex)
pPais = "'" & cboPais.ItemData(cboPais.ListIndex) & "'"
pPaisDesc = "'" & cboPais.Text & "'"
pProvincia = "'" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"
pCanton = "'" & cboCanton.ItemData(cboCanton.ListIndex) & "'"
pDistrito = "'" & cboDistrito.ItemData(cboDistrito.ListIndex) & "'"

strSQL = "exec spCajas_ROE_Actualiza " & txtROE_Id.Text & ", '" & cboTipoId_D.Text & "', '" & cboProvincia.Text _
       & "', '" & cboCanton.Text & "', '" & cboDistrito.Text & "', '" & txtD_Direccion.Text _
       & "', '" & txtD_Telefono.Text & "', '" & Format(dtpD_FechaNac.Value, "yyyy-mm-dd") & "', '" & cboTipoTransaccion.Text _
       & "', '" & cboTipoOperacion.Text & "', '" & txtC_OrigenRecursos.Text & "', '" & txtC_Observaciones.Text _
       & "', '" & txtC_BeneficiarioDatos.Text & "', '" & glogon.Usuario & "', " & pTipoId _
       & ", " & pPais & ", " & pPaisDesc & ", " & pProvincia & ", " & pCanton & ", " & pDistrito
Call OpenRecordSet(rs, strSQL)
If rs!Pass = 1 Then

        Call Bitacora("Modifica", "ROE Id: " & txtROE_Id.Text)
        MsgBox "ROE Actualizado!", vbInformation
Else
        MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnControlCierra_Click()
gbInfo.Visible = False
End Sub

Private Sub btnImprimir_Click()

On Error GoTo vError

strSQL = "select dbo.fxCajas_ROE_Imprime_Valida(" & txtROE_Id.Text & ") as 'Imprime'"
Call OpenRecordSet(rs, strSQL)

If rs!Imprime = 0 Then
   MsgBox "Este Documento no ha sido actualizado/revisado. Verifique los datos, guarde y luego imprima!", vbExclamation
   Exit Sub
End If
rs.Close

Call sbCajas_ROE(txtROE_Id.Text)

'Refresca Info
Call sbROE_Load

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub btnInfo_Click()
gbInfo.Visible = True
End Sub

Private Sub btnRefresh_Click()
Call sbROE_Load
End Sub

Private Sub cboCanton_Click()

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

If cboDistrito.ListCount = 0 Then
    cboDistrito.AddItem " "
    cboDistrito.Text = " "
End If

End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtD_Direccion.SetFocus
End Sub


Private Sub cboPais_Click()
If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias" _
           & " where cod_Pais = '" & cboPais.ItemData(cboPais.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False


If cboProvincia.ListCount = 0 Then
    cboProvincia.AddItem " "
    cboProvincia.Text = " "
    
    cboCanton.Clear
    cboCanton.AddItem " "
    cboCanton.Text = " "
    
    cboDistrito.Clear
    cboDistrito.AddItem " "
    cboDistrito.Text = " "
Else
    Call cboProvincia_Click
End If

End Sub

Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False


If cboCanton.ListCount = 0 Then
    cboCanton.AddItem " "
    cboCanton.Text = " "
Else
    Call cboCanton_Click
End If

End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub Form_Load()

vModulo = 5

strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)

'Depositante
Call sbCbo_Copia(cboTipoId, cboTipoId_D)

cboTipoTransaccion.AddItem "Ingreso"
cboTipoTransaccion.Text = "Ingreso"

cboTipoOperacion.AddItem "Deposito"
cboTipoOperacion.AddItem "Inversión o Cancelación"
cboTipoOperacion.AddItem "Prestamos"
cboTipoOperacion.AddItem "Otros"
cboTipoOperacion.Text = "Deposito"


vPaso = True
strSQL = "select cod_Pais as 'IdX', Descripcion as 'ItmX' from Paises" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboPais, strSQL, False, True)

vPaso = False



If Mid(GLOBALES.gTag, 1, 4) = "ROE_" Then
    txtROE_Id.Text = Mid(GLOBALES.gTag, 5, 20)
    Call sbROE_Load

Else
    Call cboPais_Click
End If

End Sub


