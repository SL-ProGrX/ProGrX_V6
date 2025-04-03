VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Aguinaldos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cálculo de Aguinaldos"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   16965
   LinkTopic       =   "Form8"
   ScaleHeight     =   9180
   ScaleWidth      =   16965
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lswBancos 
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   5318
      _StockProps     =   77
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
      View            =   3
      FullRowSelect   =   -1  'True
      BackColor       =   16777215
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton rbVisualizar 
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   71
      Top             =   1200
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Listado de Aguinaldos"
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
   Begin XtremeSuiteControls.GroupBox gbAjustes 
      Height          =   6615
      Left            =   4680
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   9735
      _Version        =   1441793
      _ExtentX        =   17171
      _ExtentY        =   11668
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtAjuste_Actualiza_Fecha 
         Height          =   330
         Left            =   7560
         TabIndex        =   84
         Top             =   2640
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   0
         Left            =   5760
         TabIndex        =   59
         Top             =   1200
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnAjustesAplica 
         Height          =   495
         Left            =   4080
         TabIndex        =   57
         Top             =   5880
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmRH_Aguinaldos.frx":08D1
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   0
         Left            =   1800
         TabIndex        =   20
         Top             =   1200
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   0
         Left            =   3720
         TabIndex        =   21
         Top             =   1200
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   1
         Left            =   1800
         TabIndex        =   23
         Top             =   1560
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   1
         Left            =   3720
         TabIndex        =   24
         Top             =   1560
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   26
         Top             =   1920
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   2
         Left            =   3720
         TabIndex        =   27
         Top             =   1920
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   3
         Left            =   1800
         TabIndex        =   29
         Top             =   2280
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   3
         Left            =   3720
         TabIndex        =   30
         Top             =   2280
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   4
         Left            =   1800
         TabIndex        =   32
         Top             =   2640
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   4
         Left            =   3720
         TabIndex        =   33
         Top             =   2640
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   5
         Left            =   1800
         TabIndex        =   35
         Top             =   3000
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   5
         Left            =   3720
         TabIndex        =   36
         Top             =   3000
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   6
         Left            =   1800
         TabIndex        =   38
         Top             =   3360
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   6
         Left            =   3720
         TabIndex        =   39
         Top             =   3360
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   7
         Left            =   1800
         TabIndex        =   41
         Top             =   3720
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   7
         Left            =   3720
         TabIndex        =   42
         Top             =   3720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   8
         Left            =   1800
         TabIndex        =   44
         Top             =   4080
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   8
         Left            =   3720
         TabIndex        =   45
         Top             =   4080
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   9
         Left            =   1800
         TabIndex        =   47
         Top             =   4440
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   9
         Left            =   3720
         TabIndex        =   48
         Top             =   4440
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   10
         Left            =   1800
         TabIndex        =   50
         Top             =   4800
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   10
         Left            =   3720
         TabIndex        =   51
         Top             =   4800
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_MntActual 
         Height          =   330
         Index           =   11
         Left            =   1800
         TabIndex        =   53
         Top             =   5160
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjusteMntAjustado 
         Height          =   330
         Index           =   11
         Left            =   3720
         TabIndex        =   54
         Top             =   5160
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAjustesCierra 
         Height          =   495
         Left            =   5640
         TabIndex        =   58
         Top             =   5880
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Picture         =   "frmRH_Aguinaldos.frx":0FF8
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   1
         Left            =   5760
         TabIndex        =   60
         Top             =   1560
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":1636
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   2
         Left            =   5760
         TabIndex        =   61
         Top             =   1920
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":1F07
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   3
         Left            =   5760
         TabIndex        =   62
         Top             =   2280
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":27D8
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   4
         Left            =   5760
         TabIndex        =   63
         Top             =   2640
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":30A9
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   5
         Left            =   5760
         TabIndex        =   64
         Top             =   3000
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":397A
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   6
         Left            =   5760
         TabIndex        =   65
         Top             =   3360
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":424B
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   7
         Left            =   5760
         TabIndex        =   66
         Top             =   3720
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":4B1C
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   8
         Left            =   5760
         TabIndex        =   67
         Top             =   4080
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":53ED
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   9
         Left            =   5760
         TabIndex        =   68
         Top             =   4440
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":5CBE
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   10
         Left            =   5760
         TabIndex        =   69
         Top             =   4800
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":658F
      End
      Begin XtremeSuiteControls.PushButton btnAjusteClear 
         Height          =   330
         Index           =   11
         Left            =   5760
         TabIndex        =   70
         Top             =   5160
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   79
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Aguinaldos.frx":6E60
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste_Registro_Usuario 
         Height          =   330
         Left            =   7560
         TabIndex        =   75
         Top             =   1560
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtAjuste_Registro_Fecha 
         Height          =   330
         Left            =   7560
         TabIndex        =   79
         Top             =   1200
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtAjuste_Actualiza_Usuario 
         Height          =   330
         Left            =   7560
         TabIndex        =   80
         Top             =   3000
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   19
         Left            =   7680
         TabIndex        =   83
         Top             =   2400
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Actualización"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   18
         Left            =   6360
         TabIndex        =   82
         Top             =   2640
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   17
         Left            =   6360
         TabIndex        =   81
         Top             =   3000
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   16
         Left            =   7680
         TabIndex        =   78
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registro"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   15
         Left            =   6360
         TabIndex        =   77
         Top             =   1200
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   14
         Left            =   6360
         TabIndex        =   76
         Top             =   1560
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   13
         Left            =   3720
         TabIndex        =   74
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Valor Ajustado"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   73
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Valor Actual"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scEmpleado 
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   480
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "..."
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   120
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Ajuste Manual de Aguinaldo"
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
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   52
         Top             =   5160
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Noviembre"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   49
         Top             =   4800
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Octubre"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   46
         Top             =   4440
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Septiembre"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   43
         Top             =   4080
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agosto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   40
         Top             =   3720
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Julio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   37
         Top             =   3360
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Junio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   34
         Top             =   3000
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mayo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   31
         Top             =   2640
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Abril"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   28
         Top             =   2280
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Marzo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   25
         Top             =   1920
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Febrero"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Enero"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAjusteMes 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   1200
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Diciembre"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnBoletaBancos 
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   3600
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Aguinaldos.frx":7731
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnAguinaldos 
      Height          =   735
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Consultar"
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
      Picture         =   "frmRH_Aguinaldos.frx":7E38
   End
   Begin MSComctlLib.ProgressBar prgBarX 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   9045
      Width           =   16965
      _ExtentX        =   29924
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton btnAguinaldos 
      Height          =   735
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmRH_Aguinaldos.frx":8856
   End
   Begin XtremeSuiteControls.PushButton btnAguinaldos 
      Height          =   735
      Index           =   3
      Left            =   2280
      TabIndex        =   4
      Top             =   7200
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Pago"
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
      Picture         =   "frmRH_Aguinaldos.frx":905B
   End
   Begin XtremeSuiteControls.PushButton btnAguinaldos 
      Height          =   735
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   8040
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Notificación Email"
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
      Picture         =   "frmRH_Aguinaldos.frx":94BD
   End
   Begin XtremeSuiteControls.PushButton btnAguinaldos 
      Height          =   735
      Index           =   5
      Left            =   2280
      TabIndex        =   6
      Top             =   8040
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Boletas Impresas"
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
      Picture         =   "frmRH_Aguinaldos.frx":9CDA
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7215
      Left            =   4680
      TabIndex        =   9
      Top             =   1680
      Width           =   12135
      _Version        =   524288
      _ExtentX        =   21405
      _ExtentY        =   12726
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   23
      SpreadDesigner  =   "frmRH_Aguinaldos.frx":A496
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboAnio 
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   794
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAguinaldos 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Actualizar Base"
      Top             =   2760
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   1291
      _StockProps     =   79
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
      Picture         =   "frmRH_Aguinaldos.frx":B20A
   End
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   7200
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Autorización"
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
      Picture         =   "frmRH_Aguinaldos.frx":BBCD
   End
   Begin XtremeSuiteControls.RadioButton rbVisualizar 
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   72
      Top             =   1200
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Lista de Ajustes Aplicados"
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
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Salidas por Banco.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   3975
   End
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Nómina"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   -1080
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cálculo de Aguinaldos"
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
      Height          =   600
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   16935
   End
End
Attribute VB_Name = "frmRH_Aguinaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub btnAguinaldos_Click(Index As Integer)
Dim i As Integer, pDetalle As String

On Error GoTo vError

pDetalle = "AGUINALDO > Nómina: " & cboNomina.Text & "   Periodo.: " & cboAnio.Text

Select Case Index
    Case 0 'Actualizar
        Call Bitacora("Actualiza", "Calculos de " & pDetalle)
        Call sbActualiza
        
    Case 1 'Buscar
        Call sbBuscar
    
    Case 2 'Exportar
        Call sbExportar
        
    Case 3 'Pago
    
    If lblEstado.Tag = "P" Then
        MsgBox "Los aguinaldos de este periodo ya fueron cancelados!", vbExclamation
        Exit Sub
    End If
    
    i = MsgBox("Esta seguro que desea PAGAR los Aguinaldos: " & cboAnio.Text & " ?", vbYesNo)
    If i = vbYes Then
    
        Me.MousePointer = vbHourglass
        
        strSQL = "exec spRH_Aguinaldo_Pago '" & cboNomina.ItemData(cboNomina.ListIndex) _
                & "'," & cboAnio.Text & ",'" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Pago de " & pDetalle)
        
        Me.MousePointer = vbDefault
        MsgBox "Aguinaldos han sido Reportados a Bancos para su pago!", vbInformation
        
        Call cboNomina_Click
    End If
    
    
    Case 4 'Notificacion Email
    
          Call Bitacora("Aplica", "Notificación Email de " & pDetalle)
    
          Call sbRH_Boleta_Aguinaldo_Email(cboNomina.ItemData(cboNomina.ListIndex), cboAnio.Text, "")
        
          Me.MousePointer = vbDefault
          MsgBox "Boletas de Aguinaldo, fueron enviadas por Email a los Empleados!", vbInformation
  
    
    Case 5 'Boleta Impresora
  
        Call Bitacora("Aplica", "Boletas de " & pDetalle)
        Call sbBoleta("")
    
    
End Select

Exit Sub

vError:
 Me.MousePointer = vbHourglass
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnAjusteClear_Click(Index As Integer)

txtAjusteMntAjustado(Index).Text = ""

End Sub

Private Sub btnAjustesAplica_Click()

Dim pDetalle As String, i As Integer

Dim pEnero As String, pFebrero As String, pMarzo As String, pAbril As String, pMayo As String
Dim pJunio As String, pJulio As String, pAgosto As String, pSeptiembre As String, pOctubre As String, pNoviembre As String, pDiciembre As String

On Error GoTo vError


For i = 0 To 11
    If IsNumeric(txtAjusteMntAjustado(i).Text) Then
        pDetalle = CCur(txtAjusteMntAjustado(i).Text)
    Else
        pDetalle = "Null"
    End If

    Select Case i
        Case 0
            pDiciembre = pDetalle
        Case 1
            pEnero = pDetalle
        Case 2
            pFebrero = pDetalle
        Case 3
            pMarzo = pDetalle
        Case 4
            pAbril = pDetalle
        Case 5
            pMayo = pDetalle
        Case 6
            pJunio = pDetalle
        Case 7
            pJulio = pDetalle
        Case 8
            pAgosto = pDetalle
        Case 9
            pSeptiembre = pDetalle
        Case 10
            pOctubre = pDetalle
        Case 11
            pNoviembre = pDetalle
    End Select
Next i


Me.MousePointer = vbHourglass

pDetalle = "Nómina: " & cboNomina.Text & "   Periodo.: " & cboAnio.Text & ", Empleado Id: " & scEmpleado.Tag

strSQL = "exec spRH_Aguinaldo_Ajuste '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & cboAnio.Text & ",'" & scEmpleado.Tag _
        & "', '" & glogon.Usuario & "', " & pDiciembre & ", " & pEnero & ", " & pFebrero & ", " & pMarzo _
        & ", " & pAbril & ", " & pMayo & ", " & pJunio & ", " & pJulio & ", " & pAgosto & ", " & pSeptiembre _
        & ", " & pOctubre & ", " & pNoviembre
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Ajuste de Aguinaldo, " & pDetalle)

Me.MousePointer = vbDefault

MsgBox "Ajustes aplicados correctamente!", vbInformation

Call btnAjustesCierra_Click
Call btnAguinaldos_Click(1)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAjustesCierra_Click()
gbAjustes.Visible = False
vGrid.Visible = True
End Sub

Private Sub btnAutorizacion_Click()
Dim i As Integer, pDetalle As String

    i = MsgBox("Esta seguro que desea Autorizar el Aguinaldo Nómina: " & cboNomina.Text & ", Periodo: " & cboAnio.Text & " ?", vbYesNo)
    If i = vbYes Then
        
        Me.MousePointer = vbHourglass
               
        
        pDetalle = "AGUINALDO > Nómina: " & cboNomina.Text & "   Periodo.: " & cboAnio.Text
               
               
        strSQL = "exec spRH_Aguinaldo_Autoriza '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & cboAnio.Text & ",'" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Autorización de " & pDetalle)
        
        Me.MousePointer = vbDefault
        MsgBox "El Aguinaldo para esta Tipo de Nómina ha sido Autorizado!", vbInformation
        
        Call cboNomina_Click
    
    End If

End Sub

Private Sub btnBoletaBancos_Click()

Dim pTitulo As String

On Error GoTo vError

pTitulo = "Nómina: " & cboNomina.Text & "   Periodo.: " & cboAnio.Text & "   Estado: " & lblEstado.Caption

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH, Nómina: Boleta de Control"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fxSubTitulo = '" & pTitulo & "'"
    .Formulas(3) = "fxUsuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(4) = "fxFecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Aguinaldo_Boleta_Control.rpt")
    strSQL = "{vRH_Aguinaldo_Estado_Rsm.COD_NOMINA} = '" & cboNomina.ItemData(cboNomina.ListIndex) _
            & "' AND {vRH_Aguinaldo_Estado_Rsm.PERIODO_ID} = " & cboAnio.Text
        
    .SelectionFormula = strSQL
    .SubreportToChange = "sbBancosResumen"
    
    .StoredProcParam(0) = cboNomina.ItemData(cboNomina.ListIndex)
    .StoredProcParam(1) = cboAnio.Text



    .Action = 1
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboAnio_Click()
If vPaso Then Exit Sub

Select Case Mid(cboAnio.ItemData(cboAnio.ListIndex), 1, 1)
Case "A"
    lblEstado.Caption = "Abierto"
Case "X"
    lblEstado.Caption = "Autorizada"
Case "P"
    lblEstado.Caption = "Pagado"
End Select

lblEstado.Tag = Mid(cboAnio.ItemData(cboAnio.ListIndex), 1, 1)

Call sbBuscar

End Sub

Private Sub cboNomina_Click()
If vPaso Then Exit Sub

Call sbPeriodos

End Sub

Private Sub Form_Activate()
vModulo = 23
End Sub

Private Sub Form_Load()

vModulo = 23

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

vPaso = True
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)
vPaso = False

vGrid.MaxRows = 0

With lswBancos.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Casos", 800, vbRightJustify
    .Add , , "Monto", 1800, vbRightJustify
End With
lswBancos.BackColor = RGB(214, 234, 248)



lblEstado.Caption = ""


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbBoleta(Optional pEmpleado As String = "")

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH: Boleta de Aguinaldo"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Boleta_Aguinaldo.rpt")
    strSQL = "{vRH_Aguinaldo_Boleta.COD_NOMINA} = '" & cboNomina.ItemData(cboNomina.ListIndex) _
            & "' AND {vRH_Aguinaldo_Boleta.PERIODO_ID} = " & cboAnio.Text
                
    If pEmpleado <> "" Then
        strSQL = strSQL & " AND {vRH_Aguinaldo_Boleta.EMPLEADO_ID} = '" & pEmpleado & "'"
    End If
        
        
     .SelectionFormula = strSQL
    .PrintReport
End With

End Sub



Private Sub sbPeriodos()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Aguinaldo_Control_Consulta '" & cboNomina.ItemData(cboNomina.ListIndex) & "', '" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)

vPaso = True

cboAnio.Clear

Do While Not rs.EOF
  cboAnio.AddItem CStr(rs!Periodo_Id)
  cboAnio.ItemData(cboAnio.ListCount - 1) = rs!Estado & "_" & CStr(rs!Periodo_Id)
  rs.MoveNext
Loop
rs.MoveFirst

vPaso = False


Me.MousePointer = vbDefault

cboAnio.Text = CStr(rs!Periodo_Id)


Call cboAnio_Click

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbActualiza()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Aguinaldo_Control_Calcula '" & cboNomina.ItemData(cboNomina.ListIndex) _
        & "'," & cboAnio.Text & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Call sbBuscar

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
    Case rbVisualizar(0).Value 'Listado General
        strSQL = "exec spRH_Aguinaldo_Resumen_Consulta '" & cboNomina.ItemData(cboNomina.ListIndex) & "', " & cboAnio.Text

    Case rbVisualizar(1).Value 'Ajustes
        strSQL = "exec spRH_Aguinaldo_Ajustes_Load '" & cboNomina.ItemData(cboNomina.ListIndex) & "', " & cboAnio.Text
End Select


Call sbCargaGrid_Local(vGrid, vGrid.MaxCols, strSQL)


btnAguinaldos.Item(0).Enabled = False 'Actualiza
btnAguinaldos.Item(1).Enabled = False 'Consulta
btnAguinaldos.Item(2).Enabled = False 'Exporta
btnAguinaldos.Item(3).Enabled = False 'Pago
btnAguinaldos.Item(4).Enabled = False 'Email
btnAguinaldos.Item(5).Enabled = False 'Boletas

btnAutorizacion.Enabled = False

Select Case lblEstado.Tag
    Case "A" 'Abierta
            
        If btnAguinaldos.Item(0).Tag = "1" Then
            btnAguinaldos.Item(0).Enabled = True 'Actualiza
        End If
        
        btnAguinaldos.Item(1).Enabled = True 'Consulta
        btnAguinaldos.Item(2).Enabled = True 'Exporta
    
        If btnAutorizacion.Tag = "1" Then
            btnAutorizacion.Enabled = True
        End If
    
    Case "X" 'Autorizada
        btnAguinaldos.Item(1).Enabled = True 'Consulta
        btnAguinaldos.Item(2).Enabled = True 'Exporta
        
        If btnAguinaldos.Item(3).Tag = "1" Then
            btnAguinaldos.Item(3).Enabled = True 'Paga
            btnAguinaldos.Item(4).Enabled = True 'Boleta Email
            btnAguinaldos.Item(5).Enabled = True 'Boleta Imprime
        End If
        
   
    Case "P" 'Pagada
        btnAguinaldos.Item(1).Enabled = True 'Consulta
        btnAguinaldos.Item(2).Enabled = True 'Exporta
        
        If btnAguinaldos.Item(3).Tag = "1" Then
            btnAguinaldos.Item(4).Enabled = True 'Boleta Email
            btnAguinaldos.Item(5).Enabled = True 'Boleta Imprime
        End If


End Select



'Resumen de Salidas por Banco
Dim pCasos As Long, pMonto As Currency

pCasos = 0
pMonto = 0

strSQL = "exec spRH_Aguinaldo_Pago_Banco_Rsm '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & cboAnio.Text
Call OpenRecordSet(rs, strSQL)

With lswBancos.ListItems
    .Clear
  Do While Not rs.EOF
    Set itmX = .Add(, , rs!Descripcion)
        itmX.SubItems(1) = Format(rs!Casos, "###,###0")
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
    
        pCasos = pCasos + rs!Casos
        pMonto = pMonto + rs!Monto
    rs.MoveNext
  Loop
  rs.Close

Set itmX = .Add(, , "")
    itmX.SubItems(1) = "________________"
    itmX.SubItems(2) = "________________"
Set itmX = .Add(, , "TOTAL:")
    itmX.SubItems(1) = Format(pCasos, "###,###0")
    itmX.SubItems(2) = Format(pMonto, "Standard")

End With





Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbCargaGrid_Local(pGrid As Object, pGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim i As Integer

On Error GoTo vErrorLoad

Call OpenRecordSet(rs, strSQL, 0)
  
pGrid.MaxRows = 0
Do While Not rs.EOF
  pGrid.MaxRows = pGrid.MaxRows + 1
  pGrid.Row = pGrid.MaxRows
  For i = 4 To pGrid.MaxCols
  
    pGrid.Col = i
    
    Select Case i
        Case 4
            pGrid.Text = rs!Empleado_ID
        Case 5
            pGrid.Text = rs!IDENTIFICACION
        Case 6
            pGrid.Text = rs!NOMBRE_COMPLETO
        Case 7
            pGrid.Text = Format(rs!Total_Pagar, "Standard")
        Case 8
            pGrid.Text = Format(rs!Corte, "yyyy-mm-dd")
        Case 9
            pGrid.Text = rs!CUENTA_BANCARIA & ""
    
        Case 10
            pGrid.Text = Format(rs!Diciembre, "Standard")
        Case 11
            pGrid.Text = Format(rs!Enero, "Standard")
        Case 12
            pGrid.Text = Format(rs!Febrero, "Standard")
        Case 13
            pGrid.Text = Format(rs!Marzo, "Standard")
        Case 14
            pGrid.Text = Format(rs!Abril, "Standard")
        Case 15
            pGrid.Text = Format(rs!Mayo, "Standard")
        Case 16
            pGrid.Text = Format(rs!Junio, "Standard")
        Case 17
            pGrid.Text = Format(rs!Julio, "Standard")
        Case 18
            pGrid.Text = Format(rs!Agosto, "Standard")
        Case 19
            pGrid.Text = Format(rs!Setiembre, "Standard")
        Case 20
            pGrid.Text = Format(rs!Octubre, "Standard")
        Case 21
            pGrid.Text = Format(rs!Noviembre, "Standard")
    
        Case 22
            pGrid.Text = rs!Tesoreria_NSolicitud & ""
        Case 23
            pGrid.Text = rs!Tesoreria_Fecha & ""
    
    
    End Select

  Next i
  rs.MoveNext
Loop
rs.Close

Exit Sub

vErrorLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub



Private Sub sbExportar()
 Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols

    vHeaders.Headers(1) = "@"
    vHeaders.Headers(2) = "Pr"
    vHeaders.Headers(3) = "Aj"
    
    vHeaders.Headers(4) = "Empleado Id"
    vHeaders.Headers(5) = "Identificación"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "Aguinaldo"
    vHeaders.Headers(8) = "Corte Calculo"
    vHeaders.Headers(9) = "IBAN"

    vHeaders.Headers(10) = "Diciembre"
    vHeaders.Headers(11) = "Enero"
    vHeaders.Headers(12) = "Febrero"
    vHeaders.Headers(13) = "Marzo"
    vHeaders.Headers(14) = "Abril"
    vHeaders.Headers(15) = "Mayo"
    vHeaders.Headers(16) = "Junio"
    vHeaders.Headers(17) = "Julio"
    vHeaders.Headers(18) = "Agosto"
    vHeaders.Headers(19) = "Septiembre"
    vHeaders.Headers(20) = "Octubre"
    vHeaders.Headers(21) = "Noviembre"
    vHeaders.Headers(22) = "Tesorería Id"
    vHeaders.Headers(23) = "Tesorería Fecha"

Select Case True
    Case rbVisualizar(0).Value 'General
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Aguinaldos_" & cboAnio.Text & "_" & cboNomina.Text)
    Case rbVisualizar(1).Value 'Ajustes del Periodo
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Aguinaldos_Ajustes_" & cboAnio.Text & "_" & cboNomina.Text)
    
End Select

End Sub



Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
vGrid.Width = Me.Width - (vGrid.Left + 350)
vGrid.Height = Me.Height - (vGrid.Top + 700)

End Sub

Private Sub rbVisualizar_Click(Index As Integer)
Call btnAguinaldos_Click(1) 'Buscar
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbPeriodos

End Sub

Private Sub sbAguinaldo_Ajustes(pNomina As String, pPeriodo As Long, pEmpleadoId As String, pNombre As String)

vGrid.Visible = False
gbAjustes.Visible = True

scEmpleado.Tag = pEmpleadoId
scEmpleado.Caption = pEmpleadoId & "  ¦  " & pNombre

On Error GoTo vError

Dim i As Integer


'Limpia
For i = 0 To txtAjuste_MntActual.Count - 1
    txtAjuste_MntActual(i).Text = "0.00"
    txtAjusteMntAjustado(i).Text = ""
Next i

strSQL = "EXEC spRH_Aguinaldo_Ajuste_Consulta '" & pNomina & "', " & pPeriodo & ", '" & pEmpleadoId & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 txtAjuste_MntActual(0).Text = Format(rs!Diciembre, "Standard")
 txtAjuste_MntActual(1).Text = Format(rs!Enero, "Standard")
 txtAjuste_MntActual(2).Text = Format(rs!Febrero, "Standard")
 txtAjuste_MntActual(3).Text = Format(rs!Marzo, "Standard")
 txtAjuste_MntActual(4).Text = Format(rs!Abril, "Standard")
 txtAjuste_MntActual(5).Text = Format(rs!Mayo, "Standard")
 txtAjuste_MntActual(6).Text = Format(rs!Junio, "Standard")
 txtAjuste_MntActual(7).Text = Format(rs!Julio, "Standard")
 txtAjuste_MntActual(8).Text = Format(rs!Agosto, "Standard")
 txtAjuste_MntActual(9).Text = Format(rs!Setiembre, "Standard")
 txtAjuste_MntActual(10).Text = Format(rs!Octubre, "Standard")
 txtAjuste_MntActual(11).Text = Format(rs!Noviembre, "Standard")
       
      
 txtAjusteMntAjustado(0).Text = ""
 
 txtAjusteMntAjustado(0).Text = IIf(IsNull(rs!Diciembre_Aj), "", Format(rs!Diciembre_Aj, "Standard"))
 txtAjusteMntAjustado(1).Text = IIf(IsNull(rs!Enero_Aj), "", Format(rs!Enero_Aj, "Standard"))
 txtAjusteMntAjustado(2).Text = IIf(IsNull(rs!Febrero_Aj), "", Format(rs!Febrero_Aj, "Standard"))
 txtAjusteMntAjustado(3).Text = IIf(IsNull(rs!Marzo_Aj), "", Format(rs!Marzo_Aj, "Standard"))
 txtAjusteMntAjustado(4).Text = IIf(IsNull(rs!Abril_Aj), "", Format(rs!Abril_Aj, "Standard"))
 txtAjusteMntAjustado(5).Text = IIf(IsNull(rs!Mayo_Aj), "", Format(rs!Mayo_Aj, "Standard"))
 txtAjusteMntAjustado(6).Text = IIf(IsNull(rs!Junio_Aj), "", Format(rs!Junio_Aj, "Standard"))
 txtAjusteMntAjustado(7).Text = IIf(IsNull(rs!Julio_Aj), "", Format(rs!Julio_Aj, "Standard"))
 txtAjusteMntAjustado(8).Text = IIf(IsNull(rs!Agosto_Aj), "", Format(rs!Agosto_Aj, "Standard"))
 txtAjusteMntAjustado(9).Text = IIf(IsNull(rs!Setiembre_Aj), "", Format(rs!Setiembre_Aj, "Standard"))
 txtAjusteMntAjustado(10).Text = IIf(IsNull(rs!Octubre_Aj), "", Format(rs!Octubre_Aj, "Standard"))
 txtAjusteMntAjustado(11).Text = IIf(IsNull(rs!Noviembre_Aj), "", Format(rs!Noviembre_Aj, "Standard"))
 
 
 txtAjuste_Registro_Fecha.Text = rs!Registro_Fecha & ""
 txtAjuste_Registro_Usuario.Text = rs!Registro_Usuario & ""
 
 txtAjuste_Actualiza_Fecha.Text = rs!MODIFICA_FECHA & ""
 txtAjuste_Actualiza_Usuario.Text = rs!MODIFICA_USUARIO & ""
 
'Limpia Datos no Utilizables
For i = 0 To txtAjuste_MntActual.Count - 1
    If txtAjusteMntAjustado(i).Text = "-1.00" Then
        txtAjusteMntAjustado(i).Text = ""
    End If
Next i
 
End If


rs.Close
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtAjusteMntAjustado_GotFocus(Index As Integer)
On Error GoTo vError

txtAjusteMntAjustado(Index).Text = CCur(txtAjusteMntAjustado(Index).Text)

vError:
End Sub

Private Sub txtAjusteMntAjustado_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Index = 11 Then
            txtAjusteMntAjustado(0).SetFocus
    Else
            txtAjusteMntAjustado(Index + 1).SetFocus
    End If
End If
End Sub

Private Sub txtAjusteMntAjustado_LostFocus(Index As Integer)
On Error GoTo vError

txtAjusteMntAjustado(Index).Text = Format(CCur(txtAjusteMntAjustado(Index).Text), "Standard")

Exit Sub

vError:
txtAjusteMntAjustado(Index).Text = ""

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim pEmpleadoId As String, pNombre As String

On Error GoTo vError

vGrid.Row = Row
vGrid.Col = 4
pEmpleadoId = vGrid.Text
vGrid.Col = 6
pNombre = vGrid.Text

Select Case Col
    Case 1 'Email
        Call sbRH_Boleta_Aguinaldo_Email(cboNomina.ItemData(cboNomina.ListIndex), cboAnio.Text, pEmpleadoId)
        MsgBox "Correo Electrónico Enviado al Empleado: " & pEmpleadoId, vbInformation
    Case 2 'Boleta
        Call sbBoleta(pEmpleadoId)
    Case 3 'Ajustes
         Call sbAguinaldo_Ajustes(cboNomina.ItemData(cboNomina.ListIndex), cboAnio.Text, pEmpleadoId, pNombre)
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
