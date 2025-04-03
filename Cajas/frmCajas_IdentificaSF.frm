VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_IdentificaSF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Identifica Depósitos en Tramite"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox fraIdentifica 
      Height          =   6135
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   8415
      _Version        =   1572864
      _ExtentX        =   14843
      _ExtentY        =   10821
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1935
         Left            =   0
         TabIndex        =   42
         Top             =   360
         Width           =   8415
         _Version        =   1572864
         _ExtentX        =   14843
         _ExtentY        =   3413
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
         Appearance      =   21
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_NSolicitud 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_NumDocId 
         Height          =   315
         Left            =   5880
         TabIndex        =   9
         Top             =   600
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Fecha 
         Height          =   315
         Left            =   5880
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Cedula 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   3240
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtId_Nombre 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   3600
         Width           =   6495
         _Version        =   1572864
         _ExtentX        =   11456
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Banco 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Width           =   6495
         _Version        =   1572864
         _ExtentX        =   11456
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Descripcion 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   1440
         Width           =   6495
         _Version        =   1572864
         _ExtentX        =   11456
         _ExtentY        =   556
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Monto 
         Height          =   315
         Left            =   5880
         TabIndex        =   15
         Top             =   2400
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnIdentifica 
         Height          =   495
         Index           =   0
         Left            =   5640
         TabIndex        =   16
         Top             =   5280
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCajas_IdentificaSF.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnIdentifica 
         Height          =   495
         Index           =   1
         Left            =   6840
         TabIndex        =   17
         Top             =   5280
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCajas_IdentificaSF.frx":0727
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboOrigenRecursos 
         Height          =   330
         Left            =   3000
         TabIndex        =   30
         Top             =   4800
         Width           =   5055
         _Version        =   1572864
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.ComboBox cboPagadores 
         Height          =   330
         Left            =   3000
         TabIndex        =   31
         Top             =   4440
         Width           =   5055
         _Version        =   1572864
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.PushButton btnAdjuntos 
         Height          =   330
         Left            =   6840
         TabIndex        =   32
         ToolTipText     =   "Adjuntar Documentos"
         Top             =   4080
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Adjuntos"
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
         Picture         =   "frmCajas_IdentificaSF.frx":0E3D
      End
      Begin XtremeSuiteControls.FlatEdit txtDepositoId 
         Height          =   315
         Left            =   5880
         TabIndex        =   36
         Top             =   30
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   375
         Index           =   9
         Left            =   1560
         TabIndex        =   35
         Top             =   4680
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Origen Recursos"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   375
         Index           =   10
         Left            =   1560
         TabIndex        =   34
         Top             =   4320
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Pagadores"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   33
         Top             =   4080
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación de Recursos:"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   29
         Top             =   2760
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación del Caso:"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   9495
         _Version        =   1572864
         _ExtentX        =   16748
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Identificación del Propietario del Depósito"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Solicitud:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuenta:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Descripción:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   22
         Top             =   600
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Documento:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   21
         Top             =   2400
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   19
         Top             =   3240
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   18
         Top             =   3600
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre:"
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
   Begin XtremeSuiteControls.DateTimePicker dtpId_Inicio 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpId_Corte 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   6855
      _Version        =   1572864
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.FlatEdit txtId_NumDoc 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   28
      Top             =   1320
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCajas_IdentificaSF.frx":0EC6
      ImageAlignment  =   4
   End
   Begin FPSpreadADO.fpSpread vGridId 
      Height          =   6375
      Left            =   0
      TabIndex        =   37
      Top             =   1800
      Width           =   11775
      _Version        =   524288
      _ExtentX        =   20770
      _ExtentY        =   11245
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
      MaxCols         =   11
      SpreadDesigner  =   "frmCajas_IdentificaSF.frx":15C6
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtMnt_Inicio 
      Height          =   315
      Left            =   8160
      TabIndex        =   38
      Top             =   1320
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   556
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMnt_Hasta 
      Height          =   315
      Left            =   9960
      TabIndex        =   39
      Top             =   1320
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   556
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   41
      Top             =   1320
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Identificar"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCajas_IdentificaSF.frx":1ED6
      ImageAlignment  =   4
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Montos.: "
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
      Index           =   9
      Left            =   7200
      TabIndex        =   40
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Identifica Depósitos en Tramite"
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
      Index           =   0
      Left            =   2160
      TabIndex        =   27
      Top             =   240
      Width           =   5505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta .:"
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
      Index           =   16
      Left            =   3960
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Doc.:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha .:"
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
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmCajas_IdentificaSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub sbIdentifica_Lista()

On Error GoTo vError

Dim i As Long, curTotal As Currency
Dim pDepositoId As Long, pTesoreriaId As Long, pDocumento As String, pMonto As Currency, pFecha As Date, pDescripcion As String, pCuenta As String

If vPaso Then Exit Sub

txtId_Cedula.Text = GLOBALES.gTag
txtId_Nombre.Text = GLOBALES.gTag2

lsw.ListItems.Clear
curTotal = 0

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id Deposito", 1200
    .Add , , "Id Tesoreria", 1200
    .Add , , "Documento", 1800
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Descripción", 2200
    .Add , , "Cuenta", 2200
End With

With vGridId
    
For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    If .Value = vbChecked Then
        .Col = 2
        pDepositoId = .Text
        .Col = 3
        pTesoreriaId = IIf(IsNumeric(.Text), .Text, 0)
        .Col = 4
        pCuenta = .Text
        
        .Col = 6
        pDocumento = .Text
        .Col = 7
        pFecha = .Text
        .Col = 8
        pMonto = .Text
        .Col = 9
        pDescripcion = .Text
        
        curTotal = curTotal + pMonto
        
        Set itmX = lsw.ListItems.Add(, , pDepositoId)
            itmX.SubItems(1) = pTesoreriaId
            itmX.SubItems(2) = pDocumento
            itmX.SubItems(3) = Format(pMonto, "Standard")
            itmX.SubItems(4) = Format(pFecha, "yyyy-mm-dd")
            itmX.SubItems(5) = pDescripcion
            itmX.SubItems(6) = pCuenta
            
            .Col = 4
            itmX.Tag = .CellTag 'Id Banco
    End If
    
Next i

End With

txtId_Monto.Text = Format(curTotal, "Standard")


'With vGridId
'    .Row = Row
'    .Col = 2
'    txtDepositoId.Text = .Text
'    .Col = 3
'    txtId_NSolicitud = .Text
'    .Col = 4
'    txtId_Banco.Text = .Text
'    txtId_Banco.Tag = .CellTag
'    .Col = 6
'    txtId_NumDocId.Text = .Text
'    .Col = 7
'    txtId_Fecha.Text = .Text
'    .Col = 8
'    txtId_Monto.Text = .Text
'    .Col = 9
'    txtId_Descripcion.Text = .Text
'End With


If lsw.ListItems.Count = 0 Then
    MsgBox "No se ha seleccionado ningún caso!", vbExclamation
    Exit Sub
End If

fraIdentifica.Visible = True


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAccion_Click(Index As Integer)

Select Case Index
    Case 0 'Buscar
            vPaso = True
                Call sbConsultaDPTramite
            vPaso = False
    Case 1 'Identificar
        Call sbIdentifica_Lista
    
End Select

End Sub

Private Sub btnAdjuntos_Click()

If txtId_Cedula.Text <> "" Then
 gGA.Modulo = "CAJ"
 gGA.Llave_01 = txtId_Cedula.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End If

End Sub

Private Sub btnIdentifica_Click(Index As Integer)

On Error GoTo vError

If Index = 1 Then
   fraIdentifica.Visible = False
   Exit Sub
End If

If txtId_Nombre.Text = "" Then
    MsgBox "No se ha especificado ningún Id de Cliente válido", vbExclamation
    Exit Sub
End If

If lsw.ListItems.Count = 0 Then
    MsgBox "No se ha seleccionado ningún caso!", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass

Dim i As Long
Dim pDepositoId As Long, pBancoId As Long, pDocumento As String

With lsw.ListItems
    For i = 1 To .Count
      pDepositoId = .Item(i).Text
      pBancoId = .Item(i).Tag
      pDocumento = .Item(i).SubItems(2)
        
        strSQL = "exec spCajas_Identifica_TES_Depositos " & pBancoId & ",'" & pDocumento & "','" & txtId_Cedula.Text _
               & "', '" & txtId_Nombre.Text & "', '" & glogon.Usuario & "', '" & cboPagadores.ItemData(cboPagadores.ListIndex) _
               & "', '" & cboOrigenRecursos.ItemData(cboOrigenRecursos.ListIndex) & "', " & pDepositoId
        Call ConectionExecute(strSQL)

    Next i
    
End With

Me.MousePointer = vbDefault

fraIdentifica.Visible = False

MsgBox "caso identificado correctamente!", vbInformation


vPaso = True
    Call sbConsultaDPTramite
vPaso = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
vModulo = 5

'Carga las cuentas bancarias asiganadas a la forma de pago
vPaso = True


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboBanco.Clear

strSQL = "exec spCajas_DepositosCuentasBancariasAut 'DP'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboBanco.AddItem Trim(rs!Cta) & " - " & Trim(rs!Descripcion & "")
 cboBanco.ItemData(cboBanco.ListCount - 1) = CStr(rs!Id_Banco)
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
    cboBanco.Text = Trim(rs!Cta) & " - " & Trim(rs!Descripcion & "")
End If
rs.Close


vPaso = True
    
vGridId.MaxCols = 11
vGridId.MaxRows = 0

vPaso = False


'Identificacion de Recursos
strSQL = "select COD_ENTIDAD_PAGO as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ENTIDADES_PAGO" _
       & " WHERE ACTIVA = 1 ORDER BY COD_ENTIDAD_PAGO"
Call sbCbo_Llena_New(cboPagadores, strSQL, False, True)

strSQL = "select COD_ORIGEN_RECURSOS as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ORIGEN_RECURSOS" _
       & "  WHERE ACTIVA = 1 ORDER BY COD_ORIGEN_RECURSOS"
Call sbCbo_Llena_New(cboOrigenRecursos, strSQL, False, True)



dtpId_Corte.Value = fxFechaServidor
dtpId_Inicio.Value = DateAdd("d", -10, dtpId_Corte.Value)

txtMnt_Inicio.Text = Format(0, "Standard")
txtMnt_Hasta.Text = Format(999999999999.99, "Standard")


fraIdentifica.Visible = False

Call RefrescaTags(Me)
Call Formularios(Me)

End Sub


Private Sub sbConsultaDPTramite()
Dim i As Long

On Error GoTo vError

If cboBanco.ListCount = 0 Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select Tra.*, Bn.Descripcion as 'BancoDesc'" _
        & " From TES_DEPOSITOS_TRAMITE Tra inner join Tes_Bancos Bn on Tra.id_banco = Bn.id_Banco" _
        & " Where Tra.ID_REQUERIDA = 1 And Tra.IDENTIFICADO = 0" _
        & " and  fecha between '" & Format(dtpId_Inicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpId_Corte.Value, "yyyy/mm/dd") & " 23:59:59'"


If Len(Trim(txtId_NumDoc.Text)) > 0 Then
    strSQL = strSQL & " and Tra.Documento like '%" & txtId_NumDoc.Text & "%'"
End If

strSQL = strSQL & " and Tra.Id_Banco = " & cboBanco.ItemData(cboBanco.ListIndex)

strSQL = strSQL & " and Tra.Monto between " & CCur(txtMnt_Inicio.Text) & " and " & CCur(txtMnt_Hasta.Text)

Call OpenRecordSet(rs, strSQL)

vGridId.MaxRows = 0


  Do While Not rs.EOF
    vGridId.MaxRows = vGridId.MaxRows + 1
    vGridId.Row = vGridId.MaxRows
         
    vGridId.Col = 1

    For i = 2 To vGridId.MaxCols
      vGridId.Col = i
      Select Case i
         Case 2 'Id Tramite
            vGridId.Text = CStr(rs!DP_TRAMITE_ID)
         
         Case 3 'Id
            vGridId.Text = CStr(rs!NSolicitud)
         Case 4 'Cuenta
            vGridId.Text = rs!BancoDesc & ""
            vGridId.CellTag = rs!Id_Banco
         Case 5 ' Tipo
            vGridId.Text = "DP"
         Case 6 'Num Documento
            vGridId.Text = rs!Documento
         Case 7 'Fecha del Documento
            vGridId.Text = Format(rs!fecha, "dd/mm/yyyy")
         Case 8 'Monto
            vGridId.Text = Format(rs!Monto, "Standard")
         Case 9 'Descripcion
            vGridId.Text = rs!Descripcion
         Case 10 'Registro Fecha
            vGridId.Text = rs!Registro_Fecha & ""
         Case 11 'Registro Usuario
            vGridId.Text = rs!Registro_Usuario & ""
      
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtMnt_Hasta_GotFocus()
On Error GoTo vError
  txtMnt_Hasta.Text = CCur(txtMnt_Hasta.Text)
vError:
End Sub

Private Sub txtMnt_Hasta_LostFocus()
On Error GoTo vError
  txtMnt_Hasta.Text = Format(CCur(txtMnt_Hasta.Text), "Standard")
vError:
End Sub

Private Sub txtMnt_Inicio_GotFocus()
On Error GoTo vError
  txtMnt_Inicio.Text = CCur(txtMnt_Inicio.Text)
vError:
End Sub

Private Sub txtMnt_Inicio_LostFocus()
On Error GoTo vError
  txtMnt_Inicio.Text = Format(CCur(txtMnt_Inicio.Text), "Standard")
vError:
End Sub




