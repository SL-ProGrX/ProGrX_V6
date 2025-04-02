VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.TaskPanel.v22.1.0.ocx"
Begin VB.Form frmDSB_Colaborador 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Portal del Colaborador"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   14940
   StartUpPosition =   2  'CenterScreen
   Begin XtremeTaskPanel.TaskPanel tpMain 
      Height          =   6885
      Left            =   0
      TabIndex        =   29
      Top             =   1800
      Width           =   2760
      _Version        =   1441793
      _ExtentX        =   4868
      _ExtentY        =   12144
      _StockProps     =   64
      VisualTheme     =   13
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6405
      Left            =   2760
      TabIndex        =   31
      Top             =   1800
      Width           =   8295
      _Version        =   1441793
      _ExtentX        =   14626
      _ExtentY        =   11307
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
   Begin XtremeSuiteControls.GroupBox gbCambioClave 
      Height          =   2175
      Left            =   5520
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   9135
      _Version        =   1441793
      _ExtentX        =   16113
      _ExtentY        =   3836
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtClaveNueva 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Top             =   840
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtClaveConfirma 
         Height          =   315
         Left            =   3360
         TabIndex        =   16
         Top             =   1200
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnClaveCambio 
         Height          =   495
         Index           =   0
         Left            =   6480
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cambiar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmDSB_Colaborador.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnClaveCambio 
         Height          =   495
         Index           =   1
         Left            =   7680
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmDSB_Colaborador.frx":0727
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Clave Confirmación"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   840
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Clave Nueva"
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
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   9255
         _Version        =   1441793
         _ExtentX        =   16325
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Cambio de Contraseña"
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
   Begin XtremeSuiteControls.GroupBox gbLogin 
      Height          =   3015
      Left            =   5520
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   5318
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnConectar 
         Height          =   495
         Index           =   0
         Left            =   6480
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Acceder"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmDSB_Colaborador.frx":0D65
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.CheckBox chkLoginVincular 
         Height          =   615
         Left            =   4560
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Vincular al usuario de sistema con este colaborador"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtLoginClave 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLoginIdentificacion 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.FlatEdit txtLoginNombre 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.ComboBox cboLoginEmpleadoId 
         Height          =   330
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.PushButton btnConectar 
         Height          =   495
         Index           =   1
         Left            =   6120
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reestableccer"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmDSB_Colaborador.frx":148C
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnConectar 
         Height          =   495
         Index           =   2
         Left            =   7680
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmDSB_Colaborador.frx":1B8C
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboLoginAccion 
         Height          =   330
         Left            =   1920
         TabIndex        =   27
         Top             =   1440
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.FlatEdit txtLoginEmail 
         Height          =   315
         Left            =   1920
         TabIndex        =   28
         Top             =   1920
         Width           =   6975
         _Version        =   1441793
         _ExtentX        =   12303
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Gestión"
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
      Begin XtremeSuiteControls.Label lblLoginGestion 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Clave"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Empleado Id"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   9015
         _Version        =   1441793
         _ExtentX        =   15901
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Login a su cuenta de Colaborador"
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
   Begin XtremeSuiteControls.PushButton btnMenu 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Login"
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
   Begin VB.PictureBox picFoto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   1185
      TabIndex        =   34
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   6480
      Top             =   240
   End
   Begin XtremeSuiteControls.FlatEdit txtEmpleadoId 
      Height          =   495
      Left            =   4200
      TabIndex        =   22
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   720
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   375
      Left            =   6240
      TabIndex        =   24
      Top             =   720
      Width           =   8415
      _Version        =   1441793
      _ExtentX        =   14843
      _ExtentY        =   661
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnEditarDetalle 
      Height          =   345
      Left            =   9720
      TabIndex        =   32
      Top             =   1400
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Editar"
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmDSB_Colaborador.frx":22A2
      ImageAlignment  =   0
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":289D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":2B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":2DAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":2F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":30C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":326E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":3410
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":359B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":3726
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":382A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":3AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":3BC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":3E41
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":3F07
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":40A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":4243
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnMenu 
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   36
      Top             =   1440
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Clave"
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
   Begin XtremeShortcutBar.ShortcutCaption TituloOpciones 
      Height          =   480
      Left            =   2760
      TabIndex        =   33
      Top             =   1320
      Width           =   12615
      _Version        =   1441793
      _ExtentX        =   22251
      _ExtentY        =   847
      _StockProps     =   14
      Caption         =   "Detalles:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo1 
      Height          =   480
      Left            =   0
      TabIndex        =   30
      Top             =   1320
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   847
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   21
      Top             =   720
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Identifiación"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   20
      Top             =   120
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Empleado Id"
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
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "frmDSB_Colaborador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, mTop As Long, mLeft As Long
Dim itmX As ListViewItem, mWidth As Long, mHeight As Long



Const Id_TaskItem_DatosPersonales = 0
Const Id_TaskItem_RelacionLaboral = 1
Const Id_TaskItem_Otros_Add = 2

Const Id_TaskItem_Telefonos = 3
Const Id_TaskItem_Familiares = 4
Const Id_TaskItem_Cuentas = 5

Const Id_TaskItem_Boletas_Pago = 6
Const Id_TaskItem_Plan_Carrera = 7
Const Id_TaskItem_Vacaciones = 8
Const Id_TaskItem_Permisos = 9
Const Id_TaskItem_Incapacidades = 10

Const Id_TaskItem_Tarjetas = 11
Const Id_TaskItem_Accion_Personal = 17


Private Sub sbTaskPanel_Load()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    tpMain.VisualTheme = xtpTaskPanelThemeVisualStudio2012Light
    
  
'    Set Group = tpMain.Groups.Add(0, "Registro")
'    Group.ToolTip = "Información Principal para el Registro de la Persona"
'    Group.Special = True
'
'
'    Group.Items.Add Id_TaskItem_DatosPersonales, "Datos Personales", xtpTaskItemTypeLink, 4
'    Group.Items.Add Id_TaskItem_RelacionLaboral, "Relación Laboral", xtpTaskItemTypeLink, 1
'    Group.Items.Add Id_TaskItem_Otros_Add, "Adicionales y Portal", xtpTaskItemTypeLink, 10
    
    Set Group = tpMain.Groups.Add(0, "Detalles")
    Group.ToolTip = "Datos Complementarios"
    
    Group.Items.Add Id_TaskItem_Familiares, "Familiares & Contactos", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Cuentas, "Cuentas Bancarias", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Tarjetas, "Tarjetas", xtpTaskItemTypeLink, 9
   
    
    Set Group = tpMain.Groups.Add(0, "Histórico")
    Group.Expanded = True
    Group.Items.Add Id_TaskItem_Boletas_Pago, "Boletas de Pago", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Vacaciones, "Vacaciones", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Incapacidades, "Incapacidades", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Permisos, "Permisos", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Accion_Personal, "Acciones de Personal", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Plan_Carrera, "Plan de Carrera", xtpTaskItemTypeLink, 3
    
   
    tpMain.SetImageList imlTaskPanelIcons
    
'    tpMain.SetMargins 5, 5, 5, 5, 5
    

End Sub


Private Sub sbPersona_Foto_Load()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from RH_Personas where Empleado_Id = '" & txtEmpleadoId.Text & "'"

Set picFoto.Picture = fxImagen_Leer(strSQL, "FOTO")

picFoto.PaintPicture picFoto.Picture, 0, 0, picFoto.ScaleWidth, picFoto.ScaleHeight

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault

End Sub

Private Sub sbTaskPanel_Accion(ItemId As Integer)

Dim fraX As Frame


If Trim(txtIdentificacion.Text) = "" Then Exit Sub

On Error GoTo vError

'
'Select Case ItemId
'  Case Id_TaskItem_DatosPersonales  'Datos de Contato
'
'
'    TituloOpcion.Caption = "Datos de Contacto"
'
'
'
'    Exit Sub
'
'  Case Id_TaskItem_RelacionLaboral  'Relación Laboral
'    TituloOpcion.Caption = "Datos Laborales"
'
'    tcMain.Item(1).Selected = True
'    txtCentroCod.SetFocus
'
'    Exit Sub
'
'  Case Id_TaskItem_Otros_Add 'Información Adicional
'
'    TituloOpcion.Caption = "Adicionales y Portal..."
'    tcMain.Item(2).Selected = True
'
'    DoEvents
'
'    Call sbPersona_Foto_Load
'
'    Exit Sub
'
'End Select
'
'If Not vEditar Then
'    MsgBox "Se encuentra en modo de Registro, guarde los datos de la persona y luego ingrese a esta opción!", vbInformation
'    Exit Sub
'End If


lsw.ColumnHeaders.Clear
lsw.ListItems.Clear
lsw.Checkboxes = False

btnEditarDetalle.Visible = False


Select Case ItemId
  Case Id_TaskItem_Telefonos  'Telefonos
        
    TituloOpciones.Caption = "Lista de Teléfonos..:"
    TituloOpciones.Tag = "Telefonos"
    
    btnEditarDetalle.Visible = True
        
    lsw.ColumnHeaders.Add 1, , "Numero", 1500
    lsw.ColumnHeaders.Add 2, , "Tipo", 1500
    lsw.ColumnHeaders.Add 3, , "Extension", 1500
    lsw.ColumnHeaders.Add 4, , "Contacto", 2500
    
    
    strSQL = "Select * From Telefonos where Cedula='" & Trim(txtIdentificacion) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , Trim(rs!Numero))
           itmX.SubItems(1) = (rs!Tipo)
           itmX.SubItems(2) = Trim(rs!Ext) & ""
           itmX.SubItems(3) = Trim(rs!contacto) & ""
       rs.MoveNext
    Loop
    rs.Close
    
    
  Case Id_TaskItem_Familiares 'Familiares
    btnEditarDetalle.Visible = True
 
     
    TituloOpciones.Caption = "Lista de Familiares..:"
    TituloOpciones.Tag = "Familiares"
    
    lsw.ColumnHeaders.Add , , "Identificación", 1500
    lsw.ColumnHeaders.Add , , "Nombre", 3500
    lsw.ColumnHeaders.Add , , "Parentesco", 1100, vbCenter
    

    
    strSQL = "select Identificacion, Nombre, Parentesco_Desc from vRH_Personas_Familiares where Empleado_Id = '" & Trim(txtEmpleadoId.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!Identificacion)
           itmX.SubItems(1) = Trim(rs!Nombre)
           itmX.SubItems(2) = Trim(rs!parentesco_Desc)
       
       rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Cuentas 'Cuentas Bancarias
    
    btnEditarDetalle.Visible = True
    
    
    TituloOpciones.Caption = "Cuentas bancarias..:"
    TituloOpciones.Tag = "Cuentas"
    
    lsw.ColumnHeaders.Add 1, , "Cuenta", 2500
    lsw.ColumnHeaders.Add 2, , "Banco", 3500
    lsw.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
    lsw.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
    lsw.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
    lsw.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
    lsw.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
    lsw.ColumnHeaders.Add 8, , "Fecha", 2500
    lsw.ColumnHeaders.Add 9, , "Usuario", 2500

        strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
               & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
               & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
               & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
               & " where C.Identificacion = '" & Trim(txtIdentificacion) & "'" 'and C.Modulo = 'AFI'
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!COD_DIVISA
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!Activa = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!REGISTRO_FECHA & ""
           itmX.SubItems(8) = rs!REGISTRO_USUARIO & ""
     
       rs.MoveNext
    Loop
    rs.Close
    
    
    
  
  Case Id_TaskItem_Boletas_Pago 'Boletas de Pago
  
    
    TituloOpciones.Caption = "Boletas de Pago..:"
    TituloOpciones.Tag = "BoletaPago"
    
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "No. Nómina", 1200
    lsw.ColumnHeaders.Add , , "No. Pago", 1000, vbCenter
    lsw.ColumnHeaders.Add , , "Nomina", 1000, vbCenter
    lsw.ColumnHeaders.Add , , "Inicio", 1500, vbCenter
    lsw.ColumnHeaders.Add , , "Corte", 1500, vbCenter
    lsw.ColumnHeaders.Add , , "Salario", 1500, vbRightJustify
    lsw.ColumnHeaders.Add , , "Ingresos", 1500, vbRightJustify
    lsw.ColumnHeaders.Add , , "Egresos", 1500, vbRightJustify
    lsw.ColumnHeaders.Add , , "A Pagar", 1500, vbRightJustify
    lsw.ColumnHeaders.Add , , "Descripción", 1500, vbRightJustify
            
    
    strSQL = "select Top 50 * from vRH_Boleta_Pago_List Where Empleado_Id = '" _
           & Trim(txtEmpleadoId.Text) & "' order by Fecha_Corte desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!Nomina_Num)
          itmX.SubItems(1) = rs!NPago_Mes
          itmX.SubItems(2) = rs!COD_NOMINA
          itmX.Tag = rs!COD_NOMINA
          
          itmX.SubItems(3) = Format(rs!Fecha_Inicio, "yyyy-mm-dd")
          itmX.SubItems(4) = Format(rs!Fecha_Corte, "yyyy-mm-dd")
          itmX.SubItems(5) = Format(rs!SALARIO_ORDINARIO, "Standard")
          itmX.SubItems(6) = Format(rs!Ingresos, "Standard")
          itmX.SubItems(7) = Format(rs!Egresos, "Standard")
          itmX.SubItems(8) = Format(rs!Salario_Neto, "Standard")
          itmX.SubItems(9) = rs!Nomina_Desc
          
      rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Accion_Personal 'Accion
     
    TituloOpciones.Caption = "Acciones de Personal..:"
    TituloOpciones.Tag = "AccionPersonal"

    lsw.ListItems.Clear
    lsw.ColumnHeaders.Add , , "No. Boleta", 1500
    lsw.ColumnHeaders.Add , , "Fecha", 1500, vbCenter
    lsw.ColumnHeaders.Add , , "Tipo", 2000
    lsw.ColumnHeaders.Add , , "Salario", 1600, vbRightJustify
    lsw.ColumnHeaders.Add , , "Salario Ant.", 1600, vbRightJustify
    lsw.ColumnHeaders.Add , , "Puesto", 2100
    lsw.ColumnHeaders.Add , , "Puesto Ant.", 2100
    lsw.ColumnHeaders.Add , , "Centro", 2100
    lsw.ColumnHeaders.Add , , "Centro Ant.", 2100
    lsw.ColumnHeaders.Add , , "Departamento", 2100
    lsw.ColumnHeaders.Add , , "Dept. Ant.", 2100
    lsw.ColumnHeaders.Add , , "Sección", 2100
    lsw.ColumnHeaders.Add , , "Sección Ant.", 2100
    
    lsw.ColumnHeaders.Add , , "Nómina", 2100
    lsw.ColumnHeaders.Add , , "Nómina Ant.", 2100
    
    lsw.ColumnHeaders.Add , , "Estado", 2100
    lsw.ColumnHeaders.Add , , "Estado Ant.", 2100
    
    lsw.ColumnHeaders.Add , , "Notas", 2100
    
    lsw.ColumnHeaders.Add , , "Fecha", 2500
    lsw.ColumnHeaders.Add , , "Usuario", 2500


    strSQL = "select * From vRH_Accion_Personal Where Empleado_id = '" & Trim(txtEmpleadoId.Text) & "' order by cod_Accion desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!Cod_Accion)
           itmX.SubItems(1) = Format(rs!Fecha_Accion, "dd/MM/yyyy")
           itmX.SubItems(2) = rs!TipoAccionDesc
           itmX.SubItems(3) = Format(rs!Salario_Actual, "Standard")
           itmX.SubItems(4) = Format(rs!ANT_Salario, "Standard")
           itmX.SubItems(5) = rs!PuestoDesc
           itmX.SubItems(6) = rs!A_PuestoDesc
           itmX.SubItems(7) = rs!CentroDesc
           itmX.SubItems(8) = rs!A_CentroDesc
           itmX.SubItems(9) = rs!DepartamentoDesc
           itmX.SubItems(10) = rs!A_DepartamentoDesc
           itmX.SubItems(11) = rs!SeccionDesc
           itmX.SubItems(12) = rs!A_SeccionDesc
           itmX.SubItems(13) = rs!NominaDesc
           itmX.SubItems(14) = rs!NominaDesc
           itmX.SubItems(15) = rs!EstadoPersonaDesc
           itmX.SubItems(16) = rs!A_EstadoPersonaDesc
           itmX.SubItems(17) = rs!NOTAS & ""
           itmX.SubItems(18) = rs!REGISTRO_FECHA & ""
           itmX.SubItems(19) = Trim(rs!REGISTRO_USUARIO & "")

       rs.MoveNext
    Loop
    rs.Close

  
  Case Id_TaskItem_Plan_Carrera 'Plan de Carrera
  
    With lsw
            .ListItems.Clear
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Id", 900
            .ColumnHeaders.Add , , "Nivel", 2100
            .ColumnHeaders.Add , , "Curso", 2100
            .ColumnHeaders.Add , , "Estado", 1500, vbCenter
            .ColumnHeaders.Add , , "Nota", 1500, vbCenter
            .ColumnHeaders.Add , , "Usuario", 1500
            .ColumnHeaders.Add , , "Fecha", 1500
            
            TituloOpciones.Caption = "Plan de Carrera..:"
            TituloOpciones.Tag = "PlanCarrera"
            
'            strSQL = "Select I.*,P.nombre as Promotor " _
'                   & " From Afi_Ingresos I left join promotores P on I.id_promotor = P.id_promotor" _
'                   & " where I.Cedula='" & Trim(txtIdentificacion) & "'"
'            Call OpenRecordSet(rs, strSQL)
'            Do While Not rs.EOF
'               Set itmX = .ListItems.Add(, , rs!consec)
'                   itmX.SubItems(1) = rs!Usuario & ""
'                   itmX.SubItems(2) = rs!fecha & ""
'                   itmX.SubItems(3) = Format(rs!Fecha_Ingreso)
'                   itmX.SubItems(4) = rs!Boleta & ""
'                   itmX.SubItems(5) = rs!promotor & ""
'               rs.MoveNext
'            Loop
'            rs.Close
    End With
  
  
  Case Id_TaskItem_Vacaciones  'Vacaciones
    
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "Boleta", 1000
    lsw.ColumnHeaders.Add , , "Motivo", 2100
    lsw.ColumnHeaders.Add , , "Inicio", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Corte", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Días", 900, vbCenter
    lsw.ColumnHeaders.Add , , "Estado", 1500, vbCenter
    lsw.ColumnHeaders.Add , , "Usuario", 1500
    lsw.ColumnHeaders.Add , , "Fecha", 1500
    
    
    TituloOpciones.Caption = "Vacaciones..:"
    TituloOpciones.Tag = "Vacaciones"
    
    strSQL = "Select * from vRH_Boleta_Vacaciones" _
           & " where Empleado_Id = '" & Trim(txtEmpleadoId.Text) _
           & "' order by Boleta_VAC desc"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!Boleta_Id)
           itmX.SubItems(1) = Trim(rs!Motivo)
           itmX.SubItems(2) = Format(rs!Fecha_Salida, "dd/mm/yyyy")
           itmX.SubItems(3) = Format(rs!Fecha_Entrada, "dd/mm/yyyy")
           itmX.SubItems(4) = rs!Dias_Disfrutados & ""
           itmX.SubItems(5) = rs!Estado_Transaccion
           itmX.SubItems(6) = rs!REGISTRO_USUARIO & ""
           itmX.SubItems(7) = rs!REGISTRO_FECHA & ""
           
       rs.MoveNext
    Loop
    rs.Close
    
    lsw.Enabled = True
  
  
  Case Id_TaskItem_Incapacidades 'Incapacidades
  
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "Boleta", 1000
    lsw.ColumnHeaders.Add , , "Motivo", 2100
    lsw.ColumnHeaders.Add , , "Inicio", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Corte", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Días", 900, vbCenter
    lsw.ColumnHeaders.Add , , "Estado", 1500, vbCenter
    lsw.ColumnHeaders.Add , , "Usuario", 1500
    lsw.ColumnHeaders.Add , , "Fecha", 1500
    
    
    TituloOpciones.Caption = "Incapacidades..:"
    TituloOpciones.Tag = "Incapacidades"
    
    strSQL = "Select * from vRH_Boleta_Incapacidades" _
           & " where Empleado_Id = '" & Trim(txtEmpleadoId.Text) _
           & "' order by Boleta_ID desc"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!Boleta_Id)
           itmX.SubItems(1) = Trim(rs!Motivo)
           itmX.SubItems(2) = Format(rs!Fecha_Salida, "dd/mm/yyyy")
           itmX.SubItems(3) = Format(rs!Fecha_Entrada, "dd/mm/yyyy")
           itmX.SubItems(4) = rs!Dias & ""
           itmX.SubItems(5) = rs!Estado_Transaccion
           itmX.SubItems(6) = rs!REGISTRO_USUARIO & ""
           itmX.SubItems(7) = rs!REGISTRO_FECHA & ""
           
       rs.MoveNext
    Loop
    rs.Close
    
    lsw.Enabled = True
  
  Case Id_TaskItem_Permisos 'Permisos
  
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "Boleta", 1000
    lsw.ColumnHeaders.Add , , "Motivo", 2100
    lsw.ColumnHeaders.Add , , "Fecha/Permiso", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Hr. Inicio", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Hr. Corte", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Horas", 900, vbCenter
    lsw.ColumnHeaders.Add , , "Estado", 1500, vbCenter
    lsw.ColumnHeaders.Add , , "Usuario", 1500
    lsw.ColumnHeaders.Add , , "Fecha", 1500
    
    
    TituloOpciones.Caption = "Permisos..:"
    TituloOpciones.Tag = "Permisos"
    
    strSQL = "Select * from vRH_Boleta_Permisos" _
           & " where Empleado_Id = '" & Trim(txtEmpleadoId.Text) _
           & "' order by Boleta_ID desc"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!Boleta_Id)
           itmX.SubItems(1) = Trim(rs!Motivo)
           itmX.SubItems(2) = Format(rs!Hora_Inicio, "dd/mm/yyyy")
           itmX.SubItems(3) = Format(rs!Hora_Inicio, "hh:mm:ss")
           itmX.SubItems(4) = Format(rs!Hora_Corte, "hh:mm:ss")
           itmX.SubItems(5) = rs!Hrs_Total & ""
           itmX.SubItems(6) = rs!Estado_Transaccion
           itmX.SubItems(7) = rs!REGISTRO_USUARIO & ""
           itmX.SubItems(8) = rs!REGISTRO_FECHA & ""
           
       rs.MoveNext
    Loop
    rs.Close
    
    lsw.Enabled = True
    
    
  Case Id_TaskItem_Tarjetas  'Tarjetas
  
    btnEditarDetalle.Visible = True
    
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add 1, , "No. Tarjeta", 2500
    lsw.ColumnHeaders.Add 2, , "Tipo", 2100
    lsw.ColumnHeaders.Add 3, , "Vence", 1500
    
    TituloOpciones.Caption = "Tarjetas..:"
    TituloOpciones.Tag = "Tarjetas"
  
  
            strSQL = "exec spAFI_PersonaTarjetas_Consulta " & gPortal.Empresa_Id & ",'" & txtIdentificacion.Text & "',''"
            Call OpenRecordSet(rs, strSQL)
            
            With lsw.ListItems
               .Clear
               Do While Not rs.EOF
                Set itmX = .Add(, , rs!Tarjeta_Mask)
                    itmX.SubItems(1) = rs!Tarjeta_Tipo
                    itmX.SubItems(2) = Format(rs!Tarjeta_Vence, "MM/YY")
                rs.MoveNext
               Loop
               rs.Close
            End With



End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub








Private Sub btnClaveCambio_Click(Index As Integer)

Select Case Index
    Case 0 'Cambia
        If txtClaveNueva.Text = txtClaveConfirma.Text And Len(txtClaveNueva.Text) >= 3 Then
            Call sbClave_Cambia(txtEmpleadoId.Text, txtClaveNueva.Text)
        End If
        
    Case 1 'Cierra
        gbCambioClave.Visible = False
        
        Call sbOpcionesVisibles(True)



End Select

End Sub

Private Sub btnConectar_Click(Index As Integer)

If cboLoginEmpleadoId.ListCount = 0 Then
   MsgBox "Consulta a un Empleado!", vbInformation
End If



Select Case Index
    Case 0 'Login
        If fxClave_Valida(cboLoginEmpleadoId.Text, txtLoginClave.Text) Then
           Call sbEmpleado_Load(cboLoginEmpleadoId.Text)
           
           If chkLoginVincular.Value = xtpChecked Then
                Call sbEmpleado_Vincula(cboLoginEmpleadoId.Text)
           End If
        
           gbLogin.Visible = False
           
        Else
            MsgBox "La Clave registrada no es válida, verifique!", vbExclamation
        End If
        
    Case 1 'Reestablece
        Call sbClave_Reestablece(cboLoginEmpleadoId.Text)
    
    Case 2 'Cierra
        Me.Hide
End Select

End Sub

Private Sub btnMenu_Click(Index As Integer)


gbLogin.Visible = False
gbCambioClave.Visible = False

Call sbOpcionesVisibles(False)

Select Case Index
    Case 0 'Login
        gbLogin.Visible = True
    Case 1 'Cambio Clave
        gbCambioClave.Visible = True
End Select

End Sub

Private Sub cboLoginAccion_Click()
If vPaso Then Exit Sub
If cboLoginAccion.ListCount = 0 Then Exit Sub

If Mid(cboLoginAccion.Text, 1, 1) = "A" Then
    lblLoginGestion.Caption = "Clave"
    txtLoginClave.Visible = True
    txtLoginEmail.Visible = False
    btnConectar(0).Visible = True
    btnConectar(1).Visible = False
    chkLoginVincular.Visible = True
Else
    lblLoginGestion.Caption = "Email Registrado"
    txtLoginClave.Visible = False
    txtLoginEmail.Visible = True
    btnConectar(0).Visible = False
    btnConectar(1).Visible = True
    chkLoginVincular.Visible = False
End If


End Sub

Private Sub cboLoginEmpleadoId_Click()

If vPaso Then Exit Sub
If cboLoginEmpleadoId.ListCount = 0 Then Exit Sub

Call sbEmpleado_Load_Nombre(txtLoginIdentificacion.Text, cboLoginEmpleadoId.Text)

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

Select Case TituloOpciones.Tag
  Case "BoletaPago"
    Call sbRH_Boleta_Pago(txtEmpleadoId.Text, Item.Tag, Item.Text)
    
  Case "AccionPersonal"
    Call sbRH_Boleta_Accion_Personal(Item.Text, txtEmpleadoId.Text)
  
  Case "Vacaciones"
    Call sbRH_Boleta_Vacaciones(Item.Text, txtEmpleadoId.Text)
  
  Case "Incapacidades"
    Call sbRH_Boleta_Incapacidad(Item.Text, txtEmpleadoId.Text)

  Case "Permisos"
    Call sbRH_Boleta_Permisos(Item.Text, txtEmpleadoId.Text)

  Case "PlanCarrera"

  Case Else
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tpMain_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Call sbTaskPanel_Accion(Item.Id)
End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mTop = 1320
mLeft = 20


mWidth = 15060
mHeight = 8520

gbLogin.Visible = False
gbCambioClave.Visible = False


gbLogin.top = mTop
gbLogin.Left = mLeft

gbCambioClave.top = mTop
gbCambioClave.Left = mLeft

vPaso = True
cboLoginAccion.Clear
cboLoginAccion.AddItem "Acceso"
cboLoginAccion.AddItem "Reestablece"
cboLoginAccion.Text = "Acceso"
vPaso = False

Call cboLoginAccion_Click



Call Form_Resize


End Sub

Private Sub Form_Resize()
On Error Resume Next

Dim pWidth As Long, pHeight As Long


If Me.Width < mWidth Then
    pWidth = mWidth
Else
    pWidth = Me.Width
End If

If Me.Height < mHeight Then
    pHeight = mHeight
Else
    pHeight = Me.Height
End If


imgBanner.Width = pWidth

TituloOpciones.Width = pWidth

lsw.Width = pWidth - (lsw.Left + 150)

lsw.Height = pHeight - (lsw.top + 450)

tpMain.Height = pHeight - (tpMain.top + 450)



End Sub

Private Sub TimerX_Timer()
 TimerX.Interval = 0
 TimerX.Enabled = False
 
 Call sbLogin
 
End Sub




Private Sub sbEmpleado_Id(pIdentificacion As String)

On Error GoTo vError

Me.MousePointer = vbHourglass


vPaso = True

strSQL = "exec spRH_Portal_Consulta_Id  '" & pIdentificacion & "'"

Call sbCbo_Llena_New(cboLoginEmpleadoId, strSQL, False, True)

vPaso = False

Call cboLoginEmpleadoId_Click

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbEmpleado_Load_Nombre(pIdentificacion As String, pEmpleadoId As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Portal_Consulta_Id  '" & pIdentificacion & "', '" & pEmpleadoId & "'"

Call OpenRecordSet(rs, strSQL)

txtLoginNombre.Text = ""

If Not rs.EOF And Not rs.BOF Then
   txtLoginNombre.Text = rs!NOMBRE_COMPLETO
End If

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub





Private Sub sbEmpleado_Vincula(pEmpleadoId As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Portal_Empleado_Vincula  '" & pEmpleadoId & "','" & glogon.Usuario & "', '" & glogon.AppName _
       & "', '" & glogon.AppVersion & "','" & glogon.Maquina & "'"

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbOpcionesVisibles(pVisible As Boolean)


scTitulo1.Visible = pVisible
TituloOpciones.Visible = pVisible

btnMenu.Item(0).Visible = pVisible
btnMenu.Item(1).Visible = pVisible

btnEditarDetalle.Visible = pVisible

tpMain.Visible = pVisible
lsw.Visible = pVisible

End Sub


Private Sub sbEmpleado_Load(pEmpleadoId As String)

On Error GoTo vError

Me.MousePointer = vbHourglass


Call sbOpcionesVisibles(True)

strSQL = "exec spRH_Portal_Empleado_Load  '" & pEmpleadoId & "','" & glogon.Usuario & "', '" & glogon.AppName _
       & "', '" & glogon.AppVersion & "','" & glogon.Maquina & "'"

Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    txtEmpleadoId.Text = rs!Empleado_ID
    txtIdentificacion.Text = rs!Identificacion
    txtNombre.Text = rs!NOMBRE_COMPLETO

    Call sbPersona_Foto_Load
    Call sbTaskPanel_Load
End If


Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxClave_Valida(pEmpleadoId As String, pClave As String) As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spRH_Portal_Clave_Valida  '" & pEmpleadoId & "', '" & pClave & "'"
       
Call OpenRecordSet(rs, strSQL)


fxClave_Valida = IIf((rs!Existe = 1), True, False)

Me.MousePointer = vbDefault

Exit Function

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function



Private Sub sbClave_Reestablece(pEmpleadoId As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Portal_Clave_Reestablece  '" & pEmpleadoId & "', '" & txtLoginEmail.Text & "', '" & glogon.Usuario & "', '" & glogon.AppName _
       & "', '" & glogon.AppVersion & "','" & glogon.Maquina & "'"
       
Call OpenRecordSet(rs, strSQL)


Me.MousePointer = vbDefault

If rs!Cambio = 1 Then
    MsgBox "Se ha enviado un correo a su cuenta con la nueva clave de acceso, verifique!", vbInformation
Else
    MsgBox "No fue posible reestablecer su contraseña, verifique que su identificacion y correo sean los registrados en RRHH", vbExclamation

End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbClave_Cambia(pEmpleadoId As String, pClave As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Portal_Clave_Cambia  '" & pEmpleadoId & "','" & pClave & "','" & glogon.Usuario & "', '" & glogon.AppName _
       & "', '" & glogon.AppVersion & "','" & glogon.Maquina & "'"

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Se ha cambia la clave del colaborador para uso del portal!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbLogin()

On Error GoTo vError

Me.MousePointer = vbHourglass

'Verifica si el Usuario se encuentra vinculado
strSQL = "exec spRH_Portal_Vinculado '" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)

If rs!Empleado_ID <> "" Then

    Call sbEmpleado_Load(rs!Empleado_ID)
    
    Me.MousePointer = vbDefault
    Exit Sub
End If


Me.MousePointer = vbDefault


'Oculta Opciones
Call sbOpcionesVisibles(False)


'Abre Opción de Login
gbLogin.Visible = True
txtLoginIdentificacion.SetFocus



Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtLoginIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call sbEmpleado_Id(txtLoginIdentificacion.Text)
End If

End Sub
