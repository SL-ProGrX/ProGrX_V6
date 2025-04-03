VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_ReporteCierres 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Cierre Caja"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11235
   Icon            =   "frmCajas_ReporteCierres.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   11235
   Begin XtremeSuiteControls.PushButton btnRefresh 
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   2020
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCajas_ReporteCierres.frx":0ECA
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   12091
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
      Item(0).Caption =   "Aperturas"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Accesos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswH"
      Item(2).Caption =   "Recepción de Cierres"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "gbRecibe"
      Item(2).Control(1)=   "gbCierre"
      Item(2).Control(2)=   "btnRecibe"
      Item(2).Control(3)=   "btnRevisa"
      Item(2).Control(4)=   "lswDP"
      Item(2).Control(5)=   "ShortcutCaption1"
      Item(2).Control(6)=   "ShortcutCaption2"
      Item(2).Control(7)=   "txtTotalDeposito"
      Begin XtremeSuiteControls.ListView lswDP 
         Height          =   1695
         Left            =   -69880
         TabIndex        =   36
         Top             =   3840
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1572864
         _ExtentX        =   18865
         _ExtentY        =   2990
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6015
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   11055
         _Version        =   1572864
         _ExtentX        =   19500
         _ExtentY        =   10610
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
      Begin XtremeSuiteControls.ListView lswH 
         Height          =   6015
         Left            =   -70000
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1572864
         _ExtentX        =   19500
         _ExtentY        =   10610
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
      Begin XtremeSuiteControls.PushButton btnRecibe 
         Height          =   495
         Left            =   -66280
         TabIndex        =   26
         Top             =   5640
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Recibe"
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
         Picture         =   "frmCajas_ReporteCierres.frx":15CA
      End
      Begin XtremeSuiteControls.GroupBox gbRecibe 
         Height          =   1215
         Left            =   -69760
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1572864
         _ExtentX        =   18441
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Datos de Recepción y Revisión:"
         ForeColor       =   16711680
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
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtRecibe_Usuario 
            Height          =   330
            Left            =   1800
            TabIndex        =   21
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
         Begin XtremeSuiteControls.FlatEdit txtRevisa_Usuario 
            Height          =   330
            Left            =   1800
            TabIndex        =   30
            Top             =   840
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
         Begin XtremeSuiteControls.FlatEdit txtRecibe_Fecha 
            Height          =   330
            Left            =   3840
            TabIndex        =   22
            Top             =   480
            Width           =   2415
            _Version        =   1572864
            _ExtentX        =   4260
            _ExtentY        =   582
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRevisa_Fecha 
            Height          =   330
            Left            =   3840
            TabIndex        =   31
            Top             =   840
            Width           =   2415
            _Version        =   1572864
            _ExtentX        =   4260
            _ExtentY        =   582
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   840
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Revisión"
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
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Recepción"
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
      End
      Begin XtremeSuiteControls.GroupBox gbCierre 
         Height          =   1455
         Left            =   -69760
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1572864
         _ExtentX        =   18441
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Información del Cierre:"
         ForeColor       =   16711680
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
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCierre_Usuario 
            Height          =   330
            Left            =   1800
            TabIndex        =   24
            Top             =   1080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtApertura_Usuario 
            Height          =   330
            Left            =   1800
            TabIndex        =   34
            Top             =   720
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCierre_Estado 
            Height          =   330
            Left            =   8280
            TabIndex        =   39
            Top             =   1080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCierre_Id 
            Height          =   330
            Left            =   8280
            TabIndex        =   40
            Top             =   720
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCierre_Caja 
            Height          =   330
            Left            =   1800
            TabIndex        =   42
            Top             =   360
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCierre_CajaDesc 
            Height          =   330
            Left            =   3960
            TabIndex        =   43
            Top             =   360
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11245
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtApertura_Fecha 
            Height          =   330
            Left            =   3960
            TabIndex        =   35
            Top             =   720
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCierre_Fecha 
            Height          =   330
            Left            =   3960
            TabIndex        =   25
            Top             =   1080
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Caja:"
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
            Index           =   3
            Left            =   6720
            TabIndex        =   38
            Top             =   1080
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Estado:"
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
            Index           =   2
            Left            =   6720
            TabIndex        =   37
            Top             =   720
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Apertura ID:"
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
            Index           =   5
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cierre"
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
            Index           =   4
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Apertura"
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
      End
      Begin XtremeSuiteControls.PushButton btnRevisa 
         Height          =   495
         Left            =   -64480
         TabIndex        =   27
         Top             =   5640
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Revisa"
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
         Picture         =   "frmCajas_ReporteCierres.frx":1CEA
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalDeposito 
         Height          =   330
         Left            =   -61480
         TabIndex        =   46
         Top             =   3500
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   -63040
         TabIndex        =   45
         Top             =   3480
         Visible         =   0   'False
         Width           =   3855
         _Version        =   1572864
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Total Depósitos:"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -69880
         TabIndex        =   44
         Top             =   3480
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1572864
         _ExtentX        =   12091
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Depósitos registrados al cierre:"
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
   End
   Begin XtremeSuiteControls.CheckBox chkModoSupervisor 
      Height          =   372
      Left            =   4800
      TabIndex        =   9
      Top             =   1560
      Width           =   3372
      _Version        =   1572864
      _ExtentX        =   5948
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Aperturas/Cierres (Saldos Abiertos)"
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
   End
   Begin XtremeSuiteControls.PushButton btnCierre 
      Height          =   615
      Left            =   9480
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Forzar Cierre"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmCajas_ReporteCierres.frx":2411
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_ReporteCierres.frx":2DFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_ReporteCierres.frx":2F34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCaja 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.FlatEdit txtApertura 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.FlatEdit txtCajaDesc 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1080
      Width           =   5535
      _Version        =   1572864
      _ExtentX        =   9763
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1470
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Begin MSComctlLib.Toolbar tblAplicar 
         Height          =   312
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   556
         ButtonWidth     =   1799
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Reportes del Cierre"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Resumen"
                     Text            =   "Resumen"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Cierre"
                     Text            =   "Informe de Cierre"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Movimientos"
                     Text            =   "Movimientos"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   582
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
      Height          =   330
      Left            =   3120
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   2020
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCajas_ReporteCierres.frx":304E
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboFiltro 
      Height          =   330
      Left            =   4800
      TabIndex        =   28
      Top             =   2040
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fechas"
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
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Cierres de Cajas"
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
      Height          =   372
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   6612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Apertura"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmCajas_ReporteCierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vCierreCiego As Boolean, vPaso As Boolean

Private Sub btnCierre_Click()
Dim vAplicado As Boolean

If Not IsNumeric(txtApertura.Text) Then Exit Sub

On Error GoTo vError

vAplicado = False

strSQL = "select count(*) as 'Existe'" _
       & " From CAJAS_APERTURAS_MAIN" _
       & " where  COD_CAJA = '" & txtCaja.Text & "' and cod_apertura = " & txtApertura.Text _
       & " and Estado = 'A'"
       
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 1 Then
  
  strSQL = "exec spCajas_Cierre_Forzado '" & txtCaja.Text & "'," & txtApertura.Text & ",'" & glogon.Usuario & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Aplica", "Cierre Forzado a Caja: " & txtCaja.Text & " AP: " & txtApertura.Text)
    
  vAplicado = True
End If
rs.Close

Me.MousePointer = vbDefault

If vAplicado Then
    MsgBox "Cierre de Caja: " & txtCaja.Text & " Apertura: " & txtApertura.Text & ", Realizado Satisfactoriamente!", vbInformation
    Call txtCaja_LostFocus
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

If tcMain.SelectedItem = 0 Then
    Call Excel_Exportar_Lsw(lsw, ProgressBarX)
Else
    Call Excel_Exportar_Lsw(lswH, ProgressBarX)
End If

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRecibe_Click()
Dim vAplicado As Boolean

If Not IsNumeric(txtApertura.Text) Then Exit Sub


If Trim(txtCierre_Fecha.Text) = "" Then
    MsgBox "No se ha realizado el cierre de esta caja, verifique!", vbInformation
    Exit Sub
End If

If Trim(txtRecibe_Fecha.Text) <> "" Then
    MsgBox "Ya fue aplica la recepción del cierre anteriormente, verifique!", vbInformation
    Exit Sub
End If


If Trim(txtCierre_Usuario.Text) = glogon.Usuario Then
    MsgBox "El usuario actual no puede recibir el cierre de esta caja porque es el mismo que la cerró, verifique!", vbInformation
    Exit Sub
End If


On Error GoTo vError

vAplicado = False

strSQL = "UPDATE CAJAS_APERTURAS_MAIN SET RECIBE_FECHA = getdate(), RECIBE_USUARIO = '" & glogon.Usuario & "'" _
       & " where  COD_CAJA = '" & txtCaja.Text & "' and cod_apertura = " & txtApertura.Text
Call ConectionExecute(strSQL)

If Not glogon.error Then
    Call Bitacora("Aplica", "Recepción de Cierre de Caja: " & txtCaja.Text & " AP: " & txtApertura.Text)
        
    vAplicado = True
End If

Me.MousePointer = vbDefault

If vAplicado Then
    MsgBox "Recepción de Cierre de Caja: " & txtCaja.Text & " Apertura: " & txtApertura.Text & ", Realizado Satisfactoriamente!", vbInformation
    
    Dim pApertura As Long, pEstatus As String
    pApertura = txtApertura.Text
    pEstatus = txtApertura.Tag
    
    Call txtCaja_LostFocus
    
    txtApertura.Text = pApertura
    txtApertura.Tag = txtApertura.Text
    
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRefresh_Click()
Call txtCaja_LostFocus
End Sub

Private Sub btnRevisa_Click()
Dim vAplicado As Boolean

If Not IsNumeric(txtApertura.Text) Then Exit Sub


If Trim(txtCierre_Fecha.Text) = "" Then
    MsgBox "No se ha realizado el cierre de esta caja, verifique!", vbInformation
    Exit Sub
End If

If Trim(txtRecibe_Fecha.Text) = "" Then
    MsgBox "NO se ha recibido este cierre, verifique!", vbInformation
    Exit Sub
End If

If Trim(txtRevisa_Fecha.Text) <> "" Then
    MsgBox "Este cierre ya fue Revisado anteriormente, verifique!", vbInformation
    Exit Sub
End If


'Todo: Revisar Requerimiento si el usuario que Recibe es el mismo que Revisa

If Trim(txtCierre_Usuario.Text) = glogon.Usuario Then
    MsgBox "El usuario actual no puede revisar el cierre de esta caja porque es el mismo que la cerró, verifique!", vbInformation
    Exit Sub
End If


On Error GoTo vError

vAplicado = False

strSQL = "UPDATE CAJAS_APERTURAS_MAIN SET REVISA_FECHA = getdate(), REVISA_USUARIO = '" & glogon.Usuario & "'" _
       & " where  COD_CAJA = '" & txtCaja.Text & "' and cod_apertura = " & txtApertura.Text
Call ConectionExecute(strSQL)

If Not glogon.error Then
    Call Bitacora("Aplica", "Revisión de Cierre de Caja: " & txtCaja.Text & " AP: " & txtApertura.Text)
        
    vAplicado = True
End If

Me.MousePointer = vbDefault

If vAplicado Then
    MsgBox "Revisión de Cierre de Caja: " & txtCaja.Text & " Apertura: " & txtApertura.Text & ", Realizado Satisfactoriamente!", vbInformation
    
    Dim pApertura As Long, pEstatus As String
    pApertura = txtApertura.Text
    pEstatus = txtApertura.Tag
    
    Call txtCaja_LostFocus
    
    txtApertura.Text = pApertura
    txtApertura.Tag = txtApertura.Text
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Sub Form_Load()
 vModulo = 5
 
 Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture
 
 tcMain.Item(0).Selected = True
 
 cboFiltro.Clear
 cboFiltro.AddItem "Todos"
 cboFiltro.AddItem "Pendientes de Recibir"
 cboFiltro.AddItem "Pendientes de Revisar"
 cboFiltro.Text = "Todos"
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "No.Apertura", 1500
    .Add , , "Fecha", 2100
    .Add , , "Usuario", 1850, vbCenter
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Cierre [Fecha]", 2100
    .Add , , "Cierre [Usuario]", 1850, vbCenter
  
    .Add , , "Recibe [Fecha]", 2100
    .Add , , "Recibe [Usuario]", 1850, vbCenter
  
    .Add , , "Revisa [Fecha]", 2100
    .Add , , "Revisa [Usuario]", 1850, vbCenter
  
  
  End With
 
 
 With lswH.ColumnHeaders
    .Clear
    .Add , , "Fecha", 2100
    .Add , , "Caja", 1200, vbCenter
    .Add , , "No.Apertura", 1500, vbCenter
    .Add , , "Usuario", 1850, vbCenter
    .Add , , "Versión", 3100, vbCenter
  End With
  
  

With lswDP.ColumnHeaders
  .Clear
  .Add , , "No.Depósito", 2500
  .Add , , "Monto", 2500, vbRightJustify
  .Add , , "Estado", 1800, vbCenter
  .Add , , "Cuenta", 2500
  .Add , , "Banco", 2500
End With
  
  
  dtpCorte.Value = fxFechaServidor
  dtpInicio.Value = DateAdd("d", -90, dtpCorte.Value)
 
 Call Formularios(Me)
 
 btnRecibe.Tag = btnCierre.Tag
 btnRevisa.Tag = btnCierre.Tag
 
 Call RefrescaTags(Me)
End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count = 0 Or vPaso Then Exit Sub

txtApertura.Text = Item.Text
txtApertura.Tag = Mid(Item.SubItems(3), 1, 1)
End Sub





Private Sub tblAplicar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String

'Modo de Supervision
If chkModoSupervisor.Value = vbChecked Then
   vCierreCiego = False
End If

If Not IsNumeric(txtApertura.Text) Then Exit Sub

'Aplica un Cierre Preliminar de los datos para ver el informe
If txtApertura.Tag = "Abierta" Then
   strSQL = "exec spCajas_CierreCajaMain '" & txtCaja.Text & "'," & txtApertura.Text _
       & ",'" & glogon.Usuario & "',1"
   Call ConectionExecute(strSQL)
End If

Select Case ButtonMenu.Key
  Case "Resumen"
     Call sbCajasCierreReportes(txtCaja.Text, txtApertura.Text, "Resumen", vCierreCiego)
  Case "Cierre"
     Call sbCajasCierreReportes(txtCaja.Text, txtApertura.Text, "Cierre", vCierreCiego)
  Case "Movimientos"
     Call sbCajasCierreReportes(txtCaja.Text, txtApertura.Text, "Movimientos", vCierreCiego)
End Select

End Sub

Private Sub sbApertura_Consulta()

If Not IsNumeric(txtApertura.Text) Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

txtApertura.Tag = "X"

strSQL = "select *" _
       & " From CAJAS_APERTURAS_MAIN" _
       & " where  COD_CAJA = '" & txtCaja.Text & "' and cod_apertura = " & txtApertura.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtApertura.Text = rs!COD_APERTURA
    txtApertura.Tag = rs!Estado
Else
    MsgBox "La apertura consultada (No." & txtApertura.Text & ")no existe verifique!", vbExclamation
    txtApertura.Text = 0
    txtApertura.Tag = "X"
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub


Private Sub sbVerificacion()
Dim curMonto As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

txtCierre_Fecha.Text = ""
txtCierre_Usuario.Text = ""

txtRecibe_Fecha.Text = ""
txtRecibe_Usuario.Text = ""

txtRevisa_Fecha.Text = ""
txtRevisa_Usuario.Text = ""

strSQL = "select M.*, C.DESCRIPCION" _
       & " From CAJAS_APERTURAS_MAIN M inner join cajas_definicion C on M.COD_CAJA = C.COD_CAJA" _
       & " Where M.COD_CAJA = '" & txtCaja.Text & "' and M.COD_APERTURA = " & txtApertura.Text
Call OpenRecordSet(rs, strSQL)
 
txtCierre_Caja.Text = rs!COD_CAJA
txtCierre_CajaDesc.Text = rs!DESCRIPCION
 
txtCierre_Id.Text = rs!COD_APERTURA
txtCierre_Estado.Text = IIf((rs!Estado = "C"), "Cerrada", "Abierta")
 
txtApertura_Fecha.Text = rs!APERTURA_FECHA & ""
txtApertura_Usuario.Text = rs!APERTURA_USUARIO & ""

txtCierre_Fecha.Text = rs!CIERRE_FECHA & ""
txtCierre_Usuario.Text = rs!CIERRE_USUARIO & ""

txtRecibe_Fecha.Text = rs!Recibe_Fecha & ""
txtRecibe_Usuario.Text = rs!Recibe_Usuario & ""

txtRevisa_Fecha.Text = rs!REVISA_FECHA & ""
txtRevisa_Usuario.Text = rs!REVISA_USUARIO & ""
 
rs.Close



strSQL = "exec spCajas_CierreDepositoDivisa '" & txtCierre_Caja.Text & "'," & txtCierre_Id.Text & ", 'COL'"

With lswDP.ListItems

    Call OpenRecordSet(rs, strSQL)
    curMonto = 0
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!DP_Numero)
            itmX.SubItems(1) = Format(rs!Monto, "Standard")
            
        If rs!Estado = 1 Or rs!Estado = 2 Then
           curMonto = curMonto + rs!Monto
            itmX.SubItems(2) = IIf((rs!Estado = 1), "Activado", "En Bancos")
        Else
            itmX.SubItems(2) = "Anulado"
        End If
        
        itmX.SubItems(3) = Trim(rs!DP_Cuenta)
        itmX.SubItems(4) = rs!BancoDesc

       rs.MoveNext
    
    Loop
    rs.Close

    txtTotalDeposito.Text = Format(curMonto, "Standard")

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 2 Then
   If IsNumeric(txtApertura.Text) Then
    Call sbVerificacion
   Else
        MsgBox "Consulte una Apertura Primero!", vbExclamation
        tcMain.Item(0).Selected = True
   End If
End If

End Sub

Private Sub txtApertura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtCajaDesc.SetFocus
End Sub

Private Sub txtApertura_LostFocus()
 Call sbApertura_Consulta
End Sub

Private Sub txtCaja_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtCajaDesc.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado2 = ""
    gBusquedas.Resultado = ""
    txtCaja = ""
    txtCajaDesc = ""
    gBusquedas.Columna = "cod_caja"
    gBusquedas.Orden = "cod_caja"
    gBusquedas.Consulta = "Select cod_caja,descripcion From cajas_definicion"
    
    frmBusquedas.Show vbModal
    
    txtCaja.Text = Trim(gBusquedas.Resultado)
    txtCajaDesc.Text = gBusquedas.Resultado2
        
    If gBusquedas.Resultado <> "" Then txtCajaDesc.SetFocus
    
End If


End Sub

Private Sub txtCaja_LostFocus()

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear
lswH.ListItems.Clear

txtApertura.Text = 0
txtApertura.Tag = "X"
vCierreCiego = True

strSQL = "select descripcion,cierre_tipo from cajas_Definicion where cod_Caja = '" & txtCaja.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  If rs!Cierre_Tipo = "A" Then vCierreCiego = False
  txtCajaDesc.Text = rs!DESCRIPCION
End If
rs.Close


strSQL = "select *" _
       & " From CAJAS_APERTURAS_MAIN" _
       & " where  COD_CAJA = '" & txtCaja.Text & "'" _
       & " and Apertura_Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
       
Select Case cboFiltro.Text
    Case "Pendientes de Recibir"
        strSQL = strSQL & " and Recibe_Fecha is null"
    Case "Pendientes de Revisar"
        strSQL = strSQL & " and Revisa_Fecha is null"
End Select

strSQL = strSQL & " order by COD_APERTURA desc"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!COD_APERTURA)
      itmX.SubItems(1) = rs!APERTURA_FECHA
      itmX.SubItems(2) = rs!APERTURA_USUARIO
      itmX.SubItems(3) = IIf((rs!Estado = "C"), "Cerrada", "Abierta")
      itmX.SubItems(4) = rs!CIERRE_FECHA & ""
      itmX.SubItems(5) = rs!CIERRE_USUARIO & ""
      
      itmX.SubItems(6) = rs!Recibe_Fecha & ""
      itmX.SubItems(7) = rs!Recibe_Usuario & ""
      
      itmX.SubItems(8) = rs!REVISA_FECHA & ""
      itmX.SubItems(9) = rs!REVISA_USUARIO & ""
      
      
  If txtApertura.Text = 0 Then
        txtApertura.Text = rs!COD_APERTURA
        txtApertura.Tag = rs!Estado
  End If
  rs.MoveNext
Loop
rs.Close


'Historico de Accesos
strSQL = "select *" _
       & " From CAJAS_BITACORA_INGRESO" _
       & " where  Caja = '" & txtCaja.Text & "'" _
       & " and fechaIngreso between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
       & " order by FechaIngreso desc"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  Set itmX = lswH.ListItems.Add(, , rs!FechaIngreso)
      itmX.SubItems(1) = rs!Caja
      itmX.SubItems(2) = rs!Apertura
      itmX.SubItems(3) = rs!Usuario
      itmX.SubItems(4) = rs!SifVersion
  rs.MoveNext
Loop
rs.Close


tcMain.Item(0).Selected = True

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub txtCajaDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado2 = ""
    gBusquedas.Resultado = ""
    txtCaja = ""
    txtCajaDesc = ""
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "Select cod_caja,descripcion From cajas_definicion"
    
    frmBusquedas.Show vbModal
    
    txtCaja = Trim(gBusquedas.Resultado)
    txtCajaDesc = gBusquedas.Resultado2
    
    If gBusquedas.Resultado <> "" Then
       Call txtCaja_LostFocus
    End If
    
End If

End Sub
