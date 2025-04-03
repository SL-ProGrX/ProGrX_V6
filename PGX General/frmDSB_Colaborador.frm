VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.TaskPanel.v24.0.0.ocx"
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
      _Version        =   1572864
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
      _Version        =   1572864
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6135
      Left            =   2760
      TabIndex        =   37
      Top             =   1800
      Visible         =   0   'False
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   10821
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
      PaintManager.Position=   2
      ItemCount       =   7
      SelectedItem    =   5
      Item(0).Caption =   "Vacaciones"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "gbVacaciones"
      Item(0).Control(1)=   "scVacaciones"
      Item(1).Caption =   "Permisos"
      Item(1).Tooltip =   "Incapacidades"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "scPermiso"
      Item(1).Control(1)=   "gbPermisos"
      Item(2).Caption =   "Incapacidades"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "scIncapacidad"
      Item(2).Control(1)=   "gbIncapacidades"
      Item(3).Caption =   "Autorizaciones"
      Item(3).ControlCount=   8
      Item(3).Control(0)=   "tlbAutorizacion"
      Item(3).Control(1)=   "dtpAut_Inicio"
      Item(3).Control(2)=   "dtpAut_Corte"
      Item(3).Control(3)=   "cboAut_Autorizado"
      Item(3).Control(4)=   "cboAut_Tipo"
      Item(3).Control(5)=   "lswAut"
      Item(3).Control(6)=   "chkAut_Todos"
      Item(3).Control(7)=   "scAutorizaciones"
      Item(4).Caption =   "Traslados"
      Item(4).ControlCount=   4
      Item(4).Control(0)=   "tcTraslados"
      Item(4).Control(1)=   "scActivos"
      Item(4).Control(2)=   "lswAF_Activos"
      Item(4).Control(3)=   "vGrid"
      Item(5).Caption =   "Declaración"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "tcDeclaracion"
      Item(6).Caption =   "Compras"
      Item(6).ControlCount=   0
      Begin XtremeSuiteControls.ListView lswAF_Activos 
         Height          =   1335
         Left            =   -70000
         TabIndex        =   107
         Top             =   4440
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1572864
         _ExtentX        =   19288
         _ExtentY        =   2355
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   21
      End
      Begin XtremeSuiteControls.ListView lswAut 
         Height          =   3735
         Left            =   -70000
         TabIndex        =   94
         Top             =   720
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18230
         _ExtentY        =   6588
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
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcDeclaracion 
         Height          =   5655
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Width           =   11055
         _Version        =   1572864
         _ExtentX        =   19500
         _ExtentY        =   9975
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
         Item(0).Caption =   "Declaración"
         Item(0).ControlCount=   15
         Item(0).Control(0)=   "lswDeclara"
         Item(0).Control(1)=   "cboDeclara"
         Item(0).Control(2)=   "Label5(1)"
         Item(0).Control(3)=   "Label5(6)"
         Item(0).Control(4)=   "txtD_Placa"
         Item(0).Control(5)=   "Label5(7)"
         Item(0).Control(6)=   "Label5(8)"
         Item(0).Control(7)=   "txtD_Caracteristicas"
         Item(0).Control(8)=   "cboLocaliza"
         Item(0).Control(9)=   "Label5(22)"
         Item(0).Control(10)=   "btnAF_Declara(0)"
         Item(0).Control(11)=   "btnAF_Declara(1)"
         Item(0).Control(12)=   "btnAF_Declara(2)"
         Item(0).Control(13)=   "btnAF_Declara(3)"
         Item(0).Control(14)=   "txtD_Descripcion"
         Item(1).Caption =   "Histórico"
         Item(1).ControlCount=   4
         Item(1).Control(0)=   "lswDeclara_H"
         Item(1).Control(1)=   "lswDeclara_H_Det"
         Item(1).Control(2)=   "txtDeclara_H"
         Item(1).Control(3)=   "Label3(2)"
         Begin XtremeSuiteControls.ListView lswDeclara_H 
            Height          =   2655
            Left            =   0
            TabIndex        =   127
            Top             =   360
            Width           =   11055
            _Version        =   1572864
            _ExtentX        =   19500
            _ExtentY        =   4683
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.ListView lswDeclara_H_Det 
            Height          =   2175
            Left            =   0
            TabIndex        =   128
            Top             =   3480
            Width           =   11055
            _Version        =   1572864
            _ExtentX        =   19500
            _ExtentY        =   3836
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.ListView lswDeclara 
            Height          =   1935
            Left            =   -70000
            TabIndex        =   131
            Top             =   3720
            Visible         =   0   'False
            Width           =   11055
            _Version        =   1572864
            _ExtentX        =   19500
            _ExtentY        =   3413
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.FlatEdit txtDeclara_H 
            Height          =   330
            Left            =   1800
            TabIndex        =   129
            Top             =   3120
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.ComboBox cboDeclara 
            Height          =   315
            Left            =   -67600
            TabIndex        =   132
            Top             =   480
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtD_Placa 
            Height          =   330
            Left            =   -67600
            TabIndex        =   135
            Top             =   960
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtD_Descripcion 
            Height          =   330
            Left            =   -67600
            TabIndex        =   137
            Top             =   1320
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtD_Caracteristicas 
            Height          =   810
            Left            =   -67600
            TabIndex        =   139
            Top             =   1680
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
            _ExtentY        =   1429
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
         Begin XtremeSuiteControls.ComboBox cboLocaliza 
            Height          =   315
            Left            =   -67600
            TabIndex        =   140
            Top             =   2520
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.PushButton btnAF_Declara 
            Height          =   375
            Index           =   0
            Left            =   -67600
            TabIndex        =   142
            Top             =   3000
            Visible         =   0   'False
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   79
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnAF_Declara 
            Height          =   375
            Index           =   1
            Left            =   -67000
            TabIndex        =   143
            Top             =   3000
            Visible         =   0   'False
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   79
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":0632
         End
         Begin XtremeSuiteControls.PushButton btnAF_Declara 
            Height          =   375
            Index           =   2
            Left            =   -66280
            TabIndex        =   144
            Top             =   3000
            Visible         =   0   'False
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   79
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":0D63
         End
         Begin XtremeSuiteControls.PushButton btnAF_Declara 
            Height          =   375
            Index           =   3
            Left            =   -64120
            TabIndex        =   145
            Top             =   3000
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Procesar Declaración"
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":1307
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   22
            Left            =   -69520
            TabIndex        =   141
            Top             =   2520
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Localización"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   8
            Left            =   -69520
            TabIndex        =   138
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Otras Características"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   7
            Left            =   -69520
            TabIndex        =   136
            Top             =   1320
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Descripción"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   6
            Left            =   -69520
            TabIndex        =   134
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "No. Placa"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   1
            Left            =   -69520
            TabIndex        =   133
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "No. Declaración"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   130
            Top             =   3120
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No. Declaración"
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
      Begin XtremeSuiteControls.TabControl tcTraslados 
         Height          =   3855
         Left            =   -70000
         TabIndex        =   99
         Top             =   360
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1572864
         _ExtentX        =   19288
         _ExtentY        =   6800
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
         SelectedItem    =   1
         Item(0).Caption =   "Boletas"
         Item(0).ControlCount=   6
         Item(0).Control(0)=   "Label3(0)"
         Item(0).Control(1)=   "txtAF_Boleta"
         Item(0).Control(2)=   "txtAF_Boleta_Estado"
         Item(0).Control(3)=   "btnAF_Boleta(0)"
         Item(0).Control(4)=   "btnAF_Boleta(1)"
         Item(0).Control(5)=   "lswAF_Boletas"
         Item(1).Caption =   "Recibir"
         Item(1).ControlCount=   5
         Item(1).Control(0)=   "Label3(1)"
         Item(1).Control(1)=   "txtAF_Boleta_R"
         Item(1).Control(2)=   "btnAF_Boleta(2)"
         Item(1).Control(3)=   "btnAF_Boleta(3)"
         Item(1).Control(4)=   "lswAF_Recepcion"
         Item(2).Caption =   "Trasladar"
         Item(2).ControlCount=   12
         Item(2).Control(0)=   "txtNuevoPersona"
         Item(2).Control(1)=   "txtNuevoDepartamento"
         Item(2).Control(2)=   "txtNuevoSeccion"
         Item(2).Control(3)=   "Label5(0)"
         Item(2).Control(4)=   "Label5(2)"
         Item(2).Control(5)=   "Label5(3)"
         Item(2).Control(6)=   "cboMotivo"
         Item(2).Control(7)=   "txtNotas"
         Item(2).Control(8)=   "Label5(4)"
         Item(2).Control(9)=   "Label5(5)"
         Item(2).Control(10)=   "btnAF_Boleta(4)"
         Item(2).Control(11)=   "scActivos_Trasladar"
         Begin XtremeSuiteControls.ListView lswAF_Recepcion 
            Height          =   3015
            Left            =   0
            TabIndex        =   108
            Top             =   360
            Width           =   10935
            _Version        =   1572864
            _ExtentX        =   19288
            _ExtentY        =   5318
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.ListView lswAF_Boletas 
            Height          =   3015
            Left            =   -70000
            TabIndex        =   101
            Top             =   360
            Visible         =   0   'False
            Width           =   10935
            _Version        =   1572864
            _ExtentX        =   19288
            _ExtentY        =   5318
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnAF_Boleta 
            Height          =   375
            Index           =   0
            Left            =   -64840
            TabIndex        =   105
            Top             =   3480
            Visible         =   0   'False
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1296
            _ExtentY        =   661
            _StockProps     =   79
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":1A20
         End
         Begin XtremeSuiteControls.FlatEdit txtAF_Boleta 
            Height          =   330
            Left            =   -68800
            TabIndex        =   103
            Top             =   3480
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtAF_Boleta_Estado 
            Height          =   330
            Left            =   -66880
            TabIndex        =   104
            Top             =   3480
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
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
         Begin XtremeSuiteControls.PushButton btnAF_Boleta 
            Height          =   375
            Index           =   1
            Left            =   -64120
            TabIndex        =   106
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Descartar"
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":2127
         End
         Begin XtremeSuiteControls.FlatEdit txtAF_Boleta_R 
            Height          =   330
            Left            =   1200
            TabIndex        =   110
            Top             =   3480
            Width           =   1935
            _Version        =   1572864
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
         Begin XtremeSuiteControls.PushButton btnAF_Boleta 
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   111
            Top             =   3480
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1296
            _ExtentY        =   661
            _StockProps     =   79
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":26CB
         End
         Begin XtremeSuiteControls.PushButton btnAF_Boleta 
            Height          =   375
            Index           =   3
            Left            =   3960
            TabIndex        =   112
            Top             =   3480
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Aceptar"
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":2DD2
         End
         Begin XtremeSuiteControls.FlatEdit txtNuevoPersona 
            Height          =   315
            Left            =   -67000
            TabIndex        =   113
            Top             =   960
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11028
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNuevoDepartamento 
            Height          =   315
            Left            =   -67000
            TabIndex        =   114
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1320
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11028
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNuevoSeccion 
            Height          =   315
            Left            =   -67000
            TabIndex        =   115
            Top             =   1680
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11028
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboMotivo 
            Height          =   315
            Left            =   -67000
            TabIndex        =   120
            Top             =   2160
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   675
            Left            =   -67000
            TabIndex        =   121
            Top             =   2520
            Visible         =   0   'False
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
            _ExtentY        =   1191
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
         Begin XtremeSuiteControls.PushButton btnAF_Boleta 
            Height          =   375
            Index           =   4
            Left            =   -63400
            TabIndex        =   125
            Top             =   3360
            Visible         =   0   'False
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Crear Solicitud de Traslado"
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
            Appearance      =   21
            Picture         =   "frmDSB_Colaborador.frx":34F9
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   5
            Left            =   -68920
            TabIndex        =   123
            Top             =   2160
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Motivo"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   4
            Left            =   -68920
            TabIndex        =   122
            Top             =   2520
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
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
         Begin XtremeShortcutBar.ShortcutCaption scActivos_Trasladar 
            Height          =   375
            Left            =   -70000
            TabIndex        =   119
            Top             =   360
            Visible         =   0   'False
            Width           =   11295
            _Version        =   1572864
            _ExtentX        =   19918
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Trasladar a .:"
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
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   3
            Left            =   -68920
            TabIndex        =   118
            Top             =   1320
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Departamento"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   2
            Left            =   -68920
            TabIndex        =   117
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Sección"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Index           =   0
            Left            =   -68920
            TabIndex        =   116
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Persona"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   109
            Top             =   3480
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Boleta No."
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   0
            Left            =   -69880
            TabIndex        =   102
            Top             =   3480
            Visible         =   0   'False
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Boleta No."
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
      Begin XtremeSuiteControls.GroupBox gbVacaciones 
         Height          =   3855
         Left            =   -70000
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   9495
         _Version        =   1572864
         _ExtentX        =   16748
         _ExtentY        =   6800
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.RadioButton rbVac_Accion 
            Height          =   252
            Index           =   0
            Left            =   2160
            TabIndex        =   39
            Top             =   1680
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Disfrutar"
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
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnVac_Aplicar 
            Height          =   612
            Left            =   7440
            TabIndex        =   40
            Top             =   3000
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmDSB_Colaborador.frx":3C12
         End
         Begin XtremeSuiteControls.ComboBox cboVac_Tipo 
            Height          =   312
            Left            =   2160
            TabIndex        =   41
            Top             =   240
            Width           =   6852
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtVac_Notas 
            Height          =   912
            Left            =   2160
            TabIndex        =   42
            Top             =   600
            Width           =   6852
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpVac_FechaI 
            Height          =   315
            Left            =   2160
            TabIndex        =   43
            Top             =   2160
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboVac_Estado 
            Height          =   330
            Left            =   2160
            TabIndex        =   44
            Top             =   3120
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
         End
         Begin XtremeSuiteControls.DateTimePicker dtpVac_FechaC 
            Height          =   315
            Left            =   3480
            TabIndex        =   45
            Top             =   2160
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtVac_Dias 
            Height          =   315
            Left            =   2160
            TabIndex        =   46
            ToolTipText     =   "Dias a Disfrutar"
            Top             =   2640
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.FlatEdit txtVac_DiasDisponibles 
            Height          =   315
            Left            =   3480
            TabIndex        =   47
            ToolTipText     =   "Días Disponibles"
            Top             =   2640
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.RadioButton rbVac_Accion 
            Height          =   252
            Index           =   1
            Left            =   3840
            TabIndex        =   48
            Top             =   1680
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Liquidar"
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
            Appearance      =   16
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
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
            Height          =   372
            Index           =   3
            Left            =   1080
            TabIndex        =   53
            Top             =   3120
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
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
            Left            =   1080
            TabIndex        =   52
            Top             =   2160
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
            Height          =   372
            Index           =   0
            Left            =   1080
            TabIndex        =   51
            Top             =   240
            Width           =   1092
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
            Left            =   1080
            TabIndex        =   50
            Top             =   600
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Días"
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
            Index           =   2
            Left            =   1080
            TabIndex        =   49
            Top             =   2640
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox gbPermisos 
         Height          =   3375
         Left            =   -70000
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   9495
         _Version        =   1572864
         _ExtentX        =   16748
         _ExtentY        =   5953
         _StockProps     =   79
         BackColor       =   16777215
         Appearance      =   16
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnPer_Aplicar 
            Height          =   612
            Left            =   7200
            TabIndex        =   55
            Top             =   2520
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmDSB_Colaborador.frx":43EA
         End
         Begin XtremeSuiteControls.ComboBox cboPer_Tipo 
            Height          =   312
            Left            =   2160
            TabIndex        =   56
            Top             =   240
            Width           =   6852
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPer_Notas 
            Height          =   912
            Left            =   2160
            TabIndex        =   57
            Top             =   600
            Width           =   6852
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpPer_Fecha 
            Height          =   315
            Left            =   2160
            TabIndex        =   58
            Top             =   1680
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.DateTimePicker dtpPer_HoraI 
            Height          =   315
            Left            =   2160
            TabIndex        =   59
            Top             =   2160
            Width           =   1335
            _Version        =   1572864
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
            Format          =   2
         End
         Begin XtremeSuiteControls.DateTimePicker dtpPer_HoraC 
            Height          =   315
            Left            =   3480
            TabIndex        =   60
            Top             =   2160
            Width           =   1335
            _Version        =   1572864
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
            Format          =   2
         End
         Begin XtremeSuiteControls.ComboBox cboPer_Estado 
            Height          =   312
            Left            =   6600
            TabIndex        =   61
            Top             =   1680
            Width           =   2412
            _Version        =   1572864
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.FlatEdit txtPer_Horas 
            Height          =   315
            Left            =   2160
            TabIndex        =   62
            ToolTipText     =   "Dias a Disfrutar"
            Top             =   2640
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.FlatEdit txtPer_HrsMax 
            Height          =   315
            Left            =   3480
            TabIndex        =   63
            ToolTipText     =   "Dias a Disfrutar"
            Top             =   2640
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
            Index           =   10
            Left            =   1080
            TabIndex        =   69
            Top             =   600
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
            Height          =   372
            Index           =   9
            Left            =   1080
            TabIndex        =   68
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
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
            Left            =   1080
            TabIndex        =   67
            Top             =   1680
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Horas"
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
            Left            =   1080
            TabIndex        =   66
            Top             =   2640
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
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
            Height          =   372
            Index           =   6
            Left            =   5520
            TabIndex        =   65
            Top             =   1680
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Corte"
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
            Left            =   1080
            TabIndex        =   64
            Top             =   2160
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox gbIncapacidades 
         Height          =   3375
         Left            =   -70000
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   5953
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnInc_Aplicar 
            Height          =   612
            Left            =   7440
            TabIndex        =   71
            Top             =   2400
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmDSB_Colaborador.frx":4BC2
         End
         Begin XtremeSuiteControls.ComboBox cboInc_Tipo 
            Height          =   312
            Left            =   2160
            TabIndex        =   72
            Top             =   240
            Width           =   6852
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInc_Notas 
            Height          =   912
            Left            =   2160
            TabIndex        =   73
            Top             =   600
            Width           =   6852
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpInc_FechaI 
            Height          =   315
            Left            =   2160
            TabIndex        =   74
            Top             =   1680
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboInc_Estado 
            Height          =   312
            Left            =   6840
            TabIndex        =   75
            Top             =   1680
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.DateTimePicker dtpInc_FechaC 
            Height          =   315
            Left            =   2160
            TabIndex        =   76
            Top             =   2040
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtInc_Dias 
            Height          =   315
            Left            =   4800
            TabIndex        =   77
            ToolTipText     =   "Dias a Disfrutar"
            Top             =   1680
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
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
         Begin XtremeSuiteControls.FlatEdit txtInc_Porcentaje 
            Height          =   315
            Left            =   4800
            TabIndex        =   78
            ToolTipText     =   "Dias a Disfrutar"
            Top             =   2040
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
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
            Index           =   17
            Left            =   1080
            TabIndex        =   85
            Top             =   600
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
            Height          =   372
            Index           =   16
            Left            =   1080
            TabIndex        =   84
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
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
            Left            =   1080
            TabIndex        =   83
            Top             =   1680
            Width           =   1092
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Height          =   372
            Index           =   14
            Left            =   5640
            TabIndex        =   82
            Top             =   1680
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Vence"
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
            Left            =   1080
            TabIndex        =   81
            Top             =   2040
            Width           =   1092
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Días"
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
            Index           =   12
            Left            =   3720
            TabIndex        =   80
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Porc. Patrono"
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
            Index           =   11
            Left            =   3720
            TabIndex        =   79
            Top             =   2040
            Width           =   1095
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9000
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Colaborador.frx":539A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Colaborador.frx":54B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Colaborador.frx":55E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Colaborador.frx":5708
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Colaborador.frx":5811
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAutorizacion 
         Height          =   330
         Left            =   -62680
         TabIndex        =   89
         Top             =   0
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar "
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Autorizar"
               Object.ToolTipText     =   "Autorizar Casos Marcados"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Desautorizar"
               Object.ToolTipText     =   "Desautorizar casos marcados"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Reporte"
               Object.ToolTipText     =   "Exporta a Excel"
               ImageIndex      =   5
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpAut_Inicio 
         Height          =   312
         Left            =   -67720
         TabIndex        =   90
         Top             =   0
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
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
      Begin XtremeSuiteControls.DateTimePicker dtpAut_Corte 
         Height          =   312
         Left            =   -66400
         TabIndex        =   91
         Top             =   0
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboAut_Autorizado 
         Height          =   315
         Left            =   -65080
         TabIndex        =   92
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.ComboBox cboAut_Tipo 
         Height          =   312
         Left            =   -70000
         TabIndex        =   93
         Top             =   0
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.CheckBox chkAut_Todos 
         Height          =   210
         Left            =   -69880
         TabIndex        =   95
         Top             =   480
         Visible         =   0   'False
         Width           =   210
         _Version        =   1572864
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
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
         Alignment       =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   855
         Left            =   -65320
         TabIndex        =   124
         Top             =   4200
         Visible         =   0   'False
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   1508
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmDSB_Colaborador.frx":66EB
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scActivos 
         Height          =   375
         Left            =   -70000
         TabIndex        =   100
         Top             =   0
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Control de Activos (Recepción y Traslado)"
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
      Begin XtremeShortcutBar.ShortcutCaption scAutorizaciones 
         Height          =   375
         Left            =   -70000
         TabIndex        =   96
         Top             =   360
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "                                Seleccione las Solicitudes  a Autorizar o Desautorizar"
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
      Begin XtremeShortcutBar.ShortcutCaption scVacaciones 
         Height          =   375
         Left            =   -70000
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   9495
         _Version        =   1572864
         _ExtentX        =   16748
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Solicitud de Vacaciones"
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
      Begin XtremeShortcutBar.ShortcutCaption scPermiso 
         Height          =   375
         Left            =   -70000
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   9495
         _Version        =   1572864
         _ExtentX        =   16748
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Solicitud de Permiso de ausencia"
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
      Begin XtremeShortcutBar.ShortcutCaption scIncapacidad 
         Height          =   375
         Left            =   -70000
         TabIndex        =   86
         Top             =   0
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Solicitud de Incapacidad"
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
   Begin XtremeSuiteControls.GroupBox gbCambioClave 
      Height          =   2175
      Left            =   5520
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   9135
      _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         Picture         =   "frmDSB_Colaborador.frx":6EC2
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnClaveCambio 
         Height          =   495
         Index           =   1
         Left            =   7680
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
         _Version        =   1572864
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
         Picture         =   "frmDSB_Colaborador.frx":75E9
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
      _Version        =   1572864
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
         _Version        =   1572864
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
         Picture         =   "frmDSB_Colaborador.frx":7C27
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.CheckBox chkLoginVincular 
         Height          =   615
         Left            =   4560
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         Picture         =   "frmDSB_Colaborador.frx":834E
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnConectar 
         Height          =   495
         Index           =   2
         Left            =   7680
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
         _Version        =   1572864
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
         Picture         =   "frmDSB_Colaborador.frx":8A4E
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboLoginAccion 
         Height          =   330
         Left            =   1920
         TabIndex        =   27
         Top             =   1440
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
         _Version        =   1572864
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
      _Version        =   1572864
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
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   1215
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
      _Version        =   1572864
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
      _Version        =   1572864
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
      _Version        =   1572864
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
      Left            =   12360
      TabIndex        =   32
      Top             =   1395
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
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
      Picture         =   "frmDSB_Colaborador.frx":9164
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
            Picture         =   "frmDSB_Colaborador.frx":975F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":99D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":9C6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":9DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":9F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":A130
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":A2D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":A45D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":A5E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":A6EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":A97B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":AA87
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":AD03
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":ADC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":AF69
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSB_Colaborador.frx":B105
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
      _Version        =   1572864
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
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   345
      Left            =   13560
      TabIndex        =   97
      Top             =   1395
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   609
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmDSB_Colaborador.frx":B1B0
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   98
      Top             =   1200
      Visible         =   0   'False
      Width           =   14655
      _Version        =   1572864
      _ExtentX        =   25850
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
   End
   Begin XtremeShortcutBar.ShortcutCaption TituloOpciones 
      Height          =   480
      Left            =   2760
      TabIndex        =   33
      Top             =   1320
      Width           =   12615
      _Version        =   1572864
      _ExtentX        =   22251
      _ExtentY        =   847
      _StockProps     =   14
      Caption         =   "Detalles:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.01
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
      _Version        =   1572864
      _ExtentX        =   4895
      _ExtentY        =   847
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.51
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
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Identificación"
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
      _Version        =   1572864
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

Dim mNomina As String, mId_Task_Seleted As Integer

Dim vPaso As Boolean, mTop As Long, mLeft As Long
Dim itmX As ListViewItem, mWidth As Long, mHeight As Long



Const Id_TaskItem_DatosPersonales = 0
Const Id_TaskItem_RelacionLaboral = 1
Const Id_TaskItem_Otros_Add = 2

Const Id_TaskItem_Telefonos = 3
Const Id_TaskItem_Tarjetas = 18

Const Id_TaskItem_Familiares = 4
Const Id_TaskItem_Cuentas = 5

Const Id_TaskItem_Boletas_Pago = 6
Const Id_TaskItem_Plan_Carrera = 7
Const Id_TaskItem_Vacaciones = 8
Const Id_TaskItem_Permisos = 9
Const Id_TaskItem_Incapacidades = 10

Const Id_TaskItem_Autorizaciones = 11
Const Id_TaskItem_Accion_Personal = 17

Const Id_TaskItem_Activos = 20
Const Id_TaskItem_Activos_Traslados = 21
Const Id_TaskItem_Activos_Declaracion = 22





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
   
    
    Set Group = tpMain.Groups.Add(0, "RRHH")
    Group.Expanded = True
    Group.Items.Add Id_TaskItem_Boletas_Pago, "Boletas de Pago", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Vacaciones, "Vacaciones", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Incapacidades, "Incapacidades", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Permisos, "Permisos", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Accion_Personal, "Acciones de Personal", xtpTaskItemTypeLink, 3
    
    Group.Items.Add Id_TaskItem_Autorizaciones, "Autorizaciones", xtpTaskItemTypeLink, 3
    
    Set Group = tpMain.Groups.Add(0, "Activos")
    Group.Expanded = True
    Group.Items.Add Id_TaskItem_Activos, "Activos", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Activos_Traslados, "Traslados/Recepción", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Activos_Declaracion, "Declaración", xtpTaskItemTypeLink, 3
   
    
'    Group.Items.Add Id_TaskItem_Plan_Carrera, "Plan de Carrera", xtpTaskItemTypeLink, 3
    
   
    tpMain.SetImageList imlTaskPanelIcons
    
'    tpMain.SetMargins 5, 5, 5, 5, 5
    
    mId_Task_Seleted = 0

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


mId_Task_Seleted = ItemId
btnExport.Visible = True
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
lsw.Visible = True
tcMain.Visible = False

btnEditarDetalle.Visible = True




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
           itmX.SubItems(2) = Trim(rs!Parentesco_Desc)
       
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
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           itmX.SubItems(8) = rs!Registro_Usuario & ""
     
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
          itmX.SubItems(4) = Format(rs!fecha_corte, "yyyy-mm-dd")
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
           itmX.SubItems(17) = rs!Notas & ""
           itmX.SubItems(18) = rs!Registro_Fecha & ""
           itmX.SubItems(19) = Trim(rs!Registro_Usuario & "")

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
           itmX.SubItems(3) = Format(rs!fecha_entrada, "dd/mm/yyyy")
           itmX.SubItems(4) = rs!Dias_Disfrutados & ""
           itmX.SubItems(5) = rs!Estado_Transaccion
           itmX.SubItems(6) = rs!Registro_Usuario & ""
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           
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
           itmX.SubItems(3) = Format(rs!fecha_entrada, "dd/mm/yyyy")
           itmX.SubItems(4) = rs!Dias & ""
           itmX.SubItems(5) = rs!Estado_Transaccion
           itmX.SubItems(6) = rs!Registro_Usuario & ""
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           
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
           itmX.SubItems(7) = rs!Registro_Usuario & ""
           itmX.SubItems(8) = rs!Registro_Fecha & ""
           
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


    Case Id_TaskItem_Autorizaciones
        TituloOpciones.Caption = "Transacciones Pendientes de Autorización..:"
        TituloOpciones.Tag = "Autorizaciones"
        
        Call sbOpcionesVisibles(False)
        
        Call sbAutorizaciones_Inicial
        
        scTitulo1.Visible = True
        TituloOpciones.Visible = True
        tpMain.Visible = True
        
        btnMenu.Item(0).Visible = True
        btnMenu.Item(1).Visible = True


    Case Id_TaskItem_Activos
        TituloOpciones.Caption = "Activos a su Nombre..:"
        TituloOpciones.Tag = "Activos"
    
        Call sbOpcionesVisibles(False)

        scTitulo1.Visible = True
        TituloOpciones.Visible = True
        tpMain.Visible = True
        lsw.Visible = True
        
        
        
        lsw.ListItems.Clear
        lsw.ColumnHeaders.Clear
        lsw.ColumnHeaders.Add , , "Placa", 1400
        lsw.ColumnHeaders.Add , , "Id Alterna", 1400, vbCenter
        lsw.ColumnHeaders.Add , , "Nombre", 4000
        lsw.ColumnHeaders.Add , , "Fecha Adq.", 1400
        lsw.ColumnHeaders.Add , , "Fecha Inst.", 1400
        lsw.ColumnHeaders.Add , , "Tipo", 2400
        lsw.ColumnHeaders.Add , , "Vida Util", 1400, 1
        lsw.ColumnHeaders.Add , , "Valor historico", 1400, 1
        lsw.ColumnHeaders.Add , , "Valor Rescate", 1400, 1
        lsw.ColumnHeaders.Add , , "Estado", 1400, vbCenter
        
        lsw.ColumnHeaders.Add , , "Id Responsable", 1400, vbCenter
        lsw.ColumnHeaders.Add , , "Responsable", 2500
        lsw.ColumnHeaders.Add , , "Departamento", 2400
        lsw.ColumnHeaders.Add , , "Sección", 2400
        lsw.ColumnHeaders.Add , , "Localización", 2400
        lsw.ColumnHeaders.Add , , "Proveedor", 2400
    
        lsw.ColumnHeaders.Add , , "Modelo", 2400, vbCenter
        lsw.ColumnHeaders.Add , , "Marca", 2400, vbCenter
        lsw.ColumnHeaders.Add , , "No. Serie", 2400, vbCenter
        lsw.ColumnHeaders.Add , , "Otras Señas", 3000
    
       Dim curVH As Currency, curVR As Currency
       curVH = 0
       curVR = 0
        
        strSQL = "select * " _
               & " from vActivos_General" _
               & " where Estado = 'A' and Identificacion = '" & txtIdentificacion.Text & "'" _
               & " order by  Tipo_Activo, num_placa"
    
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
         Set itmX = lsw.ListItems.Add(, , rs!NUM_PLACA)
             itmX.SubItems(1) = rs!Placa_Alterna & ""
             itmX.SubItems(2) = rs!Nombre
             itmX.SubItems(3) = Format(rs!fecha_adquisicion, "yyyy-mm-dd")
             itmX.SubItems(4) = Format(rs!fecha_instalacion, "yyyy-mm-dd")
             itmX.SubItems(5) = rs!Tipo_Activo_Desc
             itmX.SubItems(6) = rs!vida_util & IIf((rs!VIDA_UTIL_EN = "A"), " Año(s)", " Mes(es)")
             itmX.SubItems(7) = Format(rs!Valor_Historico, "Standard")
             itmX.SubItems(8) = Format(rs!Valor_Desecho, "Standard")
             itmX.SubItems(9) = rs!Estado_Desc
             
             
                itmX.SubItems(10) = rs!Identificacion
                itmX.SubItems(11) = rs!Responsable
                itmX.SubItems(12) = rs!departamento
                itmX.SubItems(13) = rs!seccion
                itmX.SubItems(14) = rs!Localizacion
                itmX.SubItems(15) = rs!Proveedor
            
                itmX.SubItems(16) = rs!Modelo & ""
                itmX.SubItems(17) = rs!Marca & ""
                itmX.SubItems(18) = rs!Num_Serie & ""
                itmX.SubItems(19) = rs!Otras_Senas
             
             curVH = curVH + rs!Valor_Historico
             curVR = curVR + rs!Valor_Desecho
             
         rs.MoveNext
        Loop
        rs.Close
         Set itmX = lsw.ListItems.Add(, "")
             itmX.SubItems(1) = "___"
             itmX.SubItems(7) = "____________________"
             itmX.SubItems(8) = "____________________"
        
         Set itmX = lsw.ListItems.Add(, "")
             itmX.SubItems(1) = Format(lsw.ListItems.Count - 2, "###,##0")
             itmX.SubItems(7) = Format(curVH, "Standard")
             itmX.SubItems(8) = Format(curVR, "Standard")
    



    Case Id_TaskItem_Activos_Traslados
        TituloOpciones.Caption = "Traslados de Activos..:"
        TituloOpciones.Tag = "Activos_Traslado"
        
        Call sbOpcionesVisibles(False)
    
        scTitulo1.Visible = True
        TituloOpciones.Visible = True
        tpMain.Visible = True
        
        btnExport.Visible = False
        
        tcMain.Visible = True
        tcMain.Item(4).Selected = True
        tcTraslados.Item(0).Selected = True
        
        lswAF_Boletas.ListItems.Clear
        lswAF_Recepcion.ListItems.Clear
        lswAF_Activos.ListItems.Clear
        vGrid.MaxRows = 0
        
        With lswAF_Boletas.ColumnHeaders
            .Clear
            .Add , , "Boleta Id", 1500
            .Add , , "Estado", 1100, vbCenter
            .Add , , "Fecha", 1800, vbCenter
            .Add , , "Usuario", 1800, vbCenter
            .Add , , "Destino Id", 1580
            .Add , , "D. Persona", 3500
            .Add , , "D. Departamento", 2500
            .Add , , "D. Sección", 2500
            .Add , , "Motivo", 2500
            .Add , , "Procesa Fecha", 1800, vbCenter
            .Add , , "Procesa Usuario", 1800, vbCenter
        End With
        
        With lswAF_Recepcion.ColumnHeaders
            .Clear
            .Add , , "Boleta Id", 1500
            .Add , , "Estado", 1100, vbCenter
            .Add , , "Fecha", 1800, vbCenter
            .Add , , "Usuario", 1800, vbCenter
            .Add , , "Origen Id", 1580
            .Add , , "O. Persona", 3500
            .Add , , "O. Departamento", 2500
            .Add , , "O. Sección", 2500
            .Add , , "Motivo", 2500
            .Add , , "Procesa Fecha", 1800, vbCenter
            .Add , , "Procesa Usuario", 1800, vbCenter
        End With
       
        
        With lswAF_Activos.ColumnHeaders
            .Clear
            .Add , , "No. Placa", 2100
            .Add , , "Descripción", 3100
            .Add , , "Depreciación", 2500, vbRightJustify
            .Add , , "Dep. Mensual", 2500, vbRightJustify
            .Add , , "Valor Libros", 2500, vbRightJustify
        End With
        
        strSQL = "select rtrim(cod_Motivo) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
               & " FROM ACTIVOS_TRASLADOS_MOTIVOS WHERE ACTIVO = 1 order by cod_Motivo"
        Call sbCbo_Llena_New(cboMotivo, strSQL, False, False)
 
        Call sbActivos_Boletas("O")


      Case Id_TaskItem_Activos_Declaracion
        TituloOpciones.Caption = "Declaración de Activos..:"
        TituloOpciones.Tag = "Activos_Declara"
        
        Call sbOpcionesVisibles(False)
    
        scTitulo1.Visible = True
        TituloOpciones.Visible = True
        tpMain.Visible = True
        
        btnExport.Visible = False
        
        tcMain.Visible = True
        tcMain.Item(5).Selected = True
        tcDeclaracion.Item(0).Selected = True
        
        lswDeclara.ListItems.Clear
        lswDeclara_H.ListItems.Clear
        lswDeclara_H_Det.ListItems.Clear
        
        With lswDeclara_H.ColumnHeaders
            .Clear
            .Add , , "Declara Id", 1500
            .Add , , "Estado", 1100, vbCenter
            .Add , , "Inicio", 1800, vbCenter
            .Add , , "Corte", 1800, vbCenter
            .Add , , "Notas", 3800
            .Add , , "Proc. Fecha", 1800, vbCenter
            .Add , , "Proc. Usuario", 1800, vbCenter
        End With
        
        
        With lswDeclara.ColumnHeaders
            .Clear
            .Add , , "Declara Id", 1500
            .Add , , "No.Placa", 2180
            .Add , , "Descripción", 3180
            .Add , , "Localización", 3180
            .Add , , "Estado", 1100, vbCenter
            .Add , , "Reg. Fecha", 1800, vbCenter
            .Add , , "Reg. Usuario", 1800, vbCenter
            .Add , , "Proc. Fecha", 1800, vbCenter
            .Add , , "Proc. Usuario", 1800, vbCenter
            .Add , , "Tras. Fecha", 1800, vbCenter
            .Add , , "Tras. Usuario", 1800, vbCenter
            .Add , , "Tras. Boleta", 1800, vbCenter
        End With

        With lswDeclara_H_Det.ColumnHeaders
            .Clear
            .Add , , "Declara Id", 1500
            .Add , , "No.Placa", 2180
            .Add , , "Descripción", 3180
            .Add , , "Localización", 3180
            .Add , , "Estado", 1100, vbCenter
            .Add , , "Reg. Fecha", 1800, vbCenter
            .Add , , "Reg. Usuario", 1800, vbCenter
            .Add , , "Proc. Fecha", 1800, vbCenter
            .Add , , "Proc. Usuario", 1800, vbCenter
            .Add , , "Tras. Fecha", 1800, vbCenter
            .Add , , "Tras. Usuario", 1800, vbCenter
            .Add , , "Tras. Boleta", 1800, vbCenter
        End With

        vPaso = True
         strSQL = "select rtrim(COD_LOCALIZA) as 'Idx', rtrim(descripcion) as 'ItmX'" _
              & " from ACTIVOS_LOCALIZACIONES Where Activa = 1 order by descripcion"
         Call sbCbo_Llena_New(cboLocaliza, strSQL, False, True)
        
        vPaso = False

End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbActivos_Boletas(pTipo As String)

On Error GoTo vError
Me.MousePointer = vbHourglass

lswAF_Activos.Visible = True
vGrid.Visible = False


lswAF_Activos.ListItems.Clear

Select Case pTipo
   Case "O" 'Origen
        tcMain.Item(4).Selected = True
        tcTraslados.Item(0).Selected = True
        
        strSQL = "select * from vActivos_Traslados_Boletas where Identificacion = '" & txtIdentificacion.Text _
               & "' Order by Registro_Fecha desc"
        Call OpenRecordSet(rs, strSQL)
        
        lswAF_Boletas.ListItems.Clear
        
        Do While Not rs.EOF
          Set itmX = lswAF_Boletas.ListItems.Add(, , rs!Cod_Traslado)
              itmX.SubItems(1) = rs!Estado_Desc
              itmX.SubItems(2) = rs!Registro_Fecha
              itmX.SubItems(3) = rs!Registro_Usuario
              itmX.SubItems(4) = rs!Identificacion_Destino
              itmX.SubItems(5) = rs!Persona_Destino
              itmX.SubItems(6) = rs!Departamento_Destino
              itmX.SubItems(7) = rs!Seccion_Destino
              itmX.SubItems(8) = rs!Motivo
              itmX.SubItems(9) = rs!Procesado_Fecha & ""
              itmX.SubItems(10) = rs!Procesado_Usuario & ""
          rs.MoveNext
        Loop
        
        txtAF_Boleta.Text = ""
        txtAF_Boleta_Estado.Text = ""
        
   Case "R" 'Recepcion
        tcMain.Item(4).Selected = True
        tcTraslados.Item(1).Selected = True
        
        strSQL = "select * from vActivos_Traslados_Boletas where Identificacion_Destino = '" & txtIdentificacion.Text _
               & "' Order by Registro_Fecha desc"
        Call OpenRecordSet(rs, strSQL)
        
        lswAF_Recepcion.ListItems.Clear
        
        Do While Not rs.EOF
          Set itmX = lswAF_Recepcion.ListItems.Add(, , rs!Cod_Traslado)
              itmX.SubItems(1) = rs!Estado_Desc
              itmX.SubItems(2) = rs!Registro_Fecha
              itmX.SubItems(3) = rs!Registro_Usuario
              itmX.SubItems(4) = rs!Identificacion
              itmX.SubItems(5) = rs!Persona
              itmX.SubItems(6) = rs!departamento
              itmX.SubItems(7) = rs!seccion
              itmX.SubItems(8) = rs!Motivo
              itmX.SubItems(9) = rs!Procesado_Fecha & ""
              itmX.SubItems(10) = rs!Procesado_Usuario & ""
          rs.MoveNext
        Loop

        txtAF_Boleta_R.Text = ""

End Select


Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbActivo_Boleta_Recibir(pBoletaId As String)
On Error GoTo vError

Dim i As Integer

i = MsgBox("Esta seguro que desea [Aceptar la Recepción] esta boleta de Cambio de Responsable?", vbYesNo)
If i = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spActivos_Responsable_Cambio_Procesa '" & pBoletaId & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Boleta de Cambio de Responsables: Procesada Satisfactoriamente!", vbInformation
   
   Call sbActivos_Boletas("R")
   Call sbActivos_Boleta_Cambio_Responsable(pBoletaId)
  
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbActivo_Descartar(pBoletaId As String)
On Error GoTo vError

Dim i As Integer

i = MsgBox("Esta seguro que desea DESCARTAR esta boleta de Cambio de Responsable?", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass

strSQL = "exec spActivos_Responsable_Cambio_Descarta '" & pBoletaId & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Boleta de Cambio de Responsables: Descartada Satisfactoriamente!", vbInformation
   Call sbActivos_Boletas("O")
   Call sbActivos_Boleta_Cambio_Responsable(pBoletaId)
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnAF_Boleta_Click(Index As Integer)

Select Case Index
    Case 0 'Boleta
        If txtAF_Boleta.Text = "" Then Exit Sub
        Call sbActivos_Boleta_Cambio_Responsable(txtAF_Boleta.Text)
    
    Case 1 'Descartar
        If txtAF_Boleta.Text = "" Then Exit Sub
        If Mid(txtAF_Boleta_Estado.Text, 1, 1) <> "S" Then Exit Sub
        
        Call sbActivo_Descartar(txtAF_Boleta.Text)
        
    Case 2 'Boleta Recepcion
        If txtAF_Boleta_R.Text = "" Then Exit Sub
        Call sbActivos_Boleta_Cambio_Responsable(txtAF_Boleta_R.Text)
        
    Case 3 'Recepcion Aceptar
        If txtAF_Boleta_R.Text = "" Then Exit Sub
        Call sbActivo_Boleta_Recibir(txtAF_Boleta_R.Text)
        
    Case 4 'Crear Boleta de Traslado
        If fxActivo_Traslado_Valida Then
            Call sbActivos_Traslado_Guardar
        End If
End Select

End Sub

Private Sub btnAF_Declara_Click(Index As Integer)
Select Case Index
    Case 0 'Nuevo
    Case 1 'Guarda
    Case 2 'Procesa
End Select

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

Private Sub sbVacaciones_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Visible = True
tcMain.Item(0).Selected = True

scVacaciones.Width = tcMain.Width
gbVacaciones.Width = tcMain.Width
gbVacaciones.Height = tcMain.Height

txtVac_Notas.Text = ""

strSQL = "select * from vRH_Vacaciones_Info" _
       & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtVac_DiasDisponibles.Text = Format(rs!Dias_Disponibles, "Standard")
    txtVac_Dias.Text = 1
    
    
    dtpVac_FechaI.MinDate = rs!Fecha_Inicio
    dtpVac_FechaC.MinDate = rs!Fecha_Inicio
    
    dtpVac_FechaI.Value = rs!fecha
    dtpVac_FechaC.Value = rs!fecha
Else
    txtVac_DiasDisponibles.Text = 0
    txtVac_Dias.Text = 0
End If
rs.Close

vPaso = True

    dtpVac_FechaI.Value = fxFechaServidor
    dtpVac_FechaC.Value = dtpInc_FechaI.Value

    strSQL = "exec spRH_Portal_Vacaciones_Tipos"
    Call sbCbo_Llena_New(cboVac_Tipo, strSQL, False, True)

vPaso = False

Call cboVac_Tipo_Click


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbPermisos_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Visible = True
tcMain.Item(1).Selected = True

scPermiso.Width = tcMain.Width
gbPermisos.Width = tcMain.Width
gbPermisos.Height = tcMain.Height

vPaso = True

txtPer_Notas.Text = ""

dtpPer_Fecha.Value = fxFechaServidor

dtpPer_HoraI.Value = dtpPer_Fecha.Value
dtpPer_HoraC.Value = dtpPer_Fecha.Value

'Tipo
strSQL = "exec spRH_Portal_Permisos_Tipos"
Call sbCbo_Llena_New(cboPer_Tipo, strSQL, False, True)

'Estado
cboPer_Estado.AddItem "Solicitado"
cboPer_Estado.Text = "Solicitado"

vPaso = False

Call cboPer_Tipo_Click


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbIncapacidades_Load()

On Error GoTo vError

tcMain.Visible = True
tcMain.Item(2).Selected = True

scIncapacidad.Width = tcMain.Width
gbIncapacidades.Width = tcMain.Width
gbIncapacidades.Height = tcMain.Height



strSQL = "select EMPLEADO_ID,IDENTIFICACION,NOMBRE_COMPLETO, dbo.Mygetdate() as 'Fecha'" _
       & ",dbo.fxRH_Nomina_Inicial_Actual(COD_NOMINA) AS 'Fecha_Inicio'" _
       & " from Rh_Personas" _
       & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then

    dtpInc_FechaI.MinDate = rs!Fecha_Inicio
    dtpInc_FechaC.MinDate = rs!Fecha_Inicio

    dtpInc_FechaI.Value = rs!fecha
    dtpInc_FechaC.Value = rs!fecha
    
    txtInc_Dias.Text = 1
Else
    dtpInc_FechaI.Value = fxFechaServidor
    dtpInc_FechaC.Value = dtpInc_FechaI.Value
    
    txtInc_Dias.Text = 0
End If
rs.Close


vPaso = True
    strSQL = "exec spRH_Portal_Incapacidades_Tipos"
    Call sbCbo_Llena_New(cboInc_Tipo, strSQL, False, True)
vPaso = False

Call cboInc_Tipo_Click

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbAutorizaciones_Inicial()

vPaso = True

tcMain.Visible = True
tcMain.Item(3).Selected = True

'scIncapacidad.Width = tcMain.Width
'gbIncapacidades.Width = tcMain.Width
'gbIncapacidades.Height = tcMain.Height


cboAut_Tipo.Clear
cboAut_Tipo.AddItem "Permisos"
cboAut_Tipo.AddItem "Vacaciones"
cboAut_Tipo.AddItem "Incapacidades"
cboAut_Tipo.Text = "Permisos"


cboAut_Autorizado.Clear
cboAut_Autorizado.AddItem "Solicitadas"
cboAut_Autorizado.AddItem "Autorizadas"
cboAut_Autorizado.AddItem "Denegadas"
cboAut_Autorizado.Text = "Solicitadas"

vPaso = False

dtpAut_Inicio.Value = fxFechaServidor
dtpAut_Corte.Value = dtpAut_Inicio.Value

Call sbAutoriza_Buscar

End Sub


Private Sub btnEditarDetalle_Click()

Select Case mId_Task_Seleted
  Case Id_TaskItem_Vacaciones, Id_TaskItem_Permisos, Id_TaskItem_Incapacidades, Id_TaskItem_Autorizaciones, Id_TaskItem_Activos_Traslados, Id_TaskItem_Activos_Declaracion
    'Ok
  Case Else
    Exit Sub
End Select

Call sbOpcionesVisibles(False)


Select Case mId_Task_Seleted
  Case Id_TaskItem_Vacaciones
        Call sbVacaciones_Load
        
  Case Id_TaskItem_Permisos
        Call sbPermisos_Load
        
  Case Id_TaskItem_Incapacidades
        Call sbIncapacidades_Load
        
  Case Id_TaskItem_Autorizaciones
        Call sbAutorizaciones_Inicial
        
  Case Id_TaskItem_Activos_Traslados
  
  Case Id_TaskItem_Activos_Declaracion
  
  Case Else

End Select

scTitulo1.Visible = True
TituloOpciones.Visible = True
tpMain.Visible = True

btnMenu.Item(0).Visible = True
btnMenu.Item(1).Visible = True

End Sub

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
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


Private Sub btnPer_Aplicar_Click()
If vPaso Then Exit Sub
If cboPer_Tipo.ListCount = 0 Then Exit Sub

Dim Boleta As String, LiquidaId As Integer

'Validacion
If dtpPer_HoraI.Value > dtpPer_HoraC.Value Then
    MsgBox "Error en Rango de Horas!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtPer_Horas.Text) Then
    MsgBox "Horas de Permisos Inválidas!", vbExclamation
    Exit Sub
End If


If Not IsNumeric(txtPer_HrsMax.Text) Then
    MsgBox "Horas de Permisos Inválidas!", vbExclamation
    Exit Sub
End If


If CCur(txtPer_Horas.Text) < 0 Then
    MsgBox "Horas de Permisos Inválidas!", vbExclamation
    Exit Sub
End If

If CCur(txtPer_Horas.Text) > CCur(txtPer_HrsMax.Text) Then
    MsgBox "Horas de Permisos Excedente el Total Permitido!", vbExclamation
    Exit Sub
End If

On Error GoTo vError

Dim pAutorizador As String

If Mid(cboPer_Estado.Text, 1, 1) = "S" Then
  pAutorizador = "Null"
Else
  pAutorizador = "Null"
End If

strSQL = "exec spRH_Permisos_Registro '" & txtEmpleadoId.Text & "','" & cboPer_Tipo.ItemData(cboPer_Tipo.ListIndex) _
        & "','" & txtPer_Notas.Text & "','" & glogon.Usuario & "'" _
        & ",'" & Format(dtpPer_Fecha.Value, "yyyy/mm/dd") & " " & Format(dtpPer_HoraI.Value, "hh:mm:ss") & "'" _
        & ",'" & Format(dtpPer_Fecha.Value, "yyyy/mm/dd") & " " & Format(dtpPer_HoraC.Value, "hh:mm:ss") & "'" _
        & "," & CCur(txtPer_Horas.Text) & ",'" & Format(dtpPer_Fecha.Value, "yyyy/mm/dd") _
        & "','" & Mid(cboPer_Estado.Text, 1, 1) & "'," & pAutorizador & ",'ProGrX'"
Call OpenRecordSet(rs, strSQL)
    Boleta = rs!BoletaId
rs.Close


Me.MousePointer = vbDefault

MsgBox "Solicitud de Permiso registrado Satisfactoriamente!", vbInformation

'Print Boleta
Call sbRH_Boleta_Permisos(Boleta, txtEmpleadoId.Text)

Call sbTaskPanel_Accion(Id_TaskItem_Permisos)


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboInc_Tipo_Click()
If vPaso Then Exit Sub
If cboInc_Tipo.ListCount = 0 Then Exit Sub

On Error GoTo vError


strSQL = "exec spRH_Portal_Incapacidades_Tipos '" & cboInc_Tipo.ItemData(cboInc_Tipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

cboInc_Estado.Clear
cboInc_Estado.AddItem "Solicitado"
cboInc_Estado.Text = "Solicitado"

txtInc_Porcentaje.Text = Format(rs!Porc_Patrono, "Standard")

If rs!REQUIERE_AUTORIZACION = 0 Then
    cboInc_Estado.AddItem "Autorizado"
End If

rs.Close

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboPer_Tipo_Click()
If vPaso Then Exit Sub
If cboPer_Tipo.ListCount = 0 Then Exit Sub

On Error GoTo vError

strSQL = "select REQUIERE_AUTORIZACION,PERMISO_HRS_MAX" _
       & " from RH_PERMISOS_TIPOS " _
       & " WHERE PERMISO_TIPO = '" & cboPer_Tipo.ItemData(cboPer_Tipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

cboPer_Estado.Clear
cboPer_Estado.AddItem "Solicitado"
cboPer_Estado.Text = "Solicitado"

txtPer_HrsMax.Text = CStr(rs!PERMISO_HRS_MAX)


If rs!REQUIERE_AUTORIZACION = 0 Then
    cboPer_Estado.AddItem "Autorizado"
End If

rs.Close

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub btnVac_Aplicar_Click()
If vPaso Then Exit Sub
If cboVac_Tipo.ListCount = 0 Then Exit Sub

Dim Boleta As String, LiquidaId As Integer

'Validacion
If dtpVac_FechaI.Value > dtpVac_FechaC.Value Then
    MsgBox "Error en Rango de Fechas!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtVac_Dias.Text) Then
    MsgBox "Dias de Vacaciones Inválido!", vbExclamation
    Exit Sub
End If

If CCur(txtVac_Dias.Text) < 0 Then
    MsgBox "Dias de Vacaciones Inválido!", vbExclamation
    Exit Sub
End If

If Not fxRH_Vacaciones_Valida(mNomina, txtEmpleadoId.Text, dtpVac_FechaI.Value, dtpVac_FechaC.Value) Then
    MsgBox "Existe conflicto de fechas de disfrute con alguna otra boleta procesada o con una Nómina ya ejecutada!", vbExclamation
    Exit Sub
End If


On Error GoTo vError


Dim pAutorizador As String

If Mid(cboPer_Estado.Text, 1, 1) = "S" Then
  pAutorizador = "Null"
Else
  pAutorizador = "Null"
End If

strSQL = "exec spRH_Vacaciones_Registro '" & txtEmpleadoId.Text & "','" & cboVac_Tipo.ItemData(cboPer_Tipo.ListIndex) _
        & "','" & txtVac_Notas.Text & "','" & glogon.Usuario & "'" _
        & ",'" & Format(dtpVac_FechaI.Value, "yyyy/mm/dd") & " 00:00:00'" _
        & ",'" & Format(dtpVac_FechaC.Value, "yyyy/mm/dd") & " 23:59:59'" _
        & "," & CInt(txtVac_Dias.Text) & "," & CCur(txtVac_DiasDisponibles.Text) & "," & LiquidaId _
        & ",'" & Mid(cboVac_Estado.Text, 1, 1) & "'," & pAutorizador & ",'ProGrX'"
Call OpenRecordSet(rs, strSQL)
    Boleta = rs!BoletaId
rs.Close

Me.MousePointer = vbDefault

MsgBox "Solicitud de Vacaciones registrada satisfactoriamente!", vbInformation

'Print Boleta
Call sbRH_Boleta_Vacaciones(Boleta, txtEmpleadoId.Text)

Call sbTaskPanel_Accion(Id_TaskItem_Vacaciones)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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



Private Sub cboVac_Tipo_Click()
If vPaso Then Exit Sub
If cboVac_Tipo.ListCount = 0 Then Exit Sub


strSQL = "exec spRH_Portal_Vacaciones_Tipos '" & cboVac_Tipo.ItemData(cboVac_Tipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

cboVac_Estado.Clear
cboVac_Estado.AddItem "Solicitado"
cboVac_Estado.Text = "Solicitado"


rbVac_Accion.Item(0).Value = True

If rs!REQUIERE_AUTORIZACION = 0 Then
    cboVac_Estado.AddItem "Autorizado"
End If

If rs!PERMITE_LIQUIDACION = 1 Then
    rbVac_Accion.Item(1).Enabled = True
Else
    rbVac_Accion.Item(1).Enabled = False
End If

Call rbVac_Accion_Click(0)

rs.Close

End Sub

Private Sub chkAut_Todos_Click()
Dim i As Long

For i = 1 To lswAut.ListItems.Count
  lswAut.ListItems.Item(i).Checked = chkAut_Todos.Value
Next i

End Sub

Private Sub dtpVac_FechaC_Change()
If txtEmpleadoId.Text <> "" Then
    txtVac_Dias.Text = fxRH_Dias_Laborales(txtEmpleadoId.Text, dtpVac_FechaI.Value, dtpVac_FechaC.Value)
End If
End Sub

Private Sub dtpVac_FechaI_Change()
If txtEmpleadoId.Text <> "" Then
    txtVac_Dias.Text = fxRH_Dias_Laborales(txtEmpleadoId.Text, dtpVac_FechaI.Value, dtpVac_FechaC.Value)
End If
End Sub





Private Sub btnInc_Aplicar_Click()
If vPaso Then Exit Sub
If cboInc_Tipo.ListCount = 0 Then Exit Sub

Dim Boleta As String, LiquidaId As Integer

'Validacion
If dtpInc_FechaI.Value > dtpInc_FechaC.Value Then
    MsgBox "Error en Rango de Fechas!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtInc_Dias.Text) Then
    MsgBox "Dias de Incapacidad Inválido!", vbExclamation
    Exit Sub
End If

If CCur(txtInc_Dias.Text) < 0 Then
    MsgBox "Dias de Incapacidad Inválido!", vbExclamation
    Exit Sub
End If

On Error GoTo vError

Dim pAutorizador As String

If Mid(cboInc_Estado.Text, 1, 1) = "S" Then
  pAutorizador = "Null"
Else
  pAutorizador = "Null"
End If

strSQL = "exec spRH_Incapacidades_Registro '" & txtEmpleadoId.Text & "','" & cboInc_Tipo.ItemData(cboInc_Tipo.ListIndex) _
        & "','" & txtInc_Notas.Text & "','" & glogon.Usuario & "'" _
        & ",'" & Format(dtpInc_FechaI.Value, "yyyy/mm/dd") & " 00:00:00'" _
        & ",'" & Format(dtpInc_FechaC.Value, "yyyy/mm/dd") & " 23:59:59'" _
        & "," & CInt(txtInc_Dias.Text) & "," & CCur(txtInc_Porcentaje.Text) _
        & ",'" & Mid(cboInc_Estado.Text, 1, 1) & "'," & pAutorizador & ",'ProGrX'"
Call OpenRecordSet(rs, strSQL)
    Boleta = rs!BoletaId
rs.Close


Me.MousePointer = vbDefault

MsgBox "Solicitud de Incapacidad Registrada Satisfactoriamente, Boleta: " & Boleta, vbInformation

'Print Boleta
Call sbRH_Boleta_Incapacidad(Boleta, txtEmpleadoId.Text)

Call sbTaskPanel_Accion(Id_TaskItem_Incapacidades)


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboAut_Tipo_Click()
If vPaso Then Exit Sub
If cboInc_Tipo.ListCount = 0 Then Exit Sub

On Error GoTo vError

strSQL = "exec spRH_Portal_Incapacidades_Tipos '" & cboInc_Tipo.ItemData(cboInc_Tipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

cboInc_Estado.Clear
cboInc_Estado.AddItem "Solicitado"
cboInc_Estado.Text = "Solicitado"

txtInc_Porcentaje.Text = Format(rs!Porc_Patrono, "Standard")

If rs!REQUIERE_AUTORIZACION = 0 Then
    cboInc_Estado.AddItem "Autorizado"
End If

rs.Close

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpInc_FechaC_Change()
If txtEmpleadoId.Text <> "" Then
    txtInc_Dias.Text = fxRH_Dias_Laborales(txtEmpleadoId.Text, dtpInc_FechaI.Value, dtpInc_FechaC.Value)
End If
End Sub

Private Sub dtpInc_FechaI_Change()
If txtEmpleadoId.Text <> "" Then
    txtInc_Dias.Text = fxRH_Dias_Laborales(txtEmpleadoId.Text, dtpInc_FechaI.Value, dtpInc_FechaC.Value)
End If
End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lswAF_Boletas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)


txtAF_Boleta.Text = Item.Text
txtAF_Boleta_Estado.Text = Item.SubItems(1)

Call sbActivos_Consulta_Placas_Listas(Item.Text)


End Sub

Private Sub lswAF_Recepcion_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtAF_Boleta_R.Text = Item.Text
Call sbActivos_Consulta_Placas_Listas(Item.Text)

End Sub


Private Sub lswAut_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswAut.SortKey = ColumnHeader.Index - 1
  If lswAut.SortOrder = 0 Then lswAut.SortOrder = 1 Else lswAut.SortOrder = 0
  lswAut.Sorted = True
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


Private Sub rbVac_Accion_Click(Index As Integer)

If Index = 0 Then
    dtpVac_FechaI.Enabled = True
    dtpVac_FechaC.Enabled = True
    
    txtVac_Dias.Locked = True
    
    Call dtpVac_FechaC_Change
    
Else
    dtpVac_FechaI.Enabled = False
    dtpVac_FechaC.Enabled = False
    txtVac_Dias.Locked = False
End If

End Sub


Private Sub sbAutoriza_Buscar()

Me.MousePointer = vbHourglass

On Error GoTo vError

lswAut.ListItems.Clear

With lswAut.ColumnHeaders
    .Clear
    .Add , , "Boleta Id", 1200
    .Add , , "Empleado Id", 1200
    .Add , , "Identificación", 1200
    .Add , , "Nombre", 3200
    .Add , , "Concepto", 2200
    .Add , , "Notas", 4000
    .Add , , "Usuario", 1600, vbCenter
    .Add , , "Fecha", 2100
End With

Select Case Mid(cboAut_Tipo.Text, 1, 1)
  Case "V"
        strSQL = "select * from vRH_Boleta_Vacaciones"
  
        lsw.ColumnHeaders.Add , , "Días", 900, vbCenter
        lsw.ColumnHeaders.Add , , "Salida", 1400, vbCenter
        lsw.ColumnHeaders.Add , , "Entrada", 1400, vbCenter
  
  Case "P"
        strSQL = "select * from vRH_Boleta_Permisos"
        lsw.ColumnHeaders.Add , , "Horas", 900, vbCenter
        lsw.ColumnHeaders.Add , , "Salida", 1800, vbCenter
        lsw.ColumnHeaders.Add , , "Entrada", 1800, vbCenter
  
  Case "I"
        strSQL = "select * from vRH_Boleta_Incapacidades"
        lsw.ColumnHeaders.Add , , "Días", 900, vbCenter
        lsw.ColumnHeaders.Add , , "Salida", 1400, vbCenter
        lsw.ColumnHeaders.Add , , "Entrada", 1400, vbCenter
End Select

strSQL = strSQL & " Where Estado = '" & Mid(cboAut_Autorizado.Text, 1, 1) _
       & "' and Registro_Fecha between '" & Format(dtpAut_Inicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpAut_Corte.Value, "yyyy/mm/dd") & " 23:59:59'"


'Autorizador> Empleados Autorizados
strSQL = strSQL & "  and dbo.fxRH_Autorizador_Valida(Empleado_ID,'" & txtEmpleadoId.Text & "') = 1"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswAut.ListItems.Add(, , rs!Boleta_Id)
     itmX.SubItems(1) = rs!Empleado_ID
     itmX.SubItems(2) = rs!Identificacion
     itmX.SubItems(3) = rs!NOMBRE_COMPLETO
     itmX.SubItems(4) = rs!TipoDesc
     itmX.SubItems(5) = rs!Motivo & ""
     itmX.SubItems(6) = rs!Registro_Usuario & ""
     itmX.SubItems(7) = Format(rs!Registro_Fecha & "", "dd/mm/yyyy")
     
    Select Case Mid(cboAut_Tipo.Text, 1, 1)
      Case "V"
        itmX.SubItems(8) = rs!Dias_Disfrutados & ""
        itmX.SubItems(9) = Format(rs!Fecha_Salida & "", "dd/mm/yyyy")
        itmX.SubItems(10) = Format(rs!fecha_entrada & "", "dd/mm/yyyy")
      
      Case "I"
        itmX.SubItems(8) = rs!Dias & ""
        itmX.SubItems(9) = Format(rs!Fecha_Salida & "", "dd/mm/yyyy")
        itmX.SubItems(10) = Format(rs!fecha_entrada & "", "dd/mm/yyyy")
      
      Case "P"
        itmX.SubItems(8) = rs!Hrs_Total & ""
        itmX.SubItems(9) = Format(rs!Hora_Inicio & "", "dd/mm/yyyy hh:mm:ss")
        itmX.SubItems(10) = Format(rs!Hora_Corte & "", "dd/mm/yyyy hh:mm:ss")
    End Select
     
     
     Select Case rs!Estado
         Case "S"
         Case "A"
              itmX.Bold = True
              itmX.TextBackColor = RGB(252, 243, 207)
         Case "D"
              itmX.ForeColor = vbRed
              itmX.Bold = True
              itmX.TextBackColor = RGB(250, 219, 216)
     End Select

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAutoriza(pGestion As String)
Dim strSQL As String, i As Long

On Error GoTo vError

vModulo = 23

Me.MousePointer = vbHourglass

'spRH_Autorizaciones_Registro(@AutorizadorId varchar(20), @Tipo varchar(10), @BoletaId varchar(30), @Usuario varchar(30)
'                , @Estado char(1) = 'A', @AppCod varchar(30) = 'ProGrX' )

If pGestion = Mid(cboAut_Autorizado.Text, 1, 1) Then Exit Sub


With lswAut.ListItems
  For i = 1 To .Count
      If .Item(i).Checked Then
         strSQL = "exec spRH_Autorizaciones_Registro '" & txtEmpleadoId.Text & "','" & Mid(cboAut_Tipo.Text, 1, 1) _
                & "','" & .Item(i).Text & "','" & glogon.Usuario & "','" & pGestion & "','ProGrX'"
         Call ConectionExecute(strSQL)

         Call Bitacora("Aplica", IIf((pGestion = "A"), "Autoriza", "Deniega") & " de Boleta Id:" & .Item(i).Text _
                 & "..Empleado Id: " & .Item(i).SubItems(1) & "..Persona Id: " & .Item(i).SubItems(2))

      End If
  Next i
End With


Me.MousePointer = vbDefault
MsgBox IIf((pGestion = "A"), "Autorización", "Denegación") & " realizada satisfactoriamente.!", vbInformation

Call sbAutoriza_Buscar

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbActivos_Consulta_Placas(Optional pCodigo As String = "")

On Error GoTo vError

Me.MousePointer = vbHourglass
    
'Carga Activos
strSQL = "exec spActivos_Responsable_Cambio_Consulta_Placas '" & pCodigo & "', '" & txtIdentificacion.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
With vGrid
  .MaxRows = 0
   
   Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      .Col = 1
      .Value = rs!Asignado
      .Col = 2
      .Text = rs!NUM_PLACA
      .Col = 3
      .Text = rs!DESCRIPCION
      .Col = 4
      .Text = Format(rs!DEPRECIACION_AC, "Standard")
      .Col = 5
      .Text = Format(rs!DEPRECIACION_MES, "Standard")
      .Col = 6
      .Text = Format(rs!VALOR_LIBROS, "Standard")
      rs.MoveNext
   Loop
   rs.Close
   
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbActivos_Consulta_Placas_Listas(pCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass
    
'Carga Activos
strSQL = "exec spActivos_Responsable_Cambio_Consulta_Placas '" & pCodigo & "', '" & txtIdentificacion.Text & "', '" & glogon.Usuario & "', 1"

Call OpenRecordSet(rs, strSQL)

        With lswAF_Activos.ColumnHeaders
            .Clear
            .Add , , "No. Placa", 2100
            .Add , , "Descripción", 3100
            .Add , , "Depreciación", 2500, vbRightJustify
            .Add , , "Dep. Mensual", 2500, vbRightJustify
            .Add , , "Valor Libros", 2500, vbRightJustify
        End With



With lswAF_Activos.ListItems
  .Clear
   
   Do While Not rs.EOF
      Set itmX = .Add(, , rs!NUM_PLACA)
          itmX.SubItems(1) = rs!DESCRIPCION
          itmX.SubItems(2) = Format(rs!DEPRECIACION_AC, "Standard")
          itmX.SubItems(3) = Format(rs!DEPRECIACION_MES, "Standard")
          itmX.SubItems(4) = Format(rs!VALOR_LIBROS, "Standard")
      rs.MoveNext
   Loop
   rs.Close
   
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Function fxActivo_Traslado_Valida() As Boolean
Dim vMensaje As String, i As Long, iPlacas As Long

vMensaje = ""
fxActivo_Traslado_Valida = True

On Error GoTo vError

If Len(txtNotas.Text) < 10 Then vMensaje = vMensaje & vbCrLf & "- No ha indicado una nota válida"
If cboMotivo.ListCount < 0 Then vMensaje = vMensaje & vbCrLf & "- No exite o no ha indicado un motivo"

If txtNuevoPersona.Tag = "" Then vMensaje = vMensaje & vbCrLf & "- No se ha indicado un responsable destino"

If vGrid.MaxRows = 0 Then vMensaje = vMensaje & vbCrLf & "- No se ha cargado ningún activo a trasladar"

iPlacas = 0
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 1
  If vGrid.Value = 1 Then
     iPlacas = iPlacas + 1
  End If
Next i

If iPlacas = 0 Then vMensaje = vMensaje & vbCrLf & "- No se ha Seleccionado ningún activo a trasladar"

vError:

If Len(vMensaje) > 0 Then
  fxActivo_Traslado_Valida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbActivos_Traslado_Guardar()
Dim i As Long, iLineaUno As Boolean

On Error GoTo vError

Dim vFecha As Date, vCodigo As String

vModulo = 36

vFecha = fxFechaServidor

strSQL = "exec spActivos_Responsable_Cambio_Boleta_Add '', '" & cboMotivo.ItemData(cboMotivo.ListIndex) _
       & "', '" & txtNotas.Text & "', '" & txtIdentificacion.Text & "', '" & txtNuevoPersona.Tag & "', '" & glogon.Usuario _
       & "', '" & Format(vFecha, "yyyy-mm-dd") & "'"

Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
    vCodigo = rs!Boleta
    
    
    Call Bitacora("Registra", "Boleta de Cambio Responsable:  " & vCodigo)

    strSQL = ""
    iLineaUno = True
    For i = 1 To vGrid.MaxRows
      vGrid.Row = i

      vGrid.Col = 1

      If vGrid.Value = 1 Then

        vGrid.Col = 2
        If iLineaUno Then
          iLineaUno = False
          strSQL = strSQL & Space(10) & "exec spActivos_Responsable_Cambio_Boleta_Placas '" & vCodigo & "', '" & vGrid.Text & "', '" & glogon.Usuario & "', 1"
        Else
          strSQL = strSQL & Space(10) & "exec spActivos_Responsable_Cambio_Boleta_Placas '" & vCodigo & "', '" & vGrid.Text & "', '" & glogon.Usuario & "', 0"
        End If

        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If

      End If
    Next i

    'Ultimo Lote
    If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL)
    End If

    MsgBox "Boleta Registrada Satisfactoriamente, puede dar seguimiento en la opción de Boletas!", vbInformation
    
    Call sbActivos_Boletas("O")

Else
    MsgBox rs!Mensaje, vbExclamation
End If



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub tcTraslados_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Boletas
        Call sbActivos_Boletas("O")
    Case 1 'Recepcion
        Call sbActivos_Boletas("R")
    Case 2 'Traslado
        vGrid.Left = lswAF_Activos.Left
        vGrid.top = lswAF_Activos.top
        
        lswAF_Activos.Visible = False
        vGrid.Visible = True
        
        txtNotas.Text = ""
        txtNuevoDepartamento.Text = ""
        txtNuevoSeccion.Text = ""
        txtNuevoPersona.Text = ""
        txtNuevoPersona.Tag = ""
        
        Call sbActivos_Consulta_Placas
        
End Select


End Sub

Private Sub tlbAutorizacion_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbAutoriza_Buscar
  Case "Autorizar"
    Call sbAutoriza("A")
  Case "Desautorizar"
    Call sbAutoriza("D")
  Case "Reporte"
    Call Excel_Exportar_Lsw(lswAut, ProgressBarX)
End Select

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

tcMain.top = lsw.top
tcMain.Left = lsw.Left
tcMain.Width = lsw.Width
tcMain.Height = lsw.Height + 450



scAutorizaciones.Width = lsw.Width
lswAut.Width = lsw.Width
lswAut.Height = lsw.Height - 450


scActivos_Trasladar.Width = lsw.Width
tcTraslados.Width = tcMain.Width
scActivos.Width = lsw.Width
lswAF_Activos.Width = lsw.Width
lswAF_Boletas.Width = lsw.Width
lswAF_Recepcion.Width = lsw.Width

lswAF_Activos.Height = pHeight - (lswAF_Activos.top + 2550)

vGrid.Height = lswAF_Activos.Height



tcDeclaracion.Width = tcMain.Width
tcDeclaracion.Height = pHeight - (tcDeclaracion.top + 2550)

lswDeclara.Width = lsw.Width
lswDeclara_H.Width = lsw.Width
lswDeclara_H_Det.Width = lsw.Width


lswDeclara.Height = pHeight - (lswDeclara.top + 2550)
lswDeclara_H_Det.Height = pHeight - (lswDeclara_H_Det.top + 2550)



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

tcMain.Visible = pVisible

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

    mNomina = rs!COD_NOMINA

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

Private Sub txtNuevoPersona_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    gBusquedas.Consulta = "select identificacion,Nombre from Activos_Personas"
    gBusquedas.Filtro = " and Identificacion <> '" & txtIdentificacion.Text & "'"
    frmBusquedas.Show vbModal
    txtNuevoPersona.Tag = gBusquedas.Resultado
    txtNuevoPersona.Text = gBusquedas.Resultado2
    Call sbResponsableNuevo(txtNuevoPersona.Tag)
    
End If

End Sub


Private Sub sbResponsableNuevo(pIdentificacion As String)
On Error GoTo vError

txtNuevoDepartamento.Text = ""
txtNuevoSeccion.Text = ""

txtNuevoDepartamento.Tag = ""
txtNuevoSeccion.Tag = ""

strSQL = "select * from vActivos_Personas where identificacion = '" & pIdentificacion & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
    txtNuevoDepartamento.Text = rs!departamento
    txtNuevoSeccion.Text = rs!seccion
    
    txtNuevoDepartamento.Tag = rs!Cod_Departamento
    txtNuevoSeccion.Tag = rs!Cod_Seccion
End If
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

