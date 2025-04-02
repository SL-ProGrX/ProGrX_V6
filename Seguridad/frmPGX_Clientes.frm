VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmPGX_Clientes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11295
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6255
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   11033
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
      ItemCount       =   5
      Item(0).Caption =   "General"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "GroupBox1(4)"
      Item(0).Control(1)=   "GroupBox1(3)"
      Item(1).Caption =   "Suscripción y Conexiones"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "GroupBox1(0)"
      Item(1).Control(1)=   "GroupBox1(2)"
      Item(1).Control(2)=   "GroupBox1(1)"
      Item(2).Caption =   "Servicios"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lsw"
      Item(3).Caption =   "Contactos"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "vGrid"
      Item(4).Caption =   "SMTP Autorizados"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "lswSMTP"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   98
         Top             =   360
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   10398
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lswSMTP 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   62
         Top             =   360
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   10398
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1575
         Index           =   0
         Left            =   -70000
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "     Suscripción"
         ForeColor       =   4210752
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   330
            Left            =   1080
            TabIndex        =   16
            Top             =   480
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
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
         Begin XtremeSuiteControls.ComboBox cboVendedor 
            Height          =   330
            Left            =   1080
            TabIndex        =   17
            Top             =   840
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
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
         Begin XtremeSuiteControls.DateTimePicker dtpSuscripcionInicia 
            Height          =   330
            Left            =   5160
            TabIndex        =   18
            Top             =   480
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
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
         Begin XtremeSuiteControls.DateTimePicker dtpSuscripcionVence 
            Height          =   330
            Left            =   5160
            TabIndex        =   19
            Top             =   840
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
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
         Begin XtremeSuiteControls.FlatEdit txtSuscripcionMensualidad 
            Height          =   330
            Left            =   8160
            TabIndex        =   20
            Top             =   480
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.FlatEdit txtSuscripcionAnualidad 
            Height          =   330
            Left            =   8160
            TabIndex        =   21
            Top             =   840
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Estado"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   26
            Top             =   480
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Inicio"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   2
            Left            =   4440
            TabIndex        =   25
            Top             =   840
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Vence"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   3
            Left            =   6960
            TabIndex        =   24
            Top             =   480
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Mensualidad"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   4
            Left            =   6960
            TabIndex        =   23
            Top             =   840
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anualidad"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Agente"
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2655
         Index           =   2
         Left            =   -70000
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   4683
         _StockProps     =   79
         Caption         =   "    Conexiones"
         ForeColor       =   4210752
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnConecTest 
            Height          =   310
            Index           =   0
            Left            =   10200
            TabIndex        =   99
            Top             =   480
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1085
            _ExtentY        =   547
            _StockProps     =   79
            Caption         =   "Test"
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
         Begin XtremeSuiteControls.CheckBox chkPruebas 
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Top             =   2040
            Width           =   3255
            _Version        =   1441793
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Activar Escenario de Pruebas"
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
         Begin XtremeSuiteControls.FlatEdit txtCoreServer 
            Height          =   330
            Left            =   1200
            TabIndex        =   37
            Top             =   480
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4895
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
         Begin XtremeSuiteControls.FlatEdit txtCoreDB 
            Height          =   330
            Left            =   3960
            TabIndex        =   38
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtCoreUser 
            Height          =   330
            Left            =   6000
            TabIndex        =   39
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtCoreClave 
            Height          =   330
            Left            =   8040
            TabIndex        =   40
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
            PasswordChar    =   "*"
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAnalisisServer 
            Height          =   330
            Left            =   1200
            TabIndex        =   41
            Top             =   840
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4895
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
         Begin XtremeSuiteControls.FlatEdit txtAnalisisDB 
            Height          =   330
            Left            =   3960
            TabIndex        =   42
            Top             =   840
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtAnalisisUser 
            Height          =   330
            Left            =   6000
            TabIndex        =   43
            Top             =   840
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtAnalisisClave 
            Height          =   330
            Left            =   8040
            TabIndex        =   44
            Top             =   840
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
            PasswordChar    =   "*"
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAuxiliarServer 
            Height          =   330
            Left            =   1200
            TabIndex        =   45
            Top             =   1200
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4895
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
         Begin XtremeSuiteControls.FlatEdit txtAuxiliarDB 
            Height          =   330
            Left            =   3960
            TabIndex        =   46
            Top             =   1200
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtAuxiliarUser 
            Height          =   330
            Left            =   6000
            TabIndex        =   47
            Top             =   1200
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtAuxiliarClave 
            Height          =   330
            Left            =   8040
            TabIndex        =   48
            Top             =   1200
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
            PasswordChar    =   "*"
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPruebasServer 
            Height          =   330
            Left            =   1200
            TabIndex        =   49
            Top             =   1560
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4895
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
         Begin XtremeSuiteControls.FlatEdit txtPruebasDB 
            Height          =   330
            Left            =   3960
            TabIndex        =   50
            Top             =   1560
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtPruebasUser 
            Height          =   330
            Left            =   6000
            TabIndex        =   51
            Top             =   1560
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtPruebasClave 
            Height          =   330
            Left            =   8040
            TabIndex        =   52
            Top             =   1560
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
            PasswordChar    =   "*"
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnConecTest 
            Height          =   315
            Index           =   1
            Left            =   10200
            TabIndex        =   100
            Top             =   840
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1085
            _ExtentY        =   547
            _StockProps     =   79
            Caption         =   "Test"
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
         Begin XtremeSuiteControls.PushButton btnConecTest 
            Height          =   315
            Index           =   2
            Left            =   10200
            TabIndex        =   101
            Top             =   1200
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1085
            _ExtentY        =   547
            _StockProps     =   79
            Caption         =   "Test"
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
         Begin XtremeSuiteControls.PushButton btnConecTest 
            Height          =   315
            Index           =   3
            Left            =   10200
            TabIndex        =   102
            Top             =   1560
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1085
            _ExtentY        =   547
            _StockProps     =   79
            Caption         =   "Test"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   36
            Top             =   1560
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Pruebas"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   35
            Top             =   1200
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Auxiliares"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   34
            Top             =   840
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Análisis"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   33
            Top             =   480
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Core"
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Servidor"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   1200
            TabIndex        =   32
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Base de Datos"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   3960
            TabIndex        =   31
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   6000
            TabIndex        =   30
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "Clave"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   8040
            TabIndex        =   29
            Top             =   240
            Width           =   2055
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1695
         Index           =   1
         Left            =   -70000
         TabIndex        =   54
         Top             =   4560
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "     URL para AutoGestión"
         ForeColor       =   4210752
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkURL_App 
            Height          =   330
            Left            =   360
            TabIndex        =   55
            Top             =   360
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Apps"
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
         Begin XtremeSuiteControls.CheckBox chkURL_Web 
            Height          =   330
            Left            =   360
            TabIndex        =   56
            Top             =   720
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Web"
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
         Begin XtremeSuiteControls.CheckBox chkURL_Logo 
            Height          =   330
            Left            =   360
            TabIndex        =   57
            Top             =   1080
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Logo"
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
         Begin XtremeSuiteControls.FlatEdit txtURL_Apps 
            Height          =   330
            Left            =   1680
            TabIndex        =   58
            Top             =   360
            Width           =   8415
            _Version        =   1441793
            _ExtentX        =   14843
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
         Begin XtremeSuiteControls.FlatEdit txtURL_Web 
            Height          =   330
            Left            =   1680
            TabIndex        =   59
            Top             =   720
            Width           =   8415
            _Version        =   1441793
            _ExtentX        =   14843
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
         Begin XtremeSuiteControls.FlatEdit txtURL_Logo 
            Height          =   330
            Left            =   1680
            TabIndex        =   60
            Top             =   1080
            Width           =   8415
            _Version        =   1441793
            _ExtentX        =   14843
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
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   10935
         _Version        =   524288
         _ExtentX        =   19288
         _ExtentY        =   10398
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
         MaxCols         =   488
         SpreadDesigner  =   "frmPGX_Clientes.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3855
         Index           =   4
         Left            =   0
         TabIndex        =   63
         Top             =   360
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   6800
         _StockProps     =   79
         Caption         =   "     Datos"
         ForeColor       =   4210752
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
         BorderStyle     =   2
         Begin XtremeSuiteControls.ComboBox cboTipoId 
            Height          =   330
            Left            =   1440
            TabIndex        =   75
            Top             =   480
            Width           =   2295
            _Version        =   1441793
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
            Height          =   330
            Left            =   1440
            TabIndex        =   78
            Top             =   840
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.ComboBox cboClasificacion 
            Height          =   330
            Left            =   1440
            TabIndex        =   79
            Top             =   1320
            Width           =   2295
            _Version        =   1441793
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono 
            Height          =   330
            Left            =   1440
            TabIndex        =   84
            Top             =   1800
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   330
            Left            =   1440
            TabIndex        =   85
            Top             =   2280
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtCelular 
            Height          =   330
            Left            =   1440
            TabIndex        =   86
            Top             =   2760
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
            Height          =   330
            Left            =   1440
            TabIndex        =   87
            Top             =   3240
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtNombreLargo 
            Height          =   1170
            Left            =   5280
            TabIndex        =   88
            Top             =   480
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEMail 
            Height          =   330
            Left            =   5280
            TabIndex        =   90
            Top             =   1800
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
         Begin XtremeSuiteControls.FlatEdit txtEMail2 
            Height          =   330
            Left            =   5280
            TabIndex        =   91
            Top             =   2280
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
         Begin XtremeSuiteControls.FlatEdit txtWebSite 
            Height          =   330
            Left            =   5280
            TabIndex        =   92
            Top             =   2760
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
         Begin XtremeSuiteControls.FlatEdit txtFacebook 
            Height          =   330
            Left            =   5280
            TabIndex        =   93
            Top             =   3240
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   11
            Left            =   4080
            TabIndex        =   97
            Top             =   3240
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Facebook"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   10
            Left            =   4080
            TabIndex        =   96
            Top             =   2760
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sitio Web"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   9
            Left            =   4080
            TabIndex        =   95
            Top             =   2280
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Email (2)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   8
            Left            =   4080
            TabIndex        =   94
            Top             =   1800
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Email (1)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   495
            Index           =   7
            Left            =   4080
            TabIndex        =   89
            Top             =   600
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Nombre Largo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   83
            Top             =   3240
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Apto.Postal"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   82
            Top             =   2760
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Móvil"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   81
            Top             =   2280
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Teléfono (2)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   80
            Top             =   1800
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Teléfono (1)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   77
            Top             =   1320
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Clasificación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   840
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Identificación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   480
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo Id"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2055
         Index           =   3
         Left            =   0
         TabIndex        =   64
         Top             =   4200
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   3625
         _StockProps     =   79
         Caption         =   "     Dirección"
         ForeColor       =   4210752
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboPais 
            Height          =   330
            Left            =   1320
            TabIndex        =   69
            Top             =   360
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   330
            Left            =   1320
            TabIndex        =   70
            Top             =   720
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   330
            Left            =   1320
            TabIndex        =   71
            Top             =   1080
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   330
            Left            =   1320
            TabIndex        =   72
            Top             =   1440
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   1410
            Left            =   3960
            TabIndex        =   73
            Top             =   360
            Width           =   6735
            _Version        =   1441793
            _ExtentX        =   11880
            _ExtentY        =   2487
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
         Begin VB.Label lblPais 
            BackStyle       =   0  'Transparent
            Caption         =   "País"
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
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblPais 
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel_1"
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
            Left            =   240
            TabIndex        =   67
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblPais 
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel_2"
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
            Index           =   2
            Left            =   240
            TabIndex        =   66
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblPais 
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel_3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   65
            Top             =   1440
            Width           =   975
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox gbSincroniza 
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   7440
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Sincronización con el CORE"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkSincroniza 
         Height          =   252
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Personeria"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnSincronizaIds 
         Height          =   372
         Left            =   7680
         TabIndex        =   6
         Top             =   240
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Sincroniza"
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
      End
      Begin XtremeSuiteControls.CheckBox chkSincroniza 
         Height          =   252
         Index           =   1
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Logos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkSincroniza 
         Height          =   252
         Index           =   2
         Left            =   5160
         TabIndex        =   9
         Top             =   360
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Id de Portal"
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
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkSincroniza 
         Height          =   252
         Index           =   3
         Left            =   7680
         TabIndex        =   11
         Top             =   720
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Sincroniza Cliente Actual"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.Label lblStatus 
         Height          =   252
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   8172
         _Version        =   1441793
         _ExtentX        =   14414
         _ExtentY        =   444
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
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9840
      TabIndex        =   0
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Clientes.frx":06CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Clientes.frx":3B5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Clientes.frx":6FEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Clientes.frx":710D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   11295
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinWidth1       =   1800
      MinHeight1      =   330
      Width1          =   1800
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   1110
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   10290
         TabIndex        =   3
         Top             =   30
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1560
      TabIndex        =   12
      Top             =   600
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   2760
      TabIndex        =   13
      Top             =   600
      Width           =   6975
      _Version        =   1441793
      _ExtentX        =   12303
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmPGX_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, mCodigo As String
Dim vScroll As Boolean, vPaso As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem


Private Sub btnConecTest_Click(Index As Integer)

Dim strSQL As String

Dim db As New ADODB.Connection

On Error GoTo vError

Me.MousePointer = vbHourglass


Dim pServer As String, pDb As String, pUser As String, pKey As String
Dim Coneccion As String

Select Case Index
 Case 0 'Core
    Coneccion = "Conección al CORE"
    pServer = txtCoreServer.Text
    pDb = txtCoreDB.Text
    pUser = txtCoreUser.Text
    pKey = txtCoreClave.Text
 Case 1 'Analisis
    Coneccion = "Conección a Analisis"
    pServer = txtAnalisisServer.Text
    pDb = txtAnalisisDB.Text
    pUser = txtAnalisisUser.Text
    pKey = txtAnalisisClave.Text
 Case 2 'Auxiliares
    Coneccion = "Conección a Auxiliares"
    pServer = txtAuxiliarServer.Text
    pDb = txtAuxiliarDB.Text
    pUser = txtAuxiliarUser.Text
    pKey = txtAuxiliarClave.Text
 Case 3 'Pruebas
    Coneccion = "Conección a Pruebas"
    pServer = txtPruebasServer.Text
    pDb = txtPruebasDB.Text
    pUser = txtPruebasUser.Text
    pKey = txtPruebasClave.Text
    
End Select


strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & Trim(pServer) _
       & ";Database=" & Trim(pDb) & ";APP=PGX_Portal;tcp:" _
       & Trim(pServer) & "," & SIFGlobal.PuertosDisponibles & ";"


'Connección Principal
db.CommandTimeout = 15
db.Mode = adModeReadWrite
db.CursorLocation = adUseClient
db.Open strSQL, Trim(pUser), Trim(pKey)
db.CommandTimeout = 60

db.Close
  
Me.MousePointer = vbDefault

MsgBox Coneccion & " --> Exitosa!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Coneccion & " --> FALLIDA!", vbInformation

End Sub


Private Sub sbSmtps_Load()

Dim db As New ADODB.Connection

lswSMTP.ListItems.Clear

If vPaso Then Exit Sub
If Not IsNumeric(txtCodigo.Text) Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

'strSQL = "select COD_EMPRESA, NOMBRE_CORTO , NOMBRE_LARGO, IDENTIFICACION" _
'       & ", PGX_CORE_DB , PGX_CORE_SERVER , PGX_CORE_USER , PGX_CORE_KEY  " _
'       & ", URL_LOGO, URL_Logo_Activo" _
'       & "  From PGX_CLIENTES where COD_EMPRESA = " & txtCodigo.Text
'
'Call OpenRecordSet(rs, strSQL)
'
'
'  strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & Trim(rs!PGX_Core_Server) _
'         & ";Database=" & Trim(rs!PGX_BASE) & ";APP=PGX_Portal;tcp:" _
'         & Trim(rs!PGX_Core_Server) & "," & SIFGlobal.PuertosDisponibles & ";"
'
'
'  'Connección Principal
'  db.CommandTimeout = 15
'  db.Mode = adModeReadWrite
'  db.CursorLocation = adUseClient
'  db.Open strSQL, Trim(rs!PGX_Core_User), Trim(rs!PGX_Core_Key)
'  db.CommandTimeout = 60
'
  
   strSQL = "exec spPGX_SMTP_Lista " & txtCodigo.Text
   Call OpenRecordSet(rs, strSQL)
   
   vPaso = True
   
   Do While Not rs.EOF
    Set itmX = lswSMTP.ListItems.Add(, , rs!cod_SMTP)
        itmX.SubItems(1) = rs!Descripcion
        
        itmX.Checked = rs!Asignado
        
     rs.MoveNext
   Loop
   rs.Close
    
   vPaso = False
  
  
  
'  db.Close
  

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical

End Sub

Private Sub btnSincronizaIds_Click()

Dim db As New ADODB.Connection, i As Long

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select COD_EMPRESA, NOMBRE_CORTO , NOMBRE_LARGO, IDENTIFICACION" _
       & ", PGX_CORE_DB , PGX_CORE_SERVER , PGX_CORE_USER , PGX_CORE_KEY  " _
       & ", URL_LOGO, URL_Logo_Activo" _
       & "  From PGX_CLIENTES where estado = 'A'"
       
If chkSincroniza.Item(3).Value = xtpChecked Then
    strSQL = strSQL & " and COD_EMPRESA = " & txtCodigo.Text
End If
       
Call OpenRecordSet(rs, strSQL)

lblStatus.Caption = "Cargando..."
i = 0
Do While Not rs.EOF
  i = i + 1
  lblStatus.Caption = "Sincronizando: " & rs!Nombre_Corto & Space(10) & " [" & i & " - " & rs.RecordCount & "]"
    
  DoEvents
   
  strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & Trim(rs!PGX_Core_Server) _
         & ";Database=" & Trim(rs!PGX_Core_Db) & ";APP=PGX_Portal;tcp:" _
         & Trim(rs!PGX_Core_Server) & "," & SIFGlobal.PuertosDisponibles & ";"
  
  
  'Connección Principal
  db.CommandTimeout = 15
  db.Mode = adModeReadWrite
  db.CursorLocation = adUseClient
  db.Open strSQL, Trim(rs!PGX_Core_User), Trim(rs!PGX_Core_Key)
  db.CommandTimeout = 60
  
  strSQL = "Update SIF_EMPRESA SET PORTAL_ID = " & rs!cod_Empresa
  
  If chkSincroniza.Item(1).Value = xtpChecked And rs!URL_Logo_Activo = 1 Then
    strSQL = strSQL & ",LOGO_WEB_SITE = '" & Trim(rs!URL_LOGO) & "'"
  End If
  
  db.Execute strSQL
  
  db.Close
  
  rs.MoveNext
Loop
rs.Close

lblStatus.Caption = ""
MsgBox "Sincronización Finalizada!", vbInformation

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical

End Sub

Private Sub cboCanton_Click()

If vPaso Then Exit Sub

strSQL = "select rtrim(cod_pais_n3) as 'IdX', rTrim(descripcion) as 'itmX' " _
       & " from PGX_PAIS_N3" _
       & " where COD_PAIS = '" & cboPais.ItemData(cboPais.ListIndex) & "'" _
       & " and COD_PAIS_N1 = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'" _
       & " and COD_PAIS_N2 = '" & cboCanton.ItemData(cboCanton.ListIndex) & "'"

vPaso = True
 Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)
vPaso = False

End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub


Private Sub cboClasificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus
End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub cboPais_Click()

If vPaso Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select * from pgx_Pais where cod_pais = '" & cboPais.ItemData(cboPais.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  lblPais(1).Caption = rs!N1_Nombre & ""
  lblPais(2).Caption = rs!N2_Nombre & ""
  lblPais(3).Caption = rs!N3_Nombre & ""
End If
rs.Close

strSQL = "select rtrim(cod_pais_n1) as 'IdX', rtrim(descripcion) as 'itmX' " _
       & " from PGX_PAIS_N1" _
       & " where COD_PAIS = '" & cboPais.ItemData(cboPais.ListIndex) & "'"

vPaso = True
 Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)

vPaso = False

Call cboProvincia_Click

End Sub

Private Sub cboPais_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

strSQL = "select rtrim(cod_pais_n2) as 'IdX',  rTrim(descripcion) as 'itmX' " _
       & " from PGX_PAIS_N2" _
       & " where COD_PAIS = '" & cboPais.ItemData(cboPais.ListIndex) & "'" _
       & " and COD_PAIS_N1 = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"

vPaso = True
 Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click

End Sub


Private Sub cboSexo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub



Private Sub cboTipoId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdentificacion.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError


If txtCodigo.Text = "" Then txtCodigo.Text = "9999"

If vScroll Then
    strSQL = "select Top 1 cod_Empresa from PGX_Clientes"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_Empresa > '" & txtCodigo.Text & "' order by cod_Empresa asc"
    Else
       strSQL = strSQL & " where cod_Empresa < '" & txtCodigo.Text & "' order by cod_Empresa desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_Empresa
      Call txtCodigo_LostFocus
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_activate()
vModulo = 31
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 31

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
 
With lsw.ColumnHeaders
     .Clear
     .Add , , "Id", 1200
     .Add , , "Descripción", 3200
     .Add , , "Monto", 1400, vbRightJustify
     .Add , , "Costo", 1400, vbRightJustify
     .Add , , "Qty User", 1200, vbCenter
     .Add , , "R.Fecha", 1800
     .Add , , "R.Usuario", 1800
End With

With lswSMTP.ColumnHeaders
    .Clear
    .Add , , "SMTP ID", 1200
    .Add , , "DESCRIPCION", 5000
End With
 
tcMain.Item(0).Selected = True

vEdita = False

'Carga Tipos de Identificacion
vPaso = True
strSQL = "select rtrim(TIPO_ID) as  'IdX',  rtrim(Descripcion) as 'ItmX' from PGX_TIPOS_ID" _
       & " Where Activa = 1"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False

'Carga Clasificaciones de Clientes
strSQL = "select rtrim(cod_Clasificacion) as 'IdX', rtrim(descripcion) as 'ItmX' from PGX_Clientes_Clasificacion" _
       & " Where Activa = 1"
Call sbCbo_Llena_New(cboClasificacion, strSQL, False, True)

'Vendedores
strSQL = "select rtrim(cod_Vendedor)  as 'IdX', rtrim(Nombre) as 'ItmX' from PGX_Vendedores" _
       & " Where Activo = 1"
Call sbCbo_Llena_New(cboVendedor, strSQL, False, True)

'Estados
cboEstado.Clear
cboEstado.AddItem "Activa"
cboEstado.AddItem "Inactiva"
cboEstado.AddItem "Congelada"
cboEstado.AddItem "Vencida"
cboEstado.Text = "Activa"

'Carga Paises
strSQL = "select rtrim(cod_Pais) as 'IdX', rtrim(descripcion) as 'ItmX' from PGX_Pais"
vPaso = True
    Call sbCbo_Llena_New(cboPais, strSQL, False, True)
vPaso = False


 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

 Call cboPais_Click

Exit Sub

vError:
  MsgBox Err.Description, vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
mCodigo = ""
txtCodigo = ""

txtNombre = ""
txtTelefono.Text = ""
txtTelefono2.Text = ""
txtCelular.Text = ""

txtWebSite.Text = ""
txtEMail.Text = ""
txtEMail2.Text = ""
txtAptoPostal.Text = ""
txtFacebook.Text = ""

txtDireccion = ""

txtNombreLargo.Text = ""

txtURL_Apps.Text = ""
txtURL_Logo.Text = ""
txtURL_Web.Text = ""

chkURL_App.Value = vbUnchecked
chkURL_Web.Value = vbUnchecked
chkURL_Logo.Value = vbUnchecked

txtCoreServer.Text = ""
txtCoreDB.Text = ""
txtCoreUser.Text = ""
txtCoreClave.Text = ""

txtAnalisisServer.Text = ""
txtAnalisisDB.Text = ""
txtAnalisisUser.Text = ""
txtAnalisisClave.Text = ""

txtPruebasServer.Text = ""
txtPruebasDB.Text = ""
txtPruebasUser.Text = ""
txtPruebasClave.Text = ""

txtAuxiliarServer.Text = ""
txtAuxiliarDB.Text = ""
txtAuxiliarUser.Text = ""
txtAuxiliarClave.Text = ""

chkPruebas.Value = vbUnchecked

txtSuscripcionAnualidad.Text = "0"
txtSuscripcionMensualidad.Text = "0"
dtpSuscripcionInicia.Value = fxFechaServidor
dtpSuscripcionVence.Value = dtpSuscripcionInicia.Value


tcMain.Item(0).Selected = True
 
tcMain.Item(1).Enabled = True
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False

tlbAux.Buttons.Item(1).Enabled = False 'Nuevo
tlbAux.Buttons.Item(3).Enabled = False 'Borrar


End Sub



Private Sub sbConsultaContratoDetalle()

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

''Carga Pagadores
'lswP.ListItems.Clear
'strSQL = " select P.nombre, C.*" _
'     & " from PGX_Clientes P inner join PGX_Clientes_Contratos_Pagadores C on P.cod_Empresa = C.cod_Empresa_pagador" _
'     & " where C.cod_contrato = '" & txtCodigo.Text & "' and C.cod_Empresa = '" & mCodigo & "'"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
'    Set itmX = lswP.ListItems.Add(, , rs!cod_Empresa_Pagador)
'        itmX.SubItems(1) = rs!nombre
'        itmX.SubItems(2) = rs!Registro_usuario & "..." & rs!Registro_Fecha
'    rs.MoveNext
'Loop
'rs.Close
'
'
''Carga Cargos de Suscripción
'lswC.ListItems.Clear
'strSQL = " select C.descripcion,S.*" _
'       & " from CxC_Cargos C inner join PGX_Clientes_Contratos_Suscripciones S on C.cod_cargo = S.cod_cargo" _
'       & " where S.cod_contrato = '" & txtCodigo.Text & "' and S.cod_Empresa = '" & mCodigo & "'"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
'    Set itmX = lswC.ListItems.Add(, , rs!COD_CARGO)
'        itmX.SubItems(1) = rs!Descripcion
'
'        Select Case rs!Tipo
'           Case "P"
'             itmX.SubItems(2) = "Porcentual"
'           Case "M"
'             itmX.SubItems(2) = "Monto"
'        End Select
'
'
'        Select Case rs!Frecuencia_Tipo
'          Case "O"
'            itmX.SubItems(4) = "Operación"
'          Case "D"
'            itmX.SubItems(4) = "Días"
'        End Select
'
'        itmX.SubItems(3) = Format(rs!Valor, "Standard")
'        itmX.SubItems(5) = rs!Frecuencia_dias
'        itmX.SubItems(6) = Format(rs!Recaudado, "Standard")
'        itmX.SubItems(7) = Format(rs!Pago_Ultimo, "dd/mm/yyyy")
'        itmX.SubItems(8) = Format(rs!Pago_Proximo, "dd/mm/yyyy")
'        itmX.SubItems(9) = rs!Modifica
'        itmX.Checked = True
'    rs.MoveNext
'Loop
'rs.Close


vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub



Private Sub lswSMTP_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spPGX_SMTP_Autoriza " & mCodigo & ",'" & Item.Text _
        & "','" & glogon.Usuario & "','" & IIf(Item.Checked, "A", "B") & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox Err.Description, vbExclamation

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim curMonto As Currency, curSaldo As Currency

tlbAux.Buttons.Item(1).Enabled = False 'Nuevo
tlbAux.Buttons.Item(3).Enabled = False 'Borrar

If mCodigo = "" Then
  tcMain.Item(0).Selected = True
  Exit Sub
End If

Me.MousePointer = vbHourglass

Select Case Item.Index

Case 2
   tlbAux.Buttons.Item(1).Enabled = True 'Nuevo
   tlbAux.Buttons.Item(3).Enabled = True 'Borrar
   
    'Servicios
   strSQL = "select S.Cod_Servicio,S.Descripcion,A.Monto,A.Costo,A.Cantidad_Usuarios,A.Registro_Fecha,A.Registro_Usuario" _
          & " from PGX_Servicios S inner join PGX_Servicios_ASG A on S.cod_Servicio = A.Cod_Servicio" _
          & " where A.cod_Empresa = " & mCodigo & " and A.Activo = 1"
   lsw.ListItems.Clear
   
   Call OpenRecordSet(rs, strSQL)
   
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!cod_Servicio)
        itmX.SubItems(1) = rs!Descripcion
        itmX.SubItems(2) = rs!Monto
        itmX.SubItems(3) = rs!Costo
        itmX.SubItems(4) = rs!Cantidad_Usuarios
        itmX.SubItems(5) = rs!Registro_Fecha
        itmX.SubItems(6) = rs!Registro_Usuario
     rs.MoveNext
   Loop
   rs.Close
   
Case 3
   'Contactos
   strSQL = "select cod_Contacto,identificacion,nombre,tel_cell,tel_trabajo,Email_01,Email_02,Activo " _
          & " from PGX_Clientes_Contactos" _
          & " where cod_Empresa = " & mCodigo
    Call sbCargaGrid(vGrid, 8, strSQL)


Case 4
    'SMPTP
   Call sbSmtps_Load

End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If mCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(mCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Col1Name = "Id"
       gBusquedas.Col2Name = "Acronimo"
       gBusquedas.Col3Name = "Nombre"
       gBusquedas.Columna = "Nombre_Corto"
       gBusquedas.Orden = "Nombre_Corto"
       gBusquedas.Consulta = "select cod_Empresa, Nombre_Corto, Nombre_Largo from PGX_Clientes"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(xCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.*" _
       & ", Case when Estado = 'A' then 'Activa' when Estado = 'I' then 'Inactiva'" _
       & " when Estado = 'C' then 'Congelada' when Estado = 'V' then 'Vencida' else 'Activa' end as 'EstadoDesc'" _
       & ", rtrim(Tid.Descripcion) as 'TipoIdDesc'" _
       & ", rtrim(Cat.Descripcion) as 'ClasificacionDesc'" _
       & ", rtrim(Ven.Nombre) as 'VendedorDesc'" _
       & ", rtrim(Pa.Descripcion) as 'Pais'" _
       & ", rtrim(P1.Descripcion) as 'Pais_N1'" _
       & ", rtrim(P2.Descripcion) as 'Pais_N2'" _
       & ", rtrim(P3.Descripcion) as 'Pais_N3'" _
       & " from PGX_Clientes C " _
       & " left join PGX_Tipos_Id Tid on C.tipo_id = Tid.tipo_id" _
       & " left join PGX_Clientes_Clasificacion Cat on C.cod_Clasificacion = Cat.cod_Clasificacion" _
       & " left join PGX_Vendedores Ven on C.cod_Vendedor = Ven.cod_Vendedor" _
       & " left join PGX_Pais Pa on C.cod_Pais = Pa.cod_Pais" _
       & " left join PGX_Pais_N1 P1 on C.cod_Pais = P1.cod_Pais and C.cod_Pais_N1 = P1.cod_Pais_N1" _
       & " left join PGX_Pais_N2 P2 on C.cod_Pais = P2.cod_Pais and C.cod_Pais_N1 = P2.cod_Pais_N1 and C.cod_Pais_N2 = P2.cod_Pais_N2" _
       & " left join PGX_Pais_N3 P3 on C.cod_Pais = P3.cod_Pais and C.cod_Pais_N1 = P3.cod_Pais_N1 and C.cod_Pais_N2 = P3.cod_Pais_N2 and C.cod_Pais_N3 = P3.cod_Pais_N3" _
       & " where C.cod_Empresa = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  mCodigo = rs!cod_Empresa
  txtCodigo = rs!cod_Empresa

  txtIdentificacion.Text = rs!Identificacion
  txtNombre.Text = rs!Nombre_Corto & ""
  
  txtNombre.SetFocus
  
  txtNombreLargo.Text = rs!Nombre_Largo & ""
    
  txtTelefono.Text = rs!Tel_Trabajo & ""
  txtTelefono2.Text = rs!Tel_Auxiliar & ""
  txtCelular.Text = rs!Tel_Cell & ""
  txtAptoPostal.Text = rs!apto_postal & ""

  txtEMail.Text = rs!Email_01 & ""
  txtEMail2.Text = rs!Email_02 & ""
  txtWebSite.Text = rs!Web_Site & ""
  txtFacebook.Text = rs!Facebook & ""
  txtDireccion.Text = rs!Direccion

  Call sbCboAsignaDato(cboTipoId, rs!TipoIdDesc & "", True, rs!Tipo_Id)
  Call sbCboAsignaDato(cboClasificacion, rs!ClasificacionDesc & "", True, rs!cod_Clasificacion)
  Call sbCboAsignaDato(cboEstado, rs!EstadoDesc & "")
  Call sbCboAsignaDato(cboVendedor, rs!VendedorDesc & "", True, rs!Cod_Vendedor)

  
  Call sbCboAsignaDato(cboPais, rs!Pais & "", True, rs!cod_Pais)
  Call sbCboAsignaDato(cboProvincia, rs!Pais_N1 & "", True, rs!Cod_Pais_N1)
  Call sbCboAsignaDato(cboCanton, rs!Pais_N2 & "", True, rs!Cod_Pais_N2)
  Call sbCboAsignaDato(cboDistrito, rs!Pais_N3 & "", True, rs!Cod_Pais_N3)

  txtURL_Apps.Text = Trim(rs!URL_App & "")
  txtURL_Logo.Text = Trim(rs!URL_LOGO & "")
  txtURL_Web.Text = Trim(rs!URL_Web & "")
  
  chkURL_App.Value = rs!URL_App_Activo
  chkURL_Web.Value = rs!URL_Web_Activo
  chkURL_Logo.Value = rs!URL_Logo_Activo
  
    txtSuscripcionAnualidad.Text = Format(rs!Suscripcion_Anual, "Standard")
    txtSuscripcionMensualidad.Text = Format(rs!Suscripcion_Mensualidad, "Standard")
    dtpSuscripcionInicia.Value = rs!Suscripcion_Inicial
    dtpSuscripcionVence.Value = rs!Suscripcion_Vence
  
    txtCoreServer.Text = rs!PGX_Core_Server
    txtCoreDB.Text = rs!PGX_Core_Db
    txtCoreUser.Text = rs!PGX_Core_User
    txtCoreClave.Text = rs!PGX_Core_Key
    
    txtAnalisisServer.Text = rs!PGX_Analisis_Server
    txtAnalisisDB.Text = rs!PGX_Analisis_Db
    txtAnalisisUser.Text = rs!PGX_Analisis_User
    txtAnalisisClave.Text = rs!PGX_Analisis_Key
    
    txtPruebasServer.Text = rs!PGX_Pruebas_Server
    txtPruebasDB.Text = rs!PGX_Pruebas_Db
    txtPruebasUser.Text = rs!PGX_Pruebas_User
    txtPruebasClave.Text = rs!PGX_Pruebas_Key
    
    txtAuxiliarServer.Text = rs!PGX_Auxiliar_Server
    txtAuxiliarDB.Text = rs!PGX_Auxiliar_Db
    txtAuxiliarUser.Text = rs!PGX_Auxiliar_User
    txtAuxiliarClave.Text = rs!PGX_Auxiliar_Key


  tcMain.Item(0).Selected = True
  tcMain.Item(1).Enabled = True
  tcMain.Item(2).Enabled = True
  tcMain.Item(3).Enabled = True

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()

On Error GoTo vError

If vEdita Then
  strSQL = "update PGX_Clientes set Nombre_Corto = '" & Trim(txtNombre.Text) & "',Nombre_Largo= '" & Trim(txtNombreLargo.Text) _
         & "',Tipo_Id = '" & cboTipoId.ItemData(cboTipoId.ListIndex) & "',Identificacion = '" & txtIdentificacion.Text & "', Estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "',Tel_Trabajo = '" & txtTelefono.Text & "',Tel_Auxiliar = '" & txtTelefono2.Text & "',Tel_Cell = '" & txtCelular.Text & "',Facebook = '" & txtFacebook.Text _
         & "',Web_Site = '" & txtWebSite.Text & "',apto_postal = '" & txtAptoPostal.Text & "',email_01 = '" & txtEMail.Text & "', email_02 = '" & txtEMail2.Text _
         & "',direccion = '" & txtDireccion & "',cod_Pais = '" & cboPais.ItemData(cboPais.ListIndex) & "',cod_Pais_N1 = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & "',Cod_Pais_N2 = '" & cboCanton.ItemData(cboCanton.ListIndex) & "', Cod_Pais_N3 = '" & cboDistrito.ItemData(cboDistrito.ListIndex) _
         & "',cod_Clasificacion = '" & cboClasificacion.ItemData(cboClasificacion.ListIndex) & "', Cod_Vendedor = '" & cboVendedor.ItemData(cboVendedor.ListIndex) _
         & "',Suscripcion_mensualidad = " & CCur(txtSuscripcionMensualidad.Text) & ", Suscripcion_Anual = " & CCur(txtSuscripcionAnualidad.Text) _
         & ", Suscripcion_Inicial = '" & Format(dtpSuscripcionInicia.Value, "yyyy/mm/dd") & "', Suscripcion_Vence = '" & Format(dtpSuscripcionVence.Value, "yyyy/mm/dd") _
         & "',PGX_Core_Server = '" & Trim(txtCoreServer.Text) & "', PGX_Core_Db = '" & Trim(txtCoreDB.Text) & "',PGX_Core_User = '" & Trim(txtCoreUser.Text) & "', PGX_Core_Key = '" & Trim(txtCoreClave.Text) _
         & "',PGX_Pruebas_Server = '" & Trim(txtPruebasServer.Text) & "', PGX_Pruebas_Db = '" & Trim(txtPruebasDB.Text) & "',PGX_Pruebas_User = '" & Trim(txtPruebasUser.Text) & "', PGX_Pruebas_Key = '" & Trim(txtPruebasClave.Text) _
         & "',PGX_Analisis_Server = '" & Trim(txtAnalisisServer.Text) & "', PGX_Analisis_Db = '" & Trim(txtAnalisisDB.Text) & "',PGX_Analisis_User = '" & Trim(txtAnalisisUser.Text) & "', PGX_Analisis_Key = '" & Trim(txtAnalisisClave.Text) _
         & "',PGX_Auxiliar_Server = '" & Trim(txtAuxiliarServer.Text) & "', PGX_Auxiliar_Db = '" & Trim(txtAuxiliarDB.Text) & "',PGX_Auxiliar_User = '" & Trim(txtAuxiliarUser.Text) & "', PGX_Auxiliar_Key = '" & Trim(txtAuxiliarClave.Text) _
         & "',PGX_PRUEBAS_ACTIVO = " & chkPruebas.Value _
         & ", URL_App = '" & Trim(txtURL_Apps.Text) & "',URL_Web = '" & Trim(txtURL_Web.Text) & "', URL_Logo = '" & Trim(txtURL_Logo.Text) & "'" _
         & ", URL_App_Activo = " & chkURL_App.Value & ", URL_Web_Activo = " & chkURL_Web.Value & ", URL_Logo_Activo = " & chkURL_Logo.Value _
         & " where cod_Empresa = " & mCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Cliente Id:" & mCodigo)

Else

  strSQL = "select isnull(max(cod_empresa),0) + 1 as 'Codigo' from PGX_Clientes"
  Call OpenRecordSet(rs, strSQL)
  mCodigo = rs!Codigo
  rs.Close
  
   strSQL = "insert into PGX_Clientes(cod_Empresa,cod_Clasificacion,cod_Vendedor,Tipo_Id,Identificacion,Nombre_Corto,Nombre_Largo" _
          & ",Tel_Cell,Tel_Trabajo,Tel_Auxiliar,Apto_postal,email_01,email_02,web_Site,facebook" _
          & ",Cod_Pais,Cod_Pais_N1,Cod_Pais_N2,Cod_Pais_N3,direccion" _
          & ",Estado,Suscripcion_Inicial,Suscripcion_Vence,Suscripcion_Anual,Suscripcion_Mensualidad" _
          & ",PGX_Core_Server,PGX_Core_Db,PGX_Core_User,PGX_Core_Key" _
          & ",PGX_Pruebas_Server,PGX_Pruebas_Db,PGX_Pruebas_User,PGX_Pruebas_Key" _
          & ",PGX_Analisis_Server,PGX_Analisis_Db,PGX_Analisis_User,PGX_Analisis_Key" _
          & ",PGX_Auxiliar_Server,PGX_Auxiliar_Db,PGX_Auxiliar_User,PGX_Auxiliar_Key" _
          & ",PGX_PRUEBAS_ACTIVO,URL_App, URL_Web, URL_Logo, URL_App_Activo, URL_Web_Activo, URL_Logo_Activo, Registro_Fecha,Registro_Usuario)" _
          & " values(" & mCodigo & ",'" & cboClasificacion.ItemData(cboClasificacion.ListIndex) & "','" & cboVendedor.ItemData(cboVendedor.ListIndex) _
          & "','" & cboTipoId.ItemData(cboTipoId.ListIndex) & "','" & txtIdentificacion.Text & "','" & txtNombre.Text & "','" & txtNombreLargo.Text _
          & "','" & txtCelular.Text & "','" & txtTelefono & "','" & txtTelefono2 & "','" & txtAptoPostal.Text & "','" & txtEMail.Text & "','" _
          & txtEMail2.Text & "','" & txtWebSite.Text & "','" & txtFacebook.Text & "','" & cboPais.ItemData(cboPais.ListIndex) _
          & "','" & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & cboCanton.ItemData(cboCanton.ListIndex) & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) _
          & "','" & txtDireccion.Text & "','I',getdate(),getdate(), 0,0" _
          & ",'" & Trim(txtCoreServer.Text) & "','" & Trim(txtCoreDB.Text) & "','" & Trim(txtCoreUser.Text) & "','" & Trim(txtCoreClave.Text) _
          & "','" & Trim(txtPruebasServer.Text) & "','" & Trim(txtPruebasDB.Text) & "','" & Trim(txtPruebasUser.Text) & "','" & Trim(txtPruebasClave.Text) _
          & "','" & Trim(txtAnalisisServer.Text) & "','" & Trim(txtAnalisisDB.Text) & "','" & Trim(txtAnalisisUser.Text) & "','" & Trim(txtAnalisisClave.Text) _
          & "','" & Trim(txtAuxiliarServer.Text) & "','" & Trim(txtAuxiliarDB.Text) & "','" & Trim(txtAuxiliarUser.Text) & "','" & Trim(txtAuxiliarClave.Text) _
          & "'," & chkPruebas.Value & ",'" & Trim(txtURL_Apps.Text) & "','" & Trim(txtURL_Web.Text) & "','" & Trim(txtURL_Logo.Text) _
          & "'," & chkURL_App.Value & "," & chkURL_Web.Value & "," & chkURL_Logo.Value _
          & ",getdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Cliente Id:" & mCodigo)
   
   txtCodigo = mCodigo
   
End If

tcMain.Item(2).Enabled = True
tcMain.Item(3).Enabled = True

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete PGX_Clientes where cod_Empresa = '" & mCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Cliente Id:" & mCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long

GLOBALES.gTag = Trim(mCodigo)
GLOBALES.gTag2 = txtNombre.Text

Select Case Button.Key
 Case "nuevo"

    Call sbFormsCall("frmPGX_Servicios_Asignados", 1, , , False)
    tcMain.Item(2).Selected = True
 
 Case "borrar"
'    Select Case ssTab.Tab
'       Case 2 'Contratos
'
'          With lswP.ListItems
'          For i = 1 To .Count
'            If .Item(i).Checked Then
'               strSQL = "delete PGX_Clientes_Contratos_Pagadores where cod_contrato = '" & txtCodigo.Text _
'                      & "' and cod_Empresa = '" & mCodigo & "' and cod_Empresa_pagador = '" & .Item(i).Text & "'"
'               Call ConectionExecute(strSQL)
'
'                Call Bitacora("Borra", "Pagador Id.:" & .Item(i).Text & " de Contrato No.:" & txtCodigo.Text & " Ced:" & mCodigo)
'
'            End If
'          Next i
'          End With
'
'          With lswC.ListItems
'          For i = 1 To .Count
'            If .Item(i).Checked Then
'               strSQL = "delete PGX_Clientes_Contratos_Suscripciones where cod_contrato = '" & txtCodigo.Text _
'                      & "' and cod_cargo = '" & .Item(i).Text & "' and cod_Empresa = '" & mCodigo & "'"
'               Call ConectionExecute(strSQL)
'
'               Call Bitacora("Borra", "Cargo Suscripción Cod:" & .Item(i).Text & " Cnt: " & txtCodigo.Text)
'            End If
'          Next i
'          End With
'
'
'       Case 3 'Cuentas Bancarias
'          With lswBancos.ListItems
'          For i = 1 To .Count
'            If .Item(i).Checked Then
'               strSQL = "delete PGX_Clientes_Bancos where cod_Empresa = '" & mCodigo _
'                      & "' and id_Cuenta = " & .Item(i).Text
'               Call ConectionExecute(strSQL)
'               Call Bitacora("Elimina", "Cuenta Ahorros: " & .Item(i).SubItems(1) & " Id: " & .Item(i).Text & "_Ced:" & mCodigo)
'
'            End If
'          Next i
'          End With
'
'    End Select
'
'    Call sbConsultaContratoDetalle


End Select 'Toolbar

End Sub

Private Sub txtCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  Call sbCliente_Consulta
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
  Call sbConsulta(txtCodigo.Text)
End Sub


Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    tcMain.Item(1).Selected = True
    cboEstado.SetFocus
End If
End Sub

Private Sub txtEMail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
End Sub

Private Sub txtFacebook_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPais.SetFocus
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombreLargo.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoId.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre_Largo"
  gBusquedas.Orden = "Nombre_Largo"
  gBusquedas.Consulta = "select cod_Empresa,nombre_Corto,Nombre_Largo from PGX_Clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtNombreLargo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboClasificacion.SetFocus
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFacebook.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCelular.SetFocus
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEMail.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEMail2.SetFocus
End Sub

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEMail.SetFocus
End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


If vGrid.Text = "" Then

    strSQL = "select isnull(sum(cod_Contacto),0) + 1 as Consecutivo from PGX_Clientes_Contactos " _
           & " where cod_empresa = " & mCodigo
    Call OpenRecordSet(rs, strSQL)
    
    vGrid.Text = rs!Consecutivo

    rs.Close
    

  strSQL = "insert into PGX_Clientes_Contactos(cod_empresa,cod_contacto,identificacion,nombre,Tel_Cell" _
         & ",Tel_Trabajo,Email_01,Email_02,activo,registro_fecha,registro_usuario) values(" & mCodigo & "," _
         & vGrid.Text & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & ",Getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Cliente Contacto: (" & mCodigo & ") -> (" & vGrid.Text & ")")

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update PGX_Clientes_Contactos set Identificacion = '" & vGrid.Text & "',Nombre = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "', Tel_Cell = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & "', Tel_Trabajo = '"
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Text & "', Email_01 = '"
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Text & "', Email_02 = '"
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Text & "', Activo = "
 vGrid.Col = 8
 strSQL = strSQL & vGrid.Value & " where cod_Empresa = " & mCodigo & " and cod_Contacto = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Cliente Contacto: (" & mCodigo & ") -> (" & vGrid.Text & ")")

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete PGX_Clientes_Contactos where cod_empresa = " & mCodigo & " and cod_contacto = " & vGrid.Text
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Cliente Contacto: (" & mCodigo & ") -> (" & vGrid.Text & ")")

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

