VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_CtaCatalogo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas Contables de la línea de Crédito"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   HelpContextID   =   3007
   Icon            =   "frmCR_CtaCatalogo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6255
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
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
      ItemCount       =   2
      Item(0).Caption =   "Bloque No.1"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "gbNormal"
      Item(0).Control(1)=   "GroupBox1"
      Item(0).Control(2)=   "Label1(2)"
      Item(0).Control(3)=   "txtPuente"
      Item(0).Control(4)=   "txtPuente_Desc"
      Item(0).Control(5)=   "txtCtaIVA"
      Item(0).Control(6)=   "txtCtaIVA_Desc"
      Item(0).Control(7)=   "Label1(7)"
      Item(0).Control(8)=   "Label1(8)"
      Item(0).Control(9)=   "chkIVA"
      Item(0).Control(10)=   "txtAnticipo"
      Item(0).Control(11)=   "txtAnticipo_Desc"
      Item(1).Caption =   "Bloque No.2"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "GroupBox2"
      Item(1).Control(1)=   "GroupBox3"
      Begin XtremeSuiteControls.GroupBox gbNormal 
         Height          =   1695
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   10695
         _Version        =   1441793
         _ExtentX        =   18865
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Cuentas Contables [Operaciones en Estado Normal]"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCorEC 
            Height          =   330
            Left            =   2160
            TabIndex        =   4
            Top             =   480
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCorEC_Desc 
            Height          =   330
            Left            =   4320
            TabIndex        =   5
            Top             =   480
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntMorEC 
            Height          =   330
            Left            =   2160
            TabIndex        =   6
            Top             =   840
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntMorEC_Desc 
            Height          =   330
            Left            =   4320
            TabIndex        =   7
            Top             =   840
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaAmortEC 
            Height          =   330
            Left            =   2160
            TabIndex        =   8
            Top             =   1200
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaAmortEC_Desc 
            Height          =   330
            Left            =   4320
            TabIndex        =   9
            Top             =   1200
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Interés Corriente"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Interés Moratorio"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Amortización"
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1695
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   10695
         _Version        =   1441793
         _ExtentX        =   18865
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Cuentas Contables [Operaciones en Estado Ex-Asociados]"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCorEO 
            Height          =   330
            Left            =   2160
            TabIndex        =   14
            Top             =   480
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCorEO_Desc 
            Height          =   330
            Left            =   4320
            TabIndex        =   15
            Top             =   480
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntMorEO 
            Height          =   330
            Left            =   2160
            TabIndex        =   16
            Top             =   840
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntMorEO_Desc 
            Height          =   330
            Left            =   4320
            TabIndex        =   17
            Top             =   840
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaAmortEO 
            Height          =   330
            Left            =   2160
            TabIndex        =   18
            Top             =   1200
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaAmortEO_Desc 
            Height          =   330
            Left            =   4320
            TabIndex        =   19
            Top             =   1200
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Amortización"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Interés Moratorio"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Interés Corriente"
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtPuente 
         Height          =   330
         Left            =   2400
         TabIndex        =   23
         Top             =   4560
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtPuente_Desc 
         Height          =   330
         Left            =   4560
         TabIndex        =   24
         Top             =   4560
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1695
         Left            =   -69640
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   10575
         _Version        =   1441793
         _ExtentX        =   18653
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Cuentas Contables [Operaciones con Estado en Cobro Judicial]"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCorEJ 
            Height          =   330
            Left            =   2040
            TabIndex        =   27
            Top             =   480
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCorEJ_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   28
            Top             =   480
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntMorEJ 
            Height          =   330
            Left            =   2040
            TabIndex        =   29
            Top             =   840
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntMorEJ_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   30
            Top             =   840
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaAmortEJ 
            Height          =   330
            Left            =   2040
            TabIndex        =   31
            Top             =   1200
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaAmortEJ_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   32
            Top             =   1200
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Amortización"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Interés Moratorio"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Interés Corriente"
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   3495
         Left            =   -69640
         TabIndex        =   36
         Top             =   2400
         Visible         =   0   'False
         Width           =   10575
         _Version        =   1441793
         _ExtentX        =   18653
         _ExtentY        =   6165
         _StockProps     =   79
         Caption         =   "Cuentas Contables [Producto Acumulado]"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPSRegistra 
            Height          =   255
            Left            =   0
            TabIndex        =   52
            Top             =   1920
            Width           =   8295
            _Version        =   1441793
            _ExtentX        =   14626
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Línea registra Contablemente Producto Acumulado en Suspenso (Cuentas de Orden)"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtCtaProdAcumCartera 
            Height          =   330
            Left            =   2040
            TabIndex        =   37
            Top             =   480
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaProdAcumCartera_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   38
            Top             =   480
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaProdAcumEfectos 
            Height          =   330
            Left            =   2040
            TabIndex        =   39
            Top             =   840
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaProdAcumEfectos_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   40
            Top             =   840
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCbrAdelantado 
            Height          =   330
            Left            =   2040
            TabIndex        =   41
            Top             =   1200
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaIntCbrAdelantado_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   42
            Top             =   1200
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaPSDeudora 
            Height          =   330
            Left            =   2040
            TabIndex        =   46
            Top             =   2400
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaPSDeudora_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   47
            Top             =   2400
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.FlatEdit txtCtaPSAcreedora 
            Height          =   330
            Left            =   2040
            TabIndex        =   48
            Top             =   2760
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCtaPSAcreedora_Desc 
            Height          =   330
            Left            =   4200
            TabIndex        =   49
            Top             =   2760
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   51
            Top             =   2400
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "P.S. Cuenta Deudora"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   50
            Top             =   2760
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "P.S. Cuenta Acreedora"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   4
            Left            =   0
            TabIndex        =   45
            Top             =   480
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Prod. Acum. Cartera"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   7
            Left            =   0
            TabIndex        =   44
            Top             =   840
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Prod. Acum. Efectos"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   6
            Left            =   0
            TabIndex        =   43
            Top             =   1200
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Int.Cbr.Adelantado"
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox chkIVA 
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   5400
         Width           =   8295
         _Version        =   1441793
         _ExtentX        =   14626
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "La Línea de Crédito es un Producto que registra Impuesto de Valor Agregado?"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaIVA 
         Height          =   330
         Left            =   2400
         TabIndex        =   56
         Top             =   5760
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtCtaIVA_Desc 
         Height          =   330
         Left            =   4560
         TabIndex        =   57
         Top             =   5760
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.FlatEdit txtAnticipo 
         Height          =   330
         Left            =   2400
         TabIndex        =   59
         Top             =   4200
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtAnticipo_Desc 
         Height          =   330
         Left            =   4560
         TabIndex        =   60
         Top             =   4200
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   61
         Top             =   4200
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pago Anticipado"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   58
         Top             =   5760
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cta IVA "
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   615
         Index           =   2
         Left            =   480
         TabIndex        =   25
         Top             =   4440
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Transitoria Desembolsos"
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   6735
      _Version        =   1441793
      _ExtentX        =   11880
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
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   615
      Left            =   9240
      TabIndex        =   53
      Top             =   7680
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   1085
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
      Picture         =   "frmCR_CtaCatalogo.frx":030A
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   6
      Left            =   1200
      TabIndex        =   54
      Top             =   360
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Línea"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmCR_CtaCatalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub cmdGuardar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Catalogo_Cuentas_Update '" & txtCodigo.Text _
      & "', '" & txtCtaIntCorEC.Text & "', '" & txtCtaIntMorEC.Text & "', '" & txtCtaAmortEC.Text _
      & "', '" & txtCtaIntCorEO.Text & "', '" & txtCtaIntMorEO.Text & "', '" & txtCtaAmortEO.Text _
      & "', '" & txtCtaIntCorEJ.Text & "', '" & txtCtaIntMorEJ.Text & "', '" & txtCtaAmortEJ.Text _
      & "', '" & txtPuente.Text & "', '" & txtAnticipo.Text & "', '" & txtCtaIVA.Text & "', " & chkIVA.Value _
      & " , '" & txtCtaIntCbrAdelantado.Text & "', '" & txtCtaProdAcumEfectos.Text _
      & "', '" & txtCtaProdAcumCartera.Text & "', " & chkPSRegistra.Value _
      & " , '" & txtCtaPSDeudora.Text & "', '" & txtCtaPSAcreedora.Text & "', '" & glogon.Usuario & "'"


Call OpenRecordSet(rs, strSQL)

If rs!Aplica = 1 Then
    Call Bitacora("Modifica", "Actualiza cuentas en catalogo:" & txtCodigo.Text)
    
    Me.MousePointer = vbDefault
    MsgBox "La Información se guardó satisfactoriamente ...", vbInformation
    
    Unload Me
Else
    Me.MousePointer = vbDefault
    MsgBox rs!Mensaje, vbExclamation
End If


Call RefrescaTags(Me)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCuentas_Init_Load()

On Error GoTo vError


strSQL = "select * from vCrd_Catalogo_Cuentas where codigo = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
    
    txtCtaIntCorEC.Text = rs!ctaNintC_Mask & ""
    txtCtaIntMorEC.Text = rs!ctaNintM_Mask & ""
    txtCtaAmortEC.Text = rs!ctaNamort_Mask & ""
    txtCtaIntCorEC_Desc.Text = rs!ctaNintC_Desc & ""
    txtCtaIntMorEC_Desc.Text = rs!ctaNintM_Desc & ""
    txtCtaAmortEC_Desc.Text = rs!ctaNamort_Desc & ""
    
    txtCtaIntCorEO.Text = rs!ctaOintC_Mask & ""
    txtCtaIntMorEO.Text = rs!ctaOintM_Mask & ""
    txtCtaAmortEO.Text = rs!CtaOamort_Mask & ""
    txtCtaIntCorEO_Desc.Text = rs!ctaOintC_Desc & ""
    txtCtaIntMorEO_Desc.Text = rs!ctaOintM_Desc & ""
    txtCtaAmortEO_Desc.Text = rs!CtaOamort_Desc & ""
        
    txtCtaIntCorEJ.Text = rs!ctacintc_Mask & ""
    txtCtaIntMorEJ.Text = rs!ctacintm_Mask & ""
    txtCtaAmortEJ.Text = rs!ctacamort_Mask & ""
    txtCtaIntCorEJ_Desc.Text = rs!ctaCintC_Desc & ""
    txtCtaIntMorEJ_Desc.Text = rs!ctaCintM_Desc & ""
    txtCtaAmortEJ_Desc.Text = rs!CtaCamort_Desc & ""
    
    
    txtCtaProdAcumCartera.Text = rs!CTA_CAR_PRODUCTO_Mask & ""
    txtCtaProdAcumEfectos.Text = rs!CTA_PROD_ACUM_Mask & ""
    txtCtaIntCbrAdelantado.Text = rs!CTA_INT_ADELANTADO_Mask & ""
    txtCtaProdAcumCartera_Desc.Text = rs!CTA_CAR_PRODUCTO_Desc & ""
    txtCtaProdAcumEfectos_Desc.Text = rs!CTA_PROD_ACUM_Desc & ""
    txtCtaIntCbrAdelantado_Desc.Text = rs!CTA_INT_ADELANTADO_Desc & ""

    txtCtaPSDeudora.Text = rs!CTA_PS_DEUDORA_Mask & ""
    txtCtaPSAcreedora.Text = rs!CTA_PS_ACREADORA_Mask & ""
    txtCtaPSDeudora_Desc.Text = rs!CTA_PS_DEUDORA_Desc & ""
    txtCtaPSAcreedora_Desc.Text = rs!CTA_PS_ACREADORA_Desc & ""
    
    txtPuente.Text = rs!ctapuente_Mask & ""
    txtPuente_Desc.Text = rs!ctapuente_Desc & ""
    
    
    txtAnticipo.Text = rs!CTA_CARGOS_ANTICIPO_Mask & ""
    txtAnticipo_Desc.Text = rs!CTA_CARGOS_ANTICIPO_Desc & ""
    
    txtCtaIVA.Text = rs!CTA_IVA_Mask & ""
    txtCtaIVA_Desc.Text = rs!CTA_IVA_Desc & ""
    
    chkPSRegistra.Value = rs!PS_REGISTRA
    chkIVA.Value = rs!IMPUESTO

rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

txtCodigo.Text = GLOBALES.gTag
txtDescripcion.Text = GLOBALES.gTag2

Call sbCuentas_Init_Load

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbCuentas_Consulta(objCuenta As Object, objCuentaDesc As Object)

On Error GoTo vError

Call sbgCntCuentaConsulta

If gBusquedas.Resultado <> "" Then
    objCuenta.Text = fxgCntCuentaFormato(True, gBusquedas.Resultado, 0)
    objCuenta.SetFocus
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtAnticipo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If fxValida_Cuentas(txtAnticipo, txtAnticipo_Desc) Then
        txtPuente.SetFocus
    End If
 End If
 
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtAnticipo, txtAnticipo_Desc)
 End If

End Sub

Private Sub txtCtaIVA_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If fxValida_Cuentas(txtCtaIVA, txtCtaIVA_Desc) Then
       'Nada
    End If
 End If
 
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIVA, txtCtaIVA_Desc)
 End If
End Sub


'Cta Corrientes
Private Sub txtCtaIntCorEC_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaIntCorEC, txtCtaIntCorEC_Desc) Then txtCtaIntMorEC.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIntCorEC, txtCtaIntCorEC_Desc)
 End If
End Sub

Private Sub txtCtaIntMorEC_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaIntMorEC, txtCtaIntMorEC_Desc) Then txtCtaAmortEC.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIntMorEC, txtCtaIntMorEC_Desc)
 End If
End Sub

Private Sub txtCtaAmortEC_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaAmortEC, txtCtaAmortEC_Desc) Then txtCtaIntCorEO.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaAmortEC, txtCtaAmortEC_Desc)
 End If
End Sub


Private Sub txtCtaIntCorEC_LostFocus()
    Call sbCuenta_Load(txtCtaIntCorEC, txtCtaIntCorEC_Desc)
End Sub

Private Sub txtCtaIntMorEC_LostFocus()
    Call sbCuenta_Load(txtCtaIntMorEC, txtCtaIntMorEC_Desc)
End Sub

Private Sub txtCtaAmortEC_LostFocus()
    Call sbCuenta_Load(txtCtaAmortEC, txtCtaAmortEC_Desc)
End Sub


' Ctas de OPEX
Private Sub txtCtaIntCorEO_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaIntCorEO, txtCtaIntCorEO_Desc) And txtCtaIntCorEO.Text <> "" Then txtCtaIntMorEO.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIntCorEO, txtCtaIntCorEO_Desc)
 End If
End Sub

Private Sub txtCtaIntMorEO_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If fxValida_Cuentas(txtCtaIntMorEO, txtCtaIntMorEO_Desc) Then
            txtCtaAmortEO.SetFocus
    End If
 End If
 
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIntMorEO, txtCtaIntMorEO)
 End If
End Sub

Private Sub txtCtaAmortEO_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If fxValida_Cuentas(txtCtaAmortEO, txtCtaAmortEO_Desc) Then
            txtPuente.SetFocus
    End If
 End If
 
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaAmortEO, txtCtaAmortEO_Desc)
 End If
End Sub


Private Sub txtCtaIntCorEO_LostFocus()
    Call sbCuenta_Load(txtCtaIntCorEO, txtCtaIntCorEO_Desc)
End Sub

Private Sub txtCtaIntMorEO_LostFocus()
    Call sbCuenta_Load(txtCtaIntMorEO, txtCtaIntMorEO_Desc)
End Sub

Private Sub txtCtaAmortEO_LostFocus()
    Call sbCuenta_Load(txtCtaAmortEO, txtCtaAmortEO_Desc)
End Sub




'CUENTAS PARA COBRO JUDICIAL
Private Sub txtCtaIntCorEJ_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaIntCorEJ, txtCtaIntCorEJ_Desc) Then txtCtaIntMorEJ.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIntCorEJ, txtCtaIntCorEJ_Desc)
 End If
End Sub

Private Sub txtCtaIntMorEJ_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaIntMorEJ, txtCtaIntMorEJ_Desc) Then txtCtaAmortEJ.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIntMorEJ, txtCtaIntMorEJ_Desc)
 End If
End Sub

Private Sub txtCtaAmortEJ_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaAmortEJ, txtCtaAmortEJ_Desc) Then txtCtaProdAcumCartera.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaAmortEJ, txtCtaAmortEJ_Desc)
 End If
End Sub

Private Sub txtCtaIntCorEJ_LostFocus()
    Call sbCuenta_Load(txtCtaIntCorEJ, txtCtaIntCorEJ_Desc)
End Sub

Private Sub txtCtaIntMorEJ_LostFocus()
    Call sbCuenta_Load(txtCtaIntMorEJ, txtCtaIntMorEJ_Desc)
End Sub

Private Sub txtCtaAmortEJ_LostFocus()
    Call sbCuenta_Load(txtCtaAmortEJ, txtCtaAmortEJ_Desc)
End Sub



'Cuentas para Producto Acumulado
Private Sub txtCtaProdAcumCartera_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaProdAcumCartera, txtCtaProdAcumCartera_Desc) Then txtCtaProdAcumEfectos.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaProdAcumCartera, txtCtaProdAcumCartera_Desc)
 End If
End Sub

Private Sub txtCtaProdAcumEfectos_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaProdAcumEfectos, txtCtaProdAcumEfectos_Desc) Then txtCtaIntCbrAdelantado.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaProdAcumEfectos, txtCtaProdAcumEfectos_Desc)
 End If
End Sub

Private Sub txtCtaIntCbrAdelantado_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If fxValida_Cuentas(txtCtaIntCbrAdelantado, txtCtaIntCbrAdelantado_Desc) Then
        tcMain.Item(0).Selected = True
        txtPuente.SetFocus
    End If
 End If
 
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaIntCbrAdelantado, txtCtaIntCbrAdelantado_Desc)
 End If
End Sub


Private Sub txtCtaProdAcumCartera_LostFocus()
    Call sbCuenta_Load(txtCtaProdAcumCartera, txtCtaProdAcumCartera_Desc)
End Sub

Private Sub txtCtaProdAcumEfectos_LostFocus()
    Call sbCuenta_Load(txtCtaProdAcumEfectos, txtCtaProdAcumEfectos_Desc)
End Sub

Private Sub txtCtaIntCbrAdelantado_LostFocus()
    Call sbCuenta_Load(txtCtaIntCbrAdelantado, txtCtaIntCbrAdelantado_Desc)
End Sub


'Cuentas para Producto Acumulado en Suspenso
Private Sub txtCtaPSDeudora_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaPSDeudora, txtCtaPSDeudora_Desc) Then txtCtaPSAcreedora.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaPSDeudora, txtCtaPSDeudora_Desc)
 End If
End Sub

Private Sub txtCtaPSAcreedora_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then If fxValida_Cuentas(txtCtaPSAcreedora, txtCtaPSAcreedora_Desc) Then txtPuente.SetFocus
 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtCtaPSDeudora, txtCtaPSDeudora_Desc)
 End If
End Sub


Private Sub txtCtaPSAcreedora_LostFocus()
    Call sbCuenta_Load(txtCtaPSAcreedora, txtCtaPSAcreedora_Desc)
End Sub

Private Sub txtCtaPSDeudora_LostFocus()
    Call sbCuenta_Load(txtCtaPSDeudora, txtCtaPSDeudora_Desc)
End Sub

'CUENTA PUENTE
Private Sub txtPuente_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If fxValida_Cuentas(txtPuente, txtPuente_Desc) And cmdGuardar.Enabled Then
      cmdGuardar.SetFocus
    End If
 End If

 If KeyCode = vbKeyF4 Then
    Call sbCuentas_Consulta(txtPuente, txtPuente_Desc)
 End If
End Sub

Private Sub txtPuente_LostFocus()
    Call sbCuenta_Load(txtPuente, txtPuente_Desc)
End Sub



Private Sub sbCuenta_Load(objCuenta As Object, objCuentaDesc As Object)

On Error GoTo vError


objCuenta.Text = fxgCntCuentaFormato(False, objCuenta.Text, 0)
objCuentaDesc.Text = fxgCntCuentaDesc(objCuenta.Text)
objCuenta.Text = fxgCntCuentaFormato(True, objCuenta.Text, 0)


Exit Sub
vError:
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Function fxValida_Cuentas(objCuenta As Object, objCuentaDesc As Object) As Boolean
Dim vCadena As String

On Error GoTo vError

fxValida_Cuentas = False
vCadena = ""

objCuenta.Text = fxgCntCuentaFormato(False, Trim(objCuenta.Text), 0)

If fxgCntCuentaValida(objCuenta.Text) Then
  fxValida_Cuentas = True
  vCadena = fxgCntCuentaDesc(objCuenta.Text)
Else
  fxValida_Cuentas = False
End If

objCuenta.Text = fxgCntCuentaFormato(True, objCuenta.Text, 0)
objCuentaDesc.Text = vCadena


Exit Function
vError:
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Function fxValida_CuentasGrabar(objCuenta As Object) As Boolean
Dim vCuenta As String

vCuenta = objCuenta.Text
vCuenta = fxgCntCuentaFormato(False, vCuenta, 0)

fxValida_CuentasGrabar = fxgCntCuentaValida(vCuenta)

End Function

