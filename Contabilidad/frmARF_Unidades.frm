VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmARF_Unidades 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrendamientos: Unidades (Oficinas/Agencias)"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11175
      _Version        =   1572864
      _ExtentX        =   19711
      _ExtentY        =   10186
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
      ItemCount       =   1
      Item(0).Caption =   "General"
      Item(0).ControlCount=   11
      Item(0).Control(0)=   "gbDireccion"
      Item(0).Control(1)=   "chkActivo"
      Item(0).Control(2)=   "GroupBox3"
      Item(0).Control(3)=   "txtContacto"
      Item(0).Control(4)=   "Label4(8)"
      Item(0).Control(5)=   "txtUnidadCod"
      Item(0).Control(6)=   "txtCentroCod"
      Item(0).Control(7)=   "txtUnidadDesc"
      Item(0).Control(8)=   "txtCentroDesc"
      Item(0).Control(9)=   "Label1(7)"
      Item(0).Control(10)=   "Label1(4)"
      Begin XtremeSuiteControls.GroupBox gbDireccion 
         Height          =   1695
         Left            =   240
         TabIndex        =   1
         Top             =   4080
         Width           =   10815
         _Version        =   1572864
         _ExtentX        =   19076
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Dirección"
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
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   312
            Left            =   1440
            TabIndex        =   2
            Top             =   480
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         End
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   312
            Left            =   1440
            TabIndex        =   3
            Top             =   840
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         End
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   312
            Left            =   1440
            TabIndex        =   4
            Top             =   1200
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         End
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   1092
            Left            =   3840
            TabIndex        =   5
            Top             =   480
            Width           =   6852
            _Version        =   1572864
            _ExtentX        =   12086
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   13
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Provincia"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   12
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cantón"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   11
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Distrito"
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activa?"
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1815
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   10815
         _Version        =   1572864
         _ExtentX        =   19076
         _ExtentY        =   3201
         _StockProps     =   79
         Caption         =   "Información de Contacto"
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
         Begin XtremeSuiteControls.FlatEdit txtWebSite 
            Height          =   312
            Left            =   5400
            TabIndex        =   11
            Top             =   360
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   312
            Left            =   5400
            TabIndex        =   12
            Top             =   720
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail2 
            Height          =   312
            Left            =   5400
            TabIndex        =   13
            Top             =   1080
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
            Height          =   312
            Left            =   5400
            TabIndex        =   14
            Top             =   1440
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono 
            Height          =   312
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   312
            Left            =   1440
            TabIndex        =   16
            Top             =   720
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Teléfono (1)"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   2
            Left            =   0
            TabIndex        =   21
            Top             =   720
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Teléfono (2)"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   4
            Left            =   3840
            TabIndex        =   20
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Web Site"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   5
            Left            =   3840
            TabIndex        =   19
            Top             =   720
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email (1)"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   6
            Left            =   3840
            TabIndex        =   18
            Top             =   1080
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email (2)"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   7
            Left            =   3840
            TabIndex        =   17
            Top             =   1440
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Apto. Postal"
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtContacto 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   1680
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16319
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidadCod 
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   720
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroCod 
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   1080
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
         Height          =   315
         Left            =   3600
         TabIndex        =   32
         Top             =   720
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
         Height          =   315
         Left            =   3600
         TabIndex        =   33
         Top             =   1080
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin VB.Label Label1 
         Caption         =   "Unidad"
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
         Left            =   360
         TabIndex        =   35
         Top             =   720
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Centro de Costo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   24
         Top             =   1680
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Contacto"
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
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   10080
      TabIndex        =   26
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   120
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
            Picture         =   "frmARF_Unidades.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Unidades.frx":3492
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Unidades.frx":6924
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Unidades.frx":6A42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3240
      TabIndex        =   27
      Top             =   600
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11874
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1200
      TabIndex        =   28
      Top             =   600
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina/ Ud"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Oficina / Agencia"
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmARF_Unidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset, strSQL As String
Dim vScroll As Boolean, vPaso As Boolean

Dim vEdita As Boolean, vCodigo As String
Dim vCantonMascara As String, vDistritoMascara As String, vFechaActual As Date



Private Sub cboCanton_Click()

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "
End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click

End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 COD_LOCAL from ARF_UNIDADES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_LOCAL > '" & txtCodigo.Text & "' order by COD_LOCAL asc"
    Else
       strSQL = strSQL & " where COD_LOCAL < '" & txtCodigo.Text & "' order by COD_LOCAL desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_LOCAL
      Call txtCodigo_LostFocus
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 vModulo = 20

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
tcMain.Item(0).Selected = True

vEdita = False
vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")

vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCodigo = ""

txtNombre = ""
txtTelefono.Text = ""
txtTelefono2.Text = ""
txtWebSite.Text = ""
txtEmail.Text = ""
txtEmail2.Text = ""
txtAptoPostal.Text = ""

txtDireccion = ""
txtContacto.Text = ""

chkActivo.Value = vbChecked

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select COD_LOCAL,nombre from ARF_UNIDADES"
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

If Not fxSIFValidaCadena(txtCodigo.Text) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & " from ARF_UNIDADES P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and P.Canton = Dist.Canton and P.distrito = Dist.distrito" _
       & " where P.COD_LOCAL = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!COD_LOCAL
  txtCodigo = rs!COD_LOCAL

  txtNombre = rs!Descripcion & ""
  
  txtTelefono.Text = rs!Telefono_01 & ""
  txtTelefono2.Text = rs!Telefono_02 & ""

  txtWebSite.Text = rs!WebSite & ""
  txtEmail.Text = rs!email_01 & ""
  txtEmail2.Text = rs!Email_02 & ""
  txtAptoPostal.Text = rs!apto_postal & ""

  txtContacto.Text = rs!Contacto_Nombre & ""

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
     
  cboDistrito.ToolTipText = Trim(rs!Distrito) & ""
  txtDireccion.Text = rs!direccion

  txtUnidadCod.Text = rs!Cod_Unidad & ""
  txtCentroCod.Text = rs!Cod_Centro_Costo & ""
  
  txtUnidadCod_LostFocus
  txtCentroCod_LostFocus


Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If Trim(txtCodigo.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Código de la Local/Agencia/Oficina no es válida ..."
If Trim(txtNombre.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Local/ Unidad no es válido ..."

If Trim(txtEmail.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - El Email Principal no es válido ..."

If Trim(txtUnidadDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - La Unidad Contable no es válida ..."
If Trim(txtCentroCod.Text) <> "" And Trim(txtCentroDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Centro de Costos no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()

On Error GoTo vError


If vEdita Then
  strSQL = "update ARF_UNIDADES set Descripcion = '" & Trim(txtNombre.Text) & "', Telefono_01 = '" & txtTelefono.Text & "', Telefono_02 = '" & txtTelefono2.Text _
         & "', WebSite = '" & txtWebSite.Text & "',apto_postal = '" & txtAptoPostal _
         & "', Email_01 = '" & txtEmail & "', email_02 = '" & txtEmail2.Text & "', direccion = '" & txtDireccion _
         & "', Distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "', canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "', Provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & "', Contacto_Nombre = '" & txtContacto.Text & "', Activo = " & chkActivo.Value _
         & ", cod_Unidad = '" & txtUnidadCod.Text & "' , cod_Centro_Costo = '" & txtCentroCod.Text _
         & "', Modifica_Fecha = getdate(), Modifica_Usuario = '" & glogon.Usuario & "'" _
         & " where COD_LOCAL = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Arrendamiento, Unidad Id: " & vCodigo)

Else
  vCodigo = txtCodigo

   strSQL = "insert into ARF_UNIDADES(COD_LOCAL, Descripcion, Telefono_01, Telefono_02, Activo, cod_Unidad, cod_Centro_Costo, Contacto_Nombre" _
          & ", apto_postal, email_01, email_02, WebSite, provincia, canton, distrito , direccion, Registro_fecha, Registro_usuario )" _
          & " values('" & vCodigo & "','" & txtNombre.Text & "','" & txtTelefono.Text & "','" & txtTelefono2.Text _
          & "', " & chkActivo.Value & ",'" & txtUnidadCod.Text & "', '" & txtCentroCod.Text & "', '" & txtContacto.Text _
          & "', '" & txtAptoPostal.Text & "', '" & txtEmail.Text & "', '" & txtEmail2.Text & "', '" & txtWebSite.Text _
          & "', '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "', '" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "', '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "', '" & txtDireccion.Text & "', getdate(), '" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Arrendamiento, Unidad Id: " & vCodigo)

End If

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete ARF_UNIDADES where COD_LOCAL = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Arrendamiento, Unidad Id: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCentroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_centro_Costo"
   gBusquedas.Orden = "cod_centro_Costo"
   gBusquedas.Consulta = "select cod_centro_Costo,descripcion from Cntx_Centro_Costos"
   gBusquedas.Filtro = " and cod_Contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   If Len(gBusquedas.Resultado) > 0 Then
        txtCentroCod.Text = gBusquedas.Resultado
        txtCentroDesc.Text = gBusquedas.Resultado2
   End If
End If

End Sub

Private Sub txtCentroCod_LostFocus()
txtCentroDesc.Text = fxgCntCentroCostos(txtCentroCod.Text)
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_LOCAL"
  gBusquedas.Orden = "COD_LOCAL"
  gBusquedas.Consulta = "select COD_LOCAL, Descripcion from ARF_UNIDADES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
  Call sbConsulta(txtCodigo.Text)
'  txtNombre.SetFocus
End Sub



Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContacto.SetFocus
End Sub

Private Sub txtEMail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select COD_LOCAL, Descripcion from ARF_UNIDADES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail2.SetFocus
End Sub

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub


Private Sub txtUnidadCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_unidad"
   gBusquedas.Orden = "cod_unidad"
   gBusquedas.Consulta = "select cod_unidad,descripcion from Cntx_Unidades"
   gBusquedas.Filtro = " and cod_Contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   If Len(gBusquedas.Resultado) > 0 Then
        txtUnidadCod.Text = gBusquedas.Resultado
        txtUnidadDesc.Text = gBusquedas.Resultado2
   End If
End If
End Sub


Private Sub txtUnidadCod_LostFocus()
txtUnidadDesc.Text = fxgCntUnidad(txtUnidadCod.Text)
End Sub

Private Sub txtUnidadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or vbKeyTab Then txtCentroCod.SetFocus
End Sub



