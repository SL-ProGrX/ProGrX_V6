VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmIVR_Cat_Adminsitradores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Administradores"
   ClientHeight    =   8568
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11028
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8568
   ScaleWidth      =   11028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7332
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10812
      _Version        =   1310720
      _ExtentX        =   19071
      _ExtentY        =   12933
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
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   5
      Item(0).Caption =   "General"
      Item(0).ControlCount=   11
      Item(0).Control(0)=   "txtDireccion"
      Item(0).Control(1)=   "Label18(3)"
      Item(0).Control(2)=   "Label6"
      Item(0).Control(3)=   "Label7(0)"
      Item(0).Control(4)=   "Label8"
      Item(0).Control(5)=   "cboEstado"
      Item(0).Control(6)=   "GroupBox3"
      Item(0).Control(7)=   "cboTipo"
      Item(0).Control(8)=   "gbCuenta"
      Item(0).Control(9)=   "GroupBox1"
      Item(0).Control(10)=   "txtIdentificacion"
      Item(1).Caption =   "Portafolios"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "scTitulo(0)"
      Item(1).Control(1)=   "gDetalle"
      Item(2).Caption =   "Instrumentos"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "scTitulo(1)"
      Item(2).Control(1)=   "lswI"
      Item(3).Caption =   "Divisas"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "scTitulo(2)"
      Item(3).Control(1)=   "lswD"
      Item(4).Caption =   "Contactos"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "scTitulo(3)"
      Item(4).Control(1)=   "lswC"
      Item(4).Control(2)=   "btnContactos"
      Begin XtremeSuiteControls.ListView lswI 
         Height          =   6012
         Left            =   -69760
         TabIndex        =   50
         Top             =   960
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1310720
         _ExtentX        =   18013
         _ExtentY        =   10604
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.ListView lswD 
         Height          =   6012
         Left            =   -69760
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1310720
         _ExtentX        =   18013
         _ExtentY        =   10604
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.ListView lswC 
         Height          =   6012
         Left            =   -69760
         TabIndex        =   52
         Top             =   960
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1310720
         _ExtentX        =   18013
         _ExtentY        =   10604
         _StockProps     =   77
         BackColor       =   -2147483643
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox gbCuenta 
         Height          =   852
         Left            =   120
         TabIndex        =   35
         Top             =   4200
         Width           =   10692
         _Version        =   1310720
         _ExtentX        =   18860
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Cuenta para Transacciones en Tránsito"
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
         Begin XtremeSuiteControls.FlatEdit txtCuenta 
            Height          =   312
            Left            =   1440
            TabIndex        =   36
            Top             =   480
            Width           =   2172
            _Version        =   1310720
            _ExtentX        =   3831
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   312
            Left            =   3600
            TabIndex        =   37
            Top             =   480
            Width           =   6852
            _Version        =   1310720
            _ExtentX        =   12086
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Appearance      =   2
         End
         Begin VB.Label Label12 
            Caption         =   "Transitoria"
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
            Index           =   1
            Left            =   0
            TabIndex        =   38
            Top             =   480
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   912
         Left            =   1560
         TabIndex        =   1
         Top             =   3120
         Width           =   9012
         _Version        =   1310720
         _ExtentX        =   15896
         _ExtentY        =   1609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1812
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   10812
         _Version        =   1310720
         _ExtentX        =   19071
         _ExtentY        =   3196
         _StockProps     =   79
         Caption         =   "Información de Contacto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtWebSite 
            Height          =   312
            Left            =   5160
            TabIndex        =   19
            Top             =   360
            Width           =   5292
            _Version        =   1310720
            _ExtentX        =   9334
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   312
            Left            =   5160
            TabIndex        =   20
            Top             =   720
            Width           =   5292
            _Version        =   1310720
            _ExtentX        =   9334
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail2 
            Height          =   312
            Left            =   5160
            TabIndex        =   21
            Top             =   1080
            Width           =   5292
            _Version        =   1310720
            _ExtentX        =   9334
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
            Height          =   312
            Left            =   5160
            TabIndex        =   22
            Top             =   1440
            Width           =   5292
            _Version        =   1310720
            _ExtentX        =   9334
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono1 
            Height          =   312
            Left            =   1440
            TabIndex        =   23
            Top             =   360
            Width           =   2052
            _Version        =   1310720
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   312
            Left            =   1440
            TabIndex        =   24
            Top             =   720
            Width           =   2052
            _Version        =   1310720
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelFax 
            Height          =   312
            Left            =   1440
            TabIndex        =   25
            Top             =   1080
            Width           =   2052
            _Version        =   1310720
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   7
            Left            =   3840
            TabIndex        =   32
            Top             =   1440
            Width           =   1332
            _Version        =   1310720
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Apto. Postal"
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
            TabIndex        =   31
            Top             =   1080
            Width           =   1332
            _Version        =   1310720
            _ExtentX        =   2350
            _ExtentY        =   444
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   5
            Left            =   3840
            TabIndex        =   30
            Top             =   720
            Width           =   1332
            _Version        =   1310720
            _ExtentX        =   2350
            _ExtentY        =   444
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   4
            Left            =   3840
            TabIndex        =   29
            Top             =   360
            Width           =   1332
            _Version        =   1310720
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Web Site"
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
            Index           =   3
            Left            =   0
            TabIndex        =   28
            Top             =   1080
            Width           =   1332
            _Version        =   1310720
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tel. Fax"
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
            TabIndex        =   27
            Top             =   720
            Width           =   1332
            _Version        =   1310720
            _ExtentX        =   2350
            _ExtentY        =   444
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   360
            Width           =   1332
            _Version        =   1310720
            _ExtentX        =   2350
            _ExtentY        =   444
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   3600
         TabIndex        =   33
         Top             =   600
         Width           =   4932
         _Version        =   1310720
         _ExtentX        =   8700
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   8520
         TabIndex        =   34
         Top             =   600
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3620
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2172
         Left            =   120
         TabIndex        =   39
         Top             =   5160
         Width           =   10692
         _Version        =   1310720
         _ExtentX        =   18860
         _ExtentY        =   3831
         _StockProps     =   79
         Caption         =   "Cuentas Bancarias"
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
         Begin XtremeSuiteControls.ListView lswCuentas 
            Height          =   1332
            Left            =   240
            TabIndex        =   40
            Top             =   756
            Width           =   10212
            _Version        =   1310720
            _ExtentX        =   18013
            _ExtentY        =   2350
            _StockProps     =   77
            BackColor       =   -2147483643
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
            Appearance      =   16
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCuentas 
            Height          =   315
            Left            =   8880
            TabIndex        =   41
            Tag             =   "1"
            Top             =   360
            Width           =   1572
            _Version        =   1310720
            _ExtentX        =   2773
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Cuentas Bancarias"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ComboBox cboEmitir 
            Height          =   312
            Left            =   6000
            TabIndex        =   43
            Top             =   360
            Width           =   2532
            _Version        =   1310720
            _ExtentX        =   4466
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboBancos 
            Height          =   312
            Left            =   1440
            TabIndex        =   44
            Top             =   360
            Width           =   4572
            _Version        =   1310720
            _ExtentX        =   8065
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin VB.Label Label3 
            Caption         =   "Cuenta / Tipo de Documento"
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
            Left            =   1440
            TabIndex        =   42
            Top             =   156
            Width           =   3372
         End
      End
      Begin FPSpreadADO.fpSpread gDetalle 
         Height          =   5652
         Left            =   -69760
         TabIndex        =   45
         Top             =   1080
         Visible         =   0   'False
         Width           =   10092
         _Version        =   524288
         _ExtentX        =   17801
         _ExtentY        =   9970
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
         MaxCols         =   482
         ScrollBars      =   2
         SpreadDesigner  =   "frmIVR_Cat_Adminsitradores.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnContactos 
         Height          =   312
         Left            =   -61120
         TabIndex        =   53
         Tag             =   "1"
         Top             =   620
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310720
         _ExtentX        =   2773
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Contactos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmIVR_Cat_Adminsitradores.frx":064F
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
         Height          =   312
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Index           =   3
         Left            =   -69760
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1310720
         _ExtentX        =   18013
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Lista de Contactos"
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Index           =   2
         Left            =   -69760
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1310720
         _ExtentX        =   18013
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Divisas autorizadas a este administrador"
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Index           =   1
         Left            =   -69760
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1310720
         _ExtentX        =   18013
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Instrumentos autorizados a este administrador"
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Index           =   0
         Left            =   -69760
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1310720
         _ExtentX        =   17801
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Vincule a este administrador uno o varios portafolios disponibles"
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
      Begin VB.Label Label18 
         Caption         =   "Identificación"
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
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label6 
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
         Height          =   252
         Left            =   8640
         TabIndex        =   5
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo"
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
         Index           =   0
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Dirección"
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
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   1332
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   10200
      TabIndex        =   7
      Top             =   720
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1680
      TabIndex        =   8
      Top             =   720
      Width           =   1572
      _Version        =   1310720
      _ExtentX        =   2773
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3240
      TabIndex        =   9
      Top             =   720
      Width           =   6852
      _Version        =   1310720
      _ExtentX        =   12086
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Nuevo"
      Top             =   40
      Width           =   1092
      _Version        =   1310720
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Nuevo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Adminsitradores.frx":0D6F
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      ToolTipText     =   "Editar"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Adminsitradores.frx":13A1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   2
      Left            =   1560
      TabIndex        =   13
      ToolTipText     =   "Eliminar"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Adminsitradores.frx":199C
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   3
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "Guardar"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Adminsitradores.frx":1F40
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   4
      Left            =   2520
      TabIndex        =   15
      ToolTipText     =   "Deshacer"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Adminsitradores.frx":2671
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   5
      Left            =   3000
      TabIndex        =   16
      ToolTipText     =   "Reporte"
      Top             =   36
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Adminsitradores.frx":2D71
      ImageAlignment  =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption scBarra 
      Height          =   372
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11772
      _Version        =   1310720
      _ExtentX        =   20764
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
      Alignment       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Administrador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1452
   End
End
Attribute VB_Name = "frmIVR_Cat_Adminsitradores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vPaso As Boolean



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
      If vCodigo = "" Then
        MsgBox "Consulte un administrador primero!", vbInformation
      Else
        vEdita = True
        txtNombre.SetFocus
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

Private Sub btnContactos_Click()
If vCodigo = "" Then Exit Sub

GLOBALES.gTag = vCodigo
GLOBALES.gTag2 = txtNombre.Text

frmIVR_Cat_Administrador_Contactos.Show vbModal

Call sbContactos_Load

End Sub

Private Sub btnCuentas_Click()

If Trim(txtIdentificacion) = "" Then
   MsgBox "Consulte un Administrador con Identificación válida!", vbExclamation
   tcMain.Item(0).Selected = True
   Exit Sub
End If

GLOBALES.gTag = Trim(txtIdentificacion)
GLOBALES.gTag2 = "IVR"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 COD_ADMINISTRADOR from IVR_ADMINISTRADOR"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_ADMINISTRADOR > '" & txtCodigo.Text & "' order by COD_ADMINISTRADOR asc"
    Else
       strSQL = strSQL & " where COD_ADMINISTRADOR < '" & txtCodigo.Text & "' order by COD_ADMINISTRADOR desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!Cod_Administrador)
    End If

End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCuentas_Load()

On Error GoTo vError

lswCuentas.ListItems.Clear
If Trim(txtIdentificacion.Text) <> "" Then
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.COD_Banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtIdentificacion.Text) & "'"
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswCuentas.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!Cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!Activa = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!registro_fecha & ""
           itmX.SubItems(8) = rs!Registro_Usuario & ""
     
       rs.MoveNext
    Loop
    rs.Close
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbInstrumentos_Load()

If vCodigo = "" Then
    lswI.ListItems.Clear
    Exit Sub
End If

On Error GoTo vError

strSQL = "exec spIVR_ADMINISTRADORES_INSTRUMENTOS '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)

lswI.ListItems.Clear

vPaso = True

Do While Not rs.EOF
   Set itmX = lswI.ListItems.Add(, , rs!Cod_Instrumento)
       itmX.SubItems(1) = rs!Descripcion
         
       itmX.Checked = IIf(rs!Asignado = 1, True, False)
  
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbContactos_Load()

If vCodigo = "" Then
    lswC.ListItems.Clear
    Exit Sub
End If

On Error GoTo vError

strSQL = "SELECT * FROM IVR_CONTACTOS WHERE COD_ADMINISTRADOR = '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)

lswC.ListItems.Clear


Do While Not rs.EOF
   Set itmX = lswC.ListItems.Add(, , rs!COD_CONTACTO)
       itmX.SubItems(1) = rs!Nombre & ""
       itmX.SubItems(2) = rs!Celular & ""
       itmX.SubItems(3) = rs!telefono & ""
       itmX.SubItems(4) = rs!Email_01 & ""
       itmX.SubItems(5) = rs!Email_02 & ""
  
  rs.MoveNext
Loop
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbGrid_Detalle_Load()

If vCodigo = "" Then
    gDetalle.MaxRows = 0
    Exit Sub
End If

vPaso = True

strSQL = "exec spIVR_ADMINISTRADORES_PORTAFOLIO '" & vCodigo & "'"
Call sbCargaGrid(gDetalle, 4, strSQL)

If gDetalle.MaxRows > 1 Then
   gDetalle.MaxRows = gDetalle.MaxRows - 1
End If

vPaso = False

End Sub

Private Sub gDetalle_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If vCodigo = "" Then Exit Sub

If Col = 4 Then
   gDetalle.Row = Row
   gDetalle.Col = 1
   
   gIVR_Cuentas.Tipo = "X"
   gIVR_Cuentas.Codigo_1 = gDetalle.Text
   gIVR_Cuentas.Codigo_2 = vCodigo
   
   gDetalle.Col = 2
   gIVR_Cuentas.Descripcion = gDetalle.Text & " / " & txtNombre.Text
    
   gDetalle.Col = 3
   If gDetalle.Value = vbChecked Then
        frmIVR_Cat_Cuentas_Contables.Show vbModal
   End If

End If

On Error GoTo vError

If Col = 3 Then

 Dim strSQL As String
 
   gDetalle.Row = Row
   gDetalle.Col = 1
    
   strSQL = "exec spIVR_PORTAFOLIO_ADMINISTRADORES_REGISTRO '" & gDetalle.Text _
        & "', '" & vCodigo
        
   gDetalle.Col = 3
    strSQL = strSQL & "', " & gDetalle.Value & ", '" & glogon.Usuario & "'"
    
   Call ConectionExecute(strSQL)

End If

vError:

End Sub


Private Sub lswI_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "insert IVR_ADMINISTRADOR_INST(COD_ADMINISTRADOR, COD_INSTRUMENTO" _
           & ", REGISTRO_USUARIO, REGISTRO_FECHA)" _
           & " VALUES('" & vCodigo & "','" & Item.Text & "','" & glogon.Usuario & "',dbo.mygetdate())"
Else
    strSQL = "delete IVR_ADMINISTRADOR_INST" _
           & " where cod_administrador = '" & vCodigo & "'" _
           & "   and cod_instrumento = '" & Item.Text & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Select Case Item.Index
   Case 1 'Portafolio
        Call sbGrid_Detalle_Load
   
   Case 2 'Instrumentos
        Call sbInstrumentos_Load
        
   Case 3 'Divisas
   
   Case 4 'Contactos
        Call sbContactos_Load
        
End Select


vError:
End Sub

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 22

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

 gDetalle.AppearanceStyle = fxGridStyle
 
 With lswI.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Descripción", 8200
 End With
 
 With lswD.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Descripción", 8200
 End With
 
 With lswC.ColumnHeaders
    .Clear
    .Add , , "Id", 1100
    .Add , , "Nombre", 3200
    .Add , , "Móvil", 1200, vbCenter
    .Add , , "Teléfono", 1200, vbCenter
    .Add , , "Email No.1", 3200
    .Add , , "Email No.2", 3200
 End With
 
 
lswCuentas.ColumnHeaders.Add 1, , "Cuenta", 2500
lswCuentas.ColumnHeaders.Add 2, , "Banco", 3500
lswCuentas.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 8, , "Fecha", 2500
lswCuentas.ColumnHeaders.Add 9, , "Usuario", 2500

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "InActivo"


cboEmitir.Clear
cboEmitir.AddItem fxTipoDocumento("CK")
cboEmitir.AddItem fxTipoDocumento("TE")
cboEmitir.AddItem fxTipoDocumento("ND")
cboEmitir.Text = fxTipoDocumento("TE")


strSQL = "exec spIVR_Bancos_Autorizados"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)

strSQL = "select Rtrim(TIPO_ADMINISTRADOR) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from IVR_ADMINISTRADOR_TIPOS order by DESCRIPCION"
Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

 
 vEdita = True
 Call sbBarra_Accion("Nuevo")
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True

vCodigo = ""
txtCodigo.Text = ""

cboEstado.Text = "Activo"

txtNombre.Text = ""
txtIdentificacion.Text = ""

txtDireccion.Text = ""
txtAptoPostal.Text = ""

txtEmail.Text = ""
txtEmail2.Text = ""

txtTelefono1.Text = ""
txtTelefono2.Text = ""
txtTelFax.Text = ""

txtCuenta.Text = ""
txtCuentaDesc.Text = ""

lswCuentas.ListItems.Clear


End Sub


Private Sub sbConsulta(pCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select *" _
       & " from vIVR_ADMINISTRADORES" _
       & " Where COD_ADMINISTRADOR = '" & pCodigo & "'"

Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    
    Call sbBarra_Accion("activo")
  
    tcMain.Item(0).Selected = True
    
    vEdita = True
    
    vCodigo = rs!Cod_Administrador
    txtCodigo.Text = CStr(rs!Cod_Administrador)
  
 
    txtNombre.Text = rs!Descripcion & ""
    txtIdentificacion.Text = rs!Identificacion & ""
    
    
    Call sbCboAsignaDato(cboBancos, rs!Banco_Desc, True, rs!Tes_Banco)
    Call sbCboAsignaDato(cboTipo, rs!Tipo_Desc, True, rs!Tipo_Administrador)
    
    If rs!Estado = "A" Then
      cboEstado.Text = "Activo"
    Else
      cboEstado.Text = "InActivo"
    End If
    
   
    txtDireccion.Text = rs!Direccion & ""
    
    txtAptoPostal.Text = rs!Apto_Postal & ""
    txtWebSite.Text = rs!WebSite
        
    txtEmail.Text = rs!Email_01 & ""
    txtEmail2.Text = rs!Email_02 & ""
    
    txtTelefono1.Text = rs!telefono1 & ""
    txtTelefono2.Text = rs!telefono2 & ""
    txtTelFax.Text = rs!fax & ""

    txtCuenta.Text = rs!COD_CUENTA_TRANSITO_MASK
    txtCuentaDesc.Text = rs!COD_CUENTA_TRANSITO_DESC
    
    cboEmitir.Text = fxTipoDocumento(rs!Tes_Emitir)
    
Else
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close


If vCodigo <> "" Then
    Call sbCuentas_Load
End If

Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vCuenta As String, vDivisa As String
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Verifica que exista ningun otro Administrador con la misma cedula juridica
strSQL = "select isnull(count(*),0) as Existe from IVR_ADMINISTRADOR" _
       & " where COD_ADMINISTRADOR not in('" & vCodigo & "') and Identificacion = '" _
       & Trim(txtIdentificacion) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   vMensaje = vMensaje & vbCrLf & " - Existe ya un Administrador registrado con la misma Cédula Jurídica ..."
End If
rs.Close


If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, txtCuenta, 0)) Then
   vMensaje = vMensaje & vbCrLf & " - No se especificó una cuenta contable válida..."
End If

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Administrador no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim vDivisa As String, vCuenta As String

On Error GoTo vError

vCuenta = fxgCntCuentaFormato(False, txtCuenta)

strSQL = "select count(*) as 'Existe' from IVR_ADMINISTRADOR Where COD_ADMINISTRADOR = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   vCodigo = txtCodigo.Text

   strSQL = "insert into IVR_ADMINISTRADOR(COD_ADMINISTRADOR, Tipo_Administrador, Identificacion, descripcion" _
          & ", estado, telefono1, telefono2, Fax, email_01, email_02, apto_postal, WebSite" _
          & ", direccion, COD_CUENTA_TRANSITO, Tes_Emitir, Tes_Banco, REGISTRO_FECHA, REGISTRO_USUARIO)" _
          & " values('" & vCodigo & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "','" & Trim(txtIdentificacion.Text) _
          & "','" & txtNombre & "','" & Mid(cboEstado.Text, 1, 1) _
          & "','" & txtTelefono1 & "','" & txtTelefono2 & "','" & txtTelFax _
          & "','" & txtEmail.Text & "','" & txtEmail2.Text & "','" & txtAptoPostal _
          & "','" & txtWebSite.Text & "','" & txtDireccion & "','" & vCuenta _
          & "','" & fxTipoDocumento(cboEmitir.Text) & "'," & cboBancos.ItemData(cboBancos.ListIndex) _
          & ", dbo.mygetdate(), '" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Administrador: " & vCodigo)

Else
   
    
  strSQL = "update IVR_ADMINISTRADOR set descripcion = '" & Trim(txtNombre.Text) _
         & "', Identificacion = '" & Trim(txtIdentificacion.Text) & "', Tipo_Administrador = '" & cboTipo.ItemData(cboTipo.ListIndex) _
         & "', WebSite = '" & txtWebSite.Text & "', estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "', direccion = '" & txtDireccion.Text & "', apto_postal = '" & txtAptoPostal.Text _
         & "', email_01 = '" & txtEmail.Text & "', telefono1 = '" & txtTelefono1 _
         & "', email_02 = '" & txtEmail2.Text & "', telefono2= '" & txtTelefono2.Text & "',fax = '" & txtTelFax _
         & "', Tes_Banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & " , Tes_Emitir = '" & fxTipoDocumento(cboEmitir.Text) & "', COD_CUENTA_TRANSITO = '" & vCuenta & "'" _
         & "  where COD_ADMINISTRADOR = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  
  
  Call Bitacora("Modifica", "Administrador: " & vCodigo)
 
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete IVR_ADMINISTRADOR where COD_ADMINISTRADOR = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Administrador: " & vCodigo)
  Call sbLimpiaPantalla
 
  Call sbBarra_Accion("NUEVO")
  
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Administrador"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = "Id. Real"
  
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Identificacion"
  gBusquedas.Orden = "Identificacion"
  gBusquedas.Consulta = "select COD_ADMINISTRADOR,descripcion,Identificacion as 'Identificación' from IVR_ADMINISTRADOR"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If


End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Administrador"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_ADMINISTRADOR"
  gBusquedas.Orden = "COD_ADMINISTRADOR"
  gBusquedas.Consulta = "select COD_ADMINISTRADOR,Identificacion,descripcion from IVR_ADMINISTRADOR"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub



Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdentificacion.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuenta = gCuenta
End If

End Sub

Private Sub txtCuenta_LostFocus()
On Error GoTo vError

   txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuenta.Text, 0))
   txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text, 0)

vError:

End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuenta.SetFocus
End Sub


Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail2.SetFocus
End Sub

Private Sub txtEmail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub


Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdentificacion.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id. Administrador"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_ADMINISTRADOR,Identificacion,descripcion from IVR_ADMINISTRADOR"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelFax.SetFocus
End Sub


Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub
