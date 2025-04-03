VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_SeguimientoRefundiciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "REF"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10455
   HelpContextID   =   3024
   Icon            =   "frmCR_SeguimientoRefundiciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswRefunde 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   4043
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnEliminaRefundiciones 
      Height          =   375
      Left            =   8040
      TabIndex        =   43
      Top             =   5760
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Elimina Refundiciones"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_SeguimientoRefundiciones.frx":030A
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8400
      Top             =   360
   End
   Begin VB.Frame fraRefunde 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   10212
      Begin XtremeSuiteControls.GroupBox gbOpciones 
         Height          =   2532
         Left            =   7680
         TabIndex        =   34
         Top             =   480
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   4466
         _StockProps     =   79
         Caption         =   "Tipo de aplicación"
         ForeColor       =   8421504
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton rbOpcion 
            Height          =   372
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cancela Crédito"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOpcion 
            Height          =   372
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cancela Morosidad"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton rbOpcion 
            Height          =   372
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cuotas Pendientes"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
      End
      Begin XtremeSuiteControls.GroupBox gbDatos 
         Height          =   3375
         Left            =   3240
         TabIndex        =   21
         Top             =   480
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   5953
         _StockProps     =   79
         Caption         =   "Datos de Cancelación:"
         ForeColor       =   8421504
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtIntCor 
            Height          =   312
            Left            =   1920
            TabIndex        =   22
            Top             =   1200
            Width           =   2052
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtIntMor 
            Height          =   312
            Left            =   1920
            TabIndex        =   23
            Top             =   1560
            Width           =   2052
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAmortizacion 
            Height          =   312
            Left            =   1920
            TabIndex        =   24
            Top             =   840
            Width           =   2052
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCargos 
            Height          =   312
            Left            =   1920
            TabIndex        =   25
            Top             =   1920
            Width           =   2052
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizas 
            Height          =   312
            Left            =   1920
            TabIndex        =   30
            Top             =   2280
            Width           =   2052
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTotal 
            Height          =   315
            Left            =   1920
            TabIndex        =   32
            Top             =   3000
            Width           =   2055
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSaldo 
            Height          =   312
            Left            =   1920
            TabIndex        =   38
            Top             =   360
            Width           =   2052
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtIVA 
            Height          =   315
            Left            =   1920
            TabIndex        =   44
            Top             =   2640
            Width           =   2055
            _Version        =   1441793
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   45
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Saldo"
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
            Height          =   252
            Index           =   8
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label Label2 
            Caption         =   "Total del Abono"
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
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   33
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Pólizas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   7
            Left            =   240
            TabIndex        =   31
            Top             =   2280
            Width           =   1572
         End
         Begin VB.Label Label2 
            Caption         =   "Int.Corriente"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   1452
         End
         Begin VB.Label Label2 
            Caption         =   "Int.Moratorio"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   28
            Top             =   1560
            Width           =   1452
         End
         Begin VB.Label Label2 
            Caption         =   "Principal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   6
            Left            =   240
            TabIndex        =   27
            Top             =   840
            Width           =   1572
         End
         Begin VB.Label Label2 
            Caption         =   "Cargos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   9
            Left            =   240
            TabIndex        =   26
            Top             =   1920
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.PushButton btnRefunde 
         Height          =   495
         Left            =   7920
         TabIndex        =   11
         Top             =   3360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Refunde"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCR_SeguimientoRefundiciones.frx":08AE
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnCerrar 
         Height          =   495
         Left            =   9360
         TabIndex        =   12
         Top             =   3360
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCR_SeguimientoRefundiciones.frx":0FD5
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   312
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   10212
         _Version        =   1441793
         _ExtentX        =   18013
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Datos de la Refundición o Abono a la operación:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
      End
      Begin VB.Label Label2 
         Caption         =   "Línea"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "Operación"
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
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   972
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18018
      _ExtentY        =   7011
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
      Item(0).Caption =   "Propias"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lswPrestamos"
      Item(1).Caption =   "Terceros"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "lswTerceros"
      Item(1).Control(1)=   "txtConCedula"
      Item(1).Control(2)=   "txtConNombre"
      Item(1).Control(3)=   "Label1(26)"
      Begin XtremeSuiteControls.ListView lswTerceros 
         Height          =   3135
         Left            =   -70000
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   10215
         _Version        =   1441793
         _ExtentX        =   18018
         _ExtentY        =   5530
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
      End
      Begin XtremeSuiteControls.ListView lswPrestamos 
         Height          =   3495
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   10215
         _Version        =   1441793
         _ExtentX        =   18018
         _ExtentY        =   6165
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
      End
      Begin XtremeSuiteControls.FlatEdit txtConCedula 
         Height          =   312
         Left            =   -67000
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtConNombre 
         Height          =   312
         Left            =   -65320
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9758
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   26
         Left            =   -68200
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.PushButton btnActualizar 
      Height          =   492
      Left            =   8880
      TabIndex        =   4
      Top             =   240
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Actualizar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCR_SeguimientoRefundiciones.frx":1579
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtTMonto 
      Height          =   315
      Left            =   3600
      TabIndex        =   40
      ToolTipText     =   "Abono Total"
      Top             =   8520
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTCuota 
      Height          =   315
      Left            =   5280
      TabIndex        =   42
      ToolTipText     =   "Cuota a Liberar"
      Top             =   8520
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totales:"
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
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   41
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   20
      Top             =   1280
      Width           =   1095
   End
   Begin VB.Label lblDisponible 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Height          =   312
      Left            =   8160
      TabIndex        =   19
      Top             =   1284
      Width           =   2052
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   10212
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Operaciones activas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Refundiciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Refundición de Créditos"
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
      Height          =   372
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   5412
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCR_SeguimientoRefundiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OpARefundir
  Operacion As Long
  Saldo     As Currency
  Amortiza  As Currency
  IntCor    As Currency
  IntMor    As Currency
  Cargos    As Currency
  Polizas   As Currency
  IVA       As Currency
  Total     As Currency
  Tipo      As String
End Type

Dim mRefunde As OpARefundir
Dim curPrimerCuota As Currency, curPoliza As Currency, curInteres As Currency


Private Sub btnActualizar_Click()
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spCrdSGTRefundicionesActualiza " & Operacion.Operacion
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Estado de las Operaciones a Refinanciar o Abonar actualizado!", vbInformation

Call sbInicializa
Call LimpiaDatos(False)



Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnCerrar_Click()
Call LimpiaDatos(False)
End Sub

Private Sub btnEliminaRefundiciones_Click()
Dim strSQL As String, i As Integer


On Error GoTo vError

strSQL = ""

With lswRefunde.ListItems

 For i = 1 To .Count
    If .Item(i).Checked = True Then

        strSQL = strSQL & Space(10) & "delete refundiciones where id_solicitud = " & .Item(i).Text _
               & " and id_solicitudr = " & Operacion.Operacion
        
        lblDisponible.Caption = CCur(lblDisponible.Caption) + CCur(.Item(i).SubItems(10))
        lblDisponible.Caption = Format(lblDisponible, "Standard")
    
    End If
 Next i

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
End If

End With


Call sbCargaRefundiciones
Call sbCargaPrestamos

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub btnRefunde_Click()
Call sbRefunde
End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

With lswRefunde.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Garantía", 1400, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Tipo", 1800, vbCenter
    .Add , , "IVA", 1800, vbRightJustify
End With


With lswPrestamos.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Garantía", 1400, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Tipo", 1800, vbCenter
    .Add , , "IVA", 1800, vbRightJustify
End With

With lswTerceros.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Garantía", 1400, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Tipo", 1800, vbCenter
    .Add , , "IVA", 1800, vbRightJustify
End With



tcMain.Item(0).Selected = True

fraRefunde.top = tcMain.top
fraRefunde.Left = tcMain.Left
fraRefunde.Height = tcMain.Height
fraRefunde.Width = tcMain.Width

curPrimerCuota = 0
curPoliza = 0
curInteres = 0


End Sub

Private Sub sbCargaRefundiciones()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem
Dim pCuota As Currency, pMonto As Currency


pCuota = 0
pMonto = 0

On Error GoTo vError

strSQL = "exec spCrd_SGT_Refundiciones_Lista " & Operacion.Operacion
Call OpenRecordSet(rs, strSQL, 0)

With lswRefunde
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!GarantiaX
     itmX.SubItems(3) = rs!DESCRIPCION
     itmX.SubItems(4) = Format(rs!Saldo_anterior, "Standard")
     itmX.SubItems(5) = Format(IIf(IsNull(rs!IntCor), 0, rs!IntCor), "Standard")
     itmX.SubItems(6) = Format(IIf(IsNull(rs!IntMor), 0, rs!IntMor), "Standard")
     itmX.SubItems(7) = Format(IIf(IsNull(rs!Cargos), 0, rs!Cargos), "Standard")
     itmX.SubItems(8) = Format(IIf(IsNull(rs!Polizas), 0, rs!Polizas), "Standard")
     itmX.SubItems(9) = Format(IIf(IsNull(rs!Principal), 0, rs!Principal), "Standard")
     itmX.SubItems(10) = Format(rs!Monto, "Standard")
     
     Select Case rs!Tipo
       Case "C"
            itmX.SubItems(11) = "Cancela Crédito"
       Case "P"
            itmX.SubItems(11) = "Pendientes"
       Case "M"
            itmX.SubItems(11) = "Morosidad"
     End Select
     
     itmX.SubItems(12) = "0.00" 'Format(rs!IVA, "Standard")
     
     pMonto = pMonto + rs!Monto
'     pCuota = pCuota + rs!Cuota
   
   rs.MoveNext
  Loop
End With
rs.Close

txtTCuota.Text = Format(pCuota, "Standard")
txtTMonto.Text = Format(pMonto, "Standard")


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbCargaPrestamos()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

'strSQL = "exec spCrdSGTListaCreditosPersona '" & Operacion.Cedula & "','N','S','" & Operacion.Codigo & "'"

strSQL = "exec spCrd_SGT_Persona_Creditos_Pendientes_Lista " & Operacion.Operacion & ", '" & Operacion.Cedula & "','N','S','" & Operacion.Codigo & "'"

Call OpenRecordSet(rs, strSQL, 0)
With lswPrestamos
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!GarantiaX
        itmX.SubItems(3) = rs!DESCRIPCION
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = Format(rs!IntC, "Standard")
        itmX.SubItems(6) = Format(rs!IntM, "Standard")
        itmX.SubItems(7) = Format(rs!Amortiza, "Standard")
        itmX.SubItems(8) = Format(rs!Cargos, "Standard")
        itmX.SubItems(9) = Format(rs!Polizas, "Standard")
        itmX.SubItems(10) = Format(rs!Amortiza + rs!IntC + rs!IntM + rs!Cargos + rs!Polizas, "Standard")
     
        Select Case rs!Tipo
          Case "C"
               itmX.SubItems(11) = "Cancela Crédito"
          Case "P"
               itmX.SubItems(11) = "Pendientes"
          Case "M"
               itmX.SubItems(11) = "Morosidad"
        End Select
        
        itmX.SubItems(12) = "0.00" 'Format(rs!IVA, "Standard")
   
   rs.MoveNext
  Loop
End With
rs.Close


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxExisteRefundicion(vOperacion As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from refundiciones" _
       & " where id_solicitud = " & vOperacion & " and id_solicitudr=" _
       & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)
 
fxExisteRefundicion = IIf((rs!Existe = 0), False, True)
rs.Close
End Function

Private Sub LimpiaDatos(Optional vVisible As Boolean = True)

txtCodigo.Text = ""
txtOperacion.Text = ""


txtTotal.Text = ""
txtCargos.Text = ""
txtPolizas.Text = ""
txtIntCor.Text = ""
txtIntMor.Text = ""
txtSaldo.Text = ""
txtAmortizacion.Text = ""
txtIVA.Text = ""

mRefunde.IVA = 0
mRefunde.Amortiza = 0
mRefunde.IntCor = 0
mRefunde.IntMor = 0
mRefunde.Saldo = 0
mRefunde.Cargos = 0
mRefunde.Polizas = 0
mRefunde.Total = 0
mRefunde.Tipo = "C"
mRefunde.Operacion = 0


If vVisible Then
   fraRefunde.Visible = vVisible
Else
   fraRefunde.Visible = vVisible
End If

End Sub



Private Function fxValidaRefundicion() As Boolean
Dim vMensaje As String

fxValidaRefundicion = True
vMensaje = ""

If mRefunde.Operacion = 0 Then vMensaje = vMensaje & "- No se ha seleccionado ninguna operación"

If IsNumeric(txtSaldo.Text) Then
 If txtSaldo.Text > mRefunde.Saldo Then vMensaje = vMensaje & vbCrLf & "- El saldo es mayor que el Original"
 If txtSaldo.Text < 0 Then vMensaje = vMensaje & vbCrLf & "- El saldo no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- El saldo no es válido"
End If

If Len(vMensaje) > 0 Then
 fxValidaRefundicion = False
 MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbRefunde()
Dim strSQL As String, curRefundir As Currency
Dim vTipo As String

On Error GoTo vError

If fxValidaRefundicion Then

'curRefundir = CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtAmortizacion.Text) + CCur(txtCargos.Text)

curRefundir = CCur(txtTotal.Text)

If curRefundir > CCur(lblDisponible.Caption) Then
  MsgBox "El monto a refundir de la operación es mayor al disponible...", vbCritical
  Exit Sub
End If

If fxExisteRefundicion(txtOperacion.Text) Then
  MsgBox "Esta Refundición Se encuentra Registrada VERIFIQUE...", vbInformation
  Exit Sub
Else
  
  Select Case True
    Case rbOpcion.Item(0).Value
        vTipo = "C"
    Case rbOpcion.Item(1).Value
        vTipo = "M"
    Case rbOpcion.Item(2).Value
        vTipo = "P"
  End Select
  
  strSQL = "insert refundiciones(id_solicitud,codigo,monto,fecha,codigor,id_solicitudr,intcor,intmor,saldo_anterior,cargos, polizas, principal, tipo) " _
         & "values(" & txtOperacion.Text & ",'" & txtCodigo.Text & "'," & CCur(txtTotal.Text) _
         & ",dbo.MyGetdate(),'" & Operacion.Codigo & "'," & Operacion.Operacion & "," & CCur(txtIntCor.Text) _
         & "," & CCur(txtIntMor.Text) & "," & CCur(txtSaldo.Text) & "," & CCur(txtCargos.Text) _
         & "," & CCur(txtPolizas.Text) & "," & CCur(txtAmortizacion.Text) & ",'" & vTipo & "')"
  Call ConectionExecute(strSQL)
  
  lblDisponible.Caption = CCur(lblDisponible.Caption) - CCur(txtTotal.Text)
  lblDisponible.Caption = Format(lblDisponible, "Standard")
  
  Call sbCargaRefundiciones
  Call sbCargaPrestamos
  
  Call LimpiaDatos(False)
  
End If

End If 'Verificacion de OPERACION

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswPrestamos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

   
Call LimpiaDatos(True)

txtOperacion.Text = Item.Text
txtCodigo.Text = Item.SubItems(1)

txtSaldo.Text = Format(CCur(Item.SubItems(4)), "Standard")

txtIntCor.Text = Format(CCur(Item.SubItems(5)), "Standard")
txtIntMor.Text = Format(CCur(Item.SubItems(6)), "Standard")
txtAmortizacion.Text = Format(CCur(Item.SubItems(7)), "Standard")

txtCargos.Text = Format(CCur(Item.SubItems(8)), "Standard")
txtPolizas.Text = Format(CCur(Item.SubItems(9)), "Standard")


txtIVA.Text = Format(CCur(Item.SubItems(12)), "Standard")
txtTotal.Text = Format(CCur(Item.SubItems(10)), "Standard")

Select Case Mid(Item.SubItems(11), 1, 1)
    Case "C"
        rbOpcion.Item(0).Value = True
        Call rbOpcion_Click(0)
    Case "M"
        rbOpcion.Item(1).Value = True
        Call rbOpcion_Click(1)
    Case "P"
        rbOpcion.Item(2).Value = True
        Call rbOpcion_Click(2)
End Select


mRefunde.Operacion = txtOperacion.Text

mRefunde.Saldo = txtSaldo.Text

mRefunde.Amortiza = txtAmortizacion.Text
mRefunde.IntCor = txtIntCor.Text
mRefunde.IntMor = txtIntMor.Text
mRefunde.Cargos = txtCargos.Text
mRefunde.Polizas = txtPolizas.Text

mRefunde.IVA = txtIVA.Text
mRefunde.Total = txtTotal.Text
mRefunde.Tipo = Mid(Item.SubItems(11), 1, 1)


fraRefunde.Visible = True



Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaLswTerceros(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

strSQL = "exec spCrdSGTListaCreditosPersona '" & vCedula & "','N','S','" & Operacion.Codigo & "'"
Call OpenRecordSet(rs, strSQL, 0)
With lswTerceros
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!GarantiaX
        itmX.SubItems(3) = rs!DESCRIPCION
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = Format(rs!IntC, "Standard")
        itmX.SubItems(6) = Format(rs!IntM, "Standard")
        itmX.SubItems(7) = Format(rs!Amortiza, "Standard")
        itmX.SubItems(8) = Format(rs!Cargos, "Standard")
        
        itmX.SubItems(9) = Format(rs!Polizas, "Standard")
        itmX.SubItems(10) = Format(rs!Amortiza + rs!IntC + rs!IntM + rs!Cargos + rs!Polizas, "Standard")
     
        Select Case rs!Tipo
          Case "C"
               itmX.SubItems(11) = "Cancela Crédito"
          Case "P"
               itmX.SubItems(11) = "Pendientes"
          Case "M"
               itmX.SubItems(11) = "Morosidad"
        End Select
        
        itmX.SubItems(12) = "0.00" 'Format(rs!IVA, "Standard")
        
        itmX.Tag = itmX.Index
   rs.MoveNext
  Loop
End With
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub lswTerceros_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

   
Call LimpiaDatos(True)

txtOperacion.Text = Item.Text
txtCodigo.Text = Item.SubItems(1)

txtSaldo.Text = Format(CCur(Item.SubItems(4)), "Standard")

txtIntCor.Text = Format(CCur(Item.SubItems(5)), "Standard")
txtIntMor.Text = Format(CCur(Item.SubItems(6)), "Standard")
txtAmortizacion.Text = Format(CCur(Item.SubItems(7)), "Standard")

txtCargos.Text = Format(CCur(Item.SubItems(8)), "Standard")
txtPolizas.Text = Format(CCur(Item.SubItems(9)), "Standard")

txtTotal.Text = Format(CCur(Item.SubItems(10)), "Standard")
txtIVA.Text = Format(CCur(Item.SubItems(12)), "Standard")

Select Case Mid(Item.SubItems(11), 1, 1)
    Case "C"
        rbOpcion.Item(0).Value = True
        Call rbOpcion_Click(0)
    Case "M"
        rbOpcion.Item(1).Value = True
        Call rbOpcion_Click(1)
    Case "P"
        rbOpcion.Item(2).Value = True
        Call rbOpcion_Click(2)
End Select


mRefunde.Operacion = txtOperacion.Text

mRefunde.Saldo = txtSaldo.Text

mRefunde.Amortiza = txtAmortizacion.Text
mRefunde.IntCor = txtIntCor.Text
mRefunde.IntMor = txtIntMor.Text
mRefunde.Cargos = txtCargos.Text
mRefunde.Polizas = txtPolizas.Text
mRefunde.IVA = txtIVA.Text
mRefunde.Total = txtTotal.Text
mRefunde.Tipo = Mid(Item.SubItems(11), 1, 1)


fraRefunde.Visible = True


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub rbOpcion_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String

On Error GoTo vError


Me.MousePointer = vbHourglass

Select Case Index
    Case 0 'Cancela Credito
        vTipo = "C"
    Case 1 'Cancela Morosidad
        vTipo = "M"
    Case 2 'Cancela Pendientes
        vTipo = "P"
End Select

strSQL = "exec spCrd_SGT_Refunde_Datos " & txtOperacion.Text & ",'" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)


txtSaldo.Text = Format(rs!Saldo, "Standard")


txtAmortizacion.Text = Format(rs!Principal, "Standard")
txtIntCor.Text = Format(rs!IntCor, "Standard")
txtIntMor.Text = Format(rs!IntMor, "Standard")
txtCargos.Text = Format(rs!Cargos, "Standard")
txtPolizas.Text = Format(rs!Polizas, "Standard")

txtIVA.Text = "0.00" 'Format(rs!IV, "Standard")

txtTotal.Text = Format(rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos + rs!Polizas, "Standard")

rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
    Call sbCargaLswTerceros(txtConCedula)
End If

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select R.Primer_Cuota,R.Garantia,R.montoapr,R.cuota,R.int,C.convenio,R.cod_Destino" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " where R.id_solicitud =" & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)


If fxCobraTasaFormaliza(rs!cod_destino & "") Then
 ' curInteres = fxInteresesHastaFormalizar(Operacion.FechaDesembolso)
  curInteres = fxInteresesHastaFormalizar(Operacion.FechaDesembolso, , Operacion.PriDeduc, Operacion.DiaPago)
End If


If rs!PRIMER_CUOTA = "S" Then
 curPrimerCuota = rs!Cuota
 If curInteres > 0 Then 'No es convenio
     curInteres = fxInteresesDiasPrimerCuota(Operacion.FechaDesembolso, rs!montoapr, rs!Int)
 End If
End If

If rs!Garantia <> "H" And rs!Convenio = "N" Then curPoliza = fxCuotaPolizaVida(rs!montoapr)

rs.Close
    
'Si el credito es de excedentes validar que la tasa sea cero, para que pueda refundir el anterior
If fxCreditoExcedente(Operacion.Codigo) Then curInteres = 0
    
Me.Caption = "Refundiciones Operación : " & Operacion.Operacion
lblDisponible.Caption = Format(Operacion.MontoAprobado _
                      - (fxMontoEnGeneral(Operacion.Operacion) _
                      + curInteres + curPrimerCuota + curPoliza) _
                      , "Standard")
                      
Call sbCargaRefundiciones
Call sbCargaPrestamos
  
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub txtConCedula_Change()
lswTerceros.ListItems.Clear
End Sub

Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtConNombre.Text = fxNombre(txtConCedula)
    Call sbCargaLswTerceros(txtConCedula)
End If

End Sub

