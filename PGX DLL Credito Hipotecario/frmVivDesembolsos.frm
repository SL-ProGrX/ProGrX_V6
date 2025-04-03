VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmVivDesembolsos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control Desembolsos"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5655
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   9615
      _Version        =   1441793
      _ExtentX        =   16960
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
      ItemCount       =   4
      Item(0).Caption =   "Desembolsos"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "Label3(5)"
      Item(0).Control(2)=   "txtTotalMontoGirado"
      Item(0).Control(3)=   "txtDesembolsoId"
      Item(0).Control(4)=   "btnDesembolsoBoleta"
      Item(1).Caption =   "( + ) Nuevo"
      Item(1).ControlCount=   24
      Item(1).Control(0)=   "rbEnte(0)"
      Item(1).Control(1)=   "vGrid"
      Item(1).Control(2)=   "txtId"
      Item(1).Control(3)=   "txtBeneficiario"
      Item(1).Control(4)=   "txtDetalle"
      Item(1).Control(5)=   "cboBanco"
      Item(1).Control(6)=   "cboCuenta"
      Item(1).Control(7)=   "cboTipoDocumento"
      Item(1).Control(8)=   "rbEnte(1)"
      Item(1).Control(9)=   "rbEnte(2)"
      Item(1).Control(10)=   "btnCuentas"
      Item(1).Control(11)=   "Label1(27)"
      Item(1).Control(12)=   "Label1(17)"
      Item(1).Control(13)=   "Label1(16)"
      Item(1).Control(14)=   "Label1(10)"
      Item(1).Control(15)=   "Label1(12)"
      Item(1).Control(16)=   "Label1(6)"
      Item(1).Control(17)=   "Label1(13)"
      Item(1).Control(18)=   "Label1(14)"
      Item(1).Control(19)=   "Label1(5)"
      Item(1).Control(20)=   "txtNuevoDisponible"
      Item(1).Control(21)=   "txtMonto"
      Item(1).Control(22)=   "txtIntActuales"
      Item(1).Control(23)=   "txtAplicaIntereses"
      Item(2).Caption =   "Pendientes"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lswPendientes"
      Item(2).Control(1)=   "gbContacto"
      Item(3).Caption =   "( + ) Pendientes"
      Item(3).ControlCount=   17
      Item(3).Control(0)=   "txtIdGarantia"
      Item(3).Control(1)=   "txtIdContacto"
      Item(3).Control(2)=   "txtTipoDesmbolsoPend"
      Item(3).Control(3)=   "txtBeneficiarioPend"
      Item(3).Control(4)=   "txtMontoDesembolsoPend"
      Item(3).Control(5)=   "chkAplicaIntereses"
      Item(3).Control(6)=   "cmdAgregarPendientes"
      Item(3).Control(7)=   "Label1(26)"
      Item(3).Control(8)=   "Label1(25)"
      Item(3).Control(9)=   "Label1(24)"
      Item(3).Control(10)=   "Label1(23)"
      Item(3).Control(11)=   "Label1(22)"
      Item(3).Control(12)=   "Label1(21)"
      Item(3).Control(13)=   "txtNombreContacto"
      Item(3).Control(14)=   "txtNumeroFinca"
      Item(3).Control(15)=   "txtDescripDesembolsoPend"
      Item(3).Control(16)=   "txtNuevoDisponibleAgregarPend"
      Begin XtremeSuiteControls.ListView lswPendientes 
         Height          =   1935
         Left            =   -70000
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4575
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   8070
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
      Begin XtremeSuiteControls.GroupBox gbContacto 
         Height          =   3375
         Left            =   -70000
         TabIndex        =   70
         Top             =   2280
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   5953
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtContactoId 
            Height          =   330
            Left            =   1560
            TabIndex        =   75
            Top             =   240
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin XtremeSuiteControls.FlatEdit txtContactoCedula 
            Height          =   330
            Left            =   4800
            TabIndex        =   76
            Top             =   240
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin XtremeSuiteControls.FlatEdit txtContactoTipo 
            Height          =   330
            Left            =   7680
            TabIndex        =   77
            Top             =   240
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin XtremeSuiteControls.FlatEdit txtContactoNombre 
            Height          =   330
            Left            =   1560
            TabIndex        =   78
            Top             =   600
            Width           =   7695
            _Version        =   1441793
            _ExtentX        =   13573
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
         Begin XtremeSuiteControls.PushButton cmdActivar 
            Height          =   375
            Left            =   4200
            TabIndex        =   79
            Top             =   1200
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Activar"
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
            Appearance      =   17
            Picture         =   "frmVivDesembolsos.frx":0000
         End
         Begin XtremeSuiteControls.PushButton cmdNuevoDesembolsoPendiente 
            Height          =   375
            Left            =   5520
            TabIndex        =   80
            Top             =   1200
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Nuevo"
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
            Appearance      =   17
            Picture         =   "frmVivDesembolsos.frx":0727
         End
         Begin XtremeSuiteControls.FlatEdit txtMontoP 
            Height          =   315
            Left            =   2040
            TabIndex        =   81
            Top             =   1200
            Width           =   1815
            _Version        =   1441793
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDescuento 
            Height          =   315
            Left            =   2040
            TabIndex        =   82
            Top             =   1560
            Width           =   1815
            _Version        =   1441793
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtIntActualesPendientes 
            Height          =   315
            Left            =   2040
            TabIndex        =   83
            Top             =   1920
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMontoGirar 
            Height          =   315
            Left            =   2040
            TabIndex        =   84
            Top             =   2280
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCambiar 
            Height          =   375
            Left            =   6960
            TabIndex        =   85
            Top             =   1200
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cambiar"
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
            Appearance      =   17
            Picture         =   "frmVivDesembolsos.frx":0E47
         End
         Begin XtremeSuiteControls.FlatEdit txtNuevoDisponiblePendientes 
            Height          =   315
            Left            =   2040
            TabIndex        =   91
            Top             =   2760
            Width           =   1815
            _Version        =   1441793
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nuevo Disponible"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   20
            Left            =   360
            TabIndex        =   90
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Int. Actuales"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   19
            Left            =   360
            TabIndex        =   89
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   11
            Left            =   360
            TabIndex        =   88
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   9
            Left            =   360
            TabIndex        =   87
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   360
            TabIndex        =   86
            Top             =   1200
            Width           =   1335
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   74
            Top             =   720
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nombre"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   2
            Left            =   6360
            TabIndex        =   73
            Top             =   240
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   72
            Top             =   240
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Identificación"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   71
            Top             =   240
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Id Conctato"
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
      Begin XtremeSuiteControls.CheckBox chkAplicaIntereses 
         Height          =   375
         Left            =   -66040
         TabIndex        =   67
         Top             =   2760
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Aplica Intereses ?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoDesmbolsoPend 
         Height          =   315
         Left            =   -68080
         TabIndex        =   63
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIdGarantia 
         Height          =   315
         Left            =   -68080
         TabIndex        =   61
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIdContacto 
         Height          =   315
         Left            =   -68080
         TabIndex        =   59
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton rbEnte 
         Height          =   255
         Index           =   0
         Left            =   -65680
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Deudor"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   1575
         Left            =   -68680
         TabIndex        =   25
         Top             =   2880
         Visible         =   0   'False
         Width           =   7335
         _Version        =   524288
         _ExtentX        =   12938
         _ExtentY        =   2778
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   486
         MaxRows         =   498
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmVivDesembolsos.frx":1560
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtId 
         Height          =   315
         Left            =   -68200
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4254
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
         Height          =   315
         Left            =   -68200
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12086
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
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   675
         Left            =   -68200
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12086
         _ExtentY        =   1185
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   315
         Left            =   -68200
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12091
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   315
         Left            =   -64960
         TabIndex        =   30
         Top             =   2400
         Visible         =   0   'False
         Width           =   3615
         _Version        =   1441793
         _ExtentX        =   6376
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
         Height          =   315
         Left            =   -68200
         TabIndex        =   31
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.RadioButton rbEnte 
         Height          =   255
         Index           =   1
         Left            =   -64480
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Profesional"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.RadioButton rbEnte 
         Height          =   255
         Index           =   2
         Left            =   -63040
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Otro"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.PushButton btnCuentas 
         Height          =   315
         Left            =   -61240
         TabIndex        =   34
         ToolTipText     =   "Ingresar Cuentas Bancarias"
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton cmdAgregarPendientes 
         Height          =   375
         Left            =   -66040
         TabIndex        =   45
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Agregar"
         BackColor       =   -2147483633
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalMontoGirado 
         Height          =   315
         Left            =   7440
         TabIndex        =   53
         Top             =   5040
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoDisponible 
         Height          =   315
         Left            =   -66520
         TabIndex        =   54
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   -63400
         TabIndex        =   55
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntActuales 
         Height          =   315
         Left            =   -63400
         TabIndex        =   56
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAplicaIntereses 
         Height          =   315
         Left            =   -66520
         TabIndex        =   57
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
         Text            =   "Aplica Intereses ?"
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombreContacto 
         Height          =   315
         Left            =   -66280
         TabIndex        =   58
         Top             =   840
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
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
      Begin XtremeSuiteControls.FlatEdit txtNumeroFinca 
         Height          =   315
         Left            =   -66280
         TabIndex        =   60
         Top             =   1320
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
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
      Begin XtremeSuiteControls.FlatEdit txtDescripDesembolsoPend 
         Height          =   315
         Left            =   -66280
         TabIndex        =   62
         Top             =   1800
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiarioPend 
         Height          =   315
         Left            =   -68080
         TabIndex        =   64
         Top             =   2280
         Visible         =   0   'False
         Width           =   7455
         _Version        =   1441793
         _ExtentX        =   13150
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoDesembolsoPend 
         Height          =   315
         Left            =   -68080
         TabIndex        =   65
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
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
         Text            =   "0.00"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoDisponibleAgregarPend 
         Height          =   315
         Left            =   -68080
         TabIndex        =   66
         Top             =   3600
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDesembolsoId 
         Height          =   315
         Left            =   120
         TabIndex        =   68
         ToolTipText     =   "Seleccione un Desembolso de la Lista"
         Top             =   5040
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnDesembolsoBoleta 
         Height          =   315
         Left            =   1680
         TabIndex        =   69
         ToolTipText     =   "Ingresar Cuentas Bancarias"
         Top             =   5040
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         BackColor       =   -2147483633
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
         Appearance      =   17
         Picture         =   "frmVivDesembolsos.frx":1ADB
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   315
         Index           =   5
         Left            =   5640
         TabIndex        =   52
         Top             =   5040
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Total Girado:  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Desembolso:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   21
         Left            =   -69760
         TabIndex        =   51
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiario:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   22
         Left            =   -69760
         TabIndex        =   50
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   23
         Left            =   -69760
         TabIndex        =   49
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Disponible"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   24
         Left            =   -69760
         TabIndex        =   48
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   25
         Left            =   -69760
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Garantía:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   26
         Left            =   -69760
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiario"
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
         Index           =   5
         Left            =   -69640
         TabIndex        =   43
         Top             =   960
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " Int. Actuales"
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
         Height          =   315
         Index           =   14
         Left            =   -64600
         TabIndex        =   42
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo disponible"
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
         Height          =   315
         Index           =   13
         Left            =   -68680
         TabIndex        =   41
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
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
         Index           =   6
         Left            =   -69640
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " Monto"
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
         Height          =   315
         Index           =   12
         Left            =   -64600
         TabIndex        =   39
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
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
         Index           =   10
         Left            =   -69640
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Index           =   16
         Left            =   -65680
         TabIndex        =   37
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Emitir"
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
         Index           =   17
         Left            =   -69640
         TabIndex        =   36
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   -69640
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   9615
      _Version        =   1441793
      _ExtentX        =   16960
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Disponibilidad para desembolsos:"
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtMontoInicial 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1815
         _Version        =   1441793
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDisponibleBruto 
         Height          =   315
         Left            =   2040
         TabIndex        =   14
         Top             =   720
         Width           =   1815
         _Version        =   1441793
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntAcumulado 
         Height          =   315
         Left            =   3840
         TabIndex        =   16
         Top             =   720
         Width           =   1815
         _Version        =   1441793
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntSobreDisponible 
         Height          =   315
         Left            =   5640
         TabIndex        =   18
         Top             =   720
         Width           =   1815
         _Version        =   1441793
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGiroMaximo 
         Height          =   315
         Left            =   7440
         TabIndex        =   20
         Top             =   720
         Width           =   1815
         _Version        =   1441793
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   4
         Left            =   7440
         TabIndex        =   21
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Giro Máximo"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   19
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Int.S/Disponible"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   17
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Int. Acumulados"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Disponible Bruto"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto del Inicial"
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
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   9870
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   4260
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.ToolTipText     =   "reporte de desembolsos"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11451
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDesLineaCredito 
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Top             =   960
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11451
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLinea 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   435
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   762
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label lblMensaje 
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   480
      Width           =   6615
      _Version        =   1441793
      _ExtentX        =   11663
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
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
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   " Cédula"
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
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   " Línea"
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
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   " Operación"
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
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmVivDesembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vPaso As Boolean

Private m_cambioDatos As Boolean
Private m_IdDesembolso As Long
Private m_DiasActuales As Integer
Private m_TasaDiaria As Double
Private m_FechaCorte As Date
Private m_DiasAcum As Integer
Private m_FechaUltimoCorte As Date
Private m_AplicaInteresesDesembolsosP As Integer
Private m_IntAcomuladoPendientes As Double

Public ItemSeleccionado As ListViewItem


'***********************Rutinas de Usuario ******************************************

Private Sub sbImprimir()
 
Dim subreporte As String

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
'    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes de Desembolsos"
      
    .Connect = glogon.ConectRPT
    .Destination = crptToWindow
    
    .Formulas(1) = "fxFecha =  '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "fxUsuario = '" & glogon.Usuario & "'"
    .Formulas(3) = "fxTitulo = 'Boleta de Desembolsos Registrados'"
    .Formulas(4) = "fxSubTitulo = 'Control de Disponible'"
    .Formulas(5) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
    
    .Formulas(6) = "fxDisponibleInicial = '" & txtMontoInicial.Text & "'"
    .Formulas(7) = "fxDisponibleBruto = '" & txtDisponibleBruto.Text & "'"
    .Formulas(8) = "fxIntAcumulados = '" & txtIntAcumulado.Text & "'"
    .Formulas(9) = "fxIntSAcumulados = '" & txtIntSobreDisponible.Text & "'"
    .Formulas(10) = "fxGiroMaximo = '" & txtGiroMaximo.Text & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Credito_Hipotecario_Desembolso.rpt")
    .SelectionFormula = "{ViviendaDesembolsos.NumeroOperacion} = " & txtOperacion.Text
    
    .SubreportToChange = "DesembolsosPendientes"
    .SelectionFormula = "{ViviendaGarantia.NumeroOperacion} = " & txtOperacion.Text & "  And {ViviendaDesembolsosPendientes.ESTADO} = 'P' "

    .Action = 1
        
    
End With
 
 Me.MousePointer = vbDefault
Exit Sub
salir:

    Me.MousePointer = vbDefault
    frmContenedor.Crt.Formulas(1) = ""
    frmContenedor.Crt.Formulas(2) = ""
    frmContenedor.Crt.Formulas(3) = ""
    frmContenedor.Crt.Formulas(4) = ""
    frmContenedor.Crt.Formulas(5) = ""
    frmContenedor.Crt.Formulas(6) = ""
    frmContenedor.Crt.Formulas(7) = ""
    frmContenedor.Crt.Formulas(8) = ""
    frmContenedor.Crt.Formulas(9) = ""
    frmContenedor.Crt.Formulas(10) = ""
    frmContenedor.Crt.SelectionFormula = ""
    
    Exit Sub
    
vError:
    MsgBox "Ocurrió un error al imprimir los reportes solicitados - " & Err.Description, vbError

End Sub

Private Function sbAgregarDetalle() As Boolean
Dim vCodigo As String, i As Integer
Dim vMonto As Double

'Inicia Proceso
Me.MousePointer = vbHourglass

On Error GoTo vError
sbAgregarDetalle = True

'Inicia registro de detalle
For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 2
    If CCur(vGrid.Text) > 0 Then
        vMonto = CCur(vGrid.Text)
        vGrid.Col = 3
        vCodigo = vGrid.Text
        If Not ObjAgregar.fxViviendaDetalleDesembolso(m_IdDesembolso, vCodigo, vMonto, glogon.Usuario) Then
            Exit For
        End If
    End If
Next i

Exit Function

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

Private Function fxValidaDatos() As Boolean
Dim vMensaje As String
Dim vDisponibleBruto As Double
Dim vIntAcomulado As Double
Dim vBruto As Double
Dim vBrutoEnPantalla As Double

On Error GoTo vError

fxValidaDatos = False

ReDim gParametros(0 To 20)


vMensaje = ""
If m_IdDesembolso <> -1 Then  'Indica si ya se registro el desembolso
    fxValidaDatos = False
    MsgBox ("La información registrada para el desembolso no puede ser modificada.")
    Exit Function '
End If

If ObjConsultar.fxDesembolsoCalculo(txtOperacion.Text) Then
    With glogon.Recordset
        vDisponibleBruto = IIf(IsNull(.Fields!Bruto), 0, .Fields!Bruto)
        vIntAcomulado = IIf(IsNull(.Fields!IntAcumulado), 0, .Fields!IntAcumulado)
        vBruto = vDisponibleBruto - vIntAcomulado
    End With
    glogon.Recordset.Close
End If

'TODO Validar con Pedro esta linea
vBrutoEnPantalla = CCur(txtDisponibleBruto.Text) - CCur(txtIntAcumulado.Text)

If Format(vBruto, "Standard") <> Format(vBrutoEnPantalla, "Standard") Then vMensaje = vMensaje & " - El disponible bruto ha cambiado, es posible que ya se realizó un movimiento antes de aplicar este desembolso, Verifique los datos he intente de nuevo" & vbCrLf

If (CCur(txtDisponibleBruto.Text) - CCur(txtIntAcumulado.Text) < (CCur(txtMonto.Text) + CCur(txtIntActuales.Text))) Then
vMensaje = vMensaje & " - El monto calculado para el desembolso es mayor al disponible bruto.(Disponible bruto menos Intereses Acumulados) " & vbCrLf
End If

If Val(txtNuevoDisponible.Text) < 0 Then vMensaje = vMensaje & " - El monto disponible es inválido, no es posible continuar con el proceso." & vbCrLf

If Len(txtId.Text) = 0 Then vMensaje = vMensaje & " - Es requerido digitar la Identificación del beneficiario." & vbCrLf

If Len(txtBeneficiario.Text) = 0 Then vMensaje = vMensaje & " - Es requerido digitar un nombre para el beneficiario." & vbCrLf
If Len(txtDetalle.Text) = 0 Then vMensaje = vMensaje & " - Es requerido digitar un detalle para el desembolso." & vbCrLf
If cboBanco.ListCount = 0 Then vMensaje = vMensaje & " - Es requerido seleccionar un banco." & vbCrLf
If fxTipoDocumento(cboTipoDocumento.Text) = "TE" Then
    If Len(cboCuenta.Text) = 0 Then vMensaje = vMensaje & " - Es requerido una cuenta Bancaria para realizar la transferencia." & vbCrLf
End If
If Val(txtMonto.Text) = 0 Then vMensaje = vMensaje & " - Debe de ingresar un monto para el desembolso." & vbCrLf


If Val(txtNuevoDisponible.Text) = 0 Then
txtNuevoDisponible.Text = 0
End If

If Len(vMensaje) = 0 Then
  fxValidaDatos = True
Else
  fxValidaDatos = False
  Call ObjMensajes.deDatos("-1", vMensaje)
  Me.MousePointer = vbDefault
  Exit Function
End If


m_IdDesembolso = -1
gParametros(0) = m_IdDesembolso '@CodigoDesembolso int OUTPUT
gParametros(1) = txtOperacion.Text '@NumeroOperacion
gParametros(2) = glogon.Usuario '@RegistroUsuario
gParametros(3) = txtBeneficiario.Text ' @Beneficiario
gParametros(4) = CCur(txtMonto.Text)
gParametros(5) = IIf((Len(txtDetalle.Text) = 0), ObjNull.NullString, Trim(txtDetalle.Text))
gParametros(6) = CCur(txtNuevoDisponible.Text)   'disponible
gParametros(7) = 0  'Aplica Intereses S/N
If txtAplicaIntereses.Tag = "S" Then
gParametros(7) = 1 'disponible 'Aplica Intereses S/N
End If

gParametros(8) = Format(m_FechaCorte, "yyyy/mm/dd") '@FechaCorte
gParametros(9) = Format(m_FechaUltimoCorte, "yyyy/mm/dd") '@IntFechaCorteUltima
gParametros(10) = m_DiasActuales '@InteresesActualDias
gParametros(11) = CCur(txtIntActuales.Text) '@InteresesActualMonto
gParametros(12) = m_DiasAcum  '@InteresesAcumDias
gParametros(13) = CCur(txtIntAcumulado.Text)  '@InteresesAcumMonto

If txtAplicaIntereses.Tag = "S" Then
    gParametros(9) = Format(m_FechaCorte, "yyyy/mm/dd") '@IntFechaCorteUltima
Else 'No aplica intereses se registra fecha del corte anterior
    gParametros(12) = 0  '@InteresesAcumDias
    gParametros(13) = 0 '@InteresesAcumMonto
End If

gParametros(14) = cboBanco.ItemData(cboBanco.ListIndex)  '@BancoCodigo
gParametros(15) = fxTipoDocumento(cboTipoDocumento.Text)  '@BancoEmitir
gParametros(16) = cboCuenta.ItemData(cboCuenta.ListIndex)  '@BancoCuenta
gParametros(17) = txtId.Text

salir:
    Me.MousePointer = vbDefault
    Exit Function
    
vError:

    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbAgregar()
On Error GoTo vError

If tcMain.SelectedItem = 1 Then
    
    If fxValidaEstadoCredito = False Then
        MsgBox "El crédito debe estar activo y formalizado para realizar desembolsos"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    lblMensaje.Caption = ""
    
    If Not ObjConsultar.fxPermiteAvaluo(txtOperacion.Text) Then
        Me.MousePointer = vbDefault
        lblMensaje.Caption = "No es posible realizar el desembolso, la operación no esta formalizada"
        MsgBox ("No es posible realizar el desembolso, la operación no esta formalizada.")
        Exit Sub
    End If
    
    If fxValidaDatos() = False Then Exit Sub
    
    If (MsgBox("¿ Confirma que desea aplicar la información suministrada para este desembolso.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub
    'Inicia Transacción
    glogon.Conection.BeginTrans
    m_IdDesembolso = ObjAgregar.fxViviendaDesembolso(gParametros(0), gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                                     gParametros(5), gParametros(6), gParametros(7), gParametros(8), gParametros(9), _
                                                     gParametros(10), gParametros(11), gParametros(12), gParametros(13), _
                                                     gParametros(14), gParametros(15), gParametros(16), gParametros(17))
                        
    gParametros(0) = m_IdDesembolso
    If m_IdDesembolso <> -1 And m_IdDesembolso > 0 Then
        If sbAgregarDetalle Then
            MsgBox "Información fue registrada corretamente.", vbInformation
            Call sbLimpiaDatos
            Call sbHabilitaTab(1)
            Call txtOperacion_LostFocus
            Call sbToolBar(tlbPrincipal, "Nuevo")
        End If
    End If
    
    Call Bitacora("APLICA", "Desembolso vivienda operación: " & gParametros(1) & " Desembolso: " & m_IdDesembolso)

End If

Me.MousePointer = vbDefault
'Cierra Transacción
glogon.Conection.CommitTrans
Call sbTraerInformacionOperacion

Exit Sub

vError:

    glogon.Conection.RollbackTrans
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbMostraVentanBusqueda()

On Error GoTo vError

tcMain.Item(0).Selected = True
GLOBALES.gTag = ""

Call sbSIFForms("frmVivConsultaDesembolso", 1, , , False)
If GLOBALES.gTag <> "" Then
    txtOperacion.SetFocus
    txtOperacion.Text = GLOBALES.gTag
'   Call txtOperacion_LostFocus
    txtLinea.SetFocus
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbSumaLineas()
Dim i As Integer
Dim vIntActuales As Double
Dim vDisponible As Double


On Error GoTo vError

 
 
txtMonto.Text = Format(0, "Standard")
txtIntActuales.Text = Format(0, "Standard")
txtNuevoDisponible.Text = Format(0, "Standard")
txtAplicaIntereses.Tag = "N"
txtAplicaIntereses.Text = "No Aplica Intereses"
txtAplicaIntereses.BackColor = RGB(213, 245, 227)


For i = 1 To vGrid.MaxRows
    vGrid.Col = 2
    vGrid.Row = i
    txtMonto.Text = CCur(txtMonto.Text) + CCur(vGrid.Text)
    vGrid.Col = 4 'Columna contiene si cobrar intereses
    If vGrid.Value = True Then 'si la linea Aplica intereses
        vGrid.Col = 2
        If CCur(vGrid.Text) > 0 Then
            txtAplicaIntereses.Tag = "S"
            txtAplicaIntereses.Text = "Sí Aplica Intereses"
            txtAplicaIntereses.BackColor = RGB(250, 219, 216)
            
            vIntActuales = CCur(vGrid.Text) * m_TasaDiaria * m_DiasActuales
            txtIntActuales.Text = txtIntActuales.Text + vIntActuales
        End If
    End If
    
 Next i
  
  txtMonto.Text = Format(txtMonto.Text, "Standard")
  txtIntActuales.Text = Format(txtIntActuales.Text, "Standard")
  
  If txtAplicaIntereses.Tag = "S" Then
    vDisponible = txtDisponibleBruto.Text - (CCur(txtMonto.Text) + CCur(txtIntActuales.Text) + CCur(txtIntAcumulado.Text))
  Else
    vDisponible = txtDisponibleBruto.Text - (CCur(txtMonto.Text) + CCur(txtIntActuales.Text))
  End If
  
  txtNuevoDisponible.Text = Format(vDisponible, "Standard")
    
    
salir:
    Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)


End Sub

Private Sub sbHabilitaTab(ByVal pTab As Integer)

Select Case pTab

    Case 1 'Inicial para consulta
        tcMain.Item(0).Enabled = True
        tcMain.Item(1).Enabled = False
        tcMain.Item(3).Enabled = False
        tcMain.Item(0).Selected = True
        
    Case 2 'Hablitia cuando es nuevo
        tcMain.Item(0).Enabled = False
        tcMain.Item(1).Enabled = True
        tcMain.Item(3).Enabled = False
        tcMain.Item(1).Selected = True
        
    Case 3
        tcMain.Item(0).Enabled = True
        tcMain.Item(1).Enabled = False
        tcMain.Item(3).Enabled = False
        tcMain.Item(2).Selected = True
    
    Case 4 ' Habilita tab de +Pendientes
        tcMain.Item(0).Enabled = True
        tcMain.Item(1).Enabled = False
        tcMain.Item(3).Enabled = True
        tcMain.Item(3).Selected = True
    
    
    Case 5 'Hablitia todos
        tcMain.Item(0).Enabled = True
        tcMain.Item(1).Enabled = True


        
End Select


End Sub

Private Sub sbTraerInformacionOperacion()
On Error GoTo vError
 
If ObjConsultar.fxDesembolsoCalculo(txtOperacion.Text) Then
    With glogon.Recordset
    
    txtCedula.Text = Trim(!cedula)
    txtNombre.Text = (!Nombre)
    txtLinea.Text = (!codigo)
    txtDesLineaCredito.Text = (!Descripcion)
    
    txtDisponibleBruto.Text = IIf(IsNull(!Bruto), Format(0, "Standard"), Format(!Bruto, "Standard"))
    txtIntAcumulado.Text = IIf(IsNull(!IntAcumulado), Format(0, "Standard"), Format(!IntAcumulado, "Standard"))
    txtIntSobreDisponible.Text = IIf(IsNull(!IntSDisponible), Format(0, "Standard"), Format(!IntSDisponible, "Standard"))
    txtGiroMaximo.Text = IIf(IsNull(!GiroMaximo), Format(0, "Standard"), Format(!GiroMaximo, "Standard"))
    
    m_DiasAcum = IIf(IsNull(!DiasAcumulados), Format(0, "Standard"), !DiasAcumulados)
    m_TasaDiaria = IIf(IsNull(!TasaDiaria), Format(0, "Standard"), !TasaDiaria)
    m_FechaCorte = Format(!FechaCorte, "dd/mm/yyyy")
    m_DiasActuales = IIf(IsNull(!DiasActuales), Format(0, "Standard"), !DiasActuales)
    m_FechaUltimoCorte = Format(!FechaUltCorte, "dd/mm/yyyy")
    
    
    txtDisponibleBruto.ToolTipText = "Inicial: " & Format(!monto_Girado, "Standard")
    txtIntAcumulado.ToolTipText = "Dias: " & !DiasAcumulados
    txtMontoInicial.Text = Format(!monto_Girado, "Standard")
    txtMontoInicial.ToolTipText = "Tasa : " & !Tasa & "%"
    
    txtAplicaIntereses.ToolTipText = "Dias: " & !DiasActuales
    
    'cboTipoDocumento.Text = fxTipoDocumento(IIf(IsNull(rs!emitir), "OT", rs!emitir))
    
    End With
    m_IdDesembolso = -1
    Call sbListaDesembolsos(txtOperacion.Text)
End If

salir:
    Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub

    
Public Sub sbListaDesembolsos(ByVal pOperacion As Long)

Dim vItem As ListViewItem
Dim vLvw As XtremeSuiteControls.ListView
Dim vIdZona As Integer
Dim vMontoTotal As Double
Dim vKey As String


On Error GoTo error

m_IdDesembolso = -1
lsw.ColumnHeaders.Clear
lsw.ListItems.Clear

Set vLvw = Me.lsw

vLvw.ColumnHeaders.Add , , "Beneficiario", 4000
vLvw.ColumnHeaders.Add , , "Monto", 1800, 1
vLvw.ColumnHeaders.Add , , "Disponible", 1800, 1
vLvw.ColumnHeaders.Add , , "Fecha Corte", 2000
vLvw.ColumnHeaders.Add , , "Último Corte", 2000
vLvw.ColumnHeaders.Add , , "Int.Actuales Dias", 1000, 1
vLvw.ColumnHeaders.Add , , "Int.Actual Monto", 1500, 1
vLvw.ColumnHeaders.Add , , "Int.Acum.Dias", 1000, 1
vLvw.ColumnHeaders.Add , , "Int.Acum.Monto", 1500, 1
vLvw.ColumnHeaders.Add , , "Afecta disponible", 1000
vLvw.ColumnHeaders.Add , , "Aplica Int.", 1000
vLvw.ColumnHeaders.Add , , "Fecha Tesorería", 1800
vLvw.ColumnHeaders.Add , , "Usuario Tesorería", 1800
vLvw.ColumnHeaders.Add , , "Solicitud Tesorería", 2000
vLvw.ColumnHeaders.Add , , "Remesa Tesorería", 2000
vLvw.ColumnHeaders.Add , , "Usuario Registro", 2000
vLvw.ColumnHeaders.Add , , "Fecha Registro", 2000


If ObjConsultar.fxTraerDesembolsos(pOperacion) Then
    With glogon.Recordset
        While Not glogon.Recordset.EOF
            vKey = "(VV)" & Trim(glogon.Recordset("CodigoDesembolso")) _
                   & "(Cd)" & Trim(glogon.Recordset("NumeroOperacion")) & "(Op)"
                   
            Set vItem = lsw.ListItems.Add(, vKey, Trim(.Fields!Beneficiario))
                vItem.SubItems(1) = Format(.Fields!Monto, "Standard")
                vItem.SubItems(2) = IIf(IsNull(.Fields!disponible), 0, Format(.Fields!disponible, "Standard"))
                vItem.SubItems(3) = Format(.Fields!InteresesFechaCorte, "dd/mm/yyyy")
                vItem.SubItems(4) = Format(.Fields!InteresesFechaCorteUltima, "dd/mm/yyyy")
                vItem.SubItems(5) = IIf(IsNull(.Fields!interesesActualDias), 0, Format(.Fields!interesesActualDias, "Standard"))
                vItem.SubItems(6) = IIf(IsNull(.Fields!InteresesActualMonto), 0, Format(.Fields!InteresesActualMonto, "Standard"))
                vItem.SubItems(7) = IIf(IsNull(.Fields!InteresesAcumDias), 0, Format(.Fields!InteresesAcumDias, "Standard"))
                vItem.SubItems(8) = IIf(IsNull(.Fields!InteresesAcumMonto), 0, Format(.Fields!InteresesAcumMonto, "Standard"))
                vItem.SubItems(9) = .Fields!DescAfectaDisponible
                vItem.SubItems(10) = .Fields!DescAplicaIntereses
                vItem.SubItems(11) = IIf(IsNull(.Fields!TesoreriaFecha), ObjNull.NullString, Format(.Fields!TesoreriaFecha, "dd/mm/yyyy"))
                vItem.SubItems(12) = IIf(IsNull(.Fields!TesoreriaUsuario), ObjNull.NullString, .Fields!TesoreriaUsuario)
                vItem.SubItems(13) = IIf(IsNull(.Fields!TesoreriaSolicitud), ObjNull.NullString, .Fields!TesoreriaSolicitud)
                vItem.SubItems(14) = IIf(IsNull(.Fields!TesoreriaRemesa), ObjNull.NullString, .Fields!TesoreriaRemesa)
                vItem.SubItems(15) = IIf(IsNull(.Fields!RegistroUsuario), ObjNull.NullString, .Fields!RegistroUsuario)
                vItem.SubItems(16) = IIf(IsNull(.Fields!RegistroFecha), ObjNull.NullString, Format(.Fields!RegistroFecha, "dd/mm/yyyy"))
                vMontoTotal = vMontoTotal + .Fields!Monto
                
                vItem.Tag = !CodigoDesembolso
                
                .MoveNext
        Wend
    End With
End If
txtTotalMontoGirado.Text = Format(vMontoTotal, "Standard")

 

salir:
  Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Resume salir
End Sub

Public Sub sbListaDesembolsosPendientes(ByVal pOperacion As Long)

Dim vItem As ListViewItem
Dim vLvw As XtremeSuiteControls.ListView
Dim vIdZona As Integer
Dim vMontoTotal As Double
Dim vKey As String

On Error GoTo error

m_IdDesembolso = -1
Me.lswPendientes.ColumnHeaders.Clear
Me.lswPendientes.ListItems.Clear

Set vLvw = Me.lswPendientes
vLvw.ColumnHeaders.Add , , "Concepto Desembolso", 2000
vLvw.ColumnHeaders.Add , , "Tipo Contacto", 3000
vLvw.ColumnHeaders.Add , , "Beneficiario", 3000
vLvw.ColumnHeaders.Add , , "Monto", 1800, 1
vLvw.ColumnHeaders.Add , , "Cuenta Contable", 2100
vLvw.ColumnHeaders.Add , , "Aplica Intereses", 2000
vLvw.ColumnHeaders.Add , , "Usuario Registro", 2000
vLvw.ColumnHeaders.Add , , "Fecha Registro", 2000
vLvw.ColumnHeaders.Add , , "Identificación", 2000


If ObjConsultar.fxDesembolsoPendientes(pOperacion) Then
    With glogon.Recordset
        While Not glogon.Recordset.EOF
        
vKey = "(VV)" & Trim(glogon.Recordset("IdContacto")) _
                   & "(Ic)" & Trim(glogon.Recordset("IdGarantia")) _
                   & "(Ig)" & Trim(glogon.Recordset("Tipo")) _
                   & "(Tp)" & Trim(glogon.Recordset("linea")) & "(Li)" _
                   & pOperacion & "(Op)"
                   
            Set vItem = lswPendientes.ListItems.Add(, vKey, Trim(.Fields!Descripcion))
                vItem.SubItems(1) = Trim(.Fields!DesTipo)
                vItem.SubItems(2) = Trim(.Fields!Beneficiario)
                vItem.SubItems(3) = Format(.Fields!Monto, "Standard")
                vItem.SubItems(4) = Trim(.Fields!CodigoCuenta)
                vItem.SubItems(5) = Trim(.Fields!DescAplicaInt)
                vItem.SubItems(6) = .Fields!Usuario
                vItem.SubItems(7) = Format(.Fields!fecha, "dd/mm/yyyy")
                vItem.SubItems(8) = .Fields!Identificacion
                .MoveNext
        Wend
    End With
End If
salir:
  Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Resume salir
End Sub
Private Function fxDescLineaCredito(ByVal strCodigo As String) As String

On Error GoTo vError

glogon.strSQL = "select descripcion from catalogo where codigo = '" & Trim(strCodigo) & "'"

If execSql(glogon.strSQL, True) Then
    fxDescLineaCredito = IIf(IsNull(glogon.Recordset!Descripcion), "", glogon.Recordset!Descripcion)
Else
    MsgBox "No se encontró la descripción del código de la linea de crédito digitada. - " & strCodigo, vbCritical
End If

salir:
    Exit Function
vError:
    MsgBox "Ocurrió un error validar información digitada. " & "-" & Err.Description, vbExclamation
    Resume salir
    
End Function

Private Sub sbBusqueda(ByVal Control As String)

gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Control
  Case "txtOperacion", "txtLinea", "txtDesLineaCredito", "txtcedula", "txtnombre"
    Call sbMostraVentanBusqueda
    
  Case "TxtBeneficiario"
        gBusquedas.Resultado = Empty
        gBusquedas.Resultado2 = Empty
        gBusquedas.Consulta = "SELECT identificacion, Nombre FROM ViviendaContactos"
        gBusquedas.Orden = "Nombre"
        gBusquedas.Columna = "Nombre"
        gBusquedas.Filtro = " and TipoContacto = 'F'"
'        If ChkEmpresa.Value = 1 Then
'            gBusquedas.Filtro = " and TipoContacto = 'E'"
'        End If
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado2 = Empty Then Exit Sub

        txtBeneficiario.Text = gBusquedas.Resultado2
    
   Case "txtTipoDesmbolsoPend"
        gBusquedas.Resultado = Empty
        gBusquedas.Resultado2 = Empty
        gBusquedas.Consulta = "select Codigo,Descripcion from ViviendaTiposDesembolsos"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "descripcion"
        gBusquedas.Filtro = " and estado = 'A'"
        frmBusquedas.Show vbModal
        
        txtTipoDesmbolsoPend.Text = gBusquedas.Resultado
        
        If gBusquedas.Resultado2 <> Empty Then
            txtDescripDesembolsoPend.Text = gBusquedas.Resultado2
        End If

        txtBeneficiarioPend.SetFocus
        
    Case "txtIdContacto"
        gBusquedas.Resultado = Empty
        gBusquedas.Resultado2 = Empty
        gBusquedas.Consulta = "select IdContacto,Nombre from ViviendaContactos"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "Nombre"
        gBusquedas.Filtro = " and estado = 'A'"
        frmBusquedas.Show vbModal
        
        txtIdContacto.Text = gBusquedas.Resultado
        
        If gBusquedas.Resultado2 <> Empty Then
            txtNombreContacto.Text = gBusquedas.Resultado2
        End If

        txtIdGarantia.SetFocus
        
    Case "txtIdGarantia"
        
        gBusquedas.Resultado = Empty
        gBusquedas.Resultado2 = Empty
        gBusquedas.Consulta = "select IdGarantia,NumeroFinca from ViviendaGarantia"
        gBusquedas.Orden = "IdGarantia"
        gBusquedas.Columna = "NumeroFinca"
        gBusquedas.Filtro = " and NumeroOperacion = " & Trim(txtOperacion)
        frmBusquedas.Show vbModal
        
        txtIdGarantia.Text = gBusquedas.Resultado
        
        If gBusquedas.Resultado2 <> Empty Then
            txtNumeroFinca.Text = gBusquedas.Resultado2
        End If

        txtTipoDesmbolsoPend.SetFocus
    
    
End Select

End Sub

Private Sub sbLimpiaDatos()
m_IdDesembolso = -1

txtId.Text = ""
txtBeneficiario.Text = ""
txtDetalle.Text = ""

rbEnte.Item(0).Value = True

Call sbCargaCombos
Call rbEnte_Click(0)


End Sub

Private Sub sbCargavGridLocal(vGrid As Object, vGridMaxCol As Integer)
Dim rs As New ADODB.Recordset, i As Integer

Set rs = glogon.Recordset
vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

Do While Not rs.EOF
'If rs.Fields!Beneficiario & "" <> "" Then
'    txtBeneficiario.Text = rs.Fields!Beneficiario & ""
'    txtdetalle.Text = rs.Fields!Detalle & ""
'End If
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1).Value))
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
vGrid.MaxRows = vGrid.MaxRows - 1
rs.Close

End Sub

Function fxDescribeBanco(intCodigo As Integer) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select descripcion from bancos where id_banco = " & Trim(intCodigo), glogon.Conection, adOpenStatic
If Not rsX.EOF And Not rsX.BOF Then
 fxDescribeBanco = IIf(IsNull(rsX!Descripcion), "", rsX!Descripcion)
End If
rsX.Close
End Function

Function fxCodigoBanco(strDescripcion As String) As Long
Dim rsX As New ADODB.Recordset

rsX.Open "select id_banco from bancos where descripcion = '" & Trim(strDescripcion) & "'", glogon.Conection, adOpenStatic
If Not rsX.EOF And Not rsX.BOF Then
 fxCodigoBanco = IIf(IsNull(rsX!id_banco), 0, rsX!id_banco)
Else
 fxCodigoBanco = 0
End If
rsX.Close
End Function

Private Sub sbEncabezadoDesembolso(ByVal pCodigoDesembolso As Long)
If ObjConsultar.fxEncabezadoDesembolso(pCodigoDesembolso) Then

   With glogon.Recordset
    If .Fields!Beneficiario & "" <> "" Then
        txtBeneficiario.Text = .Fields!Beneficiario & ""
        txtDetalle.Text = .Fields!Detalle & ""
        
        txtNuevoDisponible.Text = Format(.Fields!disponible, "Standard")
        txtIntActuales.Text = Format(.Fields!InteresesActualMonto, "Standard")
        
        Call sbCboAsignaDato(cboCuenta, IIf(IsNull(.Fields!BancoCuenta), "", .Fields!BancoCuenta), True, IIf(IsNull(.Fields!BancoCuenta), "", .Fields!BancoCuenta))
        
        If Not IsNull(.Fields!BancoCodigo) Then
            vPaso = True
              Call sbCboAsignaDato(cboBanco, fxDescLineaCredito(.Fields!BancoCodigo), True, .Fields!BancoCodigo)
            vPaso = False
        End If
    End If
    
    If IsNull(.Fields!BancoCodigo) Then Exit Sub
     cboTipoDocumento.Text = fxTipoDocumento(IIf(IsNull(.Fields!BancoEmitir), "", .Fields!BancoEmitir))
    End With
    
End If

    
End Sub

Private Sub sbDetalleDesembolso(ByVal pCodigoDesembolso As Long)
If tcMain.SelectedItem = 1 Then
    vGrid.ColWidth(3) = 0
    vGrid.ColWidth(4) = 0
    vGrid.ColWidth(5) = 0
    
    If ObjConsultar.fxDetalleDesembolso(pCodigoDesembolso) Then
    
        Call sbCargavGridLocal(vGrid, 5)
        Call sbSumaLineas
        Call sbEncabezadoDesembolso(pCodigoDesembolso)
'        If glogon.Recordset.Fields!Beneficiario & "" <> "" Then
'            txtBeneficiario.Text = glogon.Recordset.Fields!Beneficiario & ""
'            txtdetalle.Text = glogon.Recordset.Fields!Detalle & ""
'            txtNuevoDisponible.Text  = glogon.Recordset.Fields!Disponible
'        End If
    End If
    
End If
End Sub
'*******************Eventos del Usuario**************************************

Private Sub btnCuentas_Click()
On Error GoTo vError

GLOBALES.gTag = txtId.Text
frmCC_Cuentas_Bancarias.Show vbModal

Call cboBanco_Click

Exit Sub

vError:

End Sub

Private Sub btnDesembolsoBoleta_Click()
 
On Error GoTo vError

If Not IsNumeric(txtDesembolsoId.Text) Then Exit Sub

Me.MousePointer = vbHourglass



With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes de Desembolsos"
      
    .Connect = glogon.ConectRPT
    .Destination = crptToWindow
    
    .Formulas(1) = "fxFecha =  '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "fxUsuario = '" & glogon.Usuario & "'"
    .Formulas(3) = "fxTitulo = 'Boleta de Desembolsos Registrados'"
    .Formulas(4) = "fxSubTitulo = 'Detalle del Desembolso'"
    .Formulas(5) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Credito_Hipotecario_Boleta_Desembolso.rpt")
    .SelectionFormula = "{vCrd_Hipotecario_Desembolso.NumeroOperacion} = " & txtOperacion.Text & " AND {vCrd_Hipotecario_Desembolso.CodigoDesembolso} = " & txtDesembolsoId.Text
    

    .Action = 1
        
    
End With
 
 Me.MousePointer = vbDefault
Exit Sub
salir:

    Me.MousePointer = vbDefault
    frmContenedor.Crt.Formulas(1) = ""
    frmContenedor.Crt.Formulas(2) = ""
    frmContenedor.Crt.Formulas(3) = ""
    frmContenedor.Crt.Formulas(4) = ""
    frmContenedor.Crt.Formulas(5) = ""
    frmContenedor.Crt.Formulas(6) = ""
    frmContenedor.Crt.Formulas(7) = ""
    frmContenedor.Crt.Formulas(8) = ""
    frmContenedor.Crt.Formulas(9) = ""
    frmContenedor.Crt.Formulas(10) = ""
    frmContenedor.Crt.SelectionFormula = ""
    
    Exit Sub
    
vError:
    MsgBox "Ocurrió un error al imprimir los reportes solicitados - " & Err.Description, vbError


End Sub

Private Sub cboBanco_Click()
m_cambioDatos = True

If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtId.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub


Private Sub cboTipoDocumento_Click()
m_cambioDatos = True
End Sub


Private Sub cmdActivar_Click()
    If ItemSeleccionado Is Nothing Then Exit Sub
    Call sbActivar
End Sub

Private Sub sbActivar()
Dim vDisponibleBruto As Double
Dim vIntAcomulado As Double
Dim vBruto As Double
Dim vBrutoEnPantalla As Double

On Error GoTo vError

If fxValidaEstadoCredito = False Then
    MsgBox "El crédito debe estar activo y formalizado para realizar desembolsos"
    Exit Sub
End If

Me.MousePointer = vbHourglass

ReDim gParametros(0 To 12)
Dim vTemp As String
If (MsgBox("¿Confirma que desea activar el desembolso seleccionado.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub

If ObjConsultar.fxDesembolsoCalculo(txtOperacion.Text) Then
    With glogon.Recordset
        vDisponibleBruto = IIf(IsNull(.Fields!Bruto), 0, .Fields!Bruto)
        vIntAcomulado = IIf(IsNull(.Fields!IntAcumulado), 0, .Fields!IntAcumulado)
        vBruto = vDisponibleBruto - vIntAcomulado
    End With
    glogon.Recordset.Close
End If

'TODO Validar con Pedro esta linea
vBrutoEnPantalla = CCur(txtDisponibleBruto.Text) - CCur(txtIntAcumulado.Text)

If Format(vBruto, "Standard") <> Format(vBrutoEnPantalla, "Standard") Then
    MsgBox "El disponible bruto ha cambiado, es posible que ya se realizó un movimiento antes de aplicar este desembolso, Verifique los datos he intente de nuevo"
    Exit Sub
End If

 

If Val(txtNuevoDisponiblePendientes.Text) < 0 Then
    MsgBox "El monto disponible es inválido, no es posible continuar con el proceso."
    Exit Sub
End If

 

gParametros(0) = fxDeCodePK(ItemSeleccionado.Key, 5, "(Ic)") '@Contacto
gParametros(1) = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)") '@Garantia
gParametros(2) = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Tp)") '@Tipo
gParametros(3) = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Li)") '@Linea
gParametros(4) = glogon.Usuario
gParametros(5) = IIf(Val(txtDescuento.Text) = 0, 0, CCur(txtDescuento.Text))
gParametros(6) = CCur(txtNuevoDisponiblePendientes.Text)
gParametros(7) = Format(m_FechaCorte, "yyyy/mm/dd") '@FechaCorte
gParametros(8) = Format(m_FechaUltimoCorte, "yyyy/mm/dd") '@IntFechaCorteUltima
gParametros(9) = m_DiasActuales '@InteresesActualDias
gParametros(10) = CCur(txtIntActualesPendientes.Text) '@InteresesActualMonto
gParametros(11) = m_DiasAcum  '@InteresesAcumDias
gParametros(12) = CCur(m_IntAcomuladoPendientes)  '@InteresesAcumMonto

'hacer la rutina de actiavar en el objeto activar
If ObjAgregar.fxActivarDesembolsoPendiente(gParametros(0), gParametros(1), gParametros(2), gParametros(3), gParametros(4), gParametros(5), gParametros(6), gParametros(7), _
                                            gParametros(8), gParametros(9), gParametros(10), gParametros(11), gParametros(12)) Then
    MsgBox "Información fue registrada corretamente.", vbInformation
    
    Call Bitacora("APLICA", "Activa desembolso pendiente vivienda Contacto: " & gParametros(0) & " Garantia: " & gParametros(1))
    
    Call sbTraerInformacionOperacion
    Call sbListaDesembolsosPendientes(txtOperacion.Text)
    txtMontoP.Text = Format(0, "Standard")
     
    txtDescuento.Text = Format(0, "Standard")
    Set ItemSeleccionado = Nothing
End If


salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Resume
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbAgregarPendiente()
Dim vDisponibleBruto As Double
Dim vIntAcomulado As Double
Dim vBruto As Double
Dim vBrutoEnPantalla As Double

On Error GoTo vError

If fxValidaEstadoCredito = False Then
    MsgBox "El crédito debe estar activo y formalizado para realizar desembolsos"
    Exit Sub
End If

Me.MousePointer = vbHourglass

ReDim gParametros(0 To 8)
Dim vTemp As String
If (MsgBox("¿Confirma que desea agregar el desembolso.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub

If ObjConsultar.fxDesembolsoCalculo(txtOperacion.Text) Then
    With glogon.Recordset
        vDisponibleBruto = IIf(IsNull(.Fields!Bruto), 0, .Fields!Bruto)
        vIntAcomulado = IIf(IsNull(.Fields!IntAcumulado), 0, .Fields!IntAcumulado)
        vBruto = vDisponibleBruto - vIntAcomulado
    End With
    glogon.Recordset.Close
End If

'TODO Validar con Pedro esta linea
vBrutoEnPantalla = CCur(txtDisponibleBruto.Text) - CCur(txtIntAcumulado.Text)

If Format(vBruto, "Standard") <> Format(vBrutoEnPantalla, "Standard") Then
    MsgBox "El disponible bruto ha cambiado, es posible que ya se realizó un movimiento antes de aplicar este desembolso, Verifique los datos he intente de nuevo"
    Me.MousePointer = vbDefault
    Exit Sub
End If

If Val(txtNuevoDisponibleAgregarPend.Text) < 0 Then
    MsgBox "El monto disponible es inválido, no es posible continuar con el proceso."
    Me.MousePointer = vbDefault
    Exit Sub
End If

If Len(txtIdContacto.Text) = 0 Then
    MsgBox "Debe elegir un contacto, no es posible continuar con el proceso."
    Me.MousePointer = vbDefault
    txtIdContacto.SetFocus
    Exit Sub
End If
If Len(txtIdGarantia.Text) = 0 Then
    MsgBox "Debe elegir una garantía, no es posible continuar con el proceso."
    Me.MousePointer = vbDefault
    txtIdGarantia.SetFocus
    Exit Sub
End If
If Len(txtTipoDesmbolsoPend.Text) = 0 Then
    MsgBox "Debe elegir un tipo de garantía, no es posible continuar con el proceso."
    Me.MousePointer = vbDefault
    txtTipoDesmbolsoPend.SetFocus
    Exit Sub
End If
If Len(txtBeneficiarioPend.Text) = 0 Then
    MsgBox "Debe completar el campo beneficiario, no es posible continuar con el proceso."
    Me.MousePointer = vbDefault
    txtBeneficiarioPend.SetFocus
    Exit Sub
End If
If Len(txtMontoDesembolsoPend.Text) = 0 Or txtMontoDesembolsoPend < 1 Then
    MsgBox "Debe completar monto del desembolso, no es posible continuar con el proceso."
    Me.MousePointer = vbDefault
    txtMontoDesembolsoPend.SetFocus
    Exit Sub
End If
If txtNuevoDisponibleAgregarPend.Text < 0 Then
    MsgBox "El nuevo disponible no puede ser negativo, no es posible continuar con el proceso."
    Me.MousePointer = vbDefault
    Exit Sub
End If

gParametros(0) = Trim(txtIdContacto) '@Contacto
gParametros(1) = Trim(txtIdGarantia) '@Garantia
gParametros(2) = Trim(txtTipoDesmbolsoPend) '@TipoDesembolso
gParametros(3) = Trim(txtBeneficiarioPend) '@Beneficiario
gParametros(4) = txtMontoDesembolsoPend.Text  '@Monto
gParametros(5) = chkAplicaIntereses.Value ' @AplicaIntereses
gParametros(6) = glogon.Usuario '@Usuario
gParametros(7) = Format(fxFechaServidor, "yyyymmdd hh:mm:ss") '@Fecha
gParametros(8) = txtNuevoDisponibleAgregarPend.Text  '@Disponible

'hacer la rutina de actiavar en el objeto activar
If ObjAgregar.fxAgregarDesembolsoPendiente(gParametros(0), gParametros(1), gParametros(2), gParametros(3), _
                                           gParametros(4), gParametros(5), gParametros(6), gParametros(7), _
                                           gParametros(8)) Then
    MsgBox "Información fue registrada corretamente.", vbInformation
    
    Call Bitacora("REGISTRA", "Desembolso pendiente vivienda Contacto: " & gParametros(0) & " Garantia: " & gParametros(1))
    
    Call sbTraerInformacionOperacion
    Call sbListaDesembolsosPendientes(txtOperacion.Text)
    txtMontoP.Text = Format(0, "Standard")
    txtDescuento.Text = Format(0, "Standard")
    Set ItemSeleccionado = Nothing
    
    Call sbHabilitaTab(3)
End If


salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub cmdAgregarPendientes_Click()
    Call sbAgregarPendiente
End Sub

Private Sub cmdNuevoDesembolsoPendiente_Click()
    
    Call sbLimpiarTXTPendientes
    Call sbHabilitaTab(4)
    txtIdContacto.SetFocus
    
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo vError


Select Case Me.ActiveControl.Name
Case "txtDescuento"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtDescuento.Text), KeyAscii)
End Select

Exit Sub

vError:
    MsgBox "Ocurrió un error validar la información de los formatos. " & "-" & Err.Description, vbExclamation

End Sub

Function fxTipoDocumento(vTipo As String) As String
Select Case vTipo
 Case "CK"
   fxTipoDocumento = "Cheque"
 Case "TE"
   fxTipoDocumento = "Transferencia"
 Case "EF", "RE"
   fxTipoDocumento = "Efectivo"
 Case "ND"
   fxTipoDocumento = "Nota Debito"
 Case "NC"
   fxTipoDocumento = "Nota Credito"
 Case "OT"
   fxTipoDocumento = "Otro..."
'-------
 Case "Cheque"
   fxTipoDocumento = "CK"
 Case "Transferencia"
   fxTipoDocumento = "TE"
 Case "Efectivo"
   fxTipoDocumento = "EF"
 Case "Nota Debito"
   fxTipoDocumento = "ND"
 Case "Nota Credito"
   fxTipoDocumento = "NC"
 Case "Otro..."
   fxTipoDocumento = "OT"
 Case Else
   fxTipoDocumento = ""
End Select
End Function


Private Function fxValidaEstadoCredito() As Boolean
Dim strSQL As String
Dim rs As New ADODB.Recordset
    
    fxValidaEstadoCredito = False

    strSQL = "select isnull(ESTADO,'') as ESTADO from REG_CREDITOS WHERE ID_SOLICITUD = " & Trim(txtOperacion)
    Call OpenRecordSet(rs, strSQL)
        
    If Not rs.EOF Then
        If rs.Fields(0) = "A" Or rs.Fields(0) = "C" Then
            fxValidaEstadoCredito = True
        End If
    End If

End Function

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.ActiveControl.Name = "txtdetalle" Then Exit Sub
If Me.ActiveControl.Name = "vGrid" Then Exit Sub

If (KeyCode = vbKeyReturn) Then 'Or KeyCode = vbKeyTab) Then
    Call gsbPulsarTecla(vbKeyTab)
ElseIf KeyCode = vbKeyF4 Then
      Call sbBusqueda(Me.ActiveControl.Name)
    End If

End Sub

Private Function fxBancoAsignado(vBanco As Integer, vUsuario As String) As Boolean
Dim strSQL As String
Dim rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from tes_banco_asg where id_banco = " _
       & vBanco & " and nombre = '" & vUsuario & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then
  fxBancoAsignado = False
Else
  fxBancoAsignado = True
End If

rs.Close

End Function

Function fxCuentaAhorros(strCedula As String, lngID_Banco As Long) As String

Dim strSQL As String
Dim rsX As New ADODB.Recordset

strSQL = "select cuenta from cuentas_ahorros where cedula = '" & strCedula & "' and id_banco=" _
       & lngID_Banco
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic

If rsX.EOF And rsX.BOF Then
  fxCuentaAhorros = ""
Else
  fxCuentaAhorros = IIf(IsNull(rsX!Cuenta), "", rsX!Cuenta)
End If
rsX.Close

End Function


Private Sub sbCargaCombos()
Dim strSQL As String

cboCuenta.Clear

vPaso = True

    strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
    Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

vPaso = False

cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.Text = fxTipoDocumento("TE")

End Sub

Private Sub Form_Load()

vGrid.AppearanceStyle = fxGridStyle

'' Carga nombre de la ternimal
If Len(glogon.Maquina) = 0 Then
    Call sbMaquina
End If

vModulo = 3 'Modulo de Credito
'Inicializa Barra
Call sbToolBarIconos(tlbPrincipal, False)
Call sbToolBar(tlbPrincipal, "nuevo")
'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)

gIconoLista = "Desembolsos" 'Carga el icono a mostrar el la lista detalle
m_IdDesembolso = -1

vPaso = False

tcMain.Item(0).Selected = True

If gOperacion = 0 Then Exit Sub

txtOperacion.Text = gOperacion

Call sbTraerInformacionOperacion

Call sbCargaCombos
Call sbHabilitaTab(1)

End Sub









Private Sub txtMontoP_Change()
    Call sbInteresesDesembolsosPendientes
End Sub

Private Sub sbInteresesDesembolsosPendientes()

    m_IntAcomuladoPendientes = CCur(txtIntAcumulado.Text)
    
    '' Calcula Intereses Actuales si aplica intereses
    If m_AplicaInteresesDesembolsosP = 1 Then
        txtIntActualesPendientes.Text = (CCur(txtMontoP.Text) - CCur(txtDescuento.Text)) * m_TasaDiaria * m_DiasActuales
    Else
        txtIntActualesPendientes.Text = 0
    End If
    
    
     
    txtMontoGirar.Text = (CCur(txtMontoP.Text) - CCur(txtDescuento) + CCur(txtIntActualesPendientes))
    '' Calcula Nuevo Disponible
    If m_AplicaInteresesDesembolsosP = 1 Then
        txtNuevoDisponiblePendientes.Text = txtDisponibleBruto.Text - (CCur(txtIntActualesPendientes.Text) + CCur(m_IntAcomuladoPendientes))
    Else
        txtNuevoDisponiblePendientes.Text = txtDisponibleBruto.Text
    End If
    
    '' Control si el disponible es menor que los intereses
    If txtNuevoDisponiblePendientes.Text < 0 Then
    
        If CCur(txtIntAcumulado.Text) > CCur(txtDisponibleBruto.Text) Then
            m_IntAcomuladoPendientes = CCur(txtDisponibleBruto.Text)
            txtIntActualesPendientes.Text = 0
        Else
            txtIntActualesPendientes.Text = CCur(txtDisponibleBruto.Text) - m_IntAcomuladoPendientes
        End If
        
        '' Recalcula Nuevo Monto a girar
        txtMontoGirar.Text = (CCur(txtMontoP.Text) - CCur(txtDescuento.Text) + CCur(txtIntActualesPendientes.Text))
        
        '' Calcula Nuevo Disponible con intereses modificados
        If m_AplicaInteresesDesembolsosP = 1 Then
            txtNuevoDisponiblePendientes.Text = txtDisponibleBruto.Text - (CCur(txtIntActualesPendientes.Text) + CCur(m_IntAcomuladoPendientes))
        Else
            txtNuevoDisponiblePendientes.Text = txtDisponibleBruto.Text
        End If
    
    End If
    
    
    txtIntActualesPendientes.Text = Format(txtIntActualesPendientes.Text, "Standard")
    txtMontoGirar.Text = Format(txtMontoGirar.Text, "Standard")
    txtNuevoDisponiblePendientes.Text = Format(txtNuevoDisponiblePendientes.Text, "Standard")
    
End Sub

Private Sub lsw_DblClick()

'Dim vTemp As String
'If lsw.ListItems.Count > 0 Then
'    Call sbHabilitaTab(5)
'    m_IdDesembolso = fxDeCodePK(ItemSeleccionado.Key, 5, "(Cd)")
'    SSTabGeneral.Tab = 1
'End If

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Set ItemSeleccionado = Item
m_IdDesembolso = fxDeCodePK(ItemSeleccionado.Key, 5, "(Cd)")

txtDesembolsoId.Text = m_IdDesembolso
End Sub

Private Sub lswPendientes_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtMontoP.Text = Format(0, "Standard")

If Item.SubItems(5) = "SI" Then
    m_AplicaInteresesDesembolsosP = 1
Else
    m_AplicaInteresesDesembolsosP = 0
End If

txtMontoP.Text = Item.SubItems(3)
txtDescuento.Text = Format(0, "Standard")

Set ItemSeleccionado = Item


txtContactoNombre.Text = Item.SubItems(2)
txtContactoId.Text = fxDeCodePK(Item.Key, 5, "(Ic)")  '@Contacto
txtContactoTipo.Text = Item.SubItems(1)  '@Tipo
txtContactoCedula.Text = Item.SubItems(8)

'Call txtDescuento_Change

End Sub

Private Sub rbEnte_Click(Index As Integer)

txtId.Text = ""
txtBeneficiario.Text = ""

txtId.Locked = True
txtBeneficiario.Locked = True

Select Case Index
    Case 0 'Deudor
        txtId.Text = txtCedula.Text
        txtBeneficiario.Text = txtNombre.Text
    Case 1 'Profesional
        gBusquedas.Col1Name = "Contacto Id"
        gBusquedas.Col2Name = "Identificación"
        gBusquedas.Col3Name = "Nombre"
        gBusquedas.Consulta = "select idContacto, Identificacion, Nombre, case " _
                            & " when TipoProfesional = 'A' then 'Abogado'" _
                            & " when TipoProfesional = 'I' then 'Ingeniero' else 'Otro' end as 'Tipo'" _
                            & " From ViviendaContactos "
        gBusquedas.Orden = "Nombre"
        gBusquedas.Columna = "Nombre"
        gBusquedas.Filtro = " and Estado = 'A'"
        frmBusquedas.Show vbModal
        
        If gBusquedas.Resultado <> "" Then
            txtId.Text = gBusquedas.Resultado2
            txtBeneficiario.Text = gBusquedas.Resultado3
        End If
        
    Case 2 'Otro
        txtId.Locked = False
        txtBeneficiario.Locked = False
        
        txtId.SetFocus
        
End Select

End Sub



Private Sub sbLimpiarTXTPendientes()

    txtIdContacto.Text = Empty
    txtNombreContacto.Text = Empty
    txtIdGarantia.Text = Empty
    txtNumeroFinca.Text = Empty
    txtTipoDesmbolsoPend.Text = Empty
    txtDescripDesembolsoPend.Text = Empty
    txtBeneficiarioPend.Text = Empty
    txtMontoDesembolsoPend.Text = Empty
    chkAplicaIntereses.Value = Unchecked
    txtNuevoDisponibleAgregarPend.Text = Empty

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)


lblMensaje.Caption = ""

If Not IsNumeric(txtOperacion.Text) Then
   Exit Sub
End If

If Not ObjConsultar.fxPermiteDesembolso(txtOperacion.Text, Item.Index) And Item.Index > 0 Then
    lblMensaje.Caption = "No es posible realizar el desembolso, la operación no esta formalizada"
End If

Select Case Item.Index
    Case 1
        Call sbDetalleDesembolso(m_IdDesembolso)
    Case 2
        txtDescuento.Text = Format(0, "Standard")
        txtMontoP.Text = Format(0, "Standard")
        
        Set ItemSeleccionado = Nothing
        txtBeneficiario.Text = Trim(txtNombre.Text)
        Call sbListaDesembolsosPendientes(txtOperacion.Text)
        Call sbHabilitaTab(3)
    
End Select

End Sub

Private Sub txtBeneficiario_Change()
m_cambioDatos = True
End Sub

Private Sub txtdetalle_Change()
m_cambioDatos = True
End Sub

Private Sub txtIdContacto_Change()
    If Val(txtIdContacto.Text) = 0 Then
        txtIdContacto.Text = Empty
    End If
End Sub

Private Sub txtIdContacto_LostFocus()

    If Len(txtIdContacto.Text) > 0 Then
    
        If ObjConsultar.fxTraerInfoContacto(txtIdContacto.Text) Then
            txtNombreContacto.Text = IIf(IsNull(glogon.Recordset.Fields!Nombre), Empty, glogon.Recordset.Fields!Nombre)
        Else
            txtIdContacto.Text = Empty
            txtNombreContacto.Text = Empty
        End If
        glogon.Recordset.Close
    
    Else
    
        txtNombreContacto.Text = Empty
        
    End If
End Sub



Private Sub txtIdGarantia_Change()
    If Val(txtIdGarantia.Text) = 0 Then
        txtIdGarantia.Text = Empty
    End If
End Sub

Private Sub txtIdGarantia_LostFocus()
    If Len(txtIdGarantia.Text) > 0 Then
    
        If ObjConsultar.fxTraerInfoGarantia(txtIdGarantia.Text, txtOperacion.Text) Then
            txtNumeroFinca.Text = IIf(IsNull(glogon.Recordset.Fields!NumeroFinca), Empty, glogon.Recordset.Fields!NumeroFinca)
        Else
            txtIdGarantia.Text = Empty
            txtNumeroFinca.Text = Empty
        End If
        glogon.Recordset.Close
    Else
        txtNumeroFinca.Text = Empty
    End If
End Sub

Private Sub txtMontoDesembolsoPend_Change()
    If Val(txtMontoDesembolsoPend.Text) = 0 Then
        txtMontoDesembolsoPend.Text = 0
    End If
    
    txtNuevoDisponibleAgregarPend.Text = Format(CCur(txtDisponibleBruto.Text) - CCur(txtMontoDesembolsoPend.Text), "Standard")
    
End Sub

Private Sub txtMontoDesembolsoPend_GotFocus()
    txtMontoDesembolsoPend.SelStart = 0
    txtMontoDesembolsoPend.SelLength = Len(txtMontoDesembolsoPend.Text)
End Sub

Private Sub txtMontoDesembolsoPend_LostFocus()
    txtMontoDesembolsoPend.Text = Format(txtMontoDesembolsoPend, "Standard")
End Sub

Private Sub txtTipoDesmbolsoPend_LostFocus()
    If Len(txtTipoDesmbolsoPend.Text) > 0 Then
    
        If ObjConsultar.fxTraerTipoDesembolso(txtTipoDesmbolsoPend.Text) Then
            With glogon.Recordset
                txtDescripDesembolsoPend.Text = IIf(IsNull(.Fields!Descripcion), 0, .Fields!Descripcion)
                
                If IIf(IsNull(.Fields!AplicaInteres), 0, .Fields!AplicaInteres) = True Then
                    chkAplicaIntereses.Value = Checked
                Else
                    chkAplicaIntereses.Value = Unchecked
                End If
                
            End With
        Else
            txtTipoDesmbolsoPend.Text = Empty
            txtDescripDesembolsoPend.Text = Empty
            chkAplicaIntereses.Value = Unchecked
        End If
        glogon.Recordset.Close
    Else
        txtDescripDesembolsoPend.Text = Empty
        chkAplicaIntereses.Value = Unchecked
    End If
End Sub

Private Sub vGrid_EditChange(ByVal Col As Long, ByVal Row As Long)
Call sbSumaLineas
End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If (NewCol = -1) Then Exit Sub
    If (m_cambioDatos = True) Then
        If (Row = NewRow) Then
        Exit Sub
    End If
End If
End Sub
Private Sub vGrid_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    m_cambioDatos = True
     
End Sub
Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
    Select Case Button.Key
        Case "nuevo"
            Call sbHabilitaTab(2)
            Call sbLimpiaDatos
            Call txtOperacion_LostFocus
            Call sbDetalleDesembolso(-1)
            Call sbToolBar(tlbPrincipal, "Edicion")
            txtBeneficiario.Text = txtNombre.Text
            Call cboBanco_Click
            
        Case "editar"
            If Not ItemSeleccionado Is Nothing Then
                Call sbHabilitaTab(5)
            End If
            
        Case "guardar"
            Call sbAgregar
            
        Case "deshacer"
            Call sbHabilitaTab(1)
            Call sbToolBar(tlbPrincipal, "Nuevo")
       Case "reportes"
        Call sbImprimir
    End Select
    

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

Private Sub txtDescuento_Change()

If Val(txtDescuento.Text) = 0 Then
    txtDescuento.Text = CCur(0)
End If
If Val(txtMontoP.Text) = 0 Then
    txtMontoP.Text = CCur(0)
End If

txtMontoGirar.Text = (CCur(txtMontoP.Text) - CCur(txtDescuento.Text))

'If Val(txtmontogirar.Text) = 0 Then Exit Sub
txtMontoGirar.Text = Format(txtMontoGirar.Text, "Standard")

If CCur(txtDescuento.Text) > CCur(txtMontoP.Text) Then
'        MsgBox "El descuento no puede ser mayor al monto"
    cmdActivar.Enabled = False
    txtMontoGirar.Text = Format(0, "Standard")
Else
    cmdActivar.Enabled = True
End If

Call sbInteresesDesembolsosPendientes

End Sub

Private Sub txtDescuento_GotFocus()

txtDescuento.SelStart = 0
txtDescuento.SelLength = Len(txtDescuento.Text)
txtDescuento.Text = CCur(txtDescuento.Text)

End Sub

Private Sub txtDescuento_LostFocus()

If Val(txtDescuento.Text) = 0 Then Exit Sub
    txtDescuento.Text = Format(txtDescuento.Text, "Standard")
    
End Sub

Private Sub txtLinea_Change()
txtDesLineaCredito.Text = Empty
End Sub

Private Sub txtLinea_LostFocus()

If Len(txtLinea.Text) = 0 Then Exit Sub
    txtDesLineaCredito.Text = fxDescLineaCredito(Trim(txtLinea.Text))
    
End Sub


Private Sub txtOperacion_LostFocus()
If Len(txtOperacion.Text) = 0 Then Exit Sub
Call sbTraerInformacionOperacion
End Sub
