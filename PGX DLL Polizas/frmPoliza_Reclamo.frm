VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmPoliza_Reclamo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reclamo de Poliza"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcAux 
      Height          =   3735
      Left            =   0
      TabIndex        =   44
      Top             =   4560
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   6588
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
      ItemCount       =   6
      Item(0).Caption =   "Ingreso"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "Label2(0)"
      Item(0).Control(1)=   "txtObservaciones"
      Item(0).Control(2)=   "Label2(1)"
      Item(0).Control(3)=   "Label2(2)"
      Item(0).Control(4)=   "txtR_Usuario"
      Item(0).Control(5)=   "txtR_Fecha"
      Item(1).Caption =   "Recepción"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "Label2(7)"
      Item(1).Control(1)=   "txtRec_Observacion"
      Item(1).Control(2)=   "Label2(8)"
      Item(1).Control(3)=   "txtRec_Usuario"
      Item(1).Control(4)=   "Label2(9)"
      Item(1).Control(5)=   "dtpRec_Fecha"
      Item(1).Control(6)=   "btnRecepcion"
      Item(2).Caption =   "Seguimiento"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "tcSeguimiento"
      Item(3).Caption =   "Fondo"
      Item(3).ControlCount=   12
      Item(3).Control(0)=   "Label2(10)"
      Item(3).Control(1)=   "txtF_MontoOperacion"
      Item(3).Control(2)=   "Label2(11)"
      Item(3).Control(3)=   "Label2(12)"
      Item(3).Control(4)=   "txtF_Monto"
      Item(3).Control(5)=   "txtF_Aprobado"
      Item(3).Control(6)=   "btnFondo(0)"
      Item(3).Control(7)=   "lswFondo"
      Item(3).Control(8)=   "Label2(13)"
      Item(3).Control(9)=   "txtF_Sobrante"
      Item(3).Control(10)=   "btnFondo(1)"
      Item(3).Control(11)=   "btnFondo(2)"
      Item(4).Caption =   "Desembolsos"
      Item(4).ControlCount=   11
      Item(4).Control(0)=   "Label2(14)"
      Item(4).Control(1)=   "txtD_Fondo"
      Item(4).Control(2)=   "Label2(15)"
      Item(4).Control(3)=   "txtD_Monto"
      Item(4).Control(4)=   "cboBanco"
      Item(4).Control(5)=   "cboCuenta"
      Item(4).Control(6)=   "Label2(16)"
      Item(4).Control(7)=   "Label2(17)"
      Item(4).Control(8)=   "btnDesembolso(0)"
      Item(4).Control(9)=   "btnDesembolso(1)"
      Item(4).Control(10)=   "lswDesembolsos"
      Item(5).Caption =   "Etiquetas"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "tcEtiqueta"
      Begin XtremeSuiteControls.ListView lswDesembolsos 
         Height          =   1935
         Left            =   -69880
         TabIndex        =   92
         Top             =   1680
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1441793
         _ExtentX        =   20770
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswFondo 
         Height          =   1695
         Left            =   -69880
         TabIndex        =   76
         Top             =   1560
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1441793
         _ExtentX        =   20770
         _ExtentY        =   2990
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtD_Monto 
         Height          =   330
         Left            =   -63280
         TabIndex        =   84
         Top             =   1200
         Visible         =   0   'False
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRec_Usuario 
         Height          =   330
         Left            =   -67840
         TabIndex        =   65
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.TabControl tcSeguimiento 
         Height          =   3255
         Left            =   -70000
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1441793
         _ExtentX        =   21193
         _ExtentY        =   5741
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
         Item(0).Caption =   "Estado"
         Item(0).ControlCount=   9
         Item(0).Control(0)=   "Label2(3)"
         Item(0).Control(1)=   "cboS_Estado"
         Item(0).Control(2)=   "Label2(4)"
         Item(0).Control(3)=   "txtS_Observacion"
         Item(0).Control(4)=   "btnSeguimiento"
         Item(0).Control(5)=   "Label2(5)"
         Item(0).Control(6)=   "txtS_Destinatarios"
         Item(0).Control(7)=   "chkS_Correo"
         Item(0).Control(8)=   "Label2(6)"
         Item(1).Caption =   "Histórico"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswSeguimiento"
         Begin XtremeSuiteControls.ListView lswSeguimiento 
            Height          =   2895
            Left            =   -69880
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   11775
            _Version        =   1441793
            _ExtentX        =   20770
            _ExtentY        =   5106
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkS_Correo 
            Height          =   255
            Left            =   10200
            TabIndex        =   60
            Top             =   2760
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Enviar Correo"
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
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ComboBox cboS_Estado 
            Height          =   330
            Left            =   2040
            TabIndex        =   54
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
         Begin XtremeSuiteControls.FlatEdit txtS_Observacion 
            Height          =   1215
            Left            =   2040
            TabIndex        =   56
            Top             =   960
            Width           =   9735
            _Version        =   1441793
            _ExtentX        =   17171
            _ExtentY        =   2143
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
         Begin XtremeSuiteControls.PushButton btnSeguimiento 
            Height          =   375
            Left            =   4440
            TabIndex        =   57
            Top             =   480
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   661
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
            Appearance      =   17
            Picture         =   "frmPoliza_Reclamo.frx":0000
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtS_Destinatarios 
            Height          =   330
            Left            =   2040
            TabIndex        =   59
            Top             =   2280
            Width           =   9735
            _Version        =   1441793
            _ExtentX        =   17171
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   6
            Left            =   2040
            TabIndex        =   61
            Top             =   2760
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Para ingresar más de un destinatario dene separarlos con punto y coma"
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
            Alignment       =   4
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   58
            Top             =   2280
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Destinatarios:"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   55
            Top             =   960
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaciones:"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   53
            Top             =   480
            Width           =   1455
            _Version        =   1441793
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
            Alignment       =   1
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   1815
         Left            =   1800
         TabIndex        =   46
         Top             =   600
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtR_Usuario 
         Height          =   330
         Left            =   3600
         TabIndex        =   49
         Top             =   2640
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.FlatEdit txtR_Fecha 
         Height          =   330
         Left            =   8520
         TabIndex        =   50
         Top             =   2640
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.FlatEdit txtRec_Observacion 
         Height          =   1815
         Left            =   -67840
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   3201
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
      Begin XtremeSuiteControls.DateTimePicker dtpRec_Fecha 
         Height          =   315
         Left            =   -67840
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.PushButton btnRecepcion 
         Height          =   375
         Left            =   -67840
         TabIndex        =   68
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Recibir"
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
         Picture         =   "frmPoliza_Reclamo.frx":0727
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtF_MontoOperacion 
         Height          =   330
         Left            =   -67480
         TabIndex        =   70
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtF_Monto 
         Height          =   330
         Left            =   -67480
         TabIndex        =   72
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtF_Aprobado 
         Height          =   330
         Left            =   -67480
         TabIndex        =   74
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnFondo 
         Height          =   330
         Index           =   0
         Left            =   -65200
         TabIndex        =   75
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Generar Fondo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmPoliza_Reclamo.frx":0E4E
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtF_Sobrante 
         Height          =   330
         Left            =   -60280
         TabIndex        =   78
         Top             =   3360
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnFondo 
         Height          =   330
         Index           =   1
         Left            =   -65200
         TabIndex        =   79
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Realizar Aporte"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmPoliza_Reclamo.frx":1284
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnFondo 
         Height          =   330
         Index           =   2
         Left            =   -58600
         TabIndex        =   80
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   582
         _StockProps     =   79
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
         Picture         =   "frmPoliza_Reclamo.frx":19A4
      End
      Begin XtremeSuiteControls.FlatEdit txtD_Fondo 
         Height          =   330
         Left            =   -63280
         TabIndex        =   82
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   330
         Left            =   -69160
         TabIndex        =   85
         Top             =   600
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
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
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   330
         Left            =   -69160
         TabIndex        =   86
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
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
      Begin XtremeSuiteControls.PushButton btnDesembolso 
         Height          =   375
         Index           =   0
         Left            =   -60760
         TabIndex        =   90
         Top             =   1150
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Appearance      =   17
         Picture         =   "frmPoliza_Reclamo.frx":20A4
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnDesembolso 
         Height          =   375
         Index           =   1
         Left            =   -59440
         TabIndex        =   91
         Top             =   1155
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "frmPoliza_Reclamo.frx":27CB
      End
      Begin XtremeSuiteControls.TabControl tcEtiqueta 
         Height          =   3255
         Left            =   -70000
         TabIndex        =   93
         Top             =   360
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1441793
         _ExtentX        =   21193
         _ExtentY        =   5741
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
         Item(0).Caption =   "Estado"
         Item(0).ControlCount=   10
         Item(0).Control(0)=   "CheckBox1"
         Item(0).Control(1)=   "ListView1"
         Item(0).Control(2)=   "cboE_Tag"
         Item(0).Control(3)=   "FlatEdit2"
         Item(0).Control(4)=   "btnEtiqueta"
         Item(0).Control(5)=   "FlatEdit3"
         Item(0).Control(6)=   "Label2(18)"
         Item(0).Control(7)=   "Label2(19)"
         Item(0).Control(8)=   "Label2(20)"
         Item(0).Control(9)=   "Label2(21)"
         Item(1).Caption =   "Histórico"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswEtiquetas"
         Begin XtremeSuiteControls.ListView lswEtiquetas 
            Height          =   2895
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   11775
            _Version        =   1441793
            _ExtentX        =   20770
            _ExtentY        =   5106
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView ListView1 
            Height          =   2895
            Left            =   -1.39880e5
            TabIndex        =   95
            Top             =   360
            Visible         =   0   'False
            Width           =   11775
            _Version        =   1441793
            _ExtentX        =   20770
            _ExtentY        =   5106
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   255
            Left            =   -59800
            TabIndex        =   94
            Top             =   2760
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Enviar Correo"
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
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ComboBox cboE_Tag 
            Height          =   330
            Left            =   -67960
            TabIndex        =   96
            Top             =   480
            Visible         =   0   'False
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit2 
            Height          =   1215
            Left            =   -67960
            TabIndex        =   97
            Top             =   960
            Visible         =   0   'False
            Width           =   9735
            _Version        =   1441793
            _ExtentX        =   17171
            _ExtentY        =   2143
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
         Begin XtremeSuiteControls.PushButton btnEtiqueta 
            Height          =   375
            Left            =   -65560
            TabIndex        =   98
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   661
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
            Appearance      =   17
            Picture         =   "frmPoliza_Reclamo.frx":2ECB
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit3 
            Height          =   330
            Left            =   -67960
            TabIndex        =   99
            Top             =   2280
            Visible         =   0   'False
            Width           =   9735
            _Version        =   1441793
            _ExtentX        =   17171
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   21
            Left            =   -69640
            TabIndex        =   87
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Etiqueta:"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   20
            Left            =   -69640
            TabIndex        =   102
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaciones:"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   19
            Left            =   -69640
            TabIndex        =   101
            Top             =   2280
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Destinatarios:"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   18
            Left            =   -67960
            TabIndex        =   100
            Top             =   2760
            Visible         =   0   'False
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Para ingresar más de un destinatario dene separarlos con punto y coma"
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
            Alignment       =   4
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   17
         Left            =   -69160
         TabIndex        =   89
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   16
         Left            =   -69160
         TabIndex        =   88
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Banco:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   15
         Left            =   -63280
         TabIndex        =   83
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto del desembolso:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   14
         Left            =   -63280
         TabIndex        =   81
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sobrante Fondo:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   -62440
         TabIndex        =   77
         Top             =   3360
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sobrante:"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   12
         Left            =   -69640
         TabIndex        =   73
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto aprobado:"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   11
         Left            =   -69640
         TabIndex        =   71
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fondo:"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   -69640
         TabIndex        =   69
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto de la Operación:"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   -69640
         TabIndex        =   66
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de Recepción:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   -69640
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario de Recepción:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   -69280
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaciones:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   48
         Top             =   2640
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de Registro:"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   47
         Top             =   2640
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario de Registro:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   45
         Top             =   600
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaciones:"
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   1335
      Left            =   0
      TabIndex        =   27
      Top             =   3120
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   2355
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
      Item(0).Caption =   "Reclamo de Póliza de Vida"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "gbPolizaVida"
      Item(1).Caption =   "Reclamo de Póliza de Incendio"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "GroupBox1"
      Begin XtremeSuiteControls.GroupBox gbPolizaVida 
         Height          =   1215
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   12015
         _Version        =   1441793
         _ExtentX        =   21193
         _ExtentY        =   2143
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnPolizaModificar 
            Height          =   375
            Index           =   0
            Left            =   10320
            TabIndex        =   35
            Top             =   120
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Modificar"
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
            Picture         =   "frmPoliza_Reclamo.frx":35F2
         End
         Begin XtremeSuiteControls.ComboBox cboPV_Motivo 
            Height          =   330
            Left            =   2160
            TabIndex        =   29
            Top             =   120
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.ComboBox cboPV_Enfermedad 
            Height          =   330
            Left            =   6720
            TabIndex        =   31
            Top             =   120
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.FlatEdit txtPV_Edad 
            Height          =   330
            Left            =   2160
            TabIndex        =   34
            Top             =   480
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Edad"
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
            Index           =   8
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enfermedad"
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
            Index           =   7
            Left            =   4680
            TabIndex        =   32
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   1815
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1215
         Left            =   -70000
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1441793
         _ExtentX        =   21193
         _ExtentY        =   2143
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnPolizaModificar 
            Height          =   375
            Index           =   1
            Left            =   10320
            TabIndex        =   37
            Top             =   120
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Modificar"
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
            Picture         =   "frmPoliza_Reclamo.frx":3BED
         End
         Begin XtremeSuiteControls.ComboBox cboPI_Tipo 
            Height          =   330
            Left            =   2160
            TabIndex        =   38
            Top             =   120
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.ComboBox cboPI_Causa 
            Height          =   330
            Left            =   6720
            TabIndex        =   39
            Top             =   120
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.FlatEdit txtPI_Finca 
            Height          =   330
            Left            =   2160
            TabIndex        =   40
            Top             =   480
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de siniestro"
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
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Causa"
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
            Index           =   10
            Left            =   4680
            TabIndex        =   42
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Finca"
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
            Index           =   9
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   1815
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox gbDatos 
      Height          =   1815
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   11895
      _Version        =   1441793
      _ExtentX        =   20981
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Datos del Reclamo"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboDesembolso 
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   1440
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
      Begin XtremeSuiteControls.DateTimePicker dtpFechaNacimiento 
         Height          =   315
         Left            =   6240
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   7920
         TabIndex        =   13
         Top             =   600
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   330
         Left            =   2160
         TabIndex        =   14
         Top             =   600
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   330
         Left            =   5040
         TabIndex        =   18
         Top             =   600
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   330
         Left            =   2160
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
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
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   330
         Left            =   9360
         TabIndex        =   22
         Top             =   1080
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.ComboBox ComboBox2 
         Height          =   330
         Left            =   6240
         TabIndex        =   25
         Top             =   1440
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo"
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
         Left            =   8160
         TabIndex        =   26
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "F. de nacimiento"
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
         Index           =   4
         Left            =   4440
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "M. de Pago"
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
         Index           =   3
         Left            =   4680
         TabIndex        =   23
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Desembolso"
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   360
         Index           =   0
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Apellido 1"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   360
         Index           =   2
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Apellido 2"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   360
         Index           =   3
         Left            =   7920
         TabIndex        =   15
         Top             =   240
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
         _ExtentY        =   635
         _StockProps     =   14
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
         Alignment       =   1
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Completo"
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
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   480
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   847
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
      Text            =   "00000"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   480
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   847
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
      Text            =   "00000"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPolizaCodigo 
      Height          =   480
      Left            =   6480
      TabIndex        =   4
      Top             =   360
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   847
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
   Begin XtremeSuiteControls.FlatEdit txtReclamoId 
      Height          =   480
      Left            =   9720
      TabIndex        =   6
      Top             =   360
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   847
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "00000"
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   9720
      TabIndex        =   7
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Reclamo"
      ForeColor       =   16777215
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo de Póliza"
      ForeColor       =   16777215
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Póliza"
      ForeColor       =   16777215
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Operación"
      ForeColor       =   16777215
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
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmPoliza_Reclamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
vModulo = 3
Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture


With lswFondo.ColumnHeaders
    .Clear
    .Add , , "Transacción", 2000
    .Add , , "T.Documento", 2000
    .Add , , "Referencia 1", 1500, vbCenter
    .Add , , "Referencia 2", 1500, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Usuario", 2500, vbCenter
End With

With lswDesembolsos.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Apl. Terceros", 2500, vbCenter
    .Add , , "Céd del Terceros", 1800, vbCenter
    .Add , , "Nombre del Terceros", 3800
End With

With lswEtiquetas.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Fecha", 2500
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Observaciones", 4800
End With

With lswSeguimiento.ColumnHeaders
    .Clear
    .Add , , "Fecha", 2500
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Estado", 1500, vbCenter
    .Add , , "Observaciones", 4800
End With

End Sub
