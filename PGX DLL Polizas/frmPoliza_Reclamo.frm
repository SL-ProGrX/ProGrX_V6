VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPoliza_Reclamo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reclamo de Poliza"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcAux 
      Height          =   3855
      Left            =   0
      TabIndex        =   26
      Top             =   4920
      Width           =   12015
      _Version        =   1572864
      _ExtentX        =   21193
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
      ItemCount       =   6
      SelectedItem    =   1
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
         TabIndex        =   74
         Top             =   1680
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1572864
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
         TabIndex        =   58
         Top             =   1560
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1572864
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
         TabIndex        =   66
         Top             =   1200
         Visible         =   0   'False
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
         Left            =   2160
         TabIndex        =   47
         Top             =   840
         Width           =   1815
         _Version        =   1572864
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
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1572864
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
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   11775
            _Version        =   1572864
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
            TabIndex        =   42
            Top             =   2760
            Width           =   1455
            _Version        =   1572864
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
            TabIndex        =   36
            Top             =   480
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
            TabIndex        =   38
            Top             =   960
            Width           =   9735
            _Version        =   1572864
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
            TabIndex        =   39
            Top             =   480
            Width           =   1335
            _Version        =   1572864
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
            TabIndex        =   41
            Top             =   2280
            Width           =   9735
            _Version        =   1572864
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
            TabIndex        =   43
            Top             =   2760
            Width           =   5895
            _Version        =   1572864
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
            TabIndex        =   40
            Top             =   2280
            Width           =   1455
            _Version        =   1572864
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
            TabIndex        =   37
            Top             =   960
            Width           =   1455
            _Version        =   1572864
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
            TabIndex        =   35
            Top             =   480
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
            Alignment       =   1
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   1815
         Left            =   -68200
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1572864
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
         Left            =   -66400
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
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
         Left            =   -61480
         TabIndex        =   32
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
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
         Left            =   2160
         TabIndex        =   45
         Top             =   1200
         Width           =   9735
         _Version        =   1572864
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
         Left            =   2160
         TabIndex        =   49
         Top             =   480
         Width           =   1815
         _Version        =   1572864
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
         Left            =   2160
         TabIndex        =   50
         Top             =   3120
         Width           =   1455
         _Version        =   1572864
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
         TabIndex        =   52
         Top             =   480
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
         TabIndex        =   54
         Top             =   840
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
         TabIndex        =   56
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnFondo 
         Height          =   330
         Index           =   0
         Left            =   -65200
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
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
         TabIndex        =   60
         Top             =   3360
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
         TabIndex        =   61
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
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
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
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
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
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
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1572864
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
         TabIndex        =   68
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1572864
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
         TabIndex        =   72
         Top             =   1150
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
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
         TabIndex        =   73
         Top             =   1155
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
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
         TabIndex        =   75
         Top             =   360
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1572864
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
         Item(0).ControlCount=   10
         Item(0).Control(0)=   "ListView1"
         Item(0).Control(1)=   "cboE_Tag"
         Item(0).Control(2)=   "btnEtiqueta"
         Item(0).Control(3)=   "Label2(18)"
         Item(0).Control(4)=   "Label2(19)"
         Item(0).Control(5)=   "Label2(20)"
         Item(0).Control(6)=   "Label2(21)"
         Item(0).Control(7)=   "chkE_Correo"
         Item(0).Control(8)=   "txtE_Observaciones"
         Item(0).Control(9)=   "txtE_Destinatarios"
         Item(1).Caption =   "Histórico"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswEtiquetas"
         Begin XtremeSuiteControls.ListView lswEtiquetas 
            Height          =   2895
            Left            =   -69880
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   11775
            _Version        =   1572864
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
            Left            =   -69880
            TabIndex        =   77
            Top             =   360
            Width           =   11775
            _Version        =   1572864
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
         Begin XtremeSuiteControls.CheckBox chkE_Correo 
            Height          =   255
            Left            =   10200
            TabIndex        =   76
            Top             =   2760
            Width           =   1455
            _Version        =   1572864
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
            Left            =   2040
            TabIndex        =   78
            Top             =   480
            Visible         =   0   'False
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
         Begin XtremeSuiteControls.FlatEdit txtE_Observaciones 
            Height          =   1215
            Left            =   2040
            TabIndex        =   79
            Top             =   960
            Width           =   9735
            _Version        =   1572864
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
            Left            =   10440
            TabIndex        =   80
            Top             =   480
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtE_Destinatarios 
            Height          =   330
            Left            =   2040
            TabIndex        =   81
            Top             =   2280
            Width           =   9735
            _Version        =   1572864
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
            Left            =   360
            TabIndex        =   69
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
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
            Left            =   360
            TabIndex        =   84
            Top             =   960
            Width           =   1455
            _Version        =   1572864
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
            Left            =   360
            TabIndex        =   83
            Top             =   2280
            Width           =   1455
            _Version        =   1572864
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
            Left            =   2040
            TabIndex        =   82
            Top             =   2760
            Width           =   5895
            _Version        =   1572864
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
         TabIndex        =   71
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         TabIndex        =   65
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         TabIndex        =   63
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         TabIndex        =   59
         Top             =   3360
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         TabIndex        =   55
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         TabIndex        =   53
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
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
         Left            =   360
         TabIndex        =   48
         Top             =   480
         Width           =   1815
         _Version        =   1572864
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
         Left            =   360
         TabIndex        =   46
         Top             =   840
         Width           =   1815
         _Version        =   1572864
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
         Left            =   720
         TabIndex        =   44
         Top             =   1200
         Width           =   1215
         _Version        =   1572864
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
         Left            =   -63640
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1572864
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
         Left            =   -68200
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
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
         Left            =   -69640
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
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
      Height          =   1455
      Left            =   0
      TabIndex        =   25
      Top             =   3480
      Width           =   12015
      _Version        =   1572864
      _ExtentX        =   21193
      _ExtentY        =   2566
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
      Item(1).Control(0)=   "gbPolizaIncendio"
      Begin XtremeSuiteControls.GroupBox gbPolizaVida 
         Height          =   1095
         Left            =   -120
         TabIndex        =   92
         Top             =   360
         Width           =   12135
         _Version        =   1572864
         _ExtentX        =   21405
         _ExtentY        =   1931
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnPolizaModificar 
            Height          =   375
            Index           =   0
            Left            =   10560
            TabIndex        =   101
            Top             =   120
            Width           =   1335
            _Version        =   1572864
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
            Left            =   2400
            TabIndex        =   102
            Top             =   120
            Width           =   3135
            _Version        =   1572864
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
            Left            =   6960
            TabIndex        =   103
            Top             =   120
            Width           =   3135
            _Version        =   1572864
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
            Left            =   2400
            TabIndex        =   104
            Top             =   600
            Width           =   1095
            _Version        =   1572864
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
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
            Left            =   360
            TabIndex        =   107
            Top             =   120
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
            Left            =   4920
            TabIndex        =   106
            Top             =   120
            Width           =   1815
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
            Left            =   360
            TabIndex        =   105
            Top             =   600
            Width           =   1815
         End
      End
      Begin XtremeSuiteControls.GroupBox gbPolizaIncendio 
         Height          =   1095
         Left            =   -70000
         TabIndex        =   93
         Top             =   360
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1572864
         _ExtentX        =   21193
         _ExtentY        =   1931
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtPI_Finca 
            Height          =   330
            Left            =   2280
            TabIndex        =   94
            Top             =   600
            Width           =   3135
            _Version        =   1572864
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
         Begin XtremeSuiteControls.PushButton btnPolizaModificar 
            Height          =   375
            Index           =   1
            Left            =   10440
            TabIndex        =   95
            Top             =   120
            Width           =   1335
            _Version        =   1572864
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
            Left            =   2280
            TabIndex        =   96
            Top             =   120
            Width           =   3135
            _Version        =   1572864
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
            Left            =   6840
            TabIndex        =   97
            Top             =   120
            Width           =   3135
            _Version        =   1572864
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
            Left            =   240
            TabIndex        =   100
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
            Left            =   4800
            TabIndex        =   99
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
            Left            =   240
            TabIndex        =   98
            Top             =   600
            Width           =   1815
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox gbDatos 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   11895
      _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   7920
         TabIndex        =   12
         Top             =   600
         Width           =   3135
         _Version        =   1572864
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
         TabIndex        =   13
         Top             =   600
         Width           =   2895
         _Version        =   1572864
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
         TabIndex        =   17
         Top             =   600
         Width           =   2895
         _Version        =   1572864
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
         TabIndex        =   20
         Top             =   1080
         Width           =   2295
         _Version        =   1572864
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboPago 
         Height          =   330
         Left            =   6240
         TabIndex        =   23
         Top             =   1440
         Width           =   1815
         _Version        =   1572864
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
      Begin XtremeSuiteControls.DateTimePicker dtpNacimiento 
         Height          =   315
         Left            =   6240
         TabIndex        =   88
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   330
         Left            =   9480
         TabIndex        =   89
         Top             =   1080
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
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
         TabIndex        =   24
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   360
         Index           =   0
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   2895
         _Version        =   1572864
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
         TabIndex        =   15
         Top             =   240
         Width           =   2895
         _Version        =   1572864
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
         TabIndex        =   14
         Top             =   240
         Width           =   3135
         _Version        =   1572864
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
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   847
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
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
      BackColor       =   16777215
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPolizaId 
      Height          =   480
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   2055
      _Version        =   1572864
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
      _Version        =   1572864
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
      Left            =   9600
      TabIndex        =   6
      Top             =   360
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   85
      Top             =   1200
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmPoliza_Reclamo.frx":41E8
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   86
      Top             =   1200
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmPoliza_Reclamo.frx":4919
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   2
      Left            =   10800
      TabIndex        =   87
      Top             =   1200
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmPoliza_Reclamo.frx":4EBD
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   375
      Left            =   11280
      TabIndex        =   108
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   1200
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmPoliza_Reclamo.frx":4F46
   End
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   255
      Left            =   4320
      TabIndex        =   91
      Top             =   1200
      Width           =   3855
      _Version        =   1572864
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "[ESTADO]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblPoliza 
      Height          =   255
      Left            =   120
      TabIndex        =   90
      Top             =   1200
      Width           =   4095
      _Version        =   1572864
      _ExtentX        =   7223
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "[POLIZA]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   9600
      TabIndex        =   7
      Top             =   120
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
      _Version        =   1572864
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
      _Version        =   1572864
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
      _Version        =   1572864
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean
Dim vCodigo As Long, vFecha As Date

Private Sub sbReclamo_Add()

On Error GoTo vError

'spPoliza_Reclamo_Add(@ReclamoId int, @Operacion int, @PolizaCodigo varchar(10), @PolizaId int
'            , @Cedula varchar(20), @Apellido1 varchar(30), @Apellido2 varchar(30), @Nombre varchar(30)
'            , @Sexo char(1), @FechaNac datetime, @Edad smallint, @Finca varchar(30)
'            , @SiniestroId int, @CausaId int, @MotivoId int, @EnfermedadId int, @DesembolsoId smallint, @PagoId smallint
'            , @Observaciones varchar(1000)
'            , @Usuario varchar(30))

Dim pSiniestroId As String, pCausaId As String
Dim pMotivoId As String, pEnfermedadId  As String


If tcMain(0).Visible Then
    pMotivoId = cboPV_Motivo.ItemData(cboPV_Motivo.ListIndex)
    pEnfermedadId = cboPV_Enfermedad.ItemData(cboPV_Enfermedad.ListIndex)
    
    pSiniestroId = "Null"
    pCausaId = "Null"
Else
    pMotivoId = "Null"
    pEnfermedadId = "Null"
    
    pSiniestroId = cboPI_Tipo.ItemData(cboPI_Tipo.ListIndex)
    pCausaId = cboPI_Causa.ItemData(cboPI_Causa.ListIndex)
    
End If

Me.MousePointer = vbHourglass

strSQL = "exec spPoliza_Reclamo_Add " & txtReclamoId.Text & ", " & txtOperacion.Text & ", '" & txtPolizaCodigo.Text & "', " & txtPolizaId.Text _
       & ", '" & txtCedula.Text & "', '" & txtApellido1.Text & "', '" & txtApellido2.Text & "', '" & txtNombre.Text & "', '" & Mid(cboSexo.Text, 1, 1) _
       & "', '" & Format(dtpNacimiento.Value, "yyyy-mm-dd") & "', " & txtPV_Edad.Text & ", '" & txtPI_Finca.Text _
       & "', " & pSiniestroId & ", " & pCausaId & ", " & pMotivoId & ", " & pEnfermedadId & ", " & cboDesembolso.ItemData(cboDesembolso.ListIndex) _
       & ", " & cboPago.ItemData(cboPago.ListIndex) & ", '" & txtObservaciones.Text & "', '" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
    MsgBox "Reclamo registrado satisfactoriamente!", vbInformation
Else
    MsgBox rs!Mensaje, vbCritical
End If

Call sbReclamo_Load(rs!ReclamoId)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnAccion_Click(Index As Integer)

Select Case Index
    Case 0 'Registro
        Call sbReclamo_Add
    Case 1 'Anular
        Call sbAnular
    Case 2 'Beneficiarios
        GLOBALES.gTag = txtOperacion.Text
        GLOBALES.gTag2 = txtPolizaId.Text
    
        Call sbFormsCall("frmCR_PolizasRegistroBeneficiarios", 1, , , False, Me)

End Select


End Sub

Private Sub btnAdjuntos_Click()
  
If CLng(txtReclamoId.Text) = 0 Then Exit Sub
  
 gGA.Modulo = "POL"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = txtReclamoId.Text
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub btnEtiqueta_Click()
If txtReclamoId.Text = "0" Then
   MsgBox "Consulte un reclamo registrado!", vbExclamation
   Exit Sub
End If

If Len(txtE_Observaciones.Text) < 10 Then
   MsgBox "Indique una observación válida! !", vbExclamation
   Exit Sub
End If

If chkE_Correo.Value = xtpChecked And Len(txtE_Destinatarios.Text) < 10 Then
   MsgBox "No ha indicado destinatarios válidos!", vbExclamation
   Exit Sub
End If

On Error GoTo vError

'spPoliza_Reclamo_Etiqueta_Manual_Add(@ReclamoId int, @Observaciones varchar(1000), @ICorreo smallint, @Destinatarios varchar(max), @Usuario varchar(30))
strSQL = "exec spPoliza_Reclamo_Etiqueta_Manual_Add " & txtReclamoId.Text _
       & ", '" & txtE_Observaciones.Text _
       & "', " & chkE_Correo.Value _
       & ", '" & txtE_Destinatarios.Text _
       & "', '" & glogon.Usuario & "'"

Me.MousePointer = vbHourglass

Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
    MsgBox "Etiqueta registrada satisfactoriamente!", vbInformation
Else
    MsgBox rs!Mensaje, vbCritical
End If


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnPolizaModificar_Click(Index As Integer)
If txtReclamoId.Text = "0" Then
   MsgBox "Consulte un reclamo registrado!", vbExclamation
   Exit Sub
End If

On Error GoTo vError


Select Case Index
    Case 0 'Modifica Vida
        'spPoliza_Reclamo_Actualiza_Datos_Vida(@ReclamoId int, @MotivoId smallint, @Enfermedad smallint, @Usuario varchar(30))
        strSQL = "exec spPoliza_Reclamo_Actualiza_Datos_Vida " & txtReclamoId.Text _
               & ", " & cboPV_Motivo.ItemData(cboPV_Motivo.ListIndex) _
               & ", " & cboPV_Enfermedad.ItemData(cboPV_Enfermedad.ListIndex) _
               & ", " & txtPV_Edad.Text _
               & ", '" & glogon.Usuario & "'"
               
    Case 1 'Modifica Incendio
'spPoliza_Reclamo_Actualiza_Datos_Incendio(@ReclamoId int, @SiniestroId smallint, @Causa smallint, @Finca varchar(30), @Usuario varchar(30))
        strSQL = "exec spPoliza_Reclamo_Actualiza_Datos_Incendio " & txtReclamoId.Text _
               & ", " & cboPI_Tipo.ItemData(cboPI_Tipo.ListIndex) _
               & ", " & cboPI_Causa.ItemData(cboPI_Causa.ListIndex) _
               & ", '" & txtPI_Finca.Text _
               & "', '" & glogon.Usuario & "'"
    
    
End Select

Me.MousePointer = vbHourglass

Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
    MsgBox "Datos de Motivos y/o Causales de Póliza registrados satisfactoriamente!", vbInformation
Else
    MsgBox rs!Mensaje, vbCritical
End If


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnRecepcion_Click()
On Error GoTo vError

If txtRec_Usuario.Text <> "" Then
    MsgBox "La recepción ya fué aplicada a este reclamo!", vbExclamation
    Exit Sub
End If

If txtReclamoId.Text = "0" Then
   MsgBox "Consulte un reclamo registrado!", vbExclamation
   Exit Sub
End If

If Len(txtRec_Observacion.Text) < 5 Then
   MsgBox "Indique una Observación válida", vbExclamation
   Exit Sub
End If

If dtpRec_Fecha.Value > vFecha Then
   MsgBox "La fecha de recepción no puede ser futura!", vbExclamation
   Exit Sub
End If

'spPoliza_Reclamo_Actualiza_Recepcion(@ReclamoId int, @Fecha datetime, @Observaciones varchar(30), @Usuario varchar(30))

strSQL = "exec spPoliza_Reclamo_Actualiza_Recepcion " & txtReclamoId.Text _
       & ", '" & Format(dtpRec_Fecha.Value, "yyyy-mm-dd") _
       & "', '" & txtRec_Observacion.Text & "', '" & glogon.Usuario & "'"


Me.MousePointer = vbHourglass

Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
    MsgBox "Recepción del Reclamo registrado satisfactoriamente!", vbInformation
Else
    MsgBox rs!Mensaje, vbCritical
End If


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub btnSeguimiento_Click()

If txtReclamoId.Text = "0" Then
   MsgBox "Consulte un reclamo registrado!", vbExclamation
   Exit Sub
End If

If Len(txtS_Observacion.Text) < 10 Then
   MsgBox "Indique una observación válida! !", vbExclamation
   Exit Sub
End If

If chkS_Correo.Value = xtpChecked And Len(txtS_Destinatarios.Text) < 10 Then
   MsgBox "No ha indicado destinatarios válidos!", vbExclamation
   Exit Sub
End If


On Error GoTo vError
'spPoliza_Reclamo_Seguimiento_Manual_Add(@ReclamoId int, @EstadoId smallint, @Observaciones varchar(1000), @ICorreo smallint, @Destinatarios varchar(max), @Usuario varchar(30))
strSQL = "exec spPoliza_Reclamo_Seguimiento_Manual_Add " & txtReclamoId.Text & ", " & cboS_Estado.ItemData(cboS_Estado.ListIndex) _
       & ", '" & txtS_Observacion.Text _
       & "', " & chkS_Correo.Value _
       & ", '" & txtS_Destinatarios.Text _
       & "', '" & glogon.Usuario & "'"

Me.MousePointer = vbHourglass

Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
    MsgBox "Seguimiento registrado satisfactoriamente!", vbInformation
Else
    MsgBox rs!Mensaje, vbCritical
End If


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:


End Sub

Private Sub Form_Load()
vModulo = 3
Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

vFecha = fxFechaServidor

dtpRec_Fecha.Value = vFecha

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


dtpNacimiento.Value = fxFechaServidor

cboSexo.AddItem "Masculino"
cboSexo.AddItem "Femenino"
cboSexo.AddItem "Otro"
cboSexo.Text = "Masculino"

cboPago.AddItem "Transferencia"
cboPago.ItemData(cboPago.ListCount - 1) = "1"
cboPago.AddItem "Cheque"
cboPago.ItemData(cboPago.ListCount - 1) = "2"
cboPago.Text = "Transferencia"


cboDesembolso.AddItem "Parcial"
cboDesembolso.ItemData(cboDesembolso.ListCount - 1) = "1"
cboDesembolso.AddItem "Total"
cboDesembolso.ItemData(cboDesembolso.ListCount - 1) = "0"
cboDesembolso.Text = "Parcial"


cboPV_Motivo.AddItem "Incapacidad Permanente"
cboPV_Motivo.ItemData(cboPV_Motivo.ListCount - 1) = "0"
cboPV_Motivo.AddItem "Muerte"
cboPV_Motivo.ItemData(cboPV_Motivo.ListCount - 1) = "1"
cboPV_Motivo.AddItem "No Especificado"
cboPV_Motivo.ItemData(cboPV_Motivo.ListCount - 1) = "2"
cboPV_Motivo.Text = "Incapacidad Permanente"


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbAnular()
On Error GoTo vError

Me.MousePointer = vbHourglass

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGuardar()
On Error GoTo vError

Me.MousePointer = vbHourglass

Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbReclamo_Load(pReclamoId As Long)
On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPoliza_Reclamo_Load " & pReclamoId
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    
    txtReclamoId.Text = rs!Id
    txtOperacion.Text = rs!Id_Solicitud
    txtPolizaCodigo.Text = rs!CODIGO_POLIZA
    txtPolizaId.Text = rs!Id_Solicitud_Poliza
    
    txtCedula.Text = rs!Cedula
    txtApellido1.Text = rs!Primer_Apellido
    txtApellido2.Text = rs!Segundo_Apellido
    txtNombre.Text = rs!Nombre
    
    dtpNacimiento.Value = rs!Fecha_Nacimiento
    cboSexo.Text = rs!Sexo_Desc
    
    lblEstado.Caption = rs!Estado_Desc
    lblPoliza.Caption = rs!Poliza_Desc
    
    txtPI_Finca.Text = rs!Finca & ""
    txtPV_Edad.Text = rs!Edad & ""
    
    Call sbCboAsignaDato(cboS_Estado, rs!Estado_Desc, True, rs!Estado_Actual)
    Call sbCboAsignaDato(cboDesembolso, rs!Forma_Desembolso_Desc, True, rs!Forma_Desembolso)
    Call sbCboAsignaDato(cboPago, rs!Metodo_Pago_Desc, True, rs!Metodo_Pago)
    
    If Not IsNull(rs!Enfermedad) Then
        Call sbCboAsignaDato(cboPV_Enfermedad, rs!Enfermedad_Desc, True, rs!Enfermedad)
    End If
    If Not IsNull(rs!Motivo_Reclamo) Then
        Call sbCboAsignaDato(cboPV_Motivo, rs!Motivo_Reclamo_Desc, True, rs!Motivo_Reclamo)
    End If
    
    
    If Not IsNull(rs!Tipo_Siniestro) Then
        Call sbCboAsignaDato(cboPI_Tipo, rs!Tipo_Siniestro_Desc, True, rs!Tipo_Siniestro)
    End If
    If Not IsNull(rs!Causa_Siniestro) Then
        Call sbCboAsignaDato(cboPI_Causa, rs!Causa_Desc, True, rs!Causa_Siniestro)
    End If
    
    
    If rs!Tipo_Poliza = "V" Then
        tcMain.Item(0).Visible = True
        tcMain.Item(1).Visible = False
        tcMain.Item(0).Selected = True
    Else
        tcMain.Item(0).Visible = False
        tcMain.Item(1).Visible = True
        tcMain.Item(1).Selected = True
    End If
    
    txtR_Fecha.Text = rs!Registro_Fecha & ""
    txtR_Usuario.Text = rs!Registro_Usuario & ""
    txtObservaciones.Text = rs!Registro_Observaciones & ""
    
    If Not IsNull(rs!Recepcion_Fecha) Then
        dtpRec_Fecha.Value = rs!Recepcion_Fecha
        txtRec_Usuario.Text = rs!Recepcion_Usuario & ""
        txtRec_Observacion.Text = rs!Recepcion_Observaciones & ""
    End If
    
    txtF_Aprobado.Text = Format(rs!Monto_Aprobado, "Standard")
    txtF_Monto.Text = Format(rs!Saldo_Fondo, "Standard")
    txtF_MontoOperacion.Text = Format(rs!Saldo_Credito, "Standard")
    
    
    tcAux.Item(0).Selected = True
    
End If
Me.MousePointer = vbDefault

If txtCedula.Text = "" Then
    MsgBox "No se encontró el Reclamo No." & pReclamoId, vbExclamation
    UnLoad Me
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Public Sub sbNuevo(pCedula As String, pOperacion As Long, pPoliza As Long, pPolizaCodigo As String)
On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPoliza_Reclamo_Nuevo '" & pCedula & "', " & pOperacion & ", " & pPoliza & ", " & pPolizaCodigo
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    
    txtReclamoId.Text = "0"
    txtOperacion.Text = rs!Id_Solicitud
    
    txtPolizaCodigo.Text = rs!POLIZA_CODIGO
    txtPolizaId.Text = rs!POLIZA_ID
    
    txtCedula.Text = rs!Cedula
    txtApellido1.Text = rs!Apellido1
    txtApellido2.Text = rs!Apellido2
    txtNombre.Text = rs!Nombrev2
    
    dtpNacimiento.Value = rs!Fecha_Nac
    cboSexo.Text = rs!Sexo_Desc
    
    lblEstado.Caption = "Borrador"
    lblPoliza.Caption = rs!Poliza_Desc
    
    txtPI_Finca.Text = rs!Finca & ""
    txtPV_Edad.Text = rs!Edad & ""
    
    If rs!Tipo_Poliza = "V" Then
        tcMain.Item(0).Visible = True
        tcMain.Item(1).Visible = False
        tcMain.Item(0).Selected = True
    Else
        tcMain.Item(0).Visible = False
        tcMain.Item(1).Visible = True
        tcMain.Item(1).Selected = True
    End If
    
    txtR_Fecha.Text = ""
    txtR_Usuario.Text = ""
    txtObservaciones.Text = ""
    
    dtpRec_Fecha.Value = vFecha
    txtRec_Usuario.Text = ""
    txtRec_Observacion.Text = ""
    
    txtF_Aprobado.Text = Format(0, "Standard")
    txtF_Monto.Text = Format(0, "Standard")
    txtF_MontoOperacion.Text = Format(rs!Saldo_Credito, "Standard")
    
    
    tcAux.Item(0).Selected = True
    
    tcAux.Item(1).Visible = False
    tcAux.Item(2).Visible = False
    tcAux.Item(3).Visible = False
    tcAux.Item(4).Visible = False
    tcAux.Item(5).Visible = False
    
    
    If rs!Reclamo_Id > 0 Then
    
        Me.MousePointer = vbDefault
        MsgBox "Existe un Reclamo en Proceso para Esta Póliza, Reclamo No." & rs!Reclamo_Id, vbExclamation
        UnLoad Me
    
    End If
End If

Me.MousePointer = vbDefault

If txtCedula.Text = "" Then
    MsgBox "No se encontró el registro de la Persona o Póliza!", vbExclamation
    UnLoad Me
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcEtiqueta_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPoliza_Reclamo_Etiquetas_List " & txtReclamoId.Text
Call OpenRecordSet(rs, strSQL)

With lswEtiquetas.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Id_Etiqueta)
            itmX.SubItems(1) = rs!Fecha
            itmX.SubItems(2) = rs!Usuario
            itmX.SubItems(3) = rs!Observaciones
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

Private Sub tcSeguimiento_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPoliza_Reclamo_Seguimiento_List " & txtReclamoId.Text
Call OpenRecordSet(rs, strSQL)

With lswSeguimiento.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Fecha)
            itmX.SubItems(1) = rs!Usuario
            itmX.SubItems(2) = rs!Estado_Desc
            itmX.SubItems(3) = rs!Observaciones
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

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True


strSQL = "select ID_SINIESTRO as 'IdX', rtrim(DESCRIPCION)  as 'itmX' from POLIZAS_SINIESTROS_TIPOS where ACTIVO = 1"
Call sbCbo_Llena_New(cboPI_Tipo, strSQL, False, True)

strSQL = "select ID as 'IdX', rtrim(NOMBRE)  as 'itmX'  from VIV_POLIZAS_VIDA_ENFERMEDAD where ACTIVO = 1"
Call sbCbo_Llena_New(cboPV_Enfermedad, strSQL, False, True)

strSQL = "select ID as 'IdX', rtrim(Descripcion)  as 'itmX'  from VIV_POLIZAS_INCENDIO_CAUSA where ACTIVO = 1"
Call sbCbo_Llena_New(cboPI_Causa, strSQL, False, True)

strSQL = "select ID_ESTADO  as 'IdX', rtrim(Descripcion)  as 'itmX'   from POLIZAS_RECLAMOS_ESTADOS where ACTIVO = 1"
Call sbCbo_Llena_New(cboS_Estado, strSQL, False, True)

'Consulta todas las cuentas Bancarias
strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)


'strSQL = "select cod_divisa as 'IdX', Descripcion as 'ItmX' from vSys_Divisas" _
'       & " order by divisa_local desc, Descripcion asc"
'Call sbCbo_Llena_New(cboSalarioDivisa, strSQL, False, True)

'
'strSQL = " select Catalogo_Id as 'IdX', Descripcion as 'ItmX' " _
'       & " from AFI_CATALOGOS Where Tipo_Id = 3 order by Descripcion"
'Call sbCbo_Llena_New(cboNivelAcademico, strSQL, False, True)
'
'
'strSQL = "select cod_nacionalidad as 'IdX', Descripcion as 'ItmX' from Sys_nacionalidades" _
'       & " where Activo = 1" _
'       & " order by Omision desc, Descripcion asc"
'Call sbCbo_Llena_New(cboNacionalidad, strSQL, False, True)
'
'strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
'       & " where Activo = 1" _
'       & " order by Descripcion asc"
'Call sbCbo_Llena_New(cboEstado, strSQL, False, True)
'
'strSQL = "select Estado_Laboral as 'IdX', Descripcion as 'ItmX' from AFI_ESTADO_LABORAL" _
'       & " where Activo = 1" _
'       & " order by Descripcion asc"
'Call sbCbo_Llena_New(cboEstadoLaboral, strSQL, False, True)

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub
