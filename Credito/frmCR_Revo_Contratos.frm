VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCR_Revo_Contratos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contratos Revolutivos"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   10380
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9960
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7770
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3069
            MinWidth        =   3069
            Object.ToolTipText     =   "Usuario"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   330
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Object.ToolTipText     =   "Imprime el listado seleccionado"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra esta ventana"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6615
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   11668
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
      Item(0).Caption =   "Recepcion"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "fraOperacion"
      Item(0).Control(1)=   "GroupBox2"
      Item(0).Control(2)=   "Label1(26)"
      Item(1).Caption =   "Formalizacion"
      Item(1).ControlCount=   20
      Item(1).Control(0)=   "chkDeducPlanilla"
      Item(1).Control(1)=   "GroupBox1"
      Item(1).Control(2)=   "fraEstadoFormalizacion"
      Item(1).Control(3)=   "cboMes"
      Item(1).Control(4)=   "txtAno"
      Item(1).Control(5)=   "Label2"
      Item(1).Control(6)=   "Label5(0)"
      Item(1).Control(7)=   "cboDeductora"
      Item(1).Control(8)=   "cboFrecuencia"
      Item(1).Control(9)=   "ShortcutCaption1(0)"
      Item(1).Control(10)=   "ShortcutCaption1(1)"
      Item(1).Control(11)=   "chkPlanAhorros"
      Item(1).Control(12)=   "txtPlan"
      Item(1).Control(13)=   "Label1(5)"
      Item(1).Control(14)=   "txtPlanMensualidad"
      Item(1).Control(15)=   "Label1(6)"
      Item(1).Control(16)=   "txtPlanDesc"
      Item(1).Control(17)=   "chkPlanFijarMensualidad"
      Item(1).Control(18)=   "Label5(1)"
      Item(1).Control(19)=   "dtpVencimiento"
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   2895
         Left            =   0
         TabIndex        =   55
         Top             =   3840
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   5106
         _StockProps     =   79
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   735
            Left            =   1320
            TabIndex        =   56
            Top             =   1920
            Width           =   8655
            _Version        =   1441793
            _ExtentX        =   15266
            _ExtentY        =   1296
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPAportes 
            Height          =   315
            Left            =   1320
            TabIndex        =   62
            Top             =   1080
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuotaCorte 
            Height          =   315
            Left            =   4800
            TabIndex        =   64
            Top             =   1440
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPMensualidad 
            Height          =   315
            Left            =   1320
            TabIndex        =   66
            Top             =   1440
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSaldo 
            Height          =   315
            Left            =   4800
            TabIndex        =   72
            Top             =   1080
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTransitoCreditos 
            Height          =   315
            Left            =   8160
            TabIndex        =   74
            Top             =   1440
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDisponible 
            Height          =   315
            Left            =   8160
            TabIndex        =   77
            Top             =   720
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTransitoDebitos 
            Height          =   315
            Left            =   8160
            TabIndex        =   79
            Top             =   1080
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPPlan 
            Height          =   315
            Left            =   1320
            TabIndex        =   58
            Top             =   360
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPContrato 
            Height          =   315
            Left            =   1320
            TabIndex        =   60
            Top             =   720
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtOperacionCrd 
            Height          =   315
            Left            =   4800
            TabIndex        =   68
            Top             =   360
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEstudio 
            Height          =   315
            Left            =   4800
            TabIndex        =   70
            Top             =   720
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   315
            Left            =   8160
            TabIndex        =   95
            Top             =   360
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
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
         Begin VB.Label Label1 
            Caption         =   "Tra. Créditos"
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
            Index           =   24
            Left            =   6720
            TabIndex        =   80
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Tra. Débitos"
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
            Index           =   23
            Left            =   6720
            TabIndex        =   78
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Disponible"
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
            Index           =   22
            Left            =   6720
            TabIndex        =   76
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   21
            Left            =   6720
            TabIndex        =   75
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Cuota al corte"
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
            Index           =   20
            Left            =   3360
            TabIndex        =   73
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Saldo"
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
            Index           =   18
            Left            =   3360
            TabIndex        =   71
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "No.Estudio"
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
            Left            =   3360
            TabIndex        =   69
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "No. Operacion"
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
            Index           =   15
            Left            =   3360
            TabIndex        =   67
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Mensualidad"
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
            Index           =   14
            Left            =   120
            TabIndex        =   65
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Aportes"
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
            Index           =   13
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "No. Contrato"
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
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Plan Ahorros"
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
            Index           =   7
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   57
            Top             =   1920
            Width           =   855
         End
      End
      Begin VB.Frame fraOperacion 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3735
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   9855
         Begin XtremeSuiteControls.CheckBox chkAjustaCuotasAlVencimiento 
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1920
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "La operación de crédito ajustará las cuotas al vencimiento de este contrato?"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboGarantia 
            Height          =   330
            Left            =   7680
            TabIndex        =   7
            Top             =   1080
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.ComboBox cboComite 
            Height          =   330
            Left            =   1320
            TabIndex        =   8
            Top             =   1440
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   315
            Left            =   4440
            TabIndex        =   9
            Top             =   0
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   556
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
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   315
            Left            =   4440
            TabIndex        =   10
            Top             =   360
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   556
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
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPromotorNombre 
            Height          =   315
            Left            =   1800
            TabIndex        =   11
            Top             =   1080
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   556
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
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaSolicitud 
            Height          =   315
            Left            =   7680
            TabIndex        =   12
            Top             =   2880
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   315
            Left            =   7680
            TabIndex        =   13
            Top             =   1440
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
            _ExtentY        =   556
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
            Alignment       =   1
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   315
            Left            =   8760
            TabIndex        =   14
            Top             =   1800
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   556
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
            Alignment       =   1
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtTasa 
            Height          =   315
            Left            =   8760
            TabIndex        =   15
            Top             =   2160
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   556
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
            Alignment       =   1
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtCuota 
            Height          =   315
            Left            =   7680
            TabIndex        =   16
            Top             =   2520
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   315
            Left            =   1320
            TabIndex        =   17
            Top             =   0
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2984
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
         Begin XtremeSuiteControls.FlatEdit txtCodigo 
            Height          =   315
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtPromotorId 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Top             =   1080
            Width           =   510
            _Version        =   1441793
            _ExtentX        =   889
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDivisa 
            Height          =   315
            Left            =   2400
            TabIndex        =   20
            Top             =   360
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1085
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkTopeDisponible 
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   2400
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Disponible de Giro =  Monto Contrato - (Retiros - Amortizado)"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkSupervision 
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   2880
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Indicar si para retiros se requiere proceso de Supervisión"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboDestino 
            Height          =   330
            Left            =   4440
            TabIndex        =   93
            Top             =   720
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   582
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
         Begin VB.Label Label1 
            Caption         =   "Destino"
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
            Index           =   28
            Left            =   3240
            TabIndex        =   94
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre"
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
            Left            =   3240
            TabIndex        =   31
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   30
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Descripción"
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
            Left            =   3240
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Garantía"
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
            Index           =   9
            Left            =   6480
            TabIndex        =   28
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   11
            Left            =   6480
            TabIndex        =   27
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Evaluado "
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
            Left            =   0
            TabIndex        =   26
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Cuota"
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
            Index           =   8
            Left            =   6480
            TabIndex        =   25
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label lblTasa 
            Caption         =   "Tasa"
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
            Left            =   6480
            TabIndex        =   24
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label lblPlazo 
            Caption         =   "Plazo"
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
            Left            =   6480
            TabIndex        =   23
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Monto"
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
            Left            =   6480
            TabIndex        =   22
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblFondoDisplay 
            Caption         =   "Ejecutivo"
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
            Index           =   3
            Left            =   0
            TabIndex        =   21
            Top             =   1080
            Width           =   735
         End
      End
      Begin XtremeSuiteControls.CheckBox chkDeducPlanilla 
         Height          =   255
         Left            =   -66040
         TabIndex        =   33
         Top             =   3720
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Deducir por Planilla"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1095
         Left            =   -69640
         TabIndex        =   34
         Top             =   5040
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441793
         _ExtentX        =   10186
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Recursos:"
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
         Begin VB.TextBox txtDisponibleRecursos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
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
            Left            =   2880
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   35
            Top             =   720
            Width           =   2535
         End
         Begin XtremeSuiteControls.ComboBox cboRecursos 
            Height          =   330
            Left            =   0
            TabIndex        =   36
            Top             =   360
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7011
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.DateTimePicker dtpDesembolso 
            Height          =   315
            Left            =   4080
            TabIndex        =   37
            Top             =   360
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   550
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
         Begin VB.Image imgGuardaFecDesembolso 
            Height          =   255
            Left            =   5520
            Picture         =   "frmCR_Revo_Contratos.frx":0000
            Stretch         =   -1  'True
            ToolTipText     =   "Guarda la Fecha de Desembolso"
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Disponible Recurso:"
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
            Height          =   336
            Index           =   1
            Left            =   0
            TabIndex        =   38
            Top             =   720
            Width           =   2412
         End
         Begin VB.Image imgRecalculoRecurso 
            Height          =   255
            Left            =   5520
            Picture         =   "frmCR_Revo_Contratos.frx":06D8
            Stretch         =   -1  'True
            ToolTipText     =   "Recalcula Monto Disponible Recursos"
            Top             =   720
            Width           =   255
         End
      End
      Begin XtremeSuiteControls.GroupBox fraEstadoFormalizacion 
         Height          =   2895
         Left            =   -63400
         TabIndex        =   39
         Top             =   3720
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   5106
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
         Appearance      =   16
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnFormalizar 
            Height          =   615
            Left            =   960
            TabIndex        =   40
            Top             =   1920
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "&Formalizar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Revo_Contratos.frx":0EA7
         End
         Begin XtremeSuiteControls.RadioButton optFormalizacion 
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   41
            Top             =   960
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Formalizar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            Value           =   -1  'True
            Alignment       =   1
         End
         Begin XtremeSuiteControls.RadioButton optFormalizacion 
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   42
            Top             =   1320
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Anular"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            Alignment       =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaFormalizacion 
            Height          =   315
            Left            =   1080
            TabIndex        =   43
            Top             =   360
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
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
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Fecha.:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   44
            Top             =   0
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.ComboBox cboMes 
         Height          =   315
         Left            =   -67720
         TabIndex        =   45
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtAno 
         Height          =   330
         Left            =   -68320
         TabIndex        =   46
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1085
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboDeductora 
         Height          =   330
         Left            =   -68320
         TabIndex        =   47
         Top             =   4080
         Visible         =   0   'False
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboFrecuencia 
         Height          =   315
         Left            =   -66400
         TabIndex        =   48
         Top             =   4440
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.CheckBox chkPlanAhorros 
         Height          =   255
         Left            =   -70000
         TabIndex        =   84
         Top             =   1800
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1441793
         _ExtentX        =   17595
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asociar un Plan de Ahorros con la Diferencia en Cuota Contrato vrs Cuota Operación de Crédito."
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
      Begin XtremeSuiteControls.FlatEdit txtPlan 
         Height          =   315
         Left            =   -68200
         TabIndex        =   85
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlanMensualidad 
         Height          =   315
         Left            =   -68200
         TabIndex        =   87
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlanDesc 
         Height          =   315
         Left            =   -66400
         TabIndex        =   89
         Top             =   2280
         Visible         =   0   'False
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkPlanFijarMensualidad 
         Height          =   255
         Left            =   -66280
         TabIndex        =   90
         Top             =   2640
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1441793
         _ExtentX        =   17595
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fijar monto mensual independiente al contrato?"
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
      Begin XtremeSuiteControls.DateTimePicker dtpVencimiento 
         Height          =   315
         Left            =   -61600
         TabIndex        =   92
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Vencimiento para el Contrato"
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
         Left            =   -65680
         TabIndex        =   91
         Top             =   600
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensualidad"
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
         Index           =   6
         Left            =   -69400
         TabIndex        =   88
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Ahorros"
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
         Index           =   5
         Left            =   -69400
         TabIndex        =   86
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   -70000
         TabIndex        =   83
         Top             =   1200
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1441793
         _ExtentX        =   17595
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Plan de Ahorros"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   82
         Top             =   3240
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1441793
         _ExtentX        =   17595
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Formalización"
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
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Deduc."
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
         Left            =   -69640
         TabIndex        =   51
         Top             =   4440
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   -69880
         TabIndex        =   50
         Top             =   120
         Width           =   1212
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Deductora"
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
         Left            =   -69640
         TabIndex        =   49
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3480
      TabIndex        =   81
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   315
      Index           =   1
      Left            =   9000
      TabIndex        =   96
      Top             =   120
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Causas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCR_Revo_Contratos.frx":167F
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   315
      Index           =   0
      Left            =   7800
      TabIndex        =   97
      Top             =   120
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Giros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCR_Revo_Contratos.frx":1C9D
      ImageAlignment  =   4
   End
   Begin ComctlLib.ImageList imgIconosEstados 
      Left            =   9840
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":23BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":2C0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":3461
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":3CB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":4505
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":4D57
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":55A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_Revo_Contratos.frx":5DFB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin VB.Image imgEstado 
      Height          =   240
      Left            =   4080
      Picture         =   "frmCR_Revo_Contratos.frx":664D
      Stretch         =   -1  'True
      Top             =   480
      Width           =   240
   End
End
Attribute VB_Name = "frmCR_Revo_Contratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim vMensaje                As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita                  As Boolean 'Indica si se esta actualizando o insertando
Dim vPasaFormalizacion      As Boolean 'Indica si una formalizacion normal se procesa o no
Dim vDocumentoFormalizacion As Boolean 'Indica si se debe de generar una nota de debito
Dim vPaso                   As Boolean 'Para que Tes_Bancos Click cbo lo ignore
Dim vScroll                 As Boolean, vOperacionLoad As Boolean
'Por incluir una formalizacion que no pasa a Tesoreria o el Monto Girado es Cero

Dim mFrecuenciaPago As String


Private Sub btnAccion_Click(Index As Integer)


If Operacion.Operacion > 0 And Operacion.Estado = "F" Then
    
    GLOBALES.gTag = txtOperacion.Text

    If Index = 0 Then
        Call sbFormsCall("frmCR_Revo_Contratos_Retiros")
    Else
        Call sbFormsCall("frmCR_Revo_Contratos_SGT")
    End If
    
End If

End Sub

Private Sub cboComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboGarantia.SetFocus
End Sub


Private Sub cboDeductora_Click()

If vPaso Then Exit Sub

On Error GoTo vError

Dim strSQL As String, rs As New ADODB.Recordset
Dim vProceso As Currency, pProcesoClean As Long

strSQL = "select rtrim(descripcion) as 'Descripcion', isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & cboDeductora.ItemData(cboDeductora.ListIndex)
Call OpenRecordSet(rs, strSQL)
    mFrecuenciaPago = rs!Frecuencia_ID
rs.Close

cboFrecuencia.Clear
Select Case mFrecuenciaPago
    Case "M" 'Mensual
        cboFrecuencia.AddItem "Mensual"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "0"
        cboFrecuencia.Text = "Mensual"
    
    Case "Q" 'Quincenal
        cboFrecuencia.AddItem "1er Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "1"
        cboFrecuencia.AddItem "2da Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "2"
End Select
  
  
vProceso = fxPrimerDeduccion(Operacion.Codigo, cboDeductora.ItemData(cboDeductora.ListIndex))
pProcesoClean = vProceso

cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
txtAno.Text = Mid(pProcesoClean, 1, 4)

If (vProceso - pProcesoClean) = 0.1 Then
    cboFrecuencia.Text = "1er Quincena"
Else
    cboFrecuencia.Text = "2da Quincena"
End If

Exit Sub

vError:


End Sub

Private Sub cboDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPromotorNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "D.cod_destino"
   gBusquedas.Orden = "D.Cod_Destino"
   gBusquedas.Consulta = "select D.cod_Destino,D.descripcion" _
                        & " from catalogo_destinos D inner join catalogo_destinosASG C on D.cod_destino = C.cod_destino"
   gBusquedas.Filtro = " and C.codigo = '" & txtCodigo.Text & "' "
   frmBusquedas.Show vbModal
   cboDestino.Text = Trim(gBusquedas.Resultado2)
End If

End Sub




Private Sub cboGarantia_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim mGarantiaForm As String

If vPaso Then Exit Sub
If cboGarantia.ListCount <= 0 Then Exit Sub
If cboGarantia.Text = "" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select FORMULARIO  From CRD_GARANTIA_TIPOS" _
       & " where garantia = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
 mGarantiaForm = Trim(rs!formulario)
rs.Close


Operacion.PlazoBono = 0

Select Case mGarantiaForm
    Case "F01" 'Sobre Ahorros
        strSQL = "select dbo.fxCrdGarantiaPatMnt('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "', 'M') as 'Monto'" _
               & ",dbo.fxCrdTasaBonifica('" & txtCedula.Text & "','" & txtCodigo.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PtsBono'" _
               & ",dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
          Operacion.TasaPtsBono = rs!PtsBono
          txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
          If rs!PlazoBono > 0 Then
            txtPlazo.Text = rs!PlazoBono
            Operacion.PlazoBono = rs!PlazoBono
          End If
        rs.Close
    
    Case "F05" 'Fondos de Ahorros

    Case "F06" 'Adelanto de Salario
        strSQL = "select dbo.fxCrdDisponibleAdelantoSalario('" & txtCedula.Text & "', 'M') as 'Monto'" _
               & ",dbo.fxCrdTasaBonifica('" & txtCedula.Text & "','" & txtCodigo.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PtsBono'" _
               & ",dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
          Operacion.TasaPtsBono = rs!PtsBono
          txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
          If rs!PlazoBono > 0 Then
            txtPlazo.Text = rs!PlazoBono
            Operacion.PlazoBono = rs!PlazoBono
          End If
        rs.Close


    Case Else     'Otras Garantias
        strSQL = "select dbo.fxCrdTasaBonifica('" & txtCedula.Text & "','" & txtCodigo.Text _
             & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PtsBono'" _
             & ",dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
        Call OpenRecordSet(rs, strSQL)
          Operacion.TasaPtsBono = rs!PtsBono
          txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
        
          If rs!PlazoBono > 0 Then
            txtPlazo.Text = rs!PlazoBono
            Operacion.PlazoBono = rs!PlazoBono
          End If
        
        rs.Close

End Select

'Valida Montos, Tasas y Plazos
Call txtMonto_KeyPress(vbKeyReturn)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 If KeyCode = vbKeyReturn Then cboDestino.SetFocus
End Sub

Private Sub cboRecursos_Click()
txtDisponibleRecursos = 0
End Sub

 

Private Sub sbFormalizar()
Dim rs As New ADODB.Recordset, strSQL As String

Dim lngPriDeduc As Currency, FechaUltima As Currency
Dim vFecha As Date
Dim iMes As Integer, lngAnio As Long
Dim vFechaCalculo As Date, vBoletaCK As Boolean, vTipoCobro As String, vDias As Long
Dim vTransac As Boolean, vPuntosAdd As String, vTasaPiso As Currency, vBaseCalculo As String, vDiaPago As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass



vBoletaCK = False
vTransac = False
vBaseCalculo = "01" '360/360
vDiaPago = 32

vFecha = fxFechaServidor


'Preguntar SI es TBP / revisa si utiliza TBP del Destino
'Extrae Dia de Pago y Base de Calculo

vPuntosAdd = "NULL"

strSQL = "select TBP_Utiliza,TBP_Adicional,Tasa_Destino,Base_Calculo,dbo.fxCRDPoliticaPago(dbo.MyGetdate()) as DiaPago" _
       & " From catalogo where Codigo = '" & Operacion.Codigo & "'"
Call OpenRecordSet(rs, strSQL)

vBaseCalculo = Trim(rs!Base_Calculo)
If chkDeducPlanilla.Value = vbChecked Then
    vDiaPago = 32
Else
    vDiaPago = rs!DiaPago
End If

rs.Close



vDocumentoFormalizacion = False
vPasaFormalizacion = True


'Calculo de Intereses de Formalizacion
lngPriDeduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

strSQL = "update CRD_REV_CONTRATOS set IND_DEDUCE_PLANILLA = " & chkDeducPlanilla.Value & ", PRIMER_DEDUCCION = " & lngPriDeduc _
       & ", COD_GRUPO = '" & cboRecursos.ItemData(cboRecursos.ListIndex) & "', cod_deductora = " & cboDeductora.ItemData(cboDeductora.ListIndex) _
       & ", PLAN_INDICADOR = " & chkPlanAhorros.Value & ", PLAN_FIJAR_MENSUALIDAD = " & chkPlanFijarMensualidad.Value _
       & ", PLAN_COD_PLAN = '" & Trim(txtPlan.Text) & "', PLAN_MENSUALIDAD = " & CCur(txtPlanMensualidad.Text) _
       & " where COD_CONTRATO = " & Operacion.Operacion
Call ConectionExecute(strSQL)


'Cambia Los Procesos Anteriores x StoreProcedures
strSQL = "exec spCrd_Revo_Contrato_Formaliza " & Operacion.Operacion & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'BITACORA
Call Bitacora("Registra", "Contrato Revolutivo No.: " & Operacion.Operacion)

'Imprime Boleta de Formalizacion
'Call sbCrdSGTBoletaFormaliza(Operacion.Operacion)

Me.MousePointer = vbDefault

MsgBox "Contrato Revolutivo Activado Satisfactoriamente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAnular()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vProcedimiento As Integer, vMensaje As String
Dim curInteres As Currency, curAmortiza As Currency
Dim rs2 As New ADODB.Recordset, vND As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

vMensaje = ""

strSQL = "exec spCRDFormalizaAnulacion " & Operacion.Operacion & ",'" & glogon.Usuario & "',1"
Call OpenRecordSet(rs, strSQL)
  vMensaje = rs!Mensaje
rs.Close

'BITACORA
Call Bitacora("Registra", "Anulación de la OP: " & Operacion.Operacion)
Call sbBitacoraCredito("13", "Monto : " & txtMonto.Text, "C", txtOperacion.Text, txtCodigo.Text, "SGT Anula Formalizacion del Día")
''Tags de Seguimiento (Se Aplica desde el Procedure.)
'Call sbCrdOperacionTags(Operacion.Operacion, Operacion.Codigo, "S09", "", "SGT Anula Formalizacion del Día")

vMensaje = vMensaje & vbCrLf & "...Anulación Realizada Satisfactoriamente..."

Me.MousePointer = vbDefault
If GLOBALES.SysDocVersion = 2 Then
    Call sbImprimeRecibo(Operacion.Operacion, "AFR")
End If

If Len(vMensaje) > 0 Then MsgBox vMensaje, vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnFormalizar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


If Me.optFormalizacion(0).Value Then
 

 If fxVerificaFormalizacion Then
    
     i = MsgBox("Esta seguro que desea >> formalizar << esta Operación", vbYesNo)
     If i = vbYes Then
         Call sbFormalizar
     End If
     
 Else 'Falla Verificacion de Formalizacion
  MsgBox vMensaje, vbCritical
 End If

Else 'Anulacion de la formalizacion
'  If fxVerificaAnulacion Then
'     i = MsgBox("Esta seguro que desea >> Anular << esta Operación", vbYesNo)
'     If i = vbYes Then
'        Call sbAnular
'     End If
'
'  Else
'    MsgBox vMensaje, vbCritical
'  End If
End If

Call sbContrato_Load

End Sub




Private Sub dtpFechaFormalizacion_Change()
Dim strSQL As String

'strSQL = "update reg_creditos set fechaforp = '" & Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd") & "'" _
'       & " where id_solicitud = " & Operacion.Operacion
'
'Call ConectionExecute(strSQL)
'
'Call Bitacora("Modifica", "Fecha Formalizacion Operacion " & Operacion.Operacion)

End Sub

Private Sub dtpFechaSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then txtMonto.SetFocus
End Sub


Private Function fxVerificaExisteCodigo(strCodigo As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
strSQL = "select isnull(count(*),0) as Existe from catalogo where codigo ='" & strCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaExisteCodigo = IIf((rsX!Existe > 0), True, False)
rsX.Close
End Function

Private Function fxVerificaExisteRangoCodigo(strCodigo As String, curMonto As Currency) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
 
strSQL = "select isnull(count(*),0) as Existe from rangos"
strSQL = strSQL & " where codigo ='" & strCodigo & "' and " & curMonto & " >=de and " _
        & curMonto & " <=  hasta"
Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaExisteRangoCodigo = IIf((rsX!Existe > 0), True, False)
rsX.Close
End Function


Private Function fxVerificaRecepcion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String, vPermiteCbrJud As Boolean
Dim Porcentaje As Currency

fxVerificaRecepcion = True
vMensaje = ""
vPermiteCbrJud = False


If Operacion.Operacion = 0 Or fxEstadoOperacion(cboEstado.Text) = "R" Then
        
        If IsNumeric(txtPlazo) Then
         If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado NO es válido"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
        End If
        
        If IsNumeric(txtTasa) Then
         If txtTasa < 0 Then vMensaje = vMensaje & vbCrLf & "- La Tasa solicitada no es válida"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Interés Solicitado es Inválido"
        End If
        
        If IsNumeric(txtMonto.Text) Then
         If txtMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado NO es válido"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado es Inválido"
        End If
        
        'Verifica Rangos
        If Len(vMensaje) = 0 Then
          strSQL = "exec spCrdFormaliza_Valida_Rangos '" & txtCedula.Text & "','" & txtCodigo.Text & "'," _
                 & CCur(txtMonto.Text) & "," & CCur(txtTasa) & "," & CInt(txtPlazo.Text) _
                 & ",'" & cboDestino.ItemData(cboDestino.ListIndex) & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) _
                 & "'," & Operacion.Operacion
          Call OpenRecordSet(rsX, strSQL)
          If Len(rsX!Mensaje) > 0 Then
              vMensaje = vMensaje & vbCrLf & rsX!Mensaje
          End If
          rsX.Close
        End If


End If 'Recepcion


'VERIFICAR SI TIENE CODIFICACION CONTABLE
'Update_2017m02d22: Simplifica los datos de respuesta y agrega la validacion del Nueva operacion para personas en Cobro Judicial

strSQL = "select ctaNintC,retencion,Poliza, activo, isnull(Permite_PersonaEnCbrJud,0) as 'Permite_Cbr' " _
       & " from catalogo where codigo ='" & txtCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
 If rsX.EOF And rsX.BOF Then
   vMensaje = vMensaje & vbCrLf & "- El código de préstamo no existe"
 Else
 
  If rsX!Permite_Cbr = 1 Then
      vPermiteCbrJud = True
  End If
  
  'Verifica si el codigo tiene codificacion contable
  'Es suficiente con evaluar cualquiera de las 9, pues el sistema
  'solo permite actualizar cuando se especifican todas.
   If IsNull(rsX!ctaNintC) Then vMensaje = vMensaje & vbCrLf & "- El código no se encuentra codificado contablemente"
   
   'No se permiten retenciones ni polizas
   If rsX!retencion = "S" Or rsX!Poliza = "S" Then vMensaje = vMensaje & vbCrLf & "- No se permite guardar porque el código pertenece a una Retencion o Poliza"
  
   'Verificar que el Codigo se encuentre Activo
   If rsX!activo = 0 Then vMensaje = vMensaje & vbCrLf & "- La Línea de Crédito no se encuentra Activa..."
  
 End If
rsX.Close

'Verifica el estado de la persona vrs los estados permitidos en esta línea de crédito
'Update_2017m02d22: Mejora Consulta
strSQL = "select isnull(count(*),0) as 'Existe'" _
       & " from CRD_CATALOGO_ESTADOS E inner join Socios S on E.cod_Estado = S.EstadoActual and S.cedula = '" & txtCedula.Text _
       & "' where codigo = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite El estado actual de la persona (verifique.!)"
rsX.Close

'VERIFICAR COMBOS
If fxCodigoDestino(cboDestino.ItemData(cboDestino.ListIndex), txtCodigo) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Destino No es válido para Esta Línea"
If fxCodigoComite(cboComite.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Comité Especificado NO EXISTE"
If fxEstadoOperacion(cboEstado.Text) = "" Then vMensaje = vMensaje & vbCrLf & "- El Estado de la Operación NO ES VALIDO"
If cboGarantia.ItemData(cboGarantia.ListIndex) = "" Then vMensaje = vMensaje & vbCrLf & "- La Garantía especificada NO ES VALIDA"


'Verificar que la persona no tenga prestamos en Cobro Judicial Activos
If Not vPermiteCbrJud Then
    strSQL = "select isnull(count(*),0) as Existe from reg_creditos" _
           & " where estado = 'A' and proceso = 'J' and cedula = '" & txtCedula & "'"
    Call OpenRecordSet(rsX, strSQL, 0)
    If rsX!Existe > 0 Then vMensaje = vMensaje & vbCrLf & "- La persona tiene créditos en Cobro Judicial"
    rsX.Close
End If


If Len(vMensaje) > 0 Then fxVerificaRecepcion = False

End Function

Private Function fxVerificaNivel()
Dim rsX As New ADODB.Recordset, rsX2 As New ADODB.Recordset, strSQL As String

strSQL = "select count(*) as Existe from nivel_miembros A, nivel_derechos B where A.nv_cod_grupo = " _
       & "B.nv_cod_grupo and nombre = '" & glogon.Usuario & "' and codigo = '" _
       & txtCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
rsX.Close

End Function

Private Function fxVerificaFormalizacion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Currency, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim curDisponible As Currency, curGiros As Currency
Dim curMontoTmp As Currency, vPriDeducCorte As Long

vMensaje = ""
fxVerificaFormalizacion = True


vFecha = fxFechaServidor
If chkDeducPlanilla.Value = vbChecked Then
        strSQL = "select MAX(proceso) as 'Proceso' From PRM_BITACORA" _
               & " where COD_INSTITUCION = " & cboDeductora.ItemData(cboDeductora.ListIndex) _
               & "  and GESTION = 'E' and TRANSACCION = '02'"
        Call OpenRecordSet(rsX, strSQL, 0)
        If IsNull(rsX!Proceso) Then
           vPriDeducCorte = GLOBALES.glngFechaCR
        Else
           vPriDeducCorte = rsX!Proceso
        End If
        rsX.Close
Else
           vPriDeducCorte = GLOBALES.glngFechaCR
End If


If DateDiff("d", Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd"), Format(dtpDesembolso.Value, "yyyy/mm/dd")) < 0 Then vMensaje = vMensaje & vbCrLf & "- La fecha del desembolsos no puede ser menor que la fecha de formalizacion"


If (Operacion.Estado = "A" Or Operacion.Estado = "C") And Me.optFormalizacion(0).Value = True _
    Then vMensaje = vMensaje & vbCrLf & "- Esta Operación ya fue procesada"


If IsNumeric(txtAno) Then
  If txtAno < Year(vFecha) Then vMensaje = vMensaje & vbCrLf & "- El Año especificado no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- El Año para la primer deduccion no es válido"
End If

If fxConvierteMES(cboMes.Text) = cboMes.Text Then vMensaje = vMensaje & vbCrLf & "- El Mes para la primer deduccion no es válido"

lngPriDeduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

If chkDeducPlanilla.Value = vbChecked Then
    If lngPriDeduc <= vPriDeducCorte Then vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque es igual o menor a la fecha de proceso actual"
Else
    If lngPriDeduc < vPriDeducCorte Then vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque menor a la fecha de proceso actual"
End If

If Month(dtpFechaFormalizacion.Value) <> Month(vFecha) Or Year(dtpFechaFormalizacion.Value) <> Year(vFecha) Then
 'Actualiza la fecha de formalizacion
' strSQL = "update reg_creditos set fechaforp = '" & Format(vFecha, "yyyy/mm/dd") _
'        & "' where id_solicitud = " & Operacion.Operacion
' Call ConectionExecute(strSQL)
' dtpFechaFormalizacion.Value = vFecha
End If


'Verifica Rangos
If Len(vMensaje) = 0 Then
  strSQL = "exec spCrdFormaliza_Valida_Rangos '" & txtCedula.Text & "','" & txtCodigo.Text & "'," _
         & CCur(txtMonto.Text) & "," & CCur(txtTasa) & "," & CInt(txtPlazo.Text) _
         & ",'" & cboDestino.ItemData(cboDestino.ListIndex) & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) _
         & "'," & Operacion.Operacion
  Call OpenRecordSet(rsX, strSQL)
  If Len(rsX!Mensaje) > 0 Then
      vMensaje = vMensaje & vbCrLf & rsX!Mensaje
  End If
  rsX.Close
End If

'
' STORE PROCEDURE - DE VERIFICACION DE FORMALIZACIONES
'

'strSQL = "exec spCRDFormalizaValidacion " & Operacion.Operacion & ", '" & glogon.Usuario & "'"
'Call OpenRecordSet(rsX, strSQL, 0)
'    If rsX!nivel = 0 Then vMensaje = vMensaje & vbCrLf & "- No existe nivel de formalización de este usuario para la línea " & txtCodigo
'    If rsX!refundicion = 0 Then vMensaje = vMensaje & vbCrLf & "- El saldo a refundir vario en la operación : " & Operacion.Operacion
'    If rsX!bloqueo = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta persona se encuentra bloqueda, hasta mañana se le podran formalizar operaciones..."
'    If rsX!GarAhorro = 0 Then vMensaje = vMensaje & vbCrLf & " - El Monto aprobado excede el porcentaje aprobado de sus ahorros"
'    If rsX!MaxOperaciones = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el número máximo de operaciones en esta linea"
'    If rsX!MaxLinea = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el monto máximo de la línea"
'    If rsX!MaxGarantia = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el monto maximo de la línea x Garantía"
'    If rsX!MaxGarantiaTotal = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el monto máximo x Garantía"
'
'    If rsX!Firmas = 0 Then vMensaje = vMensaje & vbCrLf & "- No se han registrado todas las firmas..."
'    If rsX!LineaActiva = 0 Then vMensaje = vMensaje & vbCrLf & "- La línea de crédito no se encuentra Activa..."
'    If rsX!DestinoActivo = 0 Then vMensaje = vMensaje & vbCrLf & "- El destino del crédito no se encuentra Activo..."
'    If rsX!COBERTURA = 0 Then vMensaje = vMensaje & vbCrLf & "- Cobertura de las hipotecas es inferior al monto del crédito..."
'    If rsX!EstadoPersona = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite El estado actual de la persona (verifique.!)"
'    If rsX!CongeladoCredito = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite El estado actual de la persona (verifique.!)"
'    If rsX!Requisitos = 0 Then vMensaje = vMensaje & vbCrLf & "- No se cumplieron los requisitos Obligatorios (verifique.!)"
'    If rsX!BaseCalculo = 0 Then vMensaje = vMensaje & vbCrLf & "- No se ha establecido la Base de Calculo para Cuota Balloon!"
'
'rsX.Close
'
''Verificar la posición de cada Operacion a refundir
'strSQL = "exec spCrdSGTRefundicionesValida " & Operacion.Operacion
'Call OpenRecordSet(rsX, strSQL, 0)
'    If rsX!Cambios > 0 Then vMensaje = vMensaje & vbCrLf & "- " & rsX!Cambios & " Operación a Refundir a Cambiado su Estado ---> Actualice!"
'rsX.Close
'

' hasta aqui el codigo en la base de datos
'


If Operacion.EstadoSolicitud <> "R" Then
     vMensaje = vMensaje & vbCrLf & "- Esta solicitud no se encuentra recibida..."
End If



If Len(vMensaje) = 0 Then
    curDisponible = 0
    strSQL = "exec spCRDDisponibleRecurso '" & cboRecursos.ItemData(cboRecursos.ListIndex) & "','" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'"
    Call OpenRecordSet(rsX, strSQL, 0)
    If Not rsX.EOF And Not rsX.BOF Then
        curDisponible = IIf(IsNull(rsX!Disponible), 0, rsX!Disponible)
    End If
    rsX.Close

    curGiros = CCur(txtMonto.Text)
    If curGiros > 0 Then
        If curDisponible < curGiros Then
           vMensaje = vMensaje & vbCrLf & " - No Hay disponible en el Recurso, para desembolsar esta Operación..."
           vMensaje = vMensaje & vbCrLf & " - Monto a Girar : " & Format(curGiros, "Standard") & " - Disponible :  " & Format(curDisponible, "Standard")
           vMensaje = vMensaje & vbCrLf & " - Monto Faltante para Girar: " & Format(curGiros - curDisponible, "Standard")
        End If
    End If
End If

If Len(vMensaje) > 0 Then fxVerificaFormalizacion = False


End Function

Private Function fxVerificaAnulacion() As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim rsTmp As New ADODB.Recordset

vMensaje = ""
fxVerificaAnulacion = True


If Operacion.EstadoSolicitud <> "F" Then
  vMensaje = vMensaje & vbCrLf & "- Esta Operación no ha sido formalizada! Utilice el Estado de DENEGADA!"
  If Len(vMensaje) > 0 Then fxVerificaAnulacion = False
  Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe" _
   & " from NIVEL_GRUPOS N INNER JOIN nivel_miembros A" _
   & " ON N.NV_COD_GRUPO = A.NV_COD_GRUPO INNER JOIN nivel_derechos B" _
   & " ON N.NV_COD_GRUPO = B.NV_COD_GRUPO Where A.nombre = '" & glogon.Usuario _
   & "' and B.codigo = '" & txtCodigo & "' AND N.nv_tipo = 'N'" _
   & " and (" & CCur(txtMonto.Text) & " between nv_desde and nv_hasta)"

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  vMensaje = vMensaje & vbCrLf & "- No existe nivel de anulación de este usuario para la línea.: " & txtCodigo
End If
rs.Close


''0. Verificacion base / Solo se pueden anular las formalizaciones del día
''Cambiado el: 22/2/2018 Para que sea en el mismo mes
''strSQL = "select fechaforp,datediff(d,fechaforp,dbo.MyGetdate()) as Resultado"
'strSQL = "select fechaforp,  month(fechaforp) - month(dbo.Mygetdate()) + ( year(fechaforp) - year(dbo.Mygetdate()) ) as Resultado" _
'       & " from reg_creditos where id_solicitud = " & Operacion.Operacion
'Call OpenRecordSet(rs, strSQL)
'    vFecha = rs!FechaForp
'    If Abs(rs!Resultado) > 0 Then
'      vMensaje = vMensaje & vbCrLf & "- Esta operación fue formalizada en un mes diferente..."
'    End If
'rs.Close
'
''2. Verifica que no se le registren desembolsos, Se deben de anular o eliminar
'strSQL = "select isnull(count(*),0) as Existe from Tes_Transacciones where op = " & Operacion.Operacion _
'       & " and estado <> 'A'"
'Call OpenRecordSet(rs, strSQL)
'If rs!Existe > 0 Then
'  vMensaje = vMensaje & vbCrLf & "- Existen solicitudes o documentos emitidos (Cheques/Transferencias) en Tesorería (Proceda a Anularlos)"
'End If
'rs.Close


'4. No puede anular retenciones
strSQL = "select retencion from catalogo where codigo = '" & Operacion.Codigo & "'"
Call OpenRecordSet(rs, strSQL)
If rs!retencion = "S" Then
   vMensaje = vMensaje & vbCrLf & "- Este es un código de retención No se puede Anular..."
End If
rs.Close


If Len(vMensaje) > 0 Then fxVerificaAnulacion = False


End Function


Private Sub sbConsultaX(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

'strSQL = "Select R.id_solicitud,R.codigo,R.cedula,S.Nombre,R.fechasol,R.montosol,R.estadosol,R.estado,R.proceso" _
'       & " FROM REG_CREDITOS R inner join CATALOGO C ON R.CODIGO = C.CODIGO" _
'       & " inner join Socios S on R.cedula = S.cedula" _
'       & " where C.retencion = 'N' and C.poliza = 'N' and R.cedula like '%" & Trim(txtConCedula.Text) _
'       & "%' and S.nombre like '%" & Trim(txtConNombre.Text) & "%'" _
'       & " order by R.id_solicitud desc"
'
'lswBusca.ListItems.Clear
'
'Call OpenRecordSet(rs, strSQL, 0)
'
'Do While Not rs.EOF
' Set itmX = lswBusca.ListItems.Add(, , CStr(rs!id_solicitud))
'  itmX.SubItems(1) = rs!Codigo
'  itmX.SubItems(2) = rs!Cedula
'  itmX.SubItems(3) = rs!Nombre
'
'  itmX.SubItems(4) = Format(rs!FechaSol, "yyyy/mm/dd")
'  itmX.SubItems(5) = Format(rs!montosol, "Standard")
'
'  Select Case rs!estadosol
'   Case "R"
'    itmX.SubItems(6) = "Recibida"
'   Case "P"
'    itmX.SubItems(6) = "Pendiente"
'   Case "A"
'    itmX.SubItems(6) = "Aprobada"
'   Case "D"
'    itmX.SubItems(6) = "Denegada"
'   Case "F"
'    itmX.SubItems(6) = "Formalizada"
'   Case "N"
'    itmX.SubItems(6) = "Anulada"
'  End Select
'
' Select Case rs!Estado
'   Case "A"
'    itmX.SubItems(7) = "Activa"
'   Case "C"
'    itmX.SubItems(7) = "Cancelada"
'   Case Else
'    itmX.SubItems(7) = "En Tramite"
' End Select
'
' Select Case rs!Proceso
'   Case "J"
'    itmX.SubItems(8) = "Cobro Jud"
'   Case "N"
'    itmX.SubItems(8) = "Normal"
'   Case "T"
'    itmX.SubItems(8) = "Traspaso"
'   Case Else
'    itmX.SubItems(8) = "------"
' End Select
'
' rs.MoveNext
'Loop
'rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Public Sub sbGXSegTraIniTlb()
If TimerX.Interval > 0 Then
   Call TimerX_Timer
End If
 Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(1))
 txtCedula = GLOBALES.gCedulaActual
 txtCedula_LostFocus
 txtCodigo.SetFocus

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBar.Value = 1 And txtOperacion.Text = "" Then txtOperacion.Text = "0"
If FlatScrollBar.Value = 0 And txtOperacion.Text = "" Then txtOperacion.Text = "999999999999"

If vScroll Then
    strSQL = "select Top 1 R.cod_contrato from CRD_REV_CONTRATOS R inner join Catalogo C on R.codigo = C.codigo" _
           & " and C.Retencion = 'N' and C.poliza = 'N'"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where R.cod_contrato > " & txtOperacion & " order by R.cod_contrato asc"
    Else
       strSQL = strSQL & " where R.cod_contrato < " & txtOperacion & " order by R.cod_contrato desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtOperacion = rs!COD_CONTRATO
      Call sbContrato_Load
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
 
 vModulo = 3

'Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture
 
 mFrecuenciaPago = "M"

 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 Call sbToolBarIconos(tlbPrincipal, False)
  
 
 With cboMes
    .Clear
    .AddItem "Enero"
    .AddItem "Febrero"
    .AddItem "Marzo"
    .AddItem "Abril"
    .AddItem "Mayo"
    .AddItem "Junio"
    .AddItem "Julio"
    .AddItem "Agosto"
    .AddItem "Septiembre"
    .AddItem "Octubre"
    .AddItem "Noviembre"
    .AddItem "Diciembre"
 End With

 With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With


'Inicializa
tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaDatos()
 Dim i As Integer


'If txtCedula.Text = "" And Operacion.Operacion = 0 And tcMain.Item(0).Selected Then Exit Sub

With Operacion
 .Operacion = 0
 .Cedula = ""
 .Codigo = ""
 .EstadoSolicitud = "R"
 .Documento = ""
 .TasaPtsBono = 0
 .PlazoBono = 0
End With

 tcMain.Item(0).Selected = True
 

 txtCedula = ""
 txtCodigo = ""
 
 txtCuota = ""
 txtCuota = ""
 txtDescripcion = ""
 txtTasa = ""
 txtNombre = ""
 lblNombre.Caption = txtNombre.Text
 txtNotas = ""
 txtPlazo = "1"
 txtMonto = ""
 txtPromotorId.Text = ""
 txtPromotorNombre.Text = ""
  
txtDisponible.Text = "0.00"
txtTransitoDebitos.Text = "0.00"
txtTransitoCreditos.Text = "0.00"
 
txtPPlan.Text = ""
txtPContrato.Text = ""
txtPAportes.Text = "0.00"
txtPMensualidad.Text = "0.00"
 
txtPlan.Text = ""
txtPlanDesc.Text = ""
txtPlanMensualidad.Text = "0.00"
 
chkAjustaCuotasAlVencimiento.Value = xtpUnchecked
chkTopeDisponible.Value = xtpChecked
chkSupervision.Value = xtpUnchecked
chkDeducPlanilla.Value = xtpChecked
 
txtOperacionCrd.Text = ""
txtEstudio.Text = ""
txtSaldo.Text = "0.00"
txtCuotaCorte.Text = ""
 
  
 
imgEstado.ToolTipText = "Nueva Operación!"
Set imgEstado.Picture = imgIconosEstados.ListImages.Item(3).Picture

 
 dtpFechaFormalizacion.Value = fxFechaServidor
 dtpFechaSolicitud.Value = dtpFechaFormalizacion.Value
 dtpVencimiento.Value = DateAdd("m", txtPlazo.Text, dtpFechaFormalizacion.Value)
 dtpDesembolso.Value = dtpFechaFormalizacion.Value
 
 cboEstado.Clear
 cboGarantia.Clear
 cboDestino.Clear

 
' For i = 0 To btnOpciones.Count - 1
'   btnOpciones.Item(i).Enabled = False
' Next i
 
End Sub

Private Sub sbCargaCombos()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "Select id_comite as 'IdX',descripcion as 'ItmX' from comites where estado = 1"
Call sbCbo_Llena_New(cboComite, strSQL, False, True)

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub CargaRecursos(cbo As Object, vCodigo As String, vGrupo As String)
Dim strSQL As String

On Error GoTo vError

strSQL = " select rtrim(G.cod_grupo) as 'IdX', rtrim(G.descripcion) as 'ItmX'" _
       & " from catalogo_grupos G inner join catalogo_asignaGrp A on G.cod_grupo = A.cod_grupo" _
       & " where G.estado = 1 and A.codigo = '" & vCodigo & "'"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub ActivaDesActiva(vEstadoSolicitud As String, vEstadoEC As String)
'Activa e Inactiva informacion, en los tabs

optFormalizacion(0).Enabled = True
optFormalizacion(0).Value = True
btnFormalizar.Enabled = True


If vEstadoEC = "N" Then
    Select Case UCase(vEstadoSolicitud)
     Case "R"
        tcMain.Item(0).Selected = True
        
     Case "P"
        tcMain.Item(0).Selected = True
     Case "A"
        tcMain.Item(1).Selected = True
     Case "D"
        tcMain.Item(0).Selected = True
     Case "F"
       tcMain.Item(1).Selected = True
       optFormalizacion(1).Value = True
       optFormalizacion(0).Enabled = False
     Case "N"
       tcMain.Item(1).Selected = True
        
       btnFormalizar.Enabled = False
    End Select

Else
    tcMain.Item(1).Selected = True
    
    Select Case UCase(vEstadoSolicitud)
      Case "F"
            optFormalizacion(1).Value = True
            optFormalizacion(0).Enabled = False
      Case "N"
            btnFormalizar.Enabled = False
     End Select
End If

End Sub


Private Function fxOperacionDestino(vDestino As String) As String
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select rtrim(cod_destino) + ' - ' + descripcion as ItemX from catalogo_destinos where cod_destino = '" & vDestino & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxOperacionDestino = " -"
Else
  fxOperacionDestino = rs!itemx
End If
rs.Close

End Function



Private Sub sbContrato_Load()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vFecha As Date, iMes As Integer, lngAnio As Long, vProceso As Currency, pProcesoClean As Long
Dim i As Integer, vTemp As String, dFecha As Date

On Error Resume Next

' For i = 0 To 13
'   btnOpciones.Item(i).Enabled = True
' Next i

vOperacionLoad = True
tcMain.Item(0).Selected = True


strSQL = "exec spCrd_Revo_Contrato_Consulta " & txtOperacion.Text

Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  
' Call sbCargaCombos
 vFecha = rs!Fecha_Server

 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 txtCodigo.Text = rs!Codigo
 txtDivisa.Text = rs!COD_DIVISA & ""
 
 lblNombre.Caption = txtNombre.Text
 
 mFrecuenciaPago = "M"
 
 If rs!BaseCalculo = "06" Then
    mFrecuenciaPago = "Q"
 End If
 
 
 Operacion.Operacion = rs!COD_CONTRATO
 Operacion.Cedula = rs!Cedula
 Operacion.Nombre = txtNombre
 Operacion.Estado = rs!Estado
 Operacion.Codigo = rs!Codigo
 Operacion.Estado = IIf(IsNull(rs!Estado), "N", rs!Estado)
 Operacion.MontoAprobado = IIf(IsNull(rs!Monto), 0, rs!Monto)
  Operacion.TasaPtsBono = 0
 
 txtTasa.ToolTipText = "Pts Bonificación: " & 0
 
' MsgBox Operacion.TS & vbCrLf & rs!TS
 txtDescripcion.Text = rs!LineaDesc
 
 
 txtCuota.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 txtPlazo.Text = CStr(IIf(IsNull(rs!Plazo), 0, rs!Plazo))
 txtTasa.Text = CStr(IIf(IsNull(rs!Tasa), 0, rs!Tasa))
 
 txtPromotorId.Text = rs!ID_PROMOTOR
 txtPromotorNombre.Text = rs!Ejecutivo
 
 txtMonto.Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
 
 txtNotas.Text = IIf(IsNull(rs!NOTAS), "", rs!NOTAS)

 
 dtpFechaFormalizacion.Value = IIf(IsNull(rs!FORMALIZA_FECHA), vFecha, rs!FORMALIZA_FECHA)
 dtpDesembolso.Value = IIf(IsNull(rs!FORMALIZA_FECHA), vFecha, rs!FORMALIZA_FECHA)
 dtpFechaSolicitud.Value = IIf(IsNull(rs!REGISTRO_FECHA), vFecha, rs!REGISTRO_FECHA)
 
 Call sbCboAsignaDato(cboComite, Trim(rs!Comdesc) & "", True, rs!id_Comite)
  
 'Carga Destino
 Call sbSTCargaCboDestinos(cboDestino, Operacion.Codigo)
 Call sbCboAsignaDato(cboDestino, rs!DestinoDesc, True, rs!cod_destino & "")
  

 'Si no tiene el banco asignado hay que crearlo pero no puede guardarlo
 'bajo este mismo banco hasta que lo tenga asignado o lo cambie.
 
 'Carga Deductoras por Institucion
 vPaso = True
     Call sbDeductoras_Load(rs!cod_institucion)
     Call sbCboAsignaDato(cboDeductora, rs!DeductoraDesc, True, rs!cod_Deductora)
    
     cboDeductora.Tag = CStr(rs!cod_Deductora)
 vPaso = False
 
 Call cboDeductora_Click

 If IsNull(rs!Primer_Deduccion) Then
'    vProceso = fxPrimerDeduccion(rs!id_solicitud, rs!cod_Deductora)
        vProceso = fxPrimerDeduccionCuota(Operacion.Codigo)

        pProcesoClean = vProceso

        cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
        txtAno.Text = Mid(pProcesoClean, 1, 4)

        Select Case (vProceso - pProcesoClean)
          Case 0
              cboFrecuencia.Text = "Mensual"
          Case 0.1
              cboFrecuencia.Text = "1er Quincena"
          Case 0.2
              cboFrecuencia.Text = "2da Quincena"
        End Select

 Else
    vProceso = rs!Primer_Deduccion
 
    pProcesoClean = vProceso
    
    cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
    txtAno.Text = Mid(pProcesoClean, 1, 4)
    
    Select Case (vProceso - pProcesoClean)
      Case 0
          cboFrecuencia.Text = "Mensual"
      Case 0.1
          cboFrecuencia.Text = "1er Quincena"
      Case 0.2
          cboFrecuencia.Text = "2da Quincena"
    End Select
 
 End If
 
 Call sbSTCargaCboEstado(cboEstado, rs!Estado)
 
 Call CargaRecursos(cboRecursos, rs!Codigo, rs!COD_RECURSO & "")
 Call sbCboAsignaDato(cboRecursos, rs!RecursoDesc, False, rs!COD_RECURSO & "")
 
 vPaso = True
        Call sbSTCargaCboGarantia(cboGarantia, rs!Codigo)
        Call sbCboAsignaDato(cboGarantia, rs!GarantiaDesc, False, rs!Garantia)
 vPaso = False
 
 
txtDisponible.Text = Format(rs!Disponible, "Standard")
txtTransitoDebitos.Text = Format(rs!TRANSITO_DEBITOS, "Standard")
txtTransitoCreditos.Text = Format(rs!TRANSITO_CREDITOS, "Standard")
 
txtPPlan.Text = rs!PLAN_COD_PLAN & ""
txtPContrato.Text = rs!PLAN_COD_CONTRATO & ""
txtPAportes.Text = Format(rs!PlanAportes, "Standard")
txtPMensualidad.Text = Format(rs!PlanMensualidadActual, "Standard")
 
txtPlan.Text = rs!PLAN_COD_PLAN & ""
txtPlanDesc.Text = rs!PlanDesc
txtPlanMensualidad.Text = Format(rs!PLAN_MENSUALIDAD, "Standard")
 
chkAjustaCuotasAlVencimiento.Value = rs!PLAZO_AJUSTE
chkTopeDisponible.Value = rs!TOPE_RETIRO_IND
chkSupervision.Value = rs!REQUIERE_SUPERVISION
chkDeducPlanilla.Value = rs!ind_deduce_planilla
 
txtOperacionCrd.Text = rs!Operacion & ""
txtEstudio.Text = rs!cod_PreAnalisis & ""
txtSaldo.Text = Format(rs!Saldo, "Standard")
txtCuotaCorte.Text = rs!CORTE_ULTIMO & ""
 
 Call ActivaDesActiva(rs!Estado, IIf(IsNull(rs!Estado), "N", rs!Estado))

  Select Case rs!Estado
    Case "R"
       dFecha = Format(rs!REGISTRO_FECHA & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Solicitado por " & rs!REGISTRO_USUARIO & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(3).Picture
       
    Case "P" 'Pendiente (HOLD)
       dFecha = Format(rs!REGISTRO_FECHA & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Pendiente " & rs!REGISTRO_USUARIO & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(6).Picture

    Case "D" 'Denegado
       dFecha = Format(rs!REGISTRO_FECHA & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Denegado " & rs!REGISTRO_USUARIO & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(5).Picture

    Case "F" 'Formalizado
       dFecha = Format(IIf(IsNull(rs!REGISTRO_FECHA), rs!FORMALIZA_FECHA & "", rs!REGISTRO_FECHA), "dd/mm/yyyy")
       imgEstado.ToolTipText = "Formalizado: " & rs!FORMALIZA_USUARIO & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(1).Picture
       
    Case "N" 'Anulado
       dFecha = Format(rs!REGISTRO_FECHA & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Anulado por " & vbCrLf & rs!REGISTRO_USUARIO & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(2).Picture
  End Select





 With tlbPrincipal.Buttons
   .Item(1).Enabled = True
   .Item(2).Enabled = True
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
 Me.fraOperacion.Enabled = False


Else
 MsgBox "No existe esta Solicitud", vbCritical
End If
rs.Close

vOperacionLoad = False


Call RefrescaTags(Me)

End Sub

Private Sub sbBusqueda(Index As Integer)
'Set GLOBALES.gfrmFormulario = Me
gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Index
  Case 0 'txtOperacion
    gBusquedas.Convertir = "S"
    gBusquedas.Consulta = "select COD_CONTRATO,codigo,cedula,Monto, Estado from CRD_REV_CONTRATOS"
    gBusquedas.Orden = "COD_CONTRATO"
    gBusquedas.Columna = "COD_CONTRATO"
    frmBusquedas.Show vbModal
    txtOperacion = gBusquedas.Resultado
    If Len(Trim(txtOperacion)) > 0 Then
    '  Call ConsultaOperacion
    End If
  
  Case 1 'txtCedula
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
  
  Case 2 'txtCodigo
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        gBusquedas.Filtro = " and Activo = 1 and Retencion = 'N' and REVOLUTIVA = 1"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  
  Case 3 'txtNombre
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
  
  Case 4 'txtDescripcion
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  
End Select

End Sub


Private Sub imgRecalculoRecurso_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass



strSQL = "exec spCRDDisponibleRecurso '" & cboRecursos.ItemData(cboRecursos.ListIndex) & "','" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtDisponibleRecursos = Format(rs!Disponible, "Standard")
Else
    txtDisponibleRecursos = 0
End If
rs.Close

Me.MousePointer = vbDefault

End Sub


Private Function fxInteresDiasX()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iDias As Integer



If Not fxCobraTasaFormaliza(cboDestino.ItemData(cboDestino.ListIndex)) Then
   fxInteresDiasX = 0
   Exit Function
End If

strSQL = "select R.FECHA_CALCULO_INT,isnull(R.FECHA_INICIO_CALCULO, R.fecha_Calculo_Int) as 'FECHA_INICIO_CALCULO',C.convenio,C.retencion,C.poliza,R.montoApr,R.int" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " where id_solicitud = " & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)
    If rs!fecha_calculo_int < rs!fecha_inicio_calculo Then
     iDias = 0
    Else
     iDias = rs!fecha_calculo_int - rs!fecha_inicio_calculo + 1
    End If
    
    If rs!Convenio = "S" Or rs!retencion = "S" Or rs!Poliza = "S" Then
      fxInteresDiasX = 0
    Else
      fxInteresDiasX = ((rs!montoapr * rs!Int) / (36000)) * iDias
    End If

rs.Close

End Function


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
 
 Call sbLimpiaDatos
 Call sbCargaCombos

End Sub



Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn And IsNumeric(txtMonto.Text) Then

        If Operacion.PlazoBono = 0 Then
            txtPlazo.Text = fxCatalogoRango(txtCodigo, txtMonto.Text, "P", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex))
        End If
        txtTasa.Text = fxCatalogoRango(txtCodigo, txtMonto.Text, "I", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex)) - Operacion.TasaPtsBono
        txtPlazo.SetFocus
 End If
End Sub


Private Sub Edicion(intActiva As Integer)
'Activa e inactiva partes a editar

If intActiva = 1 Then
  fraOperacion.Enabled = True
  Select Case Operacion.EstadoSolicitud
   Case "R", "P"
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
     Me.txtMonto.Enabled = True
     Me.txtPlazo.Enabled = True
     Me.txtTasa.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.cboEstado.Enabled = True
     Me.txtNotas.Enabled = True
     Me.dtpFechaSolicitud.Enabled = True

   Case "A"
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.txtNotas.Enabled = True
     
     Me.dtpFechaSolicitud.Enabled = False
     Me.txtMonto.Enabled = False
     Me.txtPlazo.Enabled = False
     Me.txtTasa.Enabled = False
     Me.cboEstado.Enabled = False
   Case "D", "N"
     Me.txtNotas.Enabled = True
     Me.txtCedula.Enabled = False
     Me.txtCodigo.Enabled = False
     Me.cboComite.Enabled = False
     Me.cboGarantia.Enabled = False
     Me.dtpFechaSolicitud.Enabled = False
     Me.txtMonto.Enabled = False
     Me.txtPlazo.Enabled = False
     Me.txtTasa.Enabled = False
     Me.cboEstado.Enabled = False
   Case "F"
     Me.txtNotas.Enabled = True
     Me.txtCedula.Enabled = False
     Me.txtCodigo.Enabled = False
     Me.cboComite.Enabled = False
     Me.cboGarantia.Enabled = False
     Me.dtpFechaSolicitud.Enabled = False
     Me.txtMonto.Enabled = False
     Me.txtPlazo.Enabled = False
     Me.txtTasa.Enabled = False
     Me.cboEstado.Enabled = False
  End Select
Else 'apaga
  fraOperacion.Enabled = False
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
     Me.txtMonto.Enabled = True
     Me.txtPlazo.Enabled = True
     Me.txtTasa.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.cboEstado.Enabled = True
     Me.txtNotas.Enabled = True
     Me.dtpFechaSolicitud.Enabled = True
  Select Case Operacion.EstadoSolicitud
   Case "A"
   Case "D"
   Case "N"
   Case "F"
  End Select
End If 'inactiva
End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xGarantiaFND As String, xGarantiaFNDContrato As Long
Dim vOperacionTemporal As Long, vActividad As String, vDestino As String, vCanal As String


On Error GoTo vError

vDestino = cboDestino.ItemData(cboDestino.ListIndex)

If vEdita Then

'spCrd_Revo_Contrato_Add(@Contrato int, @Cedula varchar(20), @Codigo varchar(10), @Divisa varchar(10), @Destino varchar(10), @Garantia varchar(10)
'        , @Estado char(1), @Monto dec(16,2), @Plazo int, @PlazoAjuste smallint, @Tasa dec(8,4), @Cuota dec(16,2), @Notas varchar(500)
'        , @PromotorId int, @ComiteId int, @TopeIndica smallint, @TopeMonto dec(16,2), @SupervisaIndica smallint
'        , @Vence datetime = Null, @PlanIndica smallint, @Plan varchar(10), @PlanMensualidad dec(14,2), @PlanFijaMensualidad smallint
'        , @Usuario varchar(30))
        
   strSQL = "exec spCrd_Revo_Contrato_Add " & txtOperacion.Text & ",'" & Trim(txtCedula.Text) & "', '" & Trim(txtCodigo.Text) & "','" & Trim(txtDivisa.Text) _
           & "','" & vDestino & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "','" _
           & fxEstadoOperacion(cboEstado.Text) & "', " & CCur(txtMonto.Text) & ", " & txtPlazo.Text & ", " & chkAjustaCuotasAlVencimiento.Value _
           & ", " & CCur(txtTasa.Text) & ", " & CCur(txtCuota.Text) & ", '" & txtNotas.Text & "', " & txtPromotorId.Text _
           & ", " & cboComite.ItemData(cboComite.ListIndex) & ", " & chkTopeDisponible.Value & ", " & CCur(txtMonto.Text) & ", " & chkSupervision.Value _
           & ", '" & Format(dtpVencimiento.Value, "yyyy/mm/dd") & "', " & chkPlanAhorros.Value & ", '" & Trim(txtPlan.Text) & "', " & CCur(txtPlanMensualidad.Text) _
           & ", " & chkPlanFijarMensualidad.Value & ", '" & glogon.Usuario & "'"
     Call ConectionExecute(strSQL)
        

     Call Bitacora("Actualiza", "Contrato Revolutivo No.: " & txtOperacion.Text)

Else 'Inserta
  

      
   strSQL = "exec spCrd_Revo_Contrato_Add 0, '" & Trim(txtCedula.Text) & "', '" & Trim(txtCodigo.Text) & "','" & Trim(txtDivisa.Text) _
           & "','" & vDestino & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "','" _
           & fxEstadoOperacion(cboEstado.Text) & "', " & CCur(txtMonto.Text) & ", " & txtPlazo.Text & ", " & chkAjustaCuotasAlVencimiento.Value _
           & ", " & CCur(txtTasa.Text) & ", " & CCur(txtCuota.Text) & ", '" & txtNotas.Text & "', " & txtPromotorId.Text _
           & ", " & cboComite.ItemData(cboComite.ListIndex) & ", " & chkTopeDisponible.Value & ", " & CCur(txtMonto.Text) & ", " & chkSupervision.Value _
           & ", '" & Format(dtpVencimiento.Value, "yyyy/mm/dd") & "', " & chkPlanAhorros.Value & ", '" & Trim(txtPlan.Text) & "', " & CCur(txtPlanMensualidad.Text) _
           & ", " & chkPlanFijarMensualidad.Value & ", '" & glogon.Usuario & "'"
    
    Call OpenRecordSet(rs, strSQL)
        txtOperacion.Text = rs!Contrato
    rs.Close
    
     Call Bitacora("Registra", "Contrato Revolutivo No.: " & txtOperacion.Text)
 
End If
 
MsgBox "Solicitud Grabada Satisfactoriamente...", vbInformation
Exit Sub


vError:



End Sub



Private Sub txtMonto_LostFocus()

On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:


End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
 Case "nuevo"
  txtOperacion.Text = ""
  txtOperacion.Enabled = False
  Call sbLimpiaDatos
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = False
  tlbPrincipal.Buttons(3).Enabled = True
  tlbPrincipal.Buttons(4).Enabled = True
  fraOperacion.Enabled = True
  txtNotas.Locked = False
  
  vEdita = False
  
  
  txtCedula.SetFocus
'  Call sbCargaCombos
  
  
  
 Case "editar"
  If Operacion.Operacion > 0 Then 'And Operacion.Estado = "A" Then
      vEdita = True
      Call Edicion(1)
    
      'Si el Estado Esta en Recepcion o Resolucion puede Cambiar Todos Los Datos
      'Si Esta en Formalización Solo puede Cambiar la Salida
      tlbPrincipal.Buttons(1).Enabled = False
      tlbPrincipal.Buttons(2).Enabled = False
      tlbPrincipal.Buttons(3).Enabled = True
      tlbPrincipal.Buttons(4).Enabled = True
      txtOperacion.Enabled = False
      fraOperacion.Enabled = True
      txtNotas.Locked = False


      txtCedula.SetFocus

  End If
 
 Case "guardar"
  
  If fxVerificaRecepcion Then

    Call sbGuardar
    Call Edicion(0)
    Call sbContrato_Load
    txtOperacion.Enabled = True
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    
    fraOperacion.Enabled = False
    txtNotas.Locked = True

    
    If vEdita = False Then
        tcMain.Item(0).Selected = True
        'Datos Personales
    '    Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)
    End If
    
    If vEdita = False And cboGarantia.ItemData(cboGarantia.ListIndex) = "F" Then
        tcMain.Item(0).Selected = True
        'Fiadores
        'Call btnOpciones_Click(1)
     '   Call sbTaskPanel_Accion(Id_TaskItem_Garantia)
    
    End If
    
    If vEdita = False Then
        'Requisitos
     '   Call sbTaskPanel_Accion(Id_TaskItem_Requisitos)
    
    End If
    
    If Operacion.EstadoSolicitud = "P" Or Operacion.EstadoSolicitud = "D" Then
        'Siempre verifica las causas, por si esta en Pendiente o Denegada
     '    Call sbTaskPanel_Accion(Id_TaskItem_Causas)
    End If
  
  Else
    MsgBox vMensaje, vbCritical
  End If
 
 Case "deshacer"
    txtOperacion.Enabled = True
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    fraOperacion.Enabled = False
    If txtOperacion <> "" Then Call sbContrato_Load
    txtOperacion.SetFocus
 
 Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
 
 Case "cerrar"
  Unload Me

End Select


End Sub



Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtCodigo.SetFocus
End Sub

Private Sub sbDeductoras_Load(pInstitucion As Long)
Dim strSQL As String

strSQL = "select COD_DEDUCTORA AS 'IdX', DESCRIPCION AS 'ItmX'" _
       & " From vAFI_Deductoras" _
       & " Where cod_institucion = " & pInstitucion

Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)

End Sub


Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select S.nombre, isnull(I.DEDUCCION_PLANILLA,0) as 'Deduccion' " _
       & ",S.cod_institucion, Ed.Cod_Institucion as 'DeductoraCod', Ed.Descripcion as 'DeductoraDesc'" _
       & " from Socios S inner join Instituciones I on S.cod_institucion = I.cod_Institucion" _
       & " left join Instituciones Ed on isnull(S.cod_deductora,S.cod_institucion) = Ed.cod_Institucion" _
       & " Where S.cedula = '" & txtCedula.Text & "'"
       
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
    txtNombre.Text = ""
    lblNombre.Caption = ""
    chkDeducPlanilla.Value = vbUnchecked
    chkDeducPlanilla.Enabled = False
Else
    txtNombre.Text = Trim(rs!Nombre)
    lblNombre.Caption = Trim(rs!Nombre)
    

    'Carga Deductoras por Institucion
    Call sbDeductoras_Load(rs!cod_institucion)
    Call sbCboAsignaDato(cboDeductora, rs!DeductoraDesc, True, rs!DeductoraCod)

    cboDeductora.Tag = CStr(rs!DeductoraCod)
    
    
    If rs!Deduccion = 0 Then
        chkDeducPlanilla.Value = vbUnchecked
        chkDeducPlanilla.Enabled = False

    Else
        chkDeducPlanilla.Value = vbChecked
        chkDeducPlanilla.Enabled = True
    End If

End If
rs.Close

 
End Sub

Private Sub DescribeCodigoComite()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

strSQL = "select isnull(id_comite,0) as id_comite from catalogo where codigo='" & txtCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  cboComite.Text = fxDescribeComite(rs!id_Comite)
End If
rs.Close

vError:

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = vbKeyReturn Then
  txtCodigo = UCase(txtCodigo)
  
  vPaso = True
        Call sbSTCargaCboGarantia(cboGarantia, txtCodigo)
        Call sbSTCargaCboEstado(cboEstado, "R")
        Call sbSTCargaCboDestinos(cboDestino, txtCodigo)

  vPaso = False
  
  
  If fxCreditoExcedente(txtCodigo) Then
    txtMonto.Text = fxExcedenteDisponible(txtCedula)
  End If
  
  cboGarantia.SetFocus

End If

End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

On Error GoTo vError

txtDivisa.Text = ""
txtDescripcion.Text = ""

strSQL = "select Cat.CODIGO, Cat.DESCRIPCION, Cat.MONEDA as 'COD_DIVISA', CAT.Base_Calculo" _
       & " , Cat.ID_COMITE, isnull(Com.DESCRIPCION,'') as 'COMITE_DESC'" _
       & " from CATALOGO Cat left join COMITES Com on Cat.ID_COMITE = Com.ID_COMITE" _
       & " where Cat.CODIGO = '" & txtCodigo.Text & "'"




mFrecuenciaPago = "M"

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.BOF Then
    txtDescripcion.Text = rs!Descripcion & ""
    txtDivisa.Text = rs!COD_DIVISA
    Call sbCboAsignaDato(cboComite, rs!Comite_Desc, True, rs!id_Comite)
    If rs!Base_Calculo = "06" Then
        mFrecuenciaPago = "Q"
    End If
End If
rs.Close

'strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "','" & txtDivisa.Text & "'"
'Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
End Sub


Private Sub txtPromotorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "ID_PROMOTOR"
   gBusquedas.Orden = "ID_PROMOTOR"
   gBusquedas.Consulta = "select ID_PROMOTOR as 'Id.',Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorId.Text = Trim(gBusquedas.Resultado)
   txtPromotorNombre.Text = Trim(gBusquedas.Resultado2)
End If
End Sub



Private Sub txtPromotorNombre_GotFocus()

If txtPromotorNombre.Text = "" Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select ID_PROMOTOR as 'Id.' ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorId.Text = Trim(gBusquedas.Resultado)
   txtPromotorNombre.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtPromotorNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select ID_PROMOTOR as 'Id.' ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorId.Text = Trim(gBusquedas.Resultado)
   txtPromotorNombre.Text = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtTasa_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If
vError:

End Sub

Private Sub txtMonto_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) > 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
 txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If

vError:
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  On Error Resume Next
    If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
        And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
      txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
    End If
   cboComite.SetFocus
End If
End Sub

Private Sub txtTasa_LostFocus()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
End Sub



Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 If TimerX.Interval > 0 Then
    Call TimerX_Timer
 End If
 Call txtOperacion_KeyDown(vbKeyReturn, 0)
End Sub


Private Sub txtOperacion_Change()
 Call sbLimpiaDatos
  With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbContrato_Load
If KeyCode = vbKeyF4 Then Call sbBusqueda(0)
End Sub

Private Sub txtPlazo_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota.Text = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If

vError:
End Sub


Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim x As Double

If KeyCode = vbKeyReturn And txtPlazo.Text <> "" Then
    

       x = fxCatalogoRangoPlz(txtCodigo, txtPlazo, cboDestino.ItemData(cboDestino.ListIndex))
       If x > 0 Then txtTasa.Text = x - Operacion.TasaPtsBono

End If
 
End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtTasa.SetFocus
End Sub

