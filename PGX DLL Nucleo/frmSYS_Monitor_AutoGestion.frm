VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmSYS_Monitor_AutoGestion 
   Caption         =   "Tramites de Auto Gestión"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   16425
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3495
      Left            =   3240
      TabIndex        =   15
      Top             =   5640
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   6165
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
      Item(0).Caption =   "Caso"
      Item(0).ControlCount=   29
      Item(0).Control(0)=   "Label2(0)"
      Item(0).Control(1)=   "Label2(1)"
      Item(0).Control(2)=   "Label2(3)"
      Item(0).Control(3)=   "txtCuota"
      Item(0).Control(4)=   "txtMonto"
      Item(0).Control(5)=   "txtTasa"
      Item(0).Control(6)=   "txtPlazo"
      Item(0).Control(7)=   "Label1(1)"
      Item(0).Control(8)=   "LblTasa"
      Item(0).Control(9)=   "LblPlazo"
      Item(0).Control(10)=   "txtTramite"
      Item(0).Control(11)=   "txtEstado"
      Item(0).Control(12)=   "txtCedula"
      Item(0).Control(13)=   "txtNombre"
      Item(0).Control(14)=   "Label3(0)"
      Item(0).Control(15)=   "Label2(2)"
      Item(0).Control(16)=   "txtResolucionUsuario"
      Item(0).Control(17)=   "txtResolucionFecha"
      Item(0).Control(18)=   "txtGarantía"
      Item(0).Control(19)=   "txtRegistraUsuario"
      Item(0).Control(20)=   "txtRegistraFecha"
      Item(0).Control(21)=   "Label2(4)"
      Item(0).Control(22)=   "Label3(1)"
      Item(0).Control(23)=   "txtDescripcion"
      Item(0).Control(24)=   "txtCodigo"
      Item(0).Control(25)=   "Label2(8)"
      Item(0).Control(26)=   "txtOperacion"
      Item(0).Control(27)=   "chkRefunde"
      Item(0).Control(28)=   "btnBoleta"
      Item(1).Caption =   "Adjuntos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsw"
      Item(2).Caption =   "Resolución"
      Item(2).ControlCount=   6
      Item(2).Control(0)=   "txtNotas"
      Item(2).Control(1)=   "cboResolucion"
      Item(2).Control(2)=   "Label2(6)"
      Item(2).Control(3)=   "btnResolucion"
      Item(2).Control(4)=   "gbAprobacion"
      Item(2).Control(5)=   "lblResolucion(5)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3132
         Left            =   -70000
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   10812
         _Version        =   1441793
         _ExtentX        =   19071
         _ExtentY        =   5524
         _StockProps     =   77
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBoleta 
         Height          =   375
         Left            =   11520
         TabIndex        =   69
         ToolTipText     =   "Boleta de Formalización"
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmSYS_Monitor_AutoGestion.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkRefunde 
         Height          =   255
         Left            =   9840
         TabIndex        =   68
         Top             =   1200
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Refunde ?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox gbAprobacion 
         Height          =   972
         Left            =   -68320
         TabIndex        =   49
         Top             =   2160
         Visible         =   0   'False
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Aprobación:"
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
         Begin XtremeSuiteControls.RadioButton rbAccion 
            Height          =   612
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Crear Solicitud en Tramite de Créditos"
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbAccion 
            Height          =   612
            Index           =   1
            Left            =   2520
            TabIndex        =   51
            Top             =   240
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Formalización del Crédito"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   1560
         TabIndex        =   45
         Top             =   1440
         Width           =   2052
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuota 
         Height          =   312
         Left            =   1560
         TabIndex        =   20
         Top             =   3120
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   1560
         TabIndex        =   21
         Top             =   1920
         Width           =   2052
         _Version        =   1441793
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   312
         Left            =   2520
         TabIndex        =   22
         Top             =   2640
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   312
         Left            =   2520
         TabIndex        =   23
         Top             =   2280
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtTramite 
         Height          =   504
         Left            =   1560
         TabIndex        =   27
         Top             =   480
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   889
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   504
         Left            =   3600
         TabIndex        =   28
         Top             =   480
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   889
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   3600
         TabIndex        =   30
         Top             =   1080
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   1560
         TabIndex        =   29
         Top             =   1080
         Width           =   2052
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtResolucionUsuario 
         Height          =   312
         Left            =   7080
         TabIndex        =   33
         Top             =   2760
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
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
      Begin XtremeSuiteControls.FlatEdit txtResolucionFecha 
         Height          =   312
         Left            =   7080
         TabIndex        =   34
         Top             =   3120
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
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
      Begin XtremeSuiteControls.ComboBox cboResolucion 
         Height          =   312
         Left            =   -68320
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   3132
         _Version        =   1441793
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   612
         Left            =   -61960
         TabIndex        =   38
         Top             =   2400
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar Resolución"
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
         Picture         =   "frmSYS_Monitor_AutoGestion.frx":0707
         ImageAlignment  =   0
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   912
         Left            =   -68320
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   8652
         _Version        =   1441793
         _ExtentX        =   15261
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGarantía 
         Height          =   504
         Left            =   5760
         TabIndex        =   39
         Top             =   480
         Width           =   3612
         _Version        =   1441793
         _ExtentX        =   6371
         _ExtentY        =   889
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
      Begin XtremeSuiteControls.FlatEdit txtRegistraUsuario 
         Height          =   312
         Left            =   7080
         TabIndex        =   40
         Top             =   1920
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
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
      Begin XtremeSuiteControls.FlatEdit txtRegistraFecha 
         Height          =   312
         Left            =   7080
         TabIndex        =   41
         Top             =   2280
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   3600
         TabIndex        =   44
         Top             =   1440
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   504
         Left            =   9360
         TabIndex        =   65
         ToolTipText     =   "No. Operación de Crédito Relacionada"
         Top             =   480
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   889
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Crédito"
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Registro"
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
         Left            =   5160
         TabIndex        =   43
         Top             =   2280
         Width           =   1932
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   4
         Left            =   5160
         TabIndex        =   42
         Top             =   1920
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Registro por:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   6
         Left            =   -69760
         TabIndex        =   37
         Top             =   1080
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
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
      Begin XtremeSuiteControls.Label lblResolucion 
         Height          =   252
         Index           =   5
         Left            =   -69760
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Resolución"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   2
         Left            =   5160
         TabIndex        =   32
         Top             =   2760
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Resuelto por:"
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Resolución"
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
         Left            =   5160
         TabIndex        =   31
         Top             =   3120
         Width           =   1932
      End
      Begin VB.Label LblPlazo 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   972
      End
      Begin VB.Label LblTasa 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   972
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   612
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Monto"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "No.Trámite"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   3000
      Top             =   120
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   12015
      _Version        =   524288
      _ExtentX        =   21193
      _ExtentY        =   8705
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
      MaxCols         =   15
      SpreadDesigner  =   "frmSYS_Monitor_AutoGestion.frx":0F0C
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmSYS_Monitor_AutoGestion.frx":185A
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmSYS_Monitor_AutoGestion.frx":1F5A
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboFecha 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.FlatEdit FlatEdit_Linea 
      Height          =   336
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   593
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit_Cedula 
      Height          =   336
      Left            =   1200
      TabIndex        =   47
      Top             =   240
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   593
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEA_Casos 
      Height          =   315
      Left            =   480
      TabIndex        =   57
      Top             =   5280
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtEA_Monto 
      Height          =   315
      Left            =   1080
      TabIndex        =   58
      Top             =   5280
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtED_Casos 
      Height          =   315
      Left            =   480
      TabIndex        =   59
      Top             =   5880
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtED_Monto 
      Height          =   315
      Left            =   1080
      TabIndex        =   60
      Top             =   5880
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtEP_Casos 
      Height          =   315
      Left            =   480
      TabIndex        =   61
      Top             =   6480
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtEP_Monto 
      Height          =   315
      Left            =   1080
      TabIndex        =   62
      Top             =   6480
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtES_Casos 
      Height          =   315
      Left            =   480
      TabIndex        =   63
      Top             =   7080
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtES_Monto 
      Height          =   315
      Left            =   1080
      TabIndex        =   64
      Top             =   7080
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox cboTramite 
      Height          =   315
      Left            =   1200
      TabIndex        =   66
      Top             =   1200
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
      Appearance      =   7
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tramite Credito:"
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
      Height          =   555
      Index           =   4
      Left            =   360
      TabIndex        =   67
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitadas:"
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
      Index           =   12
      Left            =   240
      TabIndex        =   56
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pendientes:"
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
      Index           =   11
      Left            =   240
      TabIndex        =   55
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Denegadas:"
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
      Index           =   10
      Left            =   240
      TabIndex        =   54
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Aprobadas:"
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
      Index           =   7
      Left            =   240
      TabIndex        =   53
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen de Estado de Solicitudes:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   52
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula:"
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
      Index           =   2
      Left            =   360
      TabIndex        =   48
      Top             =   240
      Width           =   972
   End
   Begin XtremeShortcutBar.ShortcutCaption scCaso 
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   5160
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Selecciones un Caso!"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
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
      Index           =   8
      Left            =   600
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Index           =   9
      Left            =   600
      TabIndex        =   10
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
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
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas:"
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
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image imgMainBanner 
      Height          =   9276
      Left            =   0
      Picture         =   "frmSYS_Monitor_AutoGestion.frx":282B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3204
   End
End
Attribute VB_Name = "frmSYS_Monitor_AutoGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vPaso As Boolean, mFecUltMovUpdate As Integer

Private Sub btnBoleta_Click()

If Not IsNumeric(txtOperacion.Text) Then Exit Sub



strSQL = "select estadoSol from reg_creditos where id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.EOF Then
    If rs!EstadoSol = "F" Then
        Call sbBoleta_Formaliza(txtOperacion.Text)
    Else
        MsgBox "Esta Operación No. " & txtOperacion.Text & ", no se encuentra formalizada!", vbExclamation
    End If
End If

End Sub

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Call sbExportar
End Sub



Private Sub sbExportar()
Dim vHeaders As vGridHeaders

    vHeaders.Columnas = 15
    vHeaders.Headers(1) = "Tramite No."
    vHeaders.Headers(2) = "Estado"
    vHeaders.Headers(3) = "Identificación"
    vHeaders.Headers(4) = "Nombre"
    vHeaders.Headers(5) = "Linea Credito"
    vHeaders.Headers(6) = "Monto"
    vHeaders.Headers(7) = "Plazo"
    vHeaders.Headers(8) = "Tasa"
    vHeaders.Headers(9) = "Cuota"
    vHeaders.Headers(10) = "Garantía"
    vHeaders.Headers(11) = "Fecha Registro"
    vHeaders.Headers(12) = "Fecha Resolución"
    vHeaders.Headers(13) = "SGT Tipo"
    vHeaders.Headers(14) = "SGT Codigo"
    
    Call sbSIFGridExportar(vGrid, vHeaders, "Solicitudes_AutoGestion_Consulta")


End Sub

Private Sub sbFiltro_Aplica(ByRef pSQL As String)
Dim pWhere As Boolean

pWhere = False

If cboEstado.Text <> "Todos" Then
   pSQL = pSQL & " Where Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
   pWhere = True
End If

If UCase(cboTramite.Text) <> UCase("Todos") Then
      If pWhere Then
        pSQL = pSQL & " and TRAMITE_ESTADO_ID = '" & Mid(cboTramite.Text, 1, 1) & "'"
      Else
        pSQL = pSQL & " Where TRAMITE_ESTADO_ID = '" & Mid(cboTramite.Text, 1, 1) & "'"
      End If
   pWhere = True
End If



Select Case cboFecha.Text
  Case "Registro"
      If pWhere Then
            pSQL = pSQL & " and Registro_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
      Else
            pSQL = pSQL & " Where Registro_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
            
            pWhere = True
      End If
  
  Case "Resolución"
      If pWhere Then
            pSQL = pSQL & " and Res_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
      Else
            pSQL = pSQL & " Where Res_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

            pWhere = True
      End If

End Select

If Trim(FlatEdit_Linea.Text) <> "" Then
   If pWhere Then
        pSQL = pSQL & " and Codigo = '" & FlatEdit_Linea.Text & "'"
   Else
        pSQL = pSQL & " Where Codigo = '" & FlatEdit_Linea.Text & "'"
        
        pWhere = True
        
   End If
End If

If Trim(FlatEdit_Cedula.Text) <> "" Then
   If pWhere Then
        pSQL = pSQL & " and Cedula = '" & FlatEdit_Cedula.Text & "'"
   Else
        pSQL = pSQL & " Where Cedula = '" & FlatEdit_Cedula.Text & "'"
        
        pWhere = True
        
   End If
End If



End Sub

Private Sub sbBuscar()
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select '' as 'Btn', COD_SOLICITUD, estado_desc, cedula, nombre, Linea_Desc" _
       & ", monto, plazo, tasa, cuota, GARANTIA_DESC, REGISTRO_FECHA, RES_FECHA" _
       & ", RES_CODIGO, TRAMITE_ESTADO_DESC" _
       & " From vCrd_Solicitudes_AutoGestion "
Call sbFiltro_Aplica(strSQL)

vPaso = True

Call sbCargaGrid(vGrid, 15, strSQL, True)
If vGrid.MaxRows > 1 Then
    vGrid.MaxRows = vGrid.MaxRows - 1
End If

vPaso = False

' Proceso Informacion de Resumen
txtEP_Casos.Text = Format(0, "###,###0")
txtEP_Monto.Text = Format(0, "Standard")

txtES_Casos.Text = Format(0, "###,###0")
txtES_Monto.Text = Format(0, "Standard")

txtEA_Casos.Text = Format(0, "###,###0")
txtEA_Monto.Text = Format(0, "Standard")

txtED_Casos.Text = Format(0, "###,###0")
txtED_Monto.Text = Format(0, "Standard")


strSQL = "exec spCrd_Solicitudes_AutoGestion_Rsm '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
       & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Select Case rs!Estado
    Case "P"
        txtEP_Casos.Text = Format(rs!Casos, "###,###0")
        txtEP_Monto.Text = Format(rs!Monto, "Standard")
    
    Case "S"
        txtES_Casos.Text = Format(rs!Casos, "###,###0")
        txtES_Monto.Text = Format(rs!Monto, "Standard")
    
    Case "A"
        txtEA_Casos.Text = Format(rs!Casos, "###,###0")
        txtEA_Monto.Text = Format(rs!Monto, "Standard")
    
    Case "D"
        txtED_Casos.Text = Format(rs!Casos, "###,###0")
        txtED_Monto.Text = Format(rs!Monto, "Standard")

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



Private Sub sbLimpia()

scCaso.Tag = 0
scCaso.Caption = "Seleccione un Caso!"

tcMain.Item(0).Selected = True

txtTramite.Text = ""
txtEstado.Text = ""
txtEstado.Tag = ""

txtGarantía.Text = ""

txtCedula.Text = ""
txtNombre.Text = ""

txtCodigo.Text = ""
txtDescripcion.Text = ""


txtMonto.Text = Format(0, "Standard")
txtPlazo.Text = CStr(1)
txtTasa.Text = Format(0, "Standard")

txtCuota.Text = Format(0, "Standard")

txtRegistraFecha.Text = ""
txtRegistraUsuario.Text = ""

txtResolucionFecha.Text = ""
txtResolucionUsuario.Text = ""


End Sub

Private Sub sbCaso_Consulta(pCaso As Long)
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * " _
       & " From vCrd_Solicitudes_AutoGestion " _
       & " Where Cod_Solicitud = " & pCaso
Call OpenRecordSet(rs, strSQL)

scCaso.Tag = rs!Cod_Solicitud

scCaso.Caption = "Id: " & rs!Cod_Solicitud & "  Cédula: " & rs!Cedula & "  " & rs!Nombre

tcMain.Item(0).Selected = True

txtTramite.Text = rs!Cod_Solicitud
txtEstado.Text = rs!Estado_Desc
txtEstado.Tag = rs!Estado

txtGarantía.Text = rs!Garantia_Desc

txtCedula.Text = rs!Cedula
txtNombre.Text = rs!Nombre

txtCodigo.Text = rs!Codigo
txtDescripcion.Text = rs!Linea_Desc


txtMonto.Text = Format(rs!Monto, "Standard")
txtPlazo.Text = CStr(rs!Plazo)
txtTasa.Text = Format(rs!Tasa, "Standard")

txtCuota.Text = Format(rs!Cuota, "Standard")

txtRegistraFecha.Text = rs!registro_Fecha & ""
txtRegistraUsuario.Text = Trim(rs!Registro_Usuario & "")

txtResolucionFecha.Text = rs!Res_Fecha & ""
txtResolucionUsuario.Text = Trim(rs!Res_Usuario & "")


txtOperacion.Text = rs!Res_Codigo & ""
txtNotas.Text = rs!Notas & ""


chkRefunde.Value = rs!Refunde_ind

rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnResolucion_Click()
Dim pGestion As String

On Error GoTo vError

Me.MousePointer = vbHourglass

pGestion = IIf(rbAccion.Item(0).Value, "S", "F")


strSQL = "exec spCrd_Solicitudes_AutoGestion_Resolucion " & scCaso.Tag & ",'" & Mid(cboResolucion.Text, 1, 1) _
            & "','" & txtNotas.Text & "','" & glogon.Usuario _
            & "','" & pGestion & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Call sbCaso_Consulta(scCaso.Tag)

If Mid(cboResolucion.Text, 1, 1) = "A" And pGestion = "S" Then
   MsgBox "Se Generó la Solicitud de Crédito No. " & txtOperacion.Text, vbInformation
End If

If Mid(cboResolucion.Text, 1, 1) = "A" And pGestion = "F" Then
   Call sbBoleta_Formaliza(txtOperacion.Text)
   
   MsgBox "Se Generó la Operacion de Crédito No. " & txtOperacion.Text, vbInformation
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub


Private Sub sbBoleta_Formaliza(pOperacion As Long)
Dim strRuta As String

Me.MousePointer = vbHourglass

strRuta = SIFGlobal.fxPathReportes("Credito_BoletaFormalizacion.rpt")

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Boleta de Formalización"
 .ReportFileName = strRuta
 
 .Connect = glogon.ConectRPT
 
 .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD}=" & pOperacion
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(3) = "fxSolicitudBarras='*" & pOperacion & "*'"

 .SubreportToChange = "sbAsiento"
 .StoredProcParam(0) = "FRM"
 .StoredProcParam(1) = pOperacion
 .StoredProcParam(2) = 0

 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub cboFecha_Click()
If vPaso Then Exit Sub

If cboFecha.Text = "Todas" Then
   dtpInicio.Enabled = False
   dtpCorte.Enabled = False
Else
   dtpInicio.Enabled = True
   dtpCorte.Enabled = True
End If

End Sub




Private Sub cboResolucion_Click()

btnResolucion.Visible = True

Select Case Mid(cboResolucion.Text, 1, 1)
    Case "P", "R", "S"
        gbAprobacion.Visible = False
 
    Case "A"
        gbAprobacion.Visible = True
    
    Case "D"
        gbAprobacion.Visible = False
        
End Select

End Sub

Private Sub FlatEdit_Cedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterna"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    FlatEdit_Cedula.Text = Trim(gBusquedas.Resultado)
End If

End Sub

Private Sub Form_Activate()
vModulo = 3

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Archivo Id", 1200
lsw.ColumnHeaders.Add , , "Tipo Adjunto", 2200
lsw.ColumnHeaders.Add , , "Nombre Archivo", 5200
lsw.ColumnHeaders.Add , , "Extensión", 1400, vbCenter




End Sub



Private Sub Form_Resize()
Dim pHeight As Long, pWidth As Long

On Error Resume Next

imgMainBanner.Height = Me.Height

pHeight = 8712
pWidth = 12264


If Me.Height < pHeight Then
   Me.Height = pHeight
End If

If Me.Width < pWidth Then
   Me.Width = pWidth
End If

vGrid.Width = Me.Width - (vGrid.Left + 160)
vGrid.Height = Me.Height - (vGrid.Top + scCaso.Height + tcMain.Height + 750)

scCaso.Top = vGrid.Top + vGrid.Height + 150
tcMain.Top = scCaso.Top + scCaso.Height + 50

scCaso.Width = vGrid.Width
tcMain.Width = vGrid.Width

lsw.Width = tcMain.Width

cboTramite.AddItem "Recibida"
cboTramite.AddItem "Formalizada"
cboTramite.AddItem "Anulada"
cboTramite.AddItem "Denegada"
cboTramite.AddItem "Pendiente"
cboTramite.AddItem "TODOS"
cboTramite.Text = "TODOS"


End Sub

Private Sub FlatEdit_Linea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select CODIGO,DESCRIPCION From CATALOGO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = " and LINEA_INTERNA = 1 AND RETENCION = 'N' AND POLIZA = 'N'" _
                      & " and WEBSITE = 1"
    frmBusquedas.Show vbModal
    FlatEdit_Linea.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Linea.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub


Private Sub sbConsultaAdjunto()

If scCaso.Tag = "" Or scCaso.Tag = "0" Then
    MsgBox "Consulte un Caso!", vbInformation
    tcMain.Item(0).Selected = True
    Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "SELECT A.ARCHIVO_ID, At.DESCRIPCION as 'TIPO_ADJUNTO', A.ARCHIVO_NOMBRE, A.ARCHIVO_TIPO  " _
       & " FROM CRD_SOLICITUDES_ADJUNTOS A INNER JOIN CRD_ADJUNTOS_TIPOS At on A.COD_ADJUNTO = At.COD_ADJUNTO" _
       & " WHERE A.TRANSAC_TIPO = 'SOL' AND A.TRANSAC_CODIGO = " & scCaso.Tag

lsw.ListItems.Clear

'Carga Resultados
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!ARCHIVO_ID)
      itmX.SubItems(1) = rs!TIPO_ADJUNTO
      itmX.SubItems(2) = rs!ARCHIVO_NOMBRE
      itmX.SubItems(3) = rs!ARCHIVO_TIPO
  rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub
  
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lsw_DblClick()
Dim sql As String, Campo_Imagen As String
Dim rs As New ADODB.Recordset, Stream As New ADODB.Stream
Dim pPath As String

Dim vArchivo As String

If lsw.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

Set itmX = lsw.SelectedItem

'  Set itmX = lsw.ListItems.Add(, , rs!ARCHIVO_ID)
'      itmX.SubItems(1) = rs!TIPO_ADJUNTO
'      itmX.SubItems(2) = rs!ARCHIVO_NOMBRE
'      itmX.SubItems(3) = rs!ARCHIVO_TIPO

vArchivo = "AG_" & Format(scCaso.Tag, "00000") & "_" & txtCedula & "_" _
          & itmX.SubItems(1) & "." & Replace(itmX.SubItems(3), "application/", "")


pPath = SIFGlobal.DirectorioDeResultados & "\Adjuntos\" & vArchivo

'------------------------------------------------------------------------
'---Crea el Archivo

  

  
Campo_Imagen = "ARCHIVO_BIT"
  
sql = "select * from CRD_SOLICITUDES_ADJUNTOS" _
       & " Where Archivo_Id = " & itmX.Text


rs.Open sql, glogon.Conection, adOpenKeyset, adLockOptimistic
   
' Si no hay registros sale de la función y retorna como _
 resultado un valor Nothing, es decir ninguna imagen

If rs.RecordCount = 0 Then
   Exit Sub
End If
   
' Especifica el tipo de datos ( binario )
Stream.Type = adTypeBinary
Stream.Open
   
' verifica con la función IsNull que el campo no tenga _
 un valor Nulo ya que si no da error, en ese caso sale de la función
If IsNull(rs.Fields(Campo_Imagen).Value) Then
    GoTo vError
End If
' Graba los datos en el objeto stream
Stream.Write rs.Fields(Campo_Imagen).Value
   
' este método graba un  archivo temporal  en disco _
 ( en el pPath que luego se elimina )
Stream.SaveToFile pPath, adSaveCreateOverWrite
   
'Cierra el recordset y el objeto Stream
If rs.State = adStateOpen Then
    rs.Close
End If
If Not rs Is Nothing Then
    Set rs = Nothing
End If
   
If Stream.State = adStateOpen Then
    Stream.Close
End If
If Not Stream Is Nothing Then
    Set Stream = Nothing
End If

MsgBox "Adjunto guardado en: " & pPath, vbInformation

'Abre el Archivo
Call Shell("Explorer.exe /e," & pPath, vbNormalFocus)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
 
 On Error Resume Next
 'Si no abre el archivo automáticamente, entonces abre el directorio
 Call Shell("Explorer.exe /select," & pPath, vbNormalFocus)

End Sub

Private Sub sbResolucion_Load()

cboResolucion.Text = txtEstado.Text

Select Case Mid(txtEstado.Text, 1, 1)
    Case "R", "P", "S"
        cboResolucion.Enabled = True
        btnResolucion.Visible = True
        gbAprobacion.Visible = False
        
    Case "A", "D"
        cboResolucion.Enabled = False
        btnResolucion.Visible = False
        gbAprobacion.Visible = False

End Select

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0
    
    Case 1 'Adjuntos
     Call sbConsultaAdjunto

    Case 2 'Resolucion
     Call sbResolucion_Load
End Select

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

On Error GoTo vError

vPaso = True

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -15, dtpCorte.Value)

cboFecha.Clear
cboFecha.AddItem "Registro"
cboFecha.AddItem "Resolución"
cboFecha.AddItem "Todas"
cboFecha.Text = "Registro"

cboEstado.Clear
cboEstado.AddItem "Todos"
cboEstado.AddItem "Solicitado"
cboEstado.AddItem "Pendiente"
cboEstado.AddItem "Aprobado"
cboEstado.AddItem "Denegado"

cboEstado.Text = "Pendiente"


cboResolucion.Clear
cboResolucion.AddItem "Solicitado"
cboResolucion.AddItem "Pendiente"
cboResolucion.AddItem "Aprobado"
cboResolucion.AddItem "Denegado"

cboResolucion.Text = "Solicitado"

vPaso = False

Call cboFecha_Click

Call Formularios(Me)
Call RefrescaTags(Me)

'Fix Temporal
strSQL = "exec spCrd_Solicitudes_Adjuntos_Fix"
Call ConectionExecute(strSQL)

Call sbBuscar

Exit Sub

vError:

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim vTramite As Long

On Error GoTo vError

vGrid.Row = Row
vGrid.Col = 2

vTramite = vGrid.Text

scCaso.Tag = ""
scCaso.Caption = "Indique un Caso!"

vGrid.Col = 5
scCaso.Tag = vTramite
scCaso.Caption = "No. " & vTramite & " ¦ " & vGrid.Text

Call sbCaso_Consulta(vTramite)

Exit Sub

vError:

End Sub


