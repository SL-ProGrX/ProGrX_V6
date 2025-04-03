VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_APA_Reportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Adminsitración de Garantías"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5175
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   9128
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
      Item(0).Caption =   "Informes"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "ArbolExp"
      Item(0).Control(1)=   "cboEstado"
      Item(0).Control(2)=   "lblReporte"
      Item(0).Control(3)=   "Label4(0)"
      Item(0).Control(4)=   "cboTipo"
      Item(0).Control(5)=   "Label4(1)"
      Item(0).Control(6)=   "chkAcreedores"
      Item(0).Control(7)=   "chkOperaciones"
      Item(0).Control(8)=   "chkOperacionesConSaldo"
      Item(0).Control(9)=   "cboTipoVencimiento"
      Item(0).Control(10)=   "Label4(2)"
      Item(0).Control(11)=   "Label4(3)"
      Item(0).Control(12)=   "dtpInicio"
      Item(0).Control(13)=   "dtpCorte"
      Item(0).Control(14)=   "ckDiaPago"
      Item(0).Control(15)=   "btnInforme"
      Item(1).Caption =   "Cubos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "fraCargaDatos"
      Item(2).Caption =   "Auxiliar"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "OptX(2)"
      Item(2).Control(1)=   "OptX(1)"
      Item(2).Control(2)=   "dtpCorteAuxiliar"
      Item(2).Control(3)=   "Label1(0)"
      Item(2).Control(4)=   "btnInformeAuxiliar"
      Begin VB.OptionButton OptX 
         Appearance      =   0  'Flat
         Caption         =   "Saldos de Operaciones Apalancadas"
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
         Height          =   465
         Index           =   1
         Left            =   -69400
         TabIndex        =   35
         Top             =   1560
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton OptX 
         Appearance      =   0  'Flat
         Caption         =   "Resumen Saldos, Tasa y Plazos Ponderados"
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
         Height          =   465
         Index           =   2
         Left            =   -69400
         TabIndex        =   34
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame fraCargaDatos 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Cargar Datos para Analisis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -68680
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   8175
         Begin VB.OptionButton OptX 
            Appearance      =   0  'Flat
            Caption         =   "Saldos de Operaciones Apalancadas"
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
            Height          =   465
            Index           =   0
            Left            =   1800
            TabIndex        =   30
            Top             =   1320
            Value           =   -1  'True
            Width           =   2175
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaCorteCubo 
            Height          =   330
            Left            =   5760
            TabIndex        =   39
            Top             =   1440
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.PushButton btnCubo 
            Height          =   495
            Left            =   6480
            TabIndex        =   8
            Top             =   2520
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Procesar"
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
            Picture         =   "frmCR_APA_Reportes.frx":0000
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "(Corte)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   7320
            TabIndex        =   33
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Proceso para cargar información para Analisis de Administración de Garantías con Saldos al Corte.  (Cubos)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   1200
            TabIndex        =   32
            Top             =   720
            Width           =   6615
         End
         Begin VB.Image Image3 
            Height          =   630
            Left            =   240
            Picture         =   "frmCR_APA_Reportes.frx":0719
            Top             =   480
            Width           =   585
         End
         Begin VB.Label lblStatus 
            Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   1200
            TabIndex        =   31
            Top             =   2520
            Width           =   4695
         End
      End
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   4800
         Left            =   0
         TabIndex        =   11
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8467
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imgArbol"
         Appearance      =   0
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   5640
         TabIndex        =   13
         Top             =   1560
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   5640
         TabIndex        =   17
         Top             =   1200
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.CheckBox chkAcreedores 
         Height          =   255
         Left            =   5640
         TabIndex        =   19
         Top             =   2280
         Width           =   3255
         _Version        =   1441793
         _ExtentX        =   5741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mostrar a todos los Acreedores"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkOperaciones 
         Height          =   255
         Left            =   5640
         TabIndex        =   20
         Top             =   2640
         Width           =   3255
         _Version        =   1441793
         _ExtentX        =   5741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mostrar todas la operaciones"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkOperacionesConSaldo 
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   3000
         Width           =   3255
         _Version        =   1441793
         _ExtentX        =   5741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo Operaciones con Saldo"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboTipoVencimiento 
         Height          =   330
         Left            =   5640
         TabIndex        =   22
         Top             =   3720
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   5640
         TabIndex        =   25
         Top             =   4080
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Height          =   330
         Left            =   7080
         TabIndex        =   26
         Top             =   4080
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.CheckBox ckDiaPago 
         Height          =   255
         Left            =   8760
         TabIndex        =   27
         Top             =   3720
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Por Dia de pago"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   495
         Left            =   8760
         TabIndex        =   28
         Top             =   4680
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmCR_APA_Reportes.frx":0BE8
      End
      Begin XtremeSuiteControls.PushButton btnInformeAuxiliar 
         Height          =   495
         Left            =   -66160
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmCR_APA_Reportes.frx":12EF
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorteAuxiliar 
         Height          =   330
         Left            =   -69400
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Corte del Auxiliar ?"
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
         Left            =   -67840
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   24
         Top             =   4080
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fechas"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   23
         Top             =   3720
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rango"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Informe"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
      Begin XtremeShortcutBar.ShortcutCaption lblReporte 
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   480
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
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
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin MSComctlLib.ImageList ImgBoton 
      Left            =   11760
      Top             =   480
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
            Picture         =   "frmCR_APA_Reportes.frx":19F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":1B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":1C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":1D47
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   11040
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":1E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":86D3
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":87F1
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":891B
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":8A41
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":8B4F
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":8C5C
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":8D75
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":8EA3
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Reportes.frx":8FB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox gbCuenta 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   2566
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkSaldo 
         Height          =   255
         Left            =   7080
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo con Saldo"
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
      Begin XtremeSuiteControls.FlatEdit txtAcreedor 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtAcreedorDesc 
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   600
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   330
         Left            =   4560
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   255
         Left            =   6120
         TabIndex        =   6
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtAcreedorSaldo 
         Height          =   330
         Left            =   8160
         TabIndex        =   7
         Top             =   600
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   375
         Index           =   6
         Left            =   8160
         TabIndex        =   9
         Top             =   240
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Saldo en Operaciones:"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Operación"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Acreedor"
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
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCR_APA_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vTitulo As String
Dim vScroll As Boolean, vPaso As Boolean

Public Sub sbConsulta(pAcreedor As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAPA_ConsultaAcreedor '" & pAcreedor & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
   txtAcreedorDesc.Text = rs!Descripcion & ""
   txtAcreedorSaldo.Text = Format(rs!Saldo, "Standard")
   txtOperacion.SetFocus
Else
   txtAcreedorDesc.Text = ""
   txtAcreedorSaldo.Text = ""
End If

rs.Close

Call sbLimpiaPantalla(True)

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbLimpiaPantalla(Optional pTotal As Boolean = True)

If pTotal Then
   txtOperacion.Text = ""
   txtEstado.Text = ""
End If

End Sub

Public Sub sbConsultaOperacion(pOperacion As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpiaPantalla(False)

strSQL = "exec spAPA_ConsultaOperacion '" & txtAcreedor.Text & "','" & pOperacion & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    txtEstado.Text = rs!Estado_desc
Else
   txtEstado.Text = ""
End If

rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub btnCubo_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case OptX.Item(0).Value 'Base Transaccional
        strSQL = "exec spCrdApaSaldosCorte '" & Format(dtpFechaCorteCubo.Value, "yyyy/mm/dd") & "'"
        vMensaje = "AdmGarantias_Corte"
End Select
Call ConectionExecute(strSQL)

lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis, cubo: " & vMensaje

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnInforme_Click()
Dim strSQL As String, vAplicaFiltroGeneral As Boolean

On Error GoTo vError

If lblReporte.Caption = Empty Then
    MsgBox "Debe de seleccionar un reporte", vbExclamation
    Exit Sub
End If


vAplicaFiltroGeneral = False

With frmContenedor.Crt
 .Reset
 .WindowTitle = "Reporte Administración de Garantías (Pasivos)"
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 
 .Connect = glogon.ConectRPT

Select Case lblReporte.Caption
     'Operaciones
      Case "Reporte General"
        .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesGeneral.rpt")
        .Formulas(0) = "fxTitulo= 'Reporte General'"
        
        vAplicaFiltroGeneral = True
        
      Case "Reporte x Acreedor"
        .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesAcreedor.rpt")
        .Formulas(0) = "fxTitulo='Operaciones por Acreedor'"
        
        vAplicaFiltroGeneral = True
         
      'Reporte por saldos
      Case "Reporte x Saldo"
         .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesSaldos.rpt")
         vAplicaFiltroGeneral = True
          
      'Reporte por tasa de interes
      Case "Reporte x Tasa Interes"
        .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesTasaInt.rpt")
         vAplicaFiltroGeneral = True
 
         
      'Vecimiento de Cuotas
      Case "Reporte x Vecimientos"
          
          .WindowTitle = "Operaciones con Cuotas Vencidas"
          .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesVencimiento.rpt")
          
          Select Case cboTipoVencimiento.Text
               Case "Vencidas Al"
                  dtpCorte.Enabled = False
                  
                  .Formulas(0) = "fxTitulo=' Reporte de Cuotas vencidas al " & Format(dtpCorte.Value, "dd/mm/yyyy") & " '"
                  .SelectionFormula = "{CRD_APA_OPERACIONES.FECHA_PROX_PAGO} <= cdate('" & Format(dtpCorte.Value, "yyyy-mm-dd") & "')"
                                          
               Case "Vencidas Entre"
                  dtpCorte.Enabled = True
                  dtpInicio.Enabled = True
                  
                  If ckDiaPago.Value <> 1 Then
                     .Formulas(0) = "fxTitulo=' Reporte de Cuotas vencidas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy") & " '"
                     .SelectionFormula = "{CRD_APA_OPERACIONES.FECHA_PROX_PAGO} >= cdate('" & Format(dtpInicio.Value, "yyyy-mm-dd") & "')"
                     .SelectionFormula = .SelectionFormula & " and{CRD_APA_OPERACIONES.FECHA_PROX_PAGO} <= cdate('" & Format(dtpCorte.Value, "yyyy-mm-dd") & "')"
            
                  
                  Else 'ckDiaPago.Value <> 1
                     .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesDiasPago.rpt")
                     .Formulas(0) = "fxTitulo=' Reporte de Cuotas vencidas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy") & " '"
                     
                     .SelectionFormula = "{CRD_APA_OPERACIONES.DIA_DE_PAGO} >= " & DatePart("d", dtpInicio.Value) & ""
                     .SelectionFormula = .SelectionFormula & " and{CRD_APA_OPERACIONES.DIA_DE_PAGO} <= " & DatePart("d", dtpCorte.Value) & " "
                     .SelectionFormula = .SelectionFormula & " and{CRD_APA_CONTROL_PAGOS.FECHA_PAGO} < cdate('" & Format(dtpInicio.Value, "yyyy-mm-dd") & "')"
                  End If
              End Select
      
      
                    'Filtros para Informes s/Operaciones
                If chkAcreedores.Value = vbUnchecked Then
                   .SelectionFormula = .SelectionFormula & " and {CRD_APA_OPERACIONES.COD_ACREEDOR} = '" & txtAcreedor.Text _
                                     & "' and {CRD_APA_OPERACIONES.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
                End If
                
                If chkOperaciones.Value = vbUnchecked Then
                    .SelectionFormula = .SelectionFormula & " and {CRD_APA_OPERACIONES.OPERACION} = '" & txtOperacion.Text & "'"
                End If
            
  
      
      'Listado de Operaciones Asignadas al Fideicomiso
      Case "Listado Operaciones"
           .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesAsignadas.rpt")
           .Formulas(0) = "fxTitulo ='Reporte de Operaciones Asigandas al Acreedor'"

           'Filtros para Historicos (Garantias en Pignoradas)
           If chkAcreedores.Value = vbUnchecked Then
              .SelectionFormula = .SelectionFormula & " and {CRD_APA_GARANTIAS_H.COD_ACREEDOR} = '" & txtAcreedor.Text & "'"
           End If
           
           If chkOperaciones.Value = vbUnchecked Then
               .SelectionFormula = .SelectionFormula & " and {CRD_APA_GARANTIAS_H.OPERACION} = '" & txtOperacion.Text & "'"
           End If
                    
     '----------------------------------------------------------------------------------------------------------------------
     'Informes de Control de Pagos
      Case "Reporte Pagos Generados"
           .WindowTitle = "Reporte Pagos"
           .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesPagos.rpt")
           .Formulas(0) = "fxTitulo =' Reporte de Pagos generados '"
           .SelectionFormula = " {CRD_APA_CONTROL_PAGOS.FECHA_REGISTRO} >= cdate('" & Format(dtpInicio.Value, "yyyy-mm-dd") & "')"
           .SelectionFormula = .SelectionFormula & " and {CRD_APA_CONTROL_PAGOS.FECHA_REGISTRO} <= cdate('" & Format(dtpCorte.Value, "yyyy-mm-dd") & "')"

           'Filtros para Control de Pagos
           If chkAcreedores.Value = vbUnchecked Then
              .SelectionFormula = .SelectionFormula & " and {CRD_APA_CONTROL_PAGOS.COD_ACREEDOR} = '" & txtAcreedor.Text _
                                & "' and {CRD_APA_OPERACIONES.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
           End If
           
           If chkOperaciones.Value = vbUnchecked Then
               .SelectionFormula = .SelectionFormula & " and {CRD_APA_CONTROL_PAGOS.OPERACION} = '" & txtOperacion.Text & "'"
           End If

      Case "Reporte Pagos Pendientes"
           .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesPagos.rpt")
           .Formulas(0) = "fxTitulo =' Reporte de Pagos por Estado '"
           
           .SelectionFormula = "{CRD_APA_CONTROL_PAGOS.ESTADO} =" & "'P'"
           
           'Filtros para Control de Pagos
           If chkAcreedores.Value = vbUnchecked Then
              .SelectionFormula = .SelectionFormula & " and {CRD_APA_CONTROL_PAGOS.COD_ACREEDOR} = '" & txtAcreedor.Text _
                                & "' and {CRD_APA_OPERACIONES.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
           End If
           
           If chkOperaciones.Value = vbUnchecked Then
               .SelectionFormula = .SelectionFormula & " and {CRD_APA_CONTROL_PAGOS.OPERACION} = '" & txtOperacion.Text & "'"
           End If
      
      Case "Reporte Pagos Tramitados"
                                
           .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionesPagos.rpt")
           .Formulas(0) = "fxTitulo =' Reporte de Pagos'"
           
           .SelectionFormula = "{CRD_APA_CONTROL_PAGOS.ESTADO} = 'T'" _
                             & " and {CRD_APA_CONTROL_PAGOS.FECHA_CORTE_REMESA} >= cdate('" & Format(dtpInicio.Value, "yyyy-mm-dd") _
                             & "') and {CRD_APA_CONTROL_PAGOS.FECHA_CORTE_REMESA} <= cdate('" & Format(dtpCorte.Value, "yyyy-mm-dd") & "')"
           
           'Filtros para Control de Pagos
           If chkAcreedores.Value = vbUnchecked Then
              .SelectionFormula = .SelectionFormula & " and {CRD_APA_CONTROL_PAGOS.COD_ACREEDOR} = '" & txtAcreedor.Text _
                                & "' and {CRD_APA_OPERACIONES.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
           End If
           
           If chkOperaciones.Value = vbUnchecked Then
               .SelectionFormula = .SelectionFormula & " and {CRD_APA_CONTROL_PAGOS.OPERACION} = '" & txtOperacion.Text & "'"
           End If
           
End Select
 
 
.Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
.Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
.Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
  
'Filtros para Informes s/Operaciones
If vAplicaFiltroGeneral Then
    strSQL = "{CRD_APA_OPERACIONES.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
    If chkAcreedores.Value = vbUnchecked Then
       strSQL = strSQL & " AND {CRD_APA_OPERACIONES.COD_ACREEDOR} = '" & txtAcreedor.Text & "'"
    End If
    
    If chkOperaciones.Value = vbUnchecked Then
        strSQL = strSQL & " and {CRD_APA_OPERACIONES.OPERACION} = '" & txtOperacion.Text & "'"
    End If
    
    If chkOperacionesConSaldo.Value = vbChecked Then
        strSQL = strSQL & " AND {CRD_APA_OPERACIONES.SALDO} > 0"
    End If
    
    .SelectionFormula = strSQL
End If
  
  
.Action = 1

cboTipoVencimiento.Enabled = True
dtpInicio.Enabled = True
dtpCorte.Enabled = True

    
End With

 
Exit Sub

vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical
      

End Sub

Private Sub btnInformeAuxiliar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
        
On Error GoTo vError

strSQL = "select count(*) as Existe from CRD_APA_ACREEDORES where CORTE_FECHA = '" & Format(dtpCorteAuxiliar.Value, "yyyy/mm/dd") & " 23:59:00'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  Me.MousePointer = vbHourglass
        strSQL = "exec spAPA_AuxiliarCorte '" & Format(dtpCorteAuxiliar.Value, "yyyy/mm/dd") & " 23:59:00'"
        Call ConectionExecute(strSQL)
  Me.MousePointer = vbDefault
End If
rs.Close

With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Administración de Garantías"
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "fxUsuario='" & glogon.Usuario & "'"
    .Formulas(1) = "fxFecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "fxCorte ='" & Format(dtpCorteAuxiliar.Value, "dd/mm/yyyy") & "'"
         
     Select Case True
        Case OptX.Item(1).Value  'Listado
           .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_AuxCorteListado.rpt")
           
           .SelectionFormula = "{CRD_APA_OPERACIONES.CORTE_SALDO} <> 0"
            If Len(txtAcreedor.Text) > 0 Then
                .SelectionFormula = .SelectionFormula & " AND {CRD_APA_OPERACIONES.COD_ACREEDOR} = '" & txtAcreedor.Text & "'"
            End If
        
        Case OptX.Item(2).Value  'Resumen Saldos + Tasa + Plazos
           .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_AuxCorteResumen.rpt")
        
     End Select
    
    
    .Action = 1
End With
  

Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 OPERACION from CRD_APA_OPERACIONES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_ACREEDOR = '" & txtAcreedor.Text & "' AND OPERACION  > '" & txtOperacion.Text & "'"
       
       If chkSaldo.Value = vbChecked Then
           strSQL = strSQL & " and Saldo > 0"
       End If
       
       strSQL = strSQL & " order by OPERACION asc"
    Else
       strSQL = strSQL & " where COD_ACREEDOR = '" & txtAcreedor.Text & "' AND OPERACION < '" & txtOperacion.Text & "'"
       
       If chkSaldo.Value = vbChecked Then
           strSQL = strSQL & " and Saldo > 0"
       End If
       
       strSQL = strSQL & " order by OPERACION desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtOperacion.Text = rs!Operacion
      Call sbConsultaOperacion(rs!Operacion)
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
   vModulo = 14
End Sub

Private Function fxIndiceCodigo(xkey As String) As String
  xkey = Mid(xkey, 4, Len(xkey))
  xkey = Mid(xkey, 1, Len(xkey) - 1)
  fxIndiceCodigo = xkey
End Function

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

  If Right(Node.Key, 1) = "Z" Then
     lblReporte.Caption = Node.Text
     lblReporte.Tag = fxIndiceCodigo(Node.Key)
     
     Select Case Node.Text
       Case "Reporte x Vecimientos"
          cboTipoVencimiento.Enabled = True
       Case "Reporte Pagos Tramitados"
          cboTipoVencimiento.Enabled = False
       Case "Vencidas Entre"
          dtpCorte.Enabled = True
          dtpInicio.Enabled = True
       Case "Vencidas Al"
          dtpCorte.Enabled = False
     End Select
  End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub Form_Load()
 vModulo = 14
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 dtpInicio.Value = fxFechaServidor
 dtpCorte.Value = dtpInicio.Value
 
 dtpFechaCorteCubo.Value = dtpInicio.Value
 dtpCorteAuxiliar.Value = dtpInicio.Value
 
 Call sbRefrescaArbol
 
 Call sbCargaCbo
 
tcMain.Item(0).Selected = True

 
End Sub

Private Sub sbCargaCbo()

  cboEstado.AddItem "Activa"
  cboEstado.AddItem "Cancelada"
  cboEstado.Text = "Activa"
  
  cboTipo.AddItem "Detalle"
  cboTipo.AddItem "Resumen"
  cboTipo.Text = "Detalle"
  
  cboTipoVencimiento.AddItem "Vencidas Al"
  cboTipoVencimiento.AddItem "Vencidas Entre"
  cboTipoVencimiento.Text = "Vencidas Al"
    
End Sub

Private Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub

Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String
Dim vPadre As String

With ArbolExp
  .Nodes.Clear
  Set vNode = .Nodes.Add(, , "Reportes", "Reportes", "imgRoot")
  Call sbCreaNodos("Reportes", "Operaciones", "imgCRD", False, "0x0OPR")
     Call sbCreaNodos("0x0OPR", "Reporte General", "imgDetalle", False, "0x0" & "GEN" & "Z")
     Call sbCreaNodos("0x0OPR", "Reporte x Acreedor", "imgDetalle", False, "0x0" & "ACR" & "Z")
     Call sbCreaNodos("0x0OPR", "Reporte x Saldo", "imgDetalle", False, "0x0" & "SLD" & "Z")
     Call sbCreaNodos("0x0OPR", "Reporte x Tasa Interes", "imgDetalle", False, "0x0" & "INT" & "Z")
     Call sbCreaNodos("0x0OPR", "Reporte x Vecimientos", "imgDetalle", False, "0x0" & "VEN" & "Z")
     Call sbCreaNodos("0x0OPR", "Listado Operaciones", "imgDetalle", False, "0x0" & "lstOp" & "Z")
  Call sbCreaNodos("Reportes", "Pago", "imgCRD", False, "0x0PAG")
     Call sbCreaNodos("0x0PAG", "Reporte Pagos Generados", "imgDetalle", False, "0x0" & "PGEN" & "Z")
     Call sbCreaNodos("0x0PAG", "Reporte Pagos Pendientes", "imgDetalle", False, "0x0" & "PPEN" & "Z")
     Call sbCreaNodos("0x0PAG", "Reporte Pagos Tramitados", "imgDetalle", False, "0x0" & "PTRA" & "Z")
  .Nodes(1).Expanded = True
End With


End Sub




Private Sub txtAcreedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOperacion.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "COD_ACREEDOR"
    gBusquedas.Orden = "COD_ACREEDOR"
    gBusquedas.Consulta = "SELECT COD_ACREEDOR AS 'ACREEDOR', DESCRIPCION FROM CRD_APA_ACREEDORES"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
       txtAcreedor.Text = gBusquedas.Resultado
       Call sbConsulta(gBusquedas.Resultado)
    End If
End If


End Sub

Private Sub txtAcreedor_LostFocus()
    Call sbConsulta(txtAcreedor.Text)
End Sub

Private Sub txtAcreedorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOperacion.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Consulta = "SELECT COD_ACREEDOR AS 'ACREEDOR', DESCRIPCION FROM CRD_APA_ACREEDORES"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
       txtAcreedor.Text = gBusquedas.Resultado
       Call sbConsulta(gBusquedas.Resultado)
    End If
End If
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

If txtAcreedorDesc.Text = "" Then
   MsgBox "Debe indicar un Acreedor antes que la operación!", vbExclamation
   txtAcreedor.SetFocus
   Exit Sub
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
       Call sbConsultaOperacion(txtOperacion.Text)
       txtEstado.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "OPERACION"
    gBusquedas.Orden = "OPERACION"
    gBusquedas.Consulta = "SELECT OPERACION,COD_ACREEDOR AS 'ACREEDOR', MONTO, SALDO, FECHA_FORMALIZA AS 'FORMALIZA'" _
                        & " FROM crd_apa_operaciones"
    gBusquedas.Filtro = " AND COD_ACREEDOR = '" & txtAcreedor.Text & "'"
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
       txtOperacion.Text = gBusquedas.Resultado
       Call sbConsultaOperacion(gBusquedas.Resultado)
    End If
End If


End Sub

