VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_Convenios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Convenios"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   255
      Left            =   3360
      TabIndex        =   41
      Top             =   1320
      Width           =   1695
      _Version        =   1310722
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Activo?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   4
   End
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   5895
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   10455
      _Version        =   1310722
      _ExtentX        =   18441
      _ExtentY        =   10398
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
      ItemCount       =   9
      Item(0).Caption =   "General"
      Item(0).ControlCount=   32
      Item(0).Control(0)=   "txtPorcReservaRecaudacion"
      Item(0).Control(1)=   "cboTipo"
      Item(0).Control(2)=   "txtReservaSaldo"
      Item(0).Control(3)=   "txtReservaEjecutado"
      Item(0).Control(4)=   "txtReservaCorte"
      Item(0).Control(5)=   "txtReservaTope"
      Item(0).Control(6)=   "txtPorcComisionCreditos"
      Item(0).Control(7)=   "txtPorcComisionRecaudacion"
      Item(0).Control(8)=   "dtpInicio"
      Item(0).Control(9)=   "txtContrato"
      Item(0).Control(10)=   "txtProveedorNombre"
      Item(0).Control(11)=   "txtProveedorId"
      Item(0).Control(12)=   "txtClienteNombre"
      Item(0).Control(13)=   "txtClienteId"
      Item(0).Control(14)=   "dtpVence"
      Item(0).Control(15)=   "Label3(25)"
      Item(0).Control(16)=   "Label2(4)"
      Item(0).Control(17)=   "Label3(9)"
      Item(0).Control(18)=   "Label3(6)"
      Item(0).Control(19)=   "Label3(5)"
      Item(0).Control(20)=   "Label3(4)"
      Item(0).Control(21)=   "Label2(3)"
      Item(0).Control(22)=   "Label3(22)"
      Item(0).Control(23)=   "Label3(21)"
      Item(0).Control(24)=   "Label3(14)"
      Item(0).Control(25)=   "Label3(13)"
      Item(0).Control(26)=   "Label3(12)"
      Item(0).Control(27)=   "Label2(2)"
      Item(0).Control(28)=   "Label2(1)"
      Item(0).Control(29)=   "Label2(0)"
      Item(0).Control(30)=   "chkComisionInformativa"
      Item(0).Control(31)=   "chkCreditosAnulados"
      Item(1).Caption =   "Recaudación"
      Item(1).ControlCount=   8
      Item(1).Control(0)=   "txtRetencionCod"
      Item(1).Control(1)=   "txtRetencionNombre"
      Item(1).Control(2)=   "lswRetencion"
      Item(1).Control(3)=   "fsbRetenciones"
      Item(1).Control(4)=   "Label3(8)"
      Item(1).Control(5)=   "Label3(7)"
      Item(1).Control(6)=   "btnRecaudacion(0)"
      Item(1).Control(7)=   "btnRecaudacion(1)"
      Item(2).Caption =   "Devoluciones"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "FlatScrollBar1"
      Item(2).Control(1)=   "Label3(26)"
      Item(2).Control(2)=   "Label3(27)"
      Item(2).Control(3)=   "txtDevolucionCod"
      Item(2).Control(4)=   "txtDevolucionDesc"
      Item(2).Control(5)=   "lswDevolucion"
      Item(2).Control(6)=   "btnDevoluciones(0)"
      Item(2).Control(7)=   "btnDevoluciones(1)"
      Item(3).Caption =   "Créditos"
      Item(3).ControlCount=   9
      Item(3).Control(0)=   "optLinea(1)"
      Item(3).Control(1)=   "optLinea(0)"
      Item(3).Control(2)=   "chkLineaCombinaDestino"
      Item(3).Control(3)=   "txtLineaCod"
      Item(3).Control(4)=   "txtLineaDesc"
      Item(3).Control(5)=   "lswDestinos"
      Item(3).Control(6)=   "fsbCredito"
      Item(3).Control(7)=   "btnDestino(0)"
      Item(3).Control(8)=   "btnDestino(1)"
      Item(4).Caption =   "Cargos"
      Item(4).ControlCount=   8
      Item(4).Control(0)=   "txtCargoCod"
      Item(4).Control(1)=   "txtCargoNombre"
      Item(4).Control(2)=   "lswCargo"
      Item(4).Control(3)=   "fsbCargos"
      Item(4).Control(4)=   "Label3(11)"
      Item(4).Control(5)=   "Label3(10)"
      Item(4).Control(6)=   "btnCargo(0)"
      Item(4).Control(7)=   "btnCargo(1)"
      Item(5).Caption =   "Rebajos de CxP"
      Item(5).ControlCount=   2
      Item(5).Control(0)=   "lswCargosCxP"
      Item(5).Control(1)=   "Label3(15)"
      Item(6).Caption =   "Ahorros"
      Item(6).ControlCount=   14
      Item(6).Control(0)=   "txtPlanContrato"
      Item(6).Control(1)=   "txtPlanMensualidad"
      Item(6).Control(2)=   "txtPlanCod"
      Item(6).Control(3)=   "txtPlanName"
      Item(6).Control(4)=   "fsbFondos"
      Item(6).Control(5)=   "Label3(24)"
      Item(6).Control(6)=   "Label3(18)"
      Item(6).Control(7)=   "Label3(17)"
      Item(6).Control(8)=   "Label3(16)"
      Item(6).Control(9)=   "lswPlan"
      Item(6).Control(10)=   "btnPlanes(0)"
      Item(6).Control(11)=   "btnPlanes(1)"
      Item(6).Control(12)=   "btnPlanes(2)"
      Item(6).Control(13)=   "btnPlanes(3)"
      Item(7).Caption =   "Ordenes"
      Item(7).ControlCount=   4
      Item(7).Control(0)=   "lswOrdenes"
      Item(7).Control(1)=   "Label3(19)"
      Item(7).Control(2)=   "feLineas"
      Item(7).Control(3)=   "Label3(28)"
      Item(8).Caption =   "Reservas"
      Item(8).ControlCount=   2
      Item(8).Control(0)=   "lswReservas"
      Item(8).Control(1)=   "Label3(23)"
      Begin XtremeSuiteControls.ListView lswReservas 
         Height          =   4692
         Left            =   -69760
         TabIndex        =   81
         Top             =   840
         Visible         =   0   'False
         Width           =   9972
         _Version        =   1310722
         _ExtentX        =   17590
         _ExtentY        =   8276
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswOrdenes 
         Height          =   4212
         Left            =   -69760
         TabIndex        =   80
         Top             =   1080
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1310722
         _ExtentX        =   17378
         _ExtentY        =   7429
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswPlan 
         Height          =   4212
         Left            =   -69640
         TabIndex        =   79
         Top             =   1560
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1310722
         _ExtentX        =   17378
         _ExtentY        =   7429
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswCargosCxP 
         Height          =   4452
         Left            =   -69640
         TabIndex        =   78
         Top             =   1080
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1310722
         _ExtentX        =   17166
         _ExtentY        =   7853
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswCargo 
         Height          =   4212
         Left            =   -69640
         TabIndex        =   77
         Top             =   1320
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1310722
         _ExtentX        =   17166
         _ExtentY        =   7429
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRetencion 
         Height          =   4212
         Left            =   -69640
         TabIndex        =   74
         Top             =   1440
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1310722
         _ExtentX        =   17166
         _ExtentY        =   7429
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswDevolucion 
         Height          =   4212
         Left            =   -69640
         TabIndex        =   75
         Top             =   1440
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1310722
         _ExtentX        =   17166
         _ExtentY        =   7429
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswDestinos 
         Height          =   4212
         Left            =   -69640
         TabIndex        =   76
         Top             =   1440
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1310722
         _ExtentX        =   17166
         _ExtentY        =   7429
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnDevoluciones 
         Height          =   252
         Index           =   0
         Left            =   -60880
         TabIndex        =   82
         Top             =   960
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":0000
      End
      Begin XtremeSuiteControls.RadioButton optLinea 
         Height          =   252
         Index           =   0
         Left            =   -68440
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1310722
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Línea de Crédito"
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
         Appearance      =   2
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkLineaCombinaDestino 
         Height          =   252
         Left            =   -62320
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1310722
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Asociar a Destino"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkComisionInformativa 
         Height          =   612
         Left            =   1320
         TabIndex        =   37
         Top             =   2280
         Width           =   2772
         _Version        =   1310722
         _ExtentX        =   4890
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Calcular Comisión para Registro Informativo"
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
         Appearance      =   2
         Alignment       =   1
      End
      Begin MSComCtl2.FlatScrollBar fsbRetenciones 
         Height          =   252
         Left            =   -61480
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar fsbCredito 
         Height          =   252
         Left            =   -62920
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar fsbCargos 
         Height          =   252
         Left            =   -61600
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar fsbFondos 
         Height          =   252
         Left            =   -61600
         TabIndex        =   27
         Top             =   1080
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
         Height          =   252
         Left            =   -61480
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.RadioButton optLinea 
         Height          =   252
         Index           =   1
         Left            =   -66160
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1310722
         _ExtentX        =   4890
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Destino de crédito"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit feLineas 
         Height          =   252
         Left            =   -60640
         TabIndex        =   42
         Top             =   5400
         Visible         =   0   'False
         Width           =   732
         _Version        =   1310722
         _ExtentX        =   1291
         _ExtentY        =   444
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "50"
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkCreditosAnulados 
         Height          =   732
         Left            =   1320
         TabIndex        =   44
         Top             =   2880
         Width           =   2772
         _Version        =   1310722
         _ExtentX        =   4890
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Calcula el Cobro por Créditos Anulados en el periodo de fechas de la Liquidación?"
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
         Appearance      =   2
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   2520
         TabIndex        =   45
         Top             =   600
         Width           =   6852
         _Version        =   1310722
         _ExtentX        =   12091
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   2520
         TabIndex        =   46
         Top             =   4800
         Width           =   1575
         _Version        =   1310722
         _ExtentX        =   2773
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
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   315
         Left            =   2520
         TabIndex        =   47
         Top             =   5160
         Width           =   1575
         _Version        =   1310722
         _ExtentX        =   2773
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtReservaTope 
         Height          =   315
         Left            =   7800
         TabIndex        =   48
         Top             =   4440
         Width           =   1575
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtReservaCorte 
         Height          =   315
         Left            =   7800
         TabIndex        =   49
         Top             =   4800
         Width           =   1575
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   550
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtReservaEjecutado 
         Height          =   315
         Left            =   7800
         TabIndex        =   50
         Top             =   5160
         Width           =   1575
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   550
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtReservaSaldo 
         Height          =   315
         Left            =   7800
         TabIndex        =   51
         Top             =   5520
         Width           =   1575
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   550
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcComisionRecaudacion 
         Height          =   312
         Left            =   8520
         TabIndex        =   52
         Top             =   2400
         Width           =   852
         _Version        =   1310722
         _ExtentX        =   1503
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcComisionCreditos 
         Height          =   312
         Left            =   8520
         TabIndex        =   53
         Top             =   2760
         Width           =   852
         _Version        =   1310722
         _ExtentX        =   1503
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcReservaRecaudacion 
         Height          =   312
         Left            =   8520
         TabIndex        =   54
         Top             =   3120
         Width           =   852
         _Version        =   1310722
         _ExtentX        =   1503
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtClienteId 
         Height          =   312
         Left            =   2520
         TabIndex        =   57
         Top             =   1200
         Width           =   1572
         _Version        =   1310722
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
      Begin XtremeSuiteControls.FlatEdit txtClienteNombre 
         Height          =   312
         Left            =   4080
         TabIndex        =   58
         Top             =   1200
         Width           =   5292
         _Version        =   1310722
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedorId 
         Height          =   312
         Left            =   2520
         TabIndex        =   59
         Top             =   1800
         Width           =   1572
         _Version        =   1310722
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
      Begin XtremeSuiteControls.FlatEdit txtProveedorNombre 
         Height          =   312
         Left            =   4080
         TabIndex        =   60
         Top             =   1800
         Width           =   5292
         _Version        =   1310722
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtContrato 
         Height          =   315
         Left            =   2520
         TabIndex        =   61
         Top             =   4440
         Width           =   1575
         _Version        =   1310722
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
      Begin XtremeSuiteControls.FlatEdit txtRetencionCod 
         Height          =   312
         Left            =   -69640
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310722
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtRetencionNombre 
         Height          =   312
         Left            =   -68560
         TabIndex        =   63
         Top             =   960
         Visible         =   0   'False
         Width           =   6972
         _Version        =   1310722
         _ExtentX        =   12298
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
      Begin XtremeSuiteControls.FlatEdit txtDevolucionCod 
         Height          =   312
         Left            =   -69640
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310722
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtDevolucionDesc 
         Height          =   312
         Left            =   -68560
         TabIndex        =   65
         Top             =   960
         Visible         =   0   'False
         Width           =   6972
         _Version        =   1310722
         _ExtentX        =   12298
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
      Begin XtremeSuiteControls.FlatEdit txtLineaCod 
         Height          =   312
         Left            =   -69640
         TabIndex        =   66
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310722
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
         Height          =   312
         Left            =   -68560
         TabIndex        =   67
         Top             =   960
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1310722
         _ExtentX        =   9758
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
      Begin XtremeSuiteControls.FlatEdit txtCargoCod 
         Height          =   312
         Left            =   -69640
         TabIndex        =   68
         Top             =   840
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310722
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtCargoNombre 
         Height          =   312
         Left            =   -68560
         TabIndex        =   69
         Top             =   840
         Visible         =   0   'False
         Width           =   6852
         _Version        =   1310722
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
      Begin XtremeSuiteControls.FlatEdit txtPlanCod 
         Height          =   312
         Left            =   -69640
         TabIndex        =   70
         Top             =   1080
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310722
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtPlanName 
         Height          =   312
         Left            =   -68560
         TabIndex        =   71
         Top             =   1080
         Visible         =   0   'False
         Width           =   4452
         _Version        =   1310722
         _ExtentX        =   7853
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
      Begin XtremeSuiteControls.FlatEdit txtPlanMensualidad 
         Height          =   312
         Left            =   -64120
         TabIndex        =   72
         Top             =   1080
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPlanContrato 
         Height          =   312
         Left            =   -62800
         TabIndex        =   73
         Top             =   1080
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310722
         _ExtentX        =   1926
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton btnDevoluciones 
         Height          =   252
         Index           =   1
         Left            =   -60520
         TabIndex        =   83
         Top             =   960
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":0720
      End
      Begin XtremeSuiteControls.PushButton btnRecaudacion 
         Height          =   252
         Index           =   0
         Left            =   -60880
         TabIndex        =   84
         Top             =   960
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":0E36
      End
      Begin XtremeSuiteControls.PushButton btnRecaudacion 
         Height          =   252
         Index           =   1
         Left            =   -60520
         TabIndex        =   85
         Top             =   960
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":1556
      End
      Begin XtremeSuiteControls.PushButton btnDestino 
         Height          =   252
         Index           =   0
         Left            =   -62320
         TabIndex        =   86
         Top             =   960
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":1C6C
      End
      Begin XtremeSuiteControls.PushButton btnDestino 
         Height          =   252
         Index           =   1
         Left            =   -61960
         TabIndex        =   87
         Top             =   960
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":238C
      End
      Begin XtremeSuiteControls.PushButton btnCargo 
         Height          =   252
         Index           =   0
         Left            =   -61000
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":2AA2
      End
      Begin XtremeSuiteControls.PushButton btnCargo 
         Height          =   252
         Index           =   1
         Left            =   -60640
         TabIndex        =   89
         Top             =   840
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":31C2
      End
      Begin XtremeSuiteControls.PushButton btnPlanes 
         Height          =   252
         Index           =   0
         Left            =   -61000
         TabIndex        =   90
         Top             =   1080
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":38D8
      End
      Begin XtremeSuiteControls.PushButton btnPlanes 
         Height          =   252
         Index           =   1
         Left            =   -60640
         TabIndex        =   91
         Top             =   1080
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":3FF8
      End
      Begin XtremeSuiteControls.PushButton btnPlanes 
         Height          =   252
         Index           =   2
         Left            =   -60280
         TabIndex        =   92
         Top             =   1080
         Visible         =   0   'False
         Width           =   492
         _Version        =   1310722
         _ExtentX        =   868
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":470E
      End
      Begin XtremeSuiteControls.PushButton btnPlanes 
         Height          =   252
         Index           =   3
         Left            =   -61000
         TabIndex        =   93
         Top             =   720
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310722
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Terceros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCR_Convenios.frx":4E0E
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Líneas"
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
         Index           =   28
         Left            =   -62080
         TabIndex        =   43
         Top             =   5400
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto de Retención (Liquidación de Planes)"
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
         Index           =   27
         Left            =   -68560
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   5052
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Index           =   26
         Left            =   -69640
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes Registrados ...:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   23
         Left            =   -69760
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Ordenes de Pago (Liquidaciones) Procesadas..."
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
         Index           =   19
         Left            =   -69760
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   4572
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan de Ahorros"
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
         Index           =   16
         Left            =   -68560
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Index           =   17
         Left            =   -69640
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   252
         Index           =   18
         Left            =   -64120
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contrato"
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
         Index           =   24
         Left            =   -62800
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Cargos Aplicables y/o Asignados"
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
         Index           =   15
         Left            =   -69520
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   3972
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   10
         Left            =   -68560
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   2652
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Index           =   11
         Left            =   -69640
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Retención / Cartera administrada"
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
         Index           =   7
         Left            =   -68560
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   3132
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Index           =   8
         Left            =   -69640
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enlace con Clientes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   0
         Left            =   1080
         TabIndex        =   18
         Top             =   1080
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enlace con Proveedores"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   1800
         Width           =   1572
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia Contractual y Operativa ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   16
         Top             =   4080
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Index           =   12
         Left            =   1080
         TabIndex        =   15
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "( % ) Comisión por Caja de Recaudación "
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
         Index           =   21
         Left            =   4920
         TabIndex        =   12
         Top             =   2400
         Width           =   3372
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "( % ) Comisión por Nuevos Créditos"
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
         Index           =   22
         Left            =   4920
         TabIndex        =   11
         Top             =   2760
         Width           =   3372
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Reservas ...:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   10
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Tope de Reserva ..:"
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
         Left            =   4920
         TabIndex        =   9
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte ..:"
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
         Left            =   4920
         TabIndex        =   8
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Ejecutado desde el corte ..:"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo actual de la reserva ..:"
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
         Left            =   4920
         TabIndex        =   6
         Top             =   5520
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Convenio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   4
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "( % ) Reserva por Recaudación"
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
         Index           =   25
         Left            =   4920
         TabIndex        =   4
         Top             =   3120
         Width           =   3372
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   7770
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   14887
            MinWidth        =   14887
         EndProperty
      EndProperty
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9720
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   2280
      TabIndex        =   55
      Top             =   480
      Width           =   1095
      _Version        =   1310722
      _ExtentX        =   1926
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   3360
      TabIndex        =   56
      Top             =   480
      Width           =   6255
      _Version        =   1310722
      _ExtentX        =   11033
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
      Appearance      =   2
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   270
      Left            =   7560
      TabIndex        =   94
      Top             =   1320
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   476
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Listado"
                  Text            =   "Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Informe"
                  Text            =   "Informe detallado"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio"
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
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmCR_Convenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vCodigo As String, vPaso As Boolean
Dim vEdita As Boolean
Dim iFila As Integer
Dim vPlan As String, vTipoLinea As String
Dim vContratoPlan As Long


Private Sub btnCargo_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or txtCargoCod.Text = "" Or txtCargoNombre.Text = "" Then Exit Sub

Select Case Index
  Case 0 'Add
    
    If Not fxValidaCodigos(txtCargoCod.Text, "crd_convenios_cargos", "Cod_Cargo") Then
        strSQL = "insert crd_convenios_cargos(cod_cargo,cod_convenio,registro_fecha,registro_usuario)" _
                & " values('" & txtCargoCod.Text & "','" & txtCodigo.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
        Call ConectionExecute(strSQL)
        
        Call sbCargaListaCargos
    
    Else
        MsgBox "El cargo " & txtCargoCod.Text & " ya esta asignado a otro convenio", vbInformation
    End If
        
  Case 1 'Cancel
    
    If fxValidaCodigos(txtCargoCod.Text, "crd_convenios_cargos", "Cod_Cargo") Then
       strSQL = "Delete crd_convenios_cargos where cod_cargo = '" & txtCargoCod.Text & "' and cod_convenio = '" & txtCodigo.Text & "'"
       Call ConectionExecute(strSQL)
       
       Call sbCargaListaCargos
    
    End If
    

End Select

txtCargoCod.Text = ""
txtCargoNombre.Text = ""

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnDestino_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or txtLineaCod.Text = "" Then Exit Sub

Select Case Index
  
  Case 0 'Add
     
     If optLinea(0).Value Then
     
        If Not fxValidaCodigoCredito(txtLineaCod.Text) Then
            strSQL = "insert crd_convenios_lineas(codigo,cod_convenio,registro_fecha,registro_usuario,combina_destino)" _
                   & " values('" & txtLineaCod.Text & "','" & txtCodigo.Text & "',dbo.MyGetdate(),'" & glogon.Usuario _
                   & "'," & chkLineaCombinaDestino.Value & ")"
            Call ConectionExecute(strSQL)
        Else
            MsgBox "La línea de crédito " & txtLineaCod.Text & " ya se encuentra asigando a otro convenio", vbInformation
        End If
     Else
        If Not fxValidaCodigos(txtLineaCod.Text, "crd_convenios_destinos", "cod_destino") Then
            strSQL = "insert crd_convenios_destinos(cod_destino,cod_convenio,registro_fecha,registro_usuario)" _
                   & " values('" & txtLineaCod.Text & "','" & txtCodigo.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
            Call ConectionExecute(strSQL)
        Else
          MsgBox "El destino " & txtLineaCod.Text & " ya se encuentra asigando a otro convenio", vbInformation
        End If
        
     End If
  
  Case 1 'Cancel
     
     If optLinea(0).Value Then
        
            strSQL = "Delete crd_convenios_lineas where codigo = '" & txtLineaCod.Text & "' and cod_convenio = '" & txtCodigo.Text & "'"
            Call ConectionExecute(strSQL)
     
     Else
            strSQL = "Delete crd_convenios_destinos where cod_destino = '" & txtLineaCod.Text & "' and cod_convenio = '" & txtCodigo.Text & "'"
            Call ConectionExecute(strSQL)
     End If

End Select

Call sbCargaListaCredito

txtLineaCod.Text = ""
txtLineaDesc.Text = ""
chkLineaCombinaDestino.Value = vbUnchecked

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnDevoluciones_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or txtDevolucionCod.Text = "" Then Exit Sub

Select Case Index
  Case 0 'Add
        
    If Not fxValidaCodigos(txtDevolucionCod.Text, "CRD_CONVENIOS_DEVOLUCION", "RETENCION_CODIGO") Then
        strSQL = "insert CRD_CONVENIOS_DEVOLUCION(RETENCION_CODIGO,cod_convenio,registro_fecha,registro_usuario)" _
                & " values('" & txtDevolucionCod.Text & "','" & txtCodigo.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
        Call ConectionExecute(strSQL)
    Else
        MsgBox "El código " & txtDevolucionCod.Text & " ya esta asignado a otro convenio", vbInformation
    End If
    
  Case 1 'Cancel
    
    If fxValidaCodigos(txtDevolucionCod.Text, "CRD_CONVENIOS_DEVOLUCION", "Codigo") Then
       strSQL = "Delete CRD_CONVENIOS_DEVOLUCION where RETENCION_CODIGO = '" & txtDevolucionCod.Text & "' and cod_convenio = '" & txtCodigo.Text & "'"
       Call ConectionExecute(strSQL)
    End If
    
End Select


Call sbCargaListaDevoluciones
       
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnPlanes_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vErrorTbPlanes

Select Case Index
  Case 0 'Add
     If txtCodigo.Text = "" Or txtPlanCod.Text = "" Then Exit Sub
     If txtPlanMensualidad.Text = "" Or Not IsNumeric(txtPlanMensualidad.Text) Or txtPlanMensualidad <= 0 Then Exit Sub

        If txtPlanContrato.Text = "" Then
            txtPlanContrato.Text = "0"
        End If
        
        strSQL = "exec spConvenios_FondosPool '" & txtCodigo.Text & "','" & txtPlanCod.Text & "'," & txtPlanContrato.Text _
                & "," & CCur(txtPlanMensualidad.Text) & ",'M'"
        Call ConectionExecute(strSQL)

        Call sbCargaListaFondos

  Case 1 'Cancel
    If txtCodigo.Text = "" Or txtPlanCod.Text = "" Then Exit Sub
    
    If txtPlanContrato.Text <> "" Then
        strSQL = "exec spConvenios_FondosPool '" & txtCodigo.Text & "','" & txtPlanCod.Text & "'," & txtPlanContrato.Text _
                & "," & CCur(txtPlanMensualidad.Text) & ",'E'"
        Call ConectionExecute(strSQL)
    End If
    Call sbCargaListaFondos
    
  Case 2 'Busca

    gBusquedas.Consulta = "select C.COD_PLAN, C.COD_CONTRATO, C.MONTO, P.DESCRIPCION" _
                        & " from CRD_CONVENIOS V inner join FND_CONTRATOS C on V.CEDULA = C.CEDULA" _
                        & " inner join FND_PLANES P on C.COD_PLAN = P.COD_PLAN and C.COD_OPERADORA = P.COD_OPERADORA"
    gBusquedas.Columna = "c.cod_plan"
    gBusquedas.Filtro = " and C.estado = 'A' and P.TIPO_CDP = 0 AND V.COD_CONVENIO  = '" & txtCodigo.Text & "'"
    gBusquedas.Orden = "P.DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtPlanCod.Text = Trim(gBusquedas.Resultado)
        txtPlanName.Text = fxFND_DesPlan(txtPlanCod.Text)
        
        txtPlanContrato.Text = gBusquedas.Resultado2
        txtPlanMensualidad.Text = Format(gBusquedas.Resultado3, "Standard")
        
    End If
    
    
    Case 3 'Terceros
     gBusquedas.Consulta = "select C.COD_PLAN, C.COD_CONTRATO, C.MONTO, P.DESCRIPCION" _
                        & " from CRD_CONVENIOS V inner join FND_CONTRATOS C on V.CEDULA <> C.CEDULA" _
                        & " inner join FND_PLANES P on C.COD_PLAN = P.COD_PLAN and C.COD_OPERADORA = P.COD_OPERADORA"
    gBusquedas.Columna = "c.cod_plan"
    gBusquedas.Filtro = " and C.estado = 'A' and P.TIPO_CDP = 0 AND V.COD_CONVENIO  = '" & txtCodigo.Text & "'"
    gBusquedas.Orden = "P.DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtPlanCod.Text = Trim(gBusquedas.Resultado)
        txtPlanName.Text = fxFND_DesPlan(txtPlanCod.Text)
        
        txtPlanContrato.Text = gBusquedas.Resultado2
        txtPlanMensualidad.Text = Format(gBusquedas.Resultado3, "Standard")
        
    End If
 
End Select


Exit Sub

vErrorTbPlanes:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRecaudacion_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or txtRetencionCod.Text = "" Then Exit Sub

Select Case Index
  Case 0 'Add
        
    If Not fxValidaCodigos(txtRetencionCod.Text, "crd_convenios_retencion", "Codigo") Then
        strSQL = "insert crd_convenios_retencion(codigo,cod_convenio,registro_fecha,registro_usuario)" _
                & " values('" & txtRetencionCod.Text & "','" & txtCodigo.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
        Call ConectionExecute(strSQL)
    Else
        MsgBox "El código " & txtRetencionCod.Text & " ya esta asignado a otro convenio", vbInformation
    End If
    
  Case 1 'Cancel
    
    If fxValidaCodigos(txtRetencionCod.Text, "crd_convenios_retencion", "Codigo") Then
       strSQL = "Delete crd_convenios_retencion where codigo = '" & txtRetencionCod.Text & "' and cod_convenio = '" & txtCodigo.Text & "'"
       Call ConectionExecute(strSQL)
    End If
    
End Select


Call sbCargaListaRetenciones
       
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub feLineas_Change()
If IsNumeric(feLineas) Then
  Call sbCargaOrdenes
End If
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 COD_CONVENIO from CRD_CONVENIOS"
    
    If Len(txtCodigo.Text) > 0 Then
        If FlatScrollBar.Value = 1 Then
           strSQL = strSQL & " where COD_CONVENIO > '" & txtCodigo.Text & "' order by COD_CONVENIO asc"
        Else
           strSQL = strSQL & " where COD_CONVENIO < '" & txtCodigo.Text & "' order by COD_CONVENIO desc"
        End If
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_CONVENIO
      Call txtCodigo_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub

Private Sub Form_Activate()
vModulo = 16
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 16

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswDestinos.ColumnHeaders
    .Clear
    .Add , , "Código", 1500
    .Add , , "Descripción", 3500
    .Add , , "Asociar?", 1200, vbCenter
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
End With

With lswRetencion.ColumnHeaders
    .Clear
    .Add , , "Código", 1500
    .Add , , "Descripción", 3500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
End With

With lswDevolucion.ColumnHeaders
    .Clear
    .Add , , "Código", 1500
    .Add , , "Descripción", 3500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
End With


With lswCargo.ColumnHeaders
    .Clear
    .Add , , "Código", 1500
    .Add , , "Descripción", 3500
    .Add , , "(%) Comisión", 2100, vbCenter
End With

With lswCargosCxP.ColumnHeaders
    .Clear
    .Add , , "Código", 1500
    .Add , , "Descripción", 3500
    .Add , , "Cuenta", 2100, vbCenter
    .Add , , "Cta: Descripción", 3500
End With

 
lswPlan.Checkboxes = True
With lswPlan.ColumnHeaders
    .Clear
    .Add , , "Código", 1500
    .Add , , "Descripción", 3500
    .Add , , "Contrato", 1200, vbCenter
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Mensualidad", 2100, vbRightJustify
    .Add , , "Acumulado", 2100, vbRightJustify
End With

With lswOrdenes.ColumnHeaders
    .Clear
    .Add , , "Orden", 1500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Inicio", 2100, vbCenter
    .Add , , "Corte", 2100, vbCenter
    .Add , , "Cargos", 2100, vbRightJustify
    .Add , , "Recaudación", 2100, vbRightJustify
    .Add , , "Cartera", 2100, vbRightJustify
    .Add , , "CxP(Rebajos)", 2100, vbRightJustify
    .Add , , "Ahorros", 2100, vbRightJustify
    .Add , , "Neto a Pagar,", 2100, vbRightJustify
    .Add , , "Documento", 2100
End With

With lswReservas.ColumnHeaders
    .Clear
    .Add , , "Corte Base", 2100, vbCenter
    .Add , , "Saldo Inicial", 2100, vbRightJustify
    .Add , , "Total Retenido", 2100, vbRightJustify
    .Add , , "Total Ejecutado", 2100, vbRightJustify
    .Add , , "Saldo Actual", 2100, vbRightJustify
    .Add , , "Actualizado", 2100
End With


vScroll = False
FlatScrollBar.Value = 0
fsbRetenciones.Value = 0
fsbCredito.Value = 0
fsbCargos.Value = 0
fsbFondos.Value = 0

vScroll = True
vTipoLinea = "C"
 
txtCodigo.Text = ""
txtDescripcion.Text = ""

strSQL = "select rtrim(TIPO_CONVENIO) as 'IdX', rtrim(descripcion) as 'ItmX'" _
        & " from  CRD_CONVENIO_TIPO where Activo = 1"
        
Call sbCbo_Llena_New(cboTipo, strSQL, False, True)


Call sbLimpiaDatos
Call sbToolBarIconos(tlb)

vEdita = False
Call sbToolBar(tlb, "nuevo")


 
End Sub

Private Sub fsbCargos_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 COD_CARGO, DESCRIPCION from CARGOS_ADICIONALES"
   
    If Len(txtCargoCod.Text) > 0 Then
        If fsbCargos.Value = 1 Then
           strSQL = strSQL & " where COD_CARGO > '" & txtCargoCod.Text & "' order by COD_CARGO asc"
        Else
           strSQL = strSQL & " where COD_CARGO < '" & txtCargoCod.Text & "' order by COD_CARGO desc"
        End If
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCargoCod.Text = rs!cod_cargo
      txtCargoNombre.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
fsbCargos.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub

Private Sub fsbCredito_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
      
      If optLinea(0) Then 'se mueve en tabla convenios lineas
        strSQL = "select Top 1 Codigo, Descripcion from Catalogo"
        If Len(txtLineaCod.Text) > 0 Then
            If fsbCredito.Value = 1 Then
               strSQL = strSQL & " where Codigo > '" & txtLineaCod.Text & "' and RETENCION = 'N' and ACTIVO = 1  order by Codigo asc"
            Else
               strSQL = strSQL & " where Codigo < '" & txtLineaCod.Text & "' and RETENCION = 'N' and ACTIVO = 1  order by Codigo desc"
            End If
            
        End If
                
        Call OpenRecordSet(rs, strSQL)
      
      Else 'Se mueve en la tabla de convenios_destinos
      
        strSQL = "select Top 1 COD_DESTINO as 'Codigo',DESCRIPCION from Catalogo_Destinos"
        If Len(txtLineaCod.Text) > 0 Then
            If fsbCredito.Value = 1 Then
               strSQL = strSQL & " where COD_DESTINO > '" & txtLineaCod.Text & "' order by COD_DESTINO asc"
            Else
               strSQL = strSQL & " where COD_DESTINO < '" & txtLineaCod.Text & "' order by COD_DESTINO desc"
            End If
        End If
        Call OpenRecordSet(rs, strSQL)
        
      End If 'Fin de optLinea(0)
      

    If Not rs.EOF And Not rs.BOF Then
      txtLineaCod.Text = rs!Codigo
      txtLineaDesc.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
fsbCredito.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation
End Sub

Private Sub fsbFondos_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 COD_PLAN, DESCRIPCION" _
           & " from FND_PLANES" _
   
    If Len(txtPlanCod.Text) > 0 Then
        If fsbCargos.Value = 1 Then
           strSQL = strSQL & " where COD_PLAN > '" & txtPlanCod.Text & "' order by COD_PLAN asc"
        Else
           strSQL = strSQL & " where COD_PLAN < '" & txtPlanCod.Text & "' order by COD_PLAN desc"
        End If
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtPlanCod.Text = rs!cod_Plan
      txtPlanName.Text = rs!Descripcion
      txtContrato.Text = ""
      txtPlanMensualidad.Text = ""
      
    End If
    rs.Close
End If

vScroll = False
fsbCargos.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub

Private Sub fsbRetenciones_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 codigo,Descripcion" _
           & " from Catalogo"
   
    If Len(txtRetencionCod.Text) > 0 Then
        If fsbRetenciones.Value = 1 Then
           strSQL = strSQL & " where Codigo > '" & txtRetencionCod.Text & "' and RETENCION = 'S' and ACTIVO = 1 and LINEA_INTERNA = 0 order by Codigo asc"
        Else
           strSQL = strSQL & " where Codigo < '" & txtRetencionCod.Text & "' and RETENCION = 'S' and ACTIVO = 1 and LINEA_INTERNA = 0 order by Codigo desc"
        End If
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtRetencionCod.Text = rs!Codigo
      txtRetencionNombre.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
fsbRetenciones.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub





Private Sub lswCargo_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 If lswCargo.ListItems.Count <= 0 Then Exit Sub

 txtCargoCod.Text = Trim(Item.Text)
 txtCargoNombre.Text = Trim(Item.SubItems(1))
End Sub


Private Sub lswCargosCxP_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

On Error GoTo vError

  If Item.Checked Then
    strSQL = "Insert CRD_CONVENIOS_CARGOS_CXP(COD_CARGO, COD_CONVENIO, REGISTRO_FECHA, REGISTRO_USUARIO)" _
           & " values ('" & Trim(Item.Text) & "' ,'" & Trim(txtCodigo.Text) & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
    Call ConectionExecute(strSQL)
  Else
    strSQL = "Delete CRD_CONVENIOS_CARGOS_CXP where COD_CONVENIO = '" & Trim(txtCodigo.Text) & "' and COD_CARGO ='" & Trim(Item.Text) & "'"
    Call ConectionExecute(strSQL)
  End If
  
 Call sbCargaCargosCxP
  
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswDestinos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lswDestinos.ListItems.Count <= 0 Then Exit Sub

 If optLinea(0) Then
    txtLineaCod.Text = Trim(Item.Text)
    txtLineaDesc.Text = Trim(Item.SubItems(1))
       
    If Trim(Item.SubItems(2)) = "S" Then
       chkLineaCombinaDestino.Value = vbChecked
    Else
      chkLineaCombinaDestino.Value = Unchecked
    End If
 
 Else
    txtLineaCod.Text = Trim(Item.Text)
    txtLineaDesc.Text = Trim(Item.SubItems(1))
 End If
End Sub

Private Sub lswDevolucion_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 If lswDevolucion.ListItems.Count <= 0 Then Exit Sub
 
 txtDevolucionCod.Text = Trim(Item.Text)
 txtDevolucionDesc.Text = Trim(Item.SubItems(1))
End Sub


Private Sub lswOrdenes_DblClick()
If lswOrdenes.ListItems.Count = 0 Then Exit Sub


Dim frm As Form

Call sbSIFForms("frmCR_ConveniosLiquidacion", , , , False, Me)

For Each frm In Forms
If UCase(frm.Name) = UCase("frmCR_ConveniosLiquidacion") Then
  Call frm.sbConsultaExterna(txtCodigo.Text, lswOrdenes.SelectedItem.Text)
  Exit For
End If
Next frm
  
  
End Sub



Private Sub lswPlan_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 If lswPlan.ListItems.Count <= 0 Then Exit Sub
 
 txtPlanCod.Text = Trim(Item.Text)
 txtPlanName.Text = Trim(Item.SubItems(1))
 txtPlanMensualidad.Text = Trim(Item.SubItems(4))
 txtPlanContrato.Text = Trim(Item.SubItems(2))
End Sub


Private Sub lswRetencion_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 If lswRetencion.ListItems.Count <= 0 Then Exit Sub
 
 txtRetencionCod.Text = Trim(Item.Text)
 txtRetencionNombre.Text = Trim(Item.SubItems(1))
End Sub

Private Sub optLinea_Click(Index As Integer)
  Select Case Index
    Case 0
       vTipoLinea = "C"
       chkLineaCombinaDestino.Enabled = True
    Case 1
       vTipoLinea = "D"
       chkLineaCombinaDestino.Enabled = False
  End Select
  
  txtLineaCod.Text = ""
  txtLineaDesc.Text = ""
  
  Call sbCargaListaCredito
End Sub


Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

  If vCodigo = "" Then Exit Sub
  
  Select Case Item.Index
       
    Case 1 'Retenciones
      Call sbCargaListaRetenciones
      
    Case 2 'Devoluciones
      Call sbCargaListaDevoluciones
      
    Case 3 'Credito
      Call sbCargaListaCredito
    
    Case 4 'cargos
      Call sbCargaListaCargos
      
    Case 5 ' Cargos CxP
      Call sbCargaCargosCxP
      
    Case 6 'Ahorros
      Call sbCargaListaFondos
      
    Case 7 'Ordenes
      Call sbCargaOrdenes
      
    Case 8 'Reservas
      Call sbCargaReservas
    
  End Select
  
End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      txtCodigo.Text = ""
      txtDescripcion.Text = ""
      Call sbLimpiaDatos
      txtCodigo.Enabled = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      Call sbToolBar(tlb, "edicion")
      
    Case "BORRAR"
      Call sbBorrar
    
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      txtCodigo.Text = ""
      txtDescripcion.Text = ""
      Call sbLimpiaDatos
      Call sbToolBar(tlb, "nuevo")
      vEdita = False
    
    Case "CONSULTAR"
    
    
    
    Case "REPORTES"

    
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select
    
End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vTitulo As String, vSubTitulo As String, i As Integer, strSQL As String

On Error GoTo vError

 i = MsgBox("Esta seguro que desea visualizar TODOS los convenios Registrados?", vbYesNo)
 If i = vbYes Then
    strSQL = ""
 Else
    strSQL = "{CRD_CONVENIOS.COD_CONVENIO} = '" & txtCodigo.Text & "'"
 End If
 

Me.MousePointer = vbHourglass

vTitulo = ""
vSubTitulo = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = False
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = False
 .WindowState = crptMaximized
 
 .Connect = glogon.ConectRPT
 .WindowTitle = "Convenios: Listados Generales"
 
 vTitulo = "Lista de Convenios Registrados"


 Select Case ButtonMenu.Key
   Case "Listado"
     .ReportFileName = SIFGlobal.fxPathReportes("Convenios_ListadoRsm.rpt")
     vSubTitulo = "Listado Resumen"
   
   Case "Informe"
     .ReportFileName = SIFGlobal.fxPathReportes("Convenios_ListadoInforme.rpt")
     vSubTitulo = "Informe detallado"
 End Select
 
 .Formulas(0) = "fxTitulo= '" & vTitulo & "'"
 .Formulas(1) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(2) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(4) = "fxUsuario='Usuario: " & glogon.Usuario & "'"

 .SelectionFormula = strSQL
 
 .PrintReport

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCargoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "Select COD_CARGO,DESCRIPCION from CARGOS_ADICIONALES"
    gBusquedas.Columna = "COD_CARGO"
    gBusquedas.Orden = "COD_CARGO"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    
    frmBusquedas.Show vbModal
    txtCargoCod.Text = Trim(gBusquedas.Resultado)
    txtCargoNombre.Text = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCargoCod_LostFocus()
  If Trim(txtCargoCod.Text) <> "" Then txtCargoNombre.Text = fxDescribeCargos(txtCargoCod.Text)
End Sub

Private Sub txtCargoNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "Select COD_CARGO,DESCRIPCION from CARGOS_ADICIONALES"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "COD_CARGO"
    frmBusquedas.Show vbModal
    txtCargoCod.Text = Trim(gBusquedas.Resultado)
    txtCargoNombre.Text = Trim(gBusquedas.Resultado2)
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
End If

End Sub

Private Sub txtClienteId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtClienteNombre.SetFocus

If KeyCode = vbKeyF4 Then
    
    gBusquedas.Consulta = "Select cedula,nombre from SOCIOS"
    gBusquedas.Columna = "cedula"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtClienteId.Text = Trim(gBusquedas.Resultado)
    txtClienteNombre.Text = Trim(gBusquedas.Resultado2)
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
End If
    
End Sub

Private Sub txtClienteId_LostFocus()
   txtClienteNombre.Text = fxDescribeCliente(txtClienteId.Text)
End Sub

Private Sub txtClienteNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedorId.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "Select cedula,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtClienteId.Text = Trim(gBusquedas.Resultado)
    txtClienteNombre.Text = Trim(gBusquedas.Resultado2)
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
End If
End Sub

Private Sub txtCodigo_Change()
   Call sbLimpiaDatos
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Consulta = "Select COD_CONVENIO,DESCRIPCION" _
                        & " from CRD_CONVENIOS"
    gBusquedas.Columna = "COD_CONVENIO"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    txtCodigo.Text = Trim(gBusquedas.Resultado)
    txtDescripcion.Text = Trim(gBusquedas.Resultado2)
    
    Call sbConsulta(txtCodigo.Text)
    txtDescripcion.SetFocus

End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call sbConsulta(txtCodigo.Text)
    txtDescripcion.SetFocus
End If
  

  
End Sub

Private Sub txtCodigo_LostFocus()
  Call sbConsulta(txtCodigo.Text)
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Consulta = "Select COD_CONVENIO,DESCRIPCION" _
                        & " from CRD_CONVENIOS"
    gBusquedas.Columna = "Descripcion"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    
    txtCodigo.Text = Trim(gBusquedas.Resultado)
    txtDescripcion.Text = Trim(gBusquedas.Resultado2)
    Call sbConsulta(txtCodigo.Text)
    
    txtClienteId.SetFocus

End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbConsulta(txtCodigo.Text)
  txtClienteId.SetFocus
End If
  

End Sub

Private Sub txtDevolucionCod_Change()
  txtDevolucionDesc.Text = ""
End Sub


Private Sub txtDevolucionCod_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select RETENCION_CODIGO,DESCRIPCION from FND_RETENCION_CONCEPTOS"
    gBusquedas.Columna = "RETENCION_CODIGO"
    gBusquedas.Orden = "RETENCION_CODIGO"
    gBusquedas.Filtro = " AND ACTIVO = 1"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    
    txtDevolucionCod.Text = Trim(gBusquedas.Resultado)
    txtDevolucionDesc.Text = Trim(gBusquedas.Resultado2)
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsultaDevolucion

End Sub

Private Sub txtDevolucionDesc_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select RETENCION_CODIGO,DESCRIPCION from FND_RETENCION_CONCEPTOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Filtro = " ACTIVO = 1"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    
    txtDevolucionCod.Text = Trim(gBusquedas.Resultado)
    txtDevolucionDesc.Text = Trim(gBusquedas.Resultado2)
End If


End Sub

Private Sub txtLineaCod_KeyDown(KeyCode As Integer, Shift As Integer)

If optLinea.Item(0).Value = True Then
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select CODIGO,DESCRIPCION from CATALOGO"
        gBusquedas.Columna = "CODIGO"
        gBusquedas.Filtro = " and RETENCION = 'N' and ACTIVO = 1 and LINEA_INTERNA = 1"
        gBusquedas.Orden = "DESCRIPCION"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)
        

    End If
Else
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select rtrim(cod_destino) as codigo ,descripcion from CATALOGO_DESTINOS"
        gBusquedas.Columna = "cod_destino"
        gBusquedas.Filtro = ""
        gBusquedas.Orden = "prioridad"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)

    End If

End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtLineaDesc.SetFocus

End Sub

Private Sub txtLineaCod_LostFocus()
  If txtLineaCod.Text <> "" Then txtLineaDesc.Text = fxLineaDescrip(txtLineaCod.Text)
End Sub

Private Sub txtLineaDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If optLinea.Item(0).Value = True Then
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select CODIGO,DESCRIPCION from CATALOGO"
        gBusquedas.Columna = "DESCRIPCION"
        gBusquedas.Filtro = " and RETENCION = 'N' and ACTIVO = 1"
        gBusquedas.Orden = "DESCRIPCION"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)
        

    End If
Else
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select rtrim(cod_destino) as codigo, descripcion from CATALOGO_DESTINOS"
        gBusquedas.Columna = "descripcion"
        gBusquedas.Filtro = ""
        gBusquedas.Orden = "prioridad"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)

    End If

End If

End Sub

Private Sub txtPlanCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Consulta = "select cod_plan,DESCRIPCION from FND_PLANES"
    gBusquedas.Columna = "cod_plan"
    gBusquedas.Filtro = " and Estado = 'A'"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    txtPlanCod.Text = Trim(gBusquedas.Resultado)
    
    txtPlanName.SetFocus
    
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlanName.SetFocus

End Sub

Private Sub txtPlanCod_LostFocus()
If Trim(txtPlanCod.Text) <> "" Then txtPlanName.Text = fxFND_DesPlan(txtPlanCod.Text)
End Sub

Private Sub txtPlanMensualidad_LostFocus()
 txtPlanMensualidad.Text = Format(txtPlanMensualidad, "standard")
End Sub

Private Sub txtPlanName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Consulta = "select cod_plan,DESCRIPCION from FND_PLANES"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Filtro = " and Estado = 'A'"
    gBusquedas.Orden = "DESCRIPCION"
    frmBusquedas.Show vbModal
    txtRetencionCod.Text = Trim(gBusquedas.Resultado)
    txtPlanName.Text = Trim(gBusquedas.Resultado2)
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlanMensualidad.SetFocus

End Sub

Private Sub txtPorcComisionCreditos_Change()
If Not IsNumeric(txtPorcComisionCreditos.Text) Then
   MsgBox "Valor invalido...", vbInformation
   txtPorcComisionCreditos.Text = 0
   txtPorcComisionCreditos.SetFocus
End If

End Sub


Private Sub txtPorcComisionCreditos_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtPorcReservaRecaudacion.SetFocus
End Sub

Private Sub txtPorcComisionCreditos_LostFocus()
  txtPorcComisionCreditos.Text = Format(txtPorcComisionCreditos.Text, "standard")
End Sub

Private Sub txtPorcComisionRecaudacion_Change()
If Not IsNumeric(txtPorcComisionRecaudacion.Text) Then
   MsgBox "Valor invalido...", vbInformation
   txtPorcComisionRecaudacion.Text = 0
   txtPorcComisionRecaudacion.SetFocus
End If
End Sub

Private Sub txtPorcComisionRecaudacion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtPorcComisionCreditos.SetFocus
End Sub

Private Sub txtPorcComisionRecaudacion_LostFocus()
  txtPorcComisionRecaudacion.Text = Format(txtPorcComisionRecaudacion.Text, "standard")
End Sub


Private Sub txtPorcReservaRecaudacion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtReservaTope.SetFocus
End Sub

Private Sub txtPorcReservaRecaudacion_LostFocus()
  txtPorcReservaRecaudacion.Text = Format(txtPorcReservaRecaudacion.Text, "standard")
End Sub

Private Sub txtProveedorId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedorNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select COD_PROVEEDOR,DESCRIPCION from CXP_PROVEEDORES"
    gBusquedas.Filtro = "  and ESTADO = 'A'"
    gBusquedas.Columna = "COD_PROVEEDOR"
    gBusquedas.Orden = "DESCRIPCION"
    frmBusquedas.Show vbModal
    txtProveedorId.Text = Trim(gBusquedas.Resultado)
    txtProveedorNombre.Text = Trim(gBusquedas.Resultado2)
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
End If

End Sub

Private Sub txtProveedorId_LostFocus()
  txtProveedorNombre.Text = fxDescribeProveedor(txtProveedorId.Text)
End Sub

Private Sub txtProveedorNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPorcComisionRecaudacion.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select COD_PROVEEDOR,DESCRIPCION from CXP_PROVEEDORES"
    gBusquedas.Filtro = " and ESTADO = 'A'"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    frmBusquedas.Show vbModal
    txtProveedorId.Text = Trim(gBusquedas.Resultado)
    txtProveedorNombre.Text = Trim(gBusquedas.Resultado2)
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
End If

End Sub

Private Sub txtReservaTope_Change()
    If Not IsNumeric(txtReservaTope.Text) Then
       MsgBox "Valor invalido...", vbInformation
       txtReservaTope.Text = 0
       txtReservaTope.SetFocus
    End If
End Sub

Private Sub txtReservaTope_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtContrato.SetFocus
End Sub

Private Sub txtReservaTope_LostFocus()
     txtReservaTope.Text = Format(txtReservaTope.Text, "standard")
End Sub

Private Sub txtRetencionCod_Change()
  txtRetencionNombre.Text = ""
End Sub

Private Sub txtRetencionCod_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CODIGO,DESCRIPCION from CATALOGO"
    gBusquedas.Columna = "CODIGO"
    gBusquedas.Filtro = " and ACTIVO = 1 and (RETENCION = 'S'or LINEA_INTERNA = 0)"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    
    txtRetencionCod.Text = Trim(gBusquedas.Resultado)
    Call sbConsultaRetencion
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsultaRetencion

End Sub

Private Sub sbConsulta(pCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

Call sbLimpiaDatos

strSQL = "Select C.COD_CONVENIO,C.DESCRIPCION as 'DesConvenio' , C.TIPO_CONVENIO, ISNULL(CT.DESCRIPCION,'') as 'TipoDesc', C.DESCRIPCION" _
       & ", C.CEDULA,ISNULL(S.NOMBRE,'') as 'NombCliente',C.COD_PROVEEDOR,ISNULL(P.DESCRIPCION,'') as 'NombProveedor', C.CONTRATO_NUMERO, C.FECHA_INICIO" _
       & ", C.FECHA_VENCIMIENTO, C.ACTIVO,C.PORC_COMISION_CREDITOS, C.PORC_COMISION_RECAUDACION, C.RESERVAS_PORC" _
       & ", C.RESERVAS_RETENIDAS, C.RESERVAS_EJECUTADO, C.RESERVAS_SALDO, C.RESERVAS_CORTE, C.RESERVA_TOPE, C.ULTIMO_CORTE" _
       & ", C.COMISION_INFORMATIVA, C.COBRA_CREDITOS_ANULADOS" _
       & " from CRD_CONVENIOS C" _
       & "  LEFT join CRD_CONVENIO_TIPO CT on C.TIPO_CONVENIO = CT.TIPO_CONVENIO" _
       & "  LEFT join SOCIOS S on C.CEDULA = S.CEDULA" _
       & "  LEFT join CXP_PROVEEDORES P on C.COD_PROVEEDOR = P.COD_PROVEEDOR" _
       & " Where C.COD_CONVENIO ='" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    
   vCodigo = pCodigo
   
   txtDescripcion.Text = rs!DesConvenio
     
   Call sbCboAsignaDato(cboTipo, rs!TipoDesc, True, rs!TIPO_CONVENIO)
   
   txtClienteId.Text = Trim(rs!Cedula)
   txtClienteNombre.Text = Trim(rs!NombCliente)
   txtProveedorId.Text = Trim(rs!cod_proveedor)
   txtProveedorNombre.Text = Trim(rs!NombProveedor)
   
   txtPorcComisionRecaudacion.Text = Format(rs!Porc_Comision_Recaudacion, "standard")
   txtPorcComisionCreditos.Text = Format(rs!Porc_Comision_Creditos, "Standard")
   txtPorcReservaRecaudacion.Text = Format(rs!RESERVAS_PORC, "Standard")
   txtContrato.Text = rs!CONTRATO_NUMERO
   dtpInicio.Value = Format(rs!FECHA_INICIO, "dd/mm/yyyy")
   dtpVence.Value = Format(rs!Fecha_Vencimiento, "dd/mm/yyyy")
   txtReservaTope.Text = IIf(IsNull(rs!RESERVA_TOPE), 0, Format(rs!RESERVA_TOPE, "standard"))
   txtReservaCorte.Text = IIf(IsNull(rs!ULTIMO_CORTE), 0, Format(rs!ULTIMO_CORTE, "dd/mm/yyyy"))
   txtReservaEjecutado.Text = IIf(IsNull(rs!RESERVAS_EJECUTADO), 0, Format(rs!RESERVAS_EJECUTADO, "Standard"))
   txtReservaSaldo.Text = IIf(IsNull(rs!RESERVAS_SALDO), 0, Format(rs!RESERVAS_SALDO, "Standard"))

       
    ssTab.Item(0).Selected = True
    For i = 1 To ssTab.ItemCount - 1
     ssTab.Item(i).Enabled = True
    Next i

    chkActivo.Value = rs!Activo
    chkComisionInformativa.Value = rs!COMISION_INFORMATIVA
    chkCreditosAnulados.Value = rs!COBRA_CREDITOS_ANULADOS
    
   vEdita = True
   Call sbToolBar(tlb, "activo")

End If

rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaListaCargos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

lswCargo.ListItems.Clear

strSQL = "select CC.COD_CARGO,CA.DESCRIPCION,CA.VALOR" _
       & " from CRD_CONVENIOS_CARGOS CC inner join CARGOS_ADICIONALES CA on CC.COD_CARGO = CA.COD_CARGO" _
       & " Where CC.Cod_Convenio ='" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lswCargo.ListItems.Add(, , rs!cod_cargo)
    itmX.SubItems(1) = rs!Descripcion
    itmX.SubItems(2) = Format(rs!Valor, "Standard")
   rs.MoveNext
Loop
     
rs.Close

End Sub

Private Sub sbCargaListaRetenciones()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

txtRetencionCod.Text = ""
txtRetencionNombre.Text = ""

lswRetencion.ListItems.Clear

strSQL = "Select CVR.CODIGO, C.DESCRIPCION, CVR.COD_CONVENIO, CVR.REGISTRO_FECHA, CVR.REGISTRO_USUARIO" _
       & " from CRD_CONVENIOS_RETENCION CVR inner join Catalogo C on CVR.CODIGO = C.CODIGO" _
       & " Where CVR.COD_CONVENIO ='" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lswRetencion.ListItems.Add(, , rs!Codigo)
    itmX.SubItems(1) = rs!Descripcion
    itmX.SubItems(2) = Format(rs!registro_Fecha, "dd/mm/yyyy")
    itmX.SubItems(3) = rs!registro_usuario
   rs.MoveNext
Loop
     
rs.Close
End Sub


Private Sub sbCargaListaDevoluciones()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

txtDevolucionCod.Text = ""
txtDevolucionDesc.Text = ""

lswDevolucion.ListItems.Clear

strSQL = "Select CVR.RETENCION_CODIGO, C.DESCRIPCION, CVR.COD_CONVENIO, CVR.REGISTRO_FECHA, CVR.REGISTRO_USUARIO" _
       & " from CRD_CONVENIOS_DEVOLUCION CVR inner join FND_RETENCION_CONCEPTOS C on CVR.RETENCION_CODIGO = C.RETENCION_CODIGO" _
       & " Where CVR.COD_CONVENIO ='" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lswDevolucion.ListItems.Add(, , rs!Retencion_Codigo)
    itmX.SubItems(1) = rs!Descripcion
    itmX.SubItems(2) = Format(rs!registro_Fecha, "dd/mm/yyyy")
    itmX.SubItems(3) = rs!registro_usuario
   rs.MoveNext
Loop
     
rs.Close
End Sub


Private Sub sbConsultaRetencion()
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

strSQL = "select CODIGO,DESCRIPCION" _
       & " from Catalogo" _
       & " where CODIGO = '" & txtRetencionCod.Text & "' and ACTIVO = 1" _
       & " and (RETENCION = 'S' or LINEA_INTERNA = 0)"
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  txtRetencionCod.Text = rs!Codigo
  txtRetencionNombre.Text = rs!Descripcion
End If

rs.Close

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsultaDevolucion()
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

strSQL = "select RETENCION_CODIGO,DESCRIPCION" _
       & " from FND_RETENCION_CONCEPTOS" _
       & " where RETENCION_CODIGO = '" & txtDevolucionCod.Text & "' and ACTIVO = 1"
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  txtDevolucionCod.Text = rs!Retencion_Codigo
  txtDevolucionDesc.Text = rs!Descripcion
End If

rs.Close

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxLineaDescrip(vCodLinea As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select C.codigo,C.descripcion" _
       & " from Catalogo C" _
       & "  left join CRD_CONVENIOS_LINEAS X On C.codigo = X.codigo" _
       & " where C.codigo = '" & vCodLinea & "' and C.RETENCION = 'N' and C.ACTIVO = 1 and C.LINEA_INTERNA = 1"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    fxLineaDescrip = rs!Descripcion
Else
   fxLineaDescrip = ""
End If
rs.Close
End Function

Private Function fxConvenioDescrip(vCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select C.codigo,C.descripcion,X.tipo" _
       & " from Catalogo C" _
       & "  left join Convenios_Codigos X On C.codigo = X.codigo" _
       & " where C.codigo = '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    fxConvenioDescrip = rs!Descripcion
Else
   fxConvenioDescrip = ""
End If
rs.Close
End Function

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer
Dim vMensaje As String


vMensaje = ""
fxValida = True

If Trim(txtCodigo.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - No se ha indicado un código único al convenio..."
If Trim(txtDescripcion.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - No se especificó el nombre del Convenio..."

If Trim(txtClienteId.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Debe indicar un cliente, verifique..."
If Trim(txtProveedorId.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Debe indicar un proveedor, verifique..."
If Trim(cboTipo.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Debe indicar un tipo de convenio, verifique..."
If Trim(txtContrato.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Debe indicar un número de contrato para el  convenio, verifique..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbLimpiaDatos()
Dim strSQL As String, i As Integer

vEdita = False

vCodigo = ""

txtRetencionCod.Text = ""
txtRetencionNombre.Text = ""
txtContrato.Text = ""
txtProveedorId.Text = ""
txtProveedorNombre.Text = ""
txtClienteId.Text = ""
txtClienteNombre.Text = ""

txtReservaSaldo.Text = 0
txtReservaTope.Text = 0
txtReservaCorte.Text = 0
txtReservaEjecutado.Text = 0
txtPorcComisionCreditos.Text = 0
txtPorcComisionRecaudacion.Text = 0
txtPorcReservaRecaudacion.Text = 0

chkActivo.Value = vbChecked
chkComisionInformativa.Value = xtpUnchecked
chkCreditosAnulados.Value = xtpUnchecked

dtpInicio.Value = fxFechaServidor()
dtpVence.Value = dtpInicio.Value


ssTab.Item(0).Selected = True
For i = 1 To ssTab.ItemCount - 1
 ssTab.Item(i).Enabled = False
Next i

End Sub

Private Function fxValidaCodigos(vCodigo As String, vTabla As String, strCampo As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as existe from " & vTabla & " where " & strCampo & " =  '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe > 0 Then
  fxValidaCodigos = True
Else
  fxValidaCodigos = False
End If

rs.Close

End Function

Private Function fxValidaCodigoCredito(vCodigo As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as existe, Combina_Destino from CRD_CONVENIOS_LINEAS where CODIGO =  '" & vCodigo & "'" _
       & " group by Combina_Destino"
Call OpenRecordSet(rs, strSQL)

If rs.EOF Then
    fxValidaCodigoCredito = False
Else
    If rs!Existe > 0 Then
      'Si permite estar en mas de 1 convenio
      If rs!Combina_destino = 0 Then
        fxValidaCodigoCredito = True
      Else
        fxValidaCodigoCredito = False
      End If
    Else
      fxValidaCodigoCredito = False
    End If
End If

rs.Close

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vVencimiento As String

On Error GoTo vError


If Not IsNull(dtpVence.Value) Then
   vVencimiento = "'" & Format(dtpVence.Value, "yyyy/mm/dd") & "'"
Else
   vVencimiento = "Null"
End If

If vEdita Then
  strSQL = "update crd_convenios set TIPO_CONVENIO = '" & cboTipo.ItemData(cboTipo.ListIndex) _
         & "', DESCRIPCION = '" & Trim(txtDescripcion.Text) & "',CEDULA ='" & txtClienteId.Text _
         & "', COD_PROVEEDOR ='" & txtProveedorId.Text & "' ,CONTRATO_NUMERO = '" & txtContrato.Text _
         & "', FECHA_INICIO = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
         & "', FECHA_VENCIMIENTO = " & vVencimiento _
         & ", ACTIVO = " & chkActivo.Value & ",MODIFICA_FECHA = dbo.MyGetdate(),MODIFICA_USUARIO = '" & glogon.Usuario _
         & "', PORC_COMISION_CREDITOS = " & CCur(txtPorcComisionCreditos.Text) & "" _
         & " , PORC_COMISION_RECAUDACION = " & CCur(txtPorcComisionRecaudacion.Text) & "" _
         & " , RESERVAS_PORC = '" & CCur(txtPorcReservaRecaudacion.Text) _
         & "',RESERVA_TOPE = '" & CCur(txtReservaTope.Text) _
         & "', COMISION_INFORMATIVA = " & chkComisionInformativa.Value & ",COBRA_CREDITOS_ANULADOS = " & chkCreditosAnulados.Value _
         & " where COD_CONVENIO = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Convenio: " & vCodigo)

Else
  strSQL = "insert crd_convenios(COD_CONVENIO,TIPO_CONVENIO,DESCRIPCION,CEDULA,COD_PROVEEDOR" _
         & ",CONTRATO_NUMERO,FECHA_INICIO,FECHA_VENCIMIENTO,ACTIVO,REGISTRO_FECHA,REGISTRO_USUARIO" _
         & ",MODIFICA_FECHA,MODIFICA_USUARIO,PORC_COMISION_CREDITOS,PORC_COMISION_RECAUDACION" _
         & ",RESERVA_TOPE,ULTIMO_CORTE,RESERVAS_PORC, COMISION_INFORMATIVA, COBRA_CREDITOS_ANULADOS)" _
         & " values('" & txtCodigo.Text & "','" & cboTipo.ItemData(cboTipo.ListIndex) _
         & "','" & txtDescripcion.Text & "','" & txtClienteId.Text & "','" & txtProveedorId.Text & "','" & txtContrato.Text _
         & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
         & "'," & vVencimiento & "," & chkActivo.Value _
         & ",dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),'" & glogon.Usuario & "'," & CCur(txtPorcComisionCreditos.Text) _
         & "," & CCur(txtPorcComisionRecaudacion.Text) & "," & CCur(txtReservaTope.Text) & ",dbo.MyGetdate(),'" _
         & txtPorcReservaRecaudacion.Text & "'," & chkComisionInformativa.Value & "," & chkCreditosAnulados.Value & ")"
  Call ConectionExecute(strSQL)

  Call Bitacora("Registra", "Convenio: " & txtCodigo.Text)

  Call sbConsulta(txtCodigo.Text)

End If

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaListaCredito()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

lswDestinos.ListItems.Clear


If vTipoLinea = "C" Then
   strSQL = "Select CVL.CODIGO, C.DESCRIPCION, CVL.COD_CONVENIO, CVL.COMBINA_DESTINO, CVL.REGISTRO_FECHA, CVL.REGISTRO_USUARIO" _
          & " from CRD_CONVENIOS_LINEAS CVL inner join Catalogo C on CVL.CODIGO = C.CODIGO" _
          & " Where cod_convenio ='" & vCodigo & "'"
   Call OpenRecordSet(rs, strSQL)
         
   Do While Not rs.EOF
     Set itmX = lswDestinos.ListItems.Add(, , rs!Codigo)
       itmX.SubItems(1) = Trim(rs!Descripcion)
       itmX.SubItems(2) = IIf(rs!Combina_destino = 1, "S", "N")
       itmX.SubItems(3) = Format(rs!registro_Fecha, "dd/mm/yyyy")
       itmX.SubItems(4) = rs!registro_usuario
       rs.MoveNext
   Loop
         
   rs.Close
Else
   strSQL = "Select CVD.COD_DESTINO,C.DESCRIPCION,CVD.COD_CONVENIO, CVD.REGISTRO_FECHA, CVD.REGISTRO_USUARIO" _
          & " from CRD_CONVENIOS_DESTINOS CVD inner join Catalogo_destinos C on CVD.COD_DESTINO = C.COD_DESTINO" _
          & " Where CVD.COD_CONVENIO ='" & vCodigo & "'"
    Call OpenRecordSet(rs, strSQL)
         
    Do While Not rs.EOF
       Set itmX = lswDestinos.ListItems.Add(, , rs!cod_destino)
        itmX.SubItems(1) = Trim(rs!Descripcion)
        itmX.SubItems(2) = "...."
        itmX.SubItems(3) = Format(rs!registro_Fecha, "dd/mm/yyyy")
        itmX.SubItems(4) = rs!registro_usuario
        
       rs.MoveNext
    Loop
         
    rs.Close


End If
End Sub


Private Sub sbBorrar()
Dim iRespuesta As Integer
Dim strSQL As String

On Error GoTo vError
 iRespuesta = MsgBox("Esta seguro de confeccionar el documento...", vbYesNo)
 
 If iRespuesta = vbYes Then
 
   strSQL = "delete crd_convenios where codigo = '" & txtCodigo.Text & "'"
   Call ConectionExecute(strSQL)
   
 End If

MsgBox "Convenio eliminado satisfactoriamente...", vbInformation
Call sbLimpiaDatos
ssTab.Item(0).Selected = True
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbCargaListaFondos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

txtPlanCod.Text = ""
txtPlanName.Text = ""
txtPlanContrato.Text = ""
txtPlanMensualidad.Text = 0

lswPlan.ListItems.Clear

strSQL = " select CVP.COD_PLAN,P.DESCRIPCION ,CVP.COD_CONTRATO, CNT.FECHA_Inicio, CVP.MENSUALIDAD, CNT.APORTES + CNT.RENDIMIENTO as 'Acumulado'" _
       & " from CRD_CONVENIOS_PLANES CVP inner join FND_PLANES P on CVP.COD_PLAN = P.COD_PLAN" _
       & " INNER JOIN FND_CONTRATOS CNT on CVP.COD_PLAN = CNT.COD_PLAN AND CVP.COD_CONTRATO = CNT.COD_CONTRATO" _
       & " where CVP.COD_CONVENIO = '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lswPlan.ListItems.Add(, , rs!cod_Plan)
    itmX.SubItems(1) = rs!Descripcion
    itmX.SubItems(2) = rs!cod_Contrato
    itmX.SubItems(3) = Format(rs!FECHA_INICIO, "dd/mm/yyyy")
    itmX.SubItems(4) = Format(rs!Mensualidad, "standard")
    itmX.SubItems(5) = Format(rs!Acumulado, "standard")
   rs.MoveNext
Loop
     
rs.Close
End Sub

Function fxDescribeCargos(strCodigo As String) As String
Dim rsX As New ADODB.Recordset

If txtCodigo.Text = "" Then Exit Function

rsX.Open "select Descripcion from Cargos_Adicionales where Cod_Cargo = '" & Trim(strCodigo) & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 MsgBox "No se encontró cargo - " & strCodigo, vbCritical
Else
 fxDescribeCargos = IIf(IsNull(rsX!Descripcion), "", rsX!Descripcion)
End If
rsX.Close
End Function

Function fxDescribeCliente(strCodigo As String) As String
Dim rsX As New ADODB.Recordset

If txtClienteId.Text = "" Then Exit Function

rsX.Open "select Nombre from Socios where Cedula = '" & Trim(strCodigo) & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 MsgBox "No se encontró la cédula buscada - " & strCodigo, vbCritical
Else
 fxDescribeCliente = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close

End Function

Function fxDescribeProveedor(strCodigo As String) As String
Dim rsX As New ADODB.Recordset

If txtProveedorId.Text = "" Or Not IsNumeric(txtProveedorId.Text) Then Exit Function

rsX.Open "select DESCRIPCION from CXP_PROVEEDORES where Cod_Proveedor = '" & Trim(strCodigo) & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 MsgBox "No se encontró la cédula buscada - " & strCodigo, vbCritical
Else
 fxDescribeProveedor = IIf(IsNull(rsX!Descripcion), "", rsX!Descripcion)
End If
rsX.Close

End Function

Private Sub sbCargaReservas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

lswReservas.ListItems.Clear

strSQL = " select COD_CONVENIO,CORTE_INICIAL, SALDO_INICIAL, TOTAL_RECAUDADO, TOTAL_EJECUTADO" _
       & ", SALDO_ACTUAL, ACTUALIZADO from CRD_CONVENIOS_RESERVAS" _
       & " where COD_CONVENIO='" & vCodigo & "'"

Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   
   Set itmX = lswReservas.ListItems.Add(, , rs!COD_CONVENIO)
    itmX.SubItems(1) = Format(rs!CORTE_INICIAL, "dd/mm/yyyy")
    itmX.SubItems(2) = Format(rs!saldo_inicial, "Standard")
    itmX.SubItems(3) = Format(rs!TOTAL_RECAUDADO, "Standard")
    itmX.SubItems(4) = Format(rs!TOTAL_EJECUTADO, "Standard")
    itmX.SubItems(5) = Format(rs!saldo_actual, "Standard")
    itmX.SubItems(5) = Format(rs!ACTUALIZADO, "dd/mm/yyyy")
   rs.MoveNext
Loop
     
rs.Close

End Sub

Private Sub sbCargaCargosCxP()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If txtCodigo.Text = "" Then Exit Sub

lswCargosCxP.ListItems.Clear
lswCargosCxP.Checkboxes = True

strSQL = " select C.COD_CARGO, C.DESCRIPCION, C.COD_CUENTA, CC.COD_CARGO as 'CargoX'" _
       & ", Cta.Cod_Cuenta_Mask, Cta.Descripcion as 'CtaDesc'" _
       & " from CXP_CARGOS C left join CRD_CONVENIOS_CARGOS_CXP CC on C.COD_CARGO = CC.COD_CARGO and CC.COD_CONVENIO = '" & vCodigo & "'" _
       & " inner join vCNTX_CUENTAS_LOCAL Cta on C.cod_cuenta = Cta.cod_Cuenta" _
       & " where C.ACTIVO = 1 " _
       & " order by isnull(CC.COD_CARGO,'ZZZZZZ') asc,C.cod_Cargo asc"

Call OpenRecordSet(rs, strSQL)
     
vPaso = True
Do While Not rs.EOF
   
   Set itmX = lswCargosCxP.ListItems.Add(, , rs!cod_cargo)
    itmX.SubItems(1) = rs!Descripcion
    itmX.SubItems(2) = rs!cod_Cuenta_Mask & ""
    itmX.SubItems(3) = rs!CtaDesc & ""
    
    
   If Not IsNull(rs!cargoX) Then
     itmX.Checked = True
   End If
   
   rs.MoveNext
Loop
rs.Close

vPaso = False

End Sub

Private Sub sbCargaOrdenes()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If txtCodigo.Text = "" Then Exit Sub

On Error GoTo vError

lswOrdenes.ListItems.Clear

strSQL = "select Top " & feLineas.Text & " COD_ORDEN,REGISTRO_FECHA,REGISTRO_USUARIO,FECHA_INICIO,FECHA_CORTE,RECAUDACION_CUOTAS" _
       & ", RECAUDACION_CARGOS,NUEVOS_CREDITOS,TOTAL_REBAJOS_CXP,RETENCION_AHORROS,TOTAL_PAGAR,DOCUMENTO" _
       & " from CRD_CONVENIOS_ORDENES where COD_CONVENIO = '" & vCodigo & "'"
  
strSQL = strSQL & " order by COD_ORDEN desc"
  
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   
   Set itmX = lswOrdenes.ListItems.Add(, , rs!cod_orden)
        itmX.SubItems(1) = Format(rs!registro_Fecha, "dd/mm/yyyy")
        itmX.SubItems(2) = rs!registro_usuario
        itmX.SubItems(3) = Format(rs!FECHA_INICIO, "dd/mm/yyyy")
        itmX.SubItems(4) = Format(rs!FECHA_CORTE, "dd/mm/yyyy")
        itmX.SubItems(5) = Format(rs!Recaudacion_Cuotas, "standard")
        itmX.SubItems(6) = Format(rs!Recaudacion_Cargos, "standard")
        itmX.SubItems(7) = Format(rs!Nuevos_Creditos, "standard") 'Cartera
        itmX.SubItems(8) = Format(rs!Total_Rebajos_CXP, "standard")
        itmX.SubItems(9) = Format(rs!Retencion_Ahorros, "standard")
        itmX.SubItems(10) = Format(rs!Total_Pagar, "standard")
        itmX.SubItems(11) = rs!Documento
    
   rs.MoveNext
Loop
     
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub txtRetencionNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   
 
 If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CODIGO,DESCRIPCION from CATALOGO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Filtro = " and RETENCION = 'S' and ACTIVO = 1 and LINEA_INTERNA = 0"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    
    txtRetencionCod.Text = Trim(gBusquedas.Resultado)
    txtRetencionNombre.Text = Trim(gBusquedas.Resultado2)
End If


End Sub
