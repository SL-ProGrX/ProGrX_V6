VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_ConsultaDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   13440
   HelpContextID   =   3009
   Icon            =   "frmCR_ConsultaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkMoraCancelada 
      Height          =   255
      Left            =   2400
      TabIndex        =   110
      Top             =   8880
      Visible         =   0   'False
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Mostrar únicamente -> Morosidad Cancelada ?"
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
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   13215
      _Version        =   524288
      _ExtentX        =   23310
      _ExtentY        =   5741
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
      MaxCols         =   16
      SpreadDesigner  =   "frmCR_ConsultaDetalle.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   2652
      Left            =   0
      TabIndex        =   28
      Top             =   2760
      Width           =   13452
      _Version        =   1441793
      _ExtentX        =   23728
      _ExtentY        =   4678
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
      Item(0).Caption =   "Formalización y Cuotas"
      Item(0).ControlCount=   57
      Item(0).Control(0)=   "txtInteresC"
      Item(0).Control(1)=   "txtPlazo"
      Item(0).Control(2)=   "Label7(0)"
      Item(0).Control(3)=   "Label13"
      Item(0).Control(4)=   "lblTasa"
      Item(0).Control(5)=   "Label8"
      Item(0).Control(6)=   "Label21"
      Item(0).Control(7)=   "Label22(0)"
      Item(0).Control(8)=   "medMontoGirado"
      Item(0).Control(9)=   "medMontoApr"
      Item(0).Control(10)=   "medCuota"
      Item(0).Control(11)=   "txtGarantia"
      Item(0).Control(12)=   "txtComprobante"
      Item(0).Control(13)=   "txtFormalizacion"
      Item(0).Control(14)=   "txtUsuario"
      Item(0).Control(15)=   "Label7(1)"
      Item(0).Control(16)=   "Label6"
      Item(0).Control(17)=   "Label3"
      Item(0).Control(18)=   "Label4"
      Item(0).Control(19)=   "lblPoliza"
      Item(0).Control(20)=   "Label2"
      Item(0).Control(21)=   "txtTasaPiso"
      Item(0).Control(22)=   "txtPtsAddMora"
      Item(0).Control(23)=   "txtPtsAddLiq"
      Item(0).Control(24)=   "txtPtsAddTBP"
      Item(0).Control(25)=   "txtIntM"
      Item(0).Control(26)=   "Label11(1)"
      Item(0).Control(27)=   "Label24"
      Item(0).Control(28)=   "Label10"
      Item(0).Control(29)=   "Label9"
      Item(0).Control(30)=   "Label11(0)"
      Item(0).Control(31)=   "txtFactor"
      Item(0).Control(32)=   "txtDiaPago"
      Item(0).Control(33)=   "txtCuotasMorosidad"
      Item(0).Control(34)=   "txtCuotasAnuladas"
      Item(0).Control(35)=   "txtCuotasPagadas"
      Item(0).Control(36)=   "txtCuotasDeducidas"
      Item(0).Control(37)=   "txtMesFecProceso"
      Item(0).Control(38)=   "txtAnoUltimoAbono"
      Item(0).Control(39)=   "txtMesUltimoAbono"
      Item(0).Control(40)=   "txtAnoPrimerAbono"
      Item(0).Control(41)=   "txtMesPrimerAbono"
      Item(0).Control(42)=   "txtAnioTerminacion"
      Item(0).Control(43)=   "txtMesTerminacion"
      Item(0).Control(44)=   "txtAnoFecProceso"
      Item(0).Control(45)=   "Label22(2)"
      Item(0).Control(46)=   "Label22(3)"
      Item(0).Control(47)=   "Label17"
      Item(0).Control(48)=   "Label16"
      Item(0).Control(49)=   "Label15"
      Item(0).Control(50)=   "Label14"
      Item(0).Control(51)=   "Label19(3)"
      Item(0).Control(52)=   "Label19(2)"
      Item(0).Control(53)=   "Label19(1)"
      Item(0).Control(54)=   "Label19(0)"
      Item(0).Control(55)=   "txtSalidaDesc"
      Item(0).Control(56)=   "txtSalidaTipo"
      Item(1).Caption =   "Datos de la Aprobación"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "lswAutorizadores"
      Item(1).Control(1)=   "rbActas(0)"
      Item(1).Control(2)=   "txtActa"
      Item(1).Control(3)=   "rbActas(1)"
      Item(1).Control(4)=   "rbActas(2)"
      Item(1).Control(5)=   "Label1(35)"
      Item(1).Control(6)=   "Label1(9)"
      Item(1).Control(7)=   "txtActaFecha"
      Item(1).Control(8)=   "cboActa"
      Item(2).Caption =   "Otros"
      Item(2).ControlCount=   14
      Item(2).Control(0)=   "txtActividadDesc"
      Item(2).Control(1)=   "txtCanalDesc"
      Item(2).Control(2)=   "txtActividad"
      Item(2).Control(3)=   "txtCanal"
      Item(2).Control(4)=   "Label22(8)"
      Item(2).Control(5)=   "Label22(9)"
      Item(2).Control(6)=   "txtComiteDesc"
      Item(2).Control(7)=   "txtComite"
      Item(2).Control(8)=   "Label22(10)"
      Item(2).Control(9)=   "txtColocadorDesc"
      Item(2).Control(10)=   "txtColocador"
      Item(2).Control(11)=   "Label22(11)"
      Item(2).Control(12)=   "txtAntiguedad"
      Item(2).Control(13)=   "Label18(1)"
      Item(3).Caption =   "Bancos"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "txtBancos"
      Begin XtremeSuiteControls.ListView lswAutorizadores 
         Height          =   2172
         Left            =   -66280
         TabIndex        =   86
         Top             =   480
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441793
         _ExtentX        =   16954
         _ExtentY        =   3831
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
         Appearance      =   16
      End
      Begin VB.TextBox txtAnoFecProceso 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   9960
         TabIndex        =   75
         ToolTipText     =   "Año de la actual fecha de proceso"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtMesTerminacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   324
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtAnioTerminacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   324
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtMesPrimerAbono 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   10560
         TabIndex        =   72
         ToolTipText     =   "Mes del Primer Abono"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtAnoPrimerAbono 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   9960
         TabIndex        =   71
         ToolTipText     =   "Año del primer Abono"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtMesUltimoAbono 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   10560
         TabIndex        =   70
         ToolTipText     =   "Mes del último abono"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtAnoUltimoAbono 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   9960
         TabIndex        =   69
         ToolTipText     =   "Año del último Abono"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtMesFecProceso 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   10560
         TabIndex        =   68
         ToolTipText     =   "Mes de la Actual fecha de proceso"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtCuotasDeducidas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   12840
         TabIndex        =   67
         ToolTipText     =   "Deducciones Realizadas al Préstamo"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCuotasPagadas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   12840
         TabIndex        =   66
         ToolTipText     =   "Cuotas Pagadas"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtCuotasAnuladas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   12840
         TabIndex        =   65
         ToolTipText     =   "Cuotas Anuladas"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtCuotasMorosidad 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   12840
         TabIndex        =   64
         ToolTipText     =   "Cuotas de Morosidad"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtDiaPago 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtFactor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtIntM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtPtsAddTBP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "Interes Moratorio del préstamo"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtPtsAddLiq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   54
         ToolTipText     =   "Interes Moratorio del préstamo"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtPtsAddMora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   53
         ToolTipText     =   "Interes Moratorio del préstamo"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTasaPiso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   324
         Left            =   4560
         TabIndex        =   45
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtFormalizacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   324
         Left            =   4560
         TabIndex        =   44
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtComprobante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   324
         Left            =   4560
         TabIndex        =   43
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtGarantia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtSalidaDesc 
         Appearance      =   0  'Flat
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
         Height          =   324
         Left            =   3240
         TabIndex        =   32
         Top             =   2280
         Width           =   5052
      End
      Begin VB.TextBox txtSalidaTipo 
         Appearance      =   0  'Flat
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
         Height          =   324
         Left            =   1560
         TabIndex        =   31
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtPlazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2400
         TabIndex        =   30
         ToolTipText     =   "Plazo del préstamo (Meses)"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtInteresC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2400
         TabIndex        =   29
         ToolTipText     =   "Interés Corriente del préstamo (%)"
         Top             =   1200
         Width           =   855
      End
      Begin XtremeSuiteControls.RadioButton rbActas 
         Height          =   252
         Index           =   0
         Left            =   -68680
         TabIndex        =   87
         Top             =   1560
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Resoluciones"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtActa 
         Height          =   312
         Left            =   -68680
         TabIndex        =   88
         Top             =   480
         Visible         =   0   'False
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton rbActas 
         Height          =   252
         Index           =   1
         Left            =   -68680
         TabIndex        =   89
         Top             =   1920
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Autorizadores"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.RadioButton rbActas 
         Height          =   252
         Index           =   2
         Left            =   -68680
         TabIndex        =   90
         Top             =   2280
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Asistencia"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtActaFecha 
         Height          =   312
         Left            =   -68680
         TabIndex        =   93
         Top             =   840
         Visible         =   0   'False
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboActa 
         Height          =   312
         Left            =   -68680
         TabIndex        =   94
         Top             =   1200
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtActividadDesc 
         Height          =   312
         Left            =   -66520
         TabIndex        =   95
         Top             =   600
         Visible         =   0   'False
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCanalDesc 
         Height          =   312
         Left            =   -66520
         TabIndex        =   96
         Top             =   960
         Visible         =   0   'False
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtActividad 
         Height          =   312
         Left            =   -68200
         TabIndex        =   97
         Top             =   600
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCanal 
         Height          =   312
         Left            =   -68200
         TabIndex        =   98
         Top             =   960
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtComiteDesc 
         Height          =   312
         Left            =   -66520
         TabIndex        =   101
         Top             =   1440
         Visible         =   0   'False
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtComite 
         Height          =   312
         Left            =   -68200
         TabIndex        =   102
         Top             =   1440
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtColocadorDesc 
         Height          =   312
         Left            =   -66520
         TabIndex        =   104
         Top             =   1800
         Visible         =   0   'False
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtColocador 
         Height          =   312
         Left            =   -68200
         TabIndex        =   105
         Top             =   1800
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAntiguedad 
         Height          =   312
         Left            =   -58960
         TabIndex        =   107
         Top             =   600
         Visible         =   0   'False
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBancos 
         Height          =   2115
         Left            =   -69760
         TabIndex        =   109
         Top             =   480
         Visible         =   0   'False
         Width           =   12975
         _Version        =   1441793
         _ExtentX        =   22886
         _ExtentY        =   3731
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Antiguedad"
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
         Index           =   1
         Left            =   -60400
         TabIndex        =   108
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Colocador"
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
         Index           =   11
         Left            =   -69760
         TabIndex        =   106
         Top             =   1800
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Comité"
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
         Index           =   10
         Left            =   -69760
         TabIndex        =   103
         Top             =   1440
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Canal"
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
         Index           =   9
         Left            =   -69760
         TabIndex        =   100
         Top             =   960
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Actividad"
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
         Index           =   8
         Left            =   -69760
         TabIndex        =   99
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "No. Acta"
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
         Index           =   9
         Left            =   -69880
         TabIndex        =   92
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
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
         Height          =   252
         Index           =   35
         Left            =   -69880
         TabIndex        =   91
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label19 
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
         Index           =   0
         Left            =   8520
         TabIndex        =   85
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Ultimo Abono"
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
         Left            =   8520
         TabIndex        =   84
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Proceso"
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
         Left            =   8520
         TabIndex        =   83
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Termina"
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
         Left            =   8520
         TabIndex        =   82
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Deducciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   11040
         TabIndex        =   81
         ToolTipText     =   "Deducciones Realizadas al Préstamo"
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label15 
         Caption         =   "Cuotas Pagadas"
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
         Left            =   11040
         TabIndex        =   80
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label Label16 
         Caption         =   "Cuotas Anuladas"
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
         Left            =   11040
         TabIndex        =   79
         Top             =   1200
         Width           =   1692
      End
      Begin VB.Label Label17 
         Caption         =   "Cuotas Atrasadas"
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
         Left            =   11040
         TabIndex        =   78
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label Label22 
         Caption         =   "Día de Pago"
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
         Left            =   8520
         TabIndex        =   77
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Factor de Cálculo"
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
         Left            =   8520
         TabIndex        =   76
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Tasa Mora"
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
         Left            =   6480
         TabIndex        =   61
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label9 
         Caption         =   "Pts Add TBP"
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
         Left            =   6480
         TabIndex        =   60
         ToolTipText     =   "Puntos Adicionales a la TBP para el credito"
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label10 
         Caption         =   "Pts Add Liq"
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
         Left            =   6480
         TabIndex        =   59
         ToolTipText     =   "Puntos Adicionales por Liquidación"
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label Label24 
         Caption         =   "Pts Add Mora"
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
         Left            =   6480
         TabIndex        =   58
         ToolTipText     =   "Puntos Adicionales por Morosidad"
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label11 
         Caption         =   "Tasa Piso"
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
         Left            =   6480
         TabIndex        =   57
         Top             =   1920
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "Póliza"
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
         Left            =   3480
         TabIndex        =   51
         Top             =   1920
         Width           =   972
      End
      Begin VB.Label lblPoliza 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   312
         Left            =   4560
         TabIndex        =   50
         Top             =   1920
         Width           =   1692
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario"
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
         Left            =   3480
         TabIndex        =   49
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label3 
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
         Height          =   252
         Left            =   3480
         TabIndex        =   48
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label6 
         Caption         =   "Documento"
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
         Left            =   3480
         TabIndex        =   47
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         Height          =   252
         Index           =   1
         Left            =   3480
         TabIndex        =   46
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label medCuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   312
         Left            =   1560
         TabIndex        =   41
         Top             =   1560
         Width           =   1692
      End
      Begin VB.Label medMontoApr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   312
         Left            =   1560
         TabIndex        =   40
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label medMontoGirado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   312
         Left            =   1560
         TabIndex        =   39
         Top             =   1920
         Width           =   1692
      End
      Begin VB.Label Label22 
         Caption         =   "Salida"
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
         Left            =   120
         TabIndex        =   38
         Top             =   2280
         Width           =   1212
      End
      Begin VB.Label Label21 
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
         Height          =   252
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label8 
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
         TabIndex        =   36
         Top             =   840
         Width           =   732
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
         Height          =   252
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   2172
      End
      Begin VB.Label Label13 
         Caption         =   "Cuota "
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
         TabIndex        =   34
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Girado"
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
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   1332
      End
   End
   Begin XtremeSuiteControls.PushButton btnDetalles 
      Height          =   492
      Left            =   11040
      TabIndex        =   23
      Top             =   960
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Detalles"
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
      Picture         =   "frmCR_ConsultaDetalle.frx":21CC
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3480
      TabIndex        =   13
      Top             =   840
      Width           =   5772
      _Version        =   1441793
      _ExtentX        =   10181
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
   Begin VB.Timer TimerInicio 
      Interval        =   10
      Left            =   240
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtDestino 
      Height          =   312
      Left            =   3480
      TabIndex        =   14
      Top             =   1560
      Width           =   5772
      _Version        =   1441793
      _ExtentX        =   10181
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
   Begin XtremeSuiteControls.FlatEdit txtOficina 
      Height          =   312
      Left            =   3480
      TabIndex        =   15
      Top             =   1920
      Width           =   5772
      _Version        =   1441793
      _ExtentX        =   10181
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
   Begin XtremeSuiteControls.FlatEdit txtRecurso 
      Height          =   312
      Left            =   3480
      TabIndex        =   16
      Top             =   2280
      Width           =   5772
      _Version        =   1441793
      _ExtentX        =   10181
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
   Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
      Height          =   312
      Left            =   3480
      TabIndex        =   17
      Top             =   1200
      Width           =   5172
      _Version        =   1441793
      _ExtentX        =   9123
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
      Height          =   312
      Left            =   1800
      TabIndex        =   18
      Top             =   840
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   19
      Top             =   1200
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDestinoCod 
      Height          =   312
      Left            =   1800
      TabIndex        =   20
      Top             =   1560
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOficinaCod 
      Height          =   312
      Left            =   1800
      TabIndex        =   21
      Top             =   1920
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRecursoCod 
      Height          =   312
      Left            =   1800
      TabIndex        =   22
      Top             =   2280
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtIntPagado 
      Height          =   312
      Left            =   11040
      TabIndex        =   24
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAmortizado 
      Height          =   312
      Left            =   11040
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
         Weight          =   700
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
      Height          =   312
      Left            =   11040
      TabIndex        =   26
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDivisa 
      Height          =   312
      Left            =   8640
      TabIndex        =   27
      Top             =   1200
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
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
   Begin XtremeSuiteControls.FlatEdit txtCuentaCliente 
      Height          =   435
      Left            =   7080
      TabIndex        =   111
      Top             =   120
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   2
      Left            =   9600
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblAmorticaCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Amortizado"
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
      Left            =   9600
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Interés Pagado"
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
      Index           =   0
      Left            =   9600
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblEstadoPrestamo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Estado del Préstamo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   10680
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Operación:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
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
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina"
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
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Recursos"
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
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IBAN:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image imgPlanPagos 
      Height          =   240
      Left            =   4920
      Picture         =   "frmCR_ConsultaDetalle.frx":298E
      ToolTipText     =   "Plan de Pagos"
      Top             =   240
      Width           =   240
   End
   Begin VB.Label lblOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00000000000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   13575
   End
End
Attribute VB_Name = "frmCR_ConsultaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngOperacion As Long, iOpcion As Integer
Dim vPaso As String


Private Sub btnDetalles_Click()
If vGrid.top > 3000 Then
   vGrid.top = 2640
   vGrid.Height = 5895
Else
   vGrid.top = 5520
   vGrid.Height = 3252
End If
End Sub

Private Sub sbEstudio_Comite_Resolucion_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pTipo As String, itmX As ListViewItem
Dim pExpediente As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lswAutorizadores.Checkboxes = False
lswAutorizadores.ListItems.Clear

With lswAutorizadores.ColumnHeaders
    .Clear
    
    Select Case True
        Case rbActas.Item(0).Value
          pTipo = "RES"
          .Add , , "Estado", 1800, vbCenter
          .Add , , "Fecha", 1800
          .Add , , "Usuario", 2100, vbCenter
          .Add , , "Notas", 3100
        
        Case rbActas.Item(1).Value
          pTipo = "AUT"
          .Add , , "Estado", 1800, vbCenter
          .Add , , "Fecha", 1800
          .Add , , "Identificación", 1800
          .Add , , "Nombre", 3800
          .Add , , "Usuario", 2100, vbCenter
          .Add , , "Notas", 2100
          
        
        Case rbActas.Item(2).Value
          pTipo = "ASI"
          .Add , , "Fecha", 1800
          .Add , , "Identificación", 1800
          .Add , , "Nombre", 3800
          .Add , , "Usuario", 2100, vbCenter
    End Select
    

End With

If cboActa.Text = "Estudio Crédito" Then
    strSQL = "select COD_PREANALISIS " _
           & " From CRD_PREA_PREANALISIS where ID_SOLICITUD = " & lblOperacion.Caption
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
            pExpediente = rs!cod_PreAnalisis
    Else
            pExpediente = ""
    End If
    rs.Close
    
    strSQL = "exec spCrd_Estudio_Resolucion_Detalle '" & pExpediente & "', '" & pTipo & "'"
Else
    pExpediente = lblOperacion.Caption
    strSQL = "exec spCrd_SGT_Resolucion_Detalle " & lblOperacion.Caption & ", '" & pTipo & "'"
End If

Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    txtActa.Text = rs!acta
    txtActaFecha.Text = Format(rs!Acta_Fecha, "dd/MM/yyyy")
End If

Do While Not rs.EOF

    Select Case True
     Case rbActas.Item(0).Value 'Resoluciones
        
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Estado)
            itmX.SubItems(1) = rs!Registro_Fecha
            itmX.SubItems(2) = rs!Registro_Usuario
            itmX.SubItems(3) = rs!Notas
     
     Case rbActas.Item(1).Value 'Autorizaciones
        
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Estado)
            itmX.SubItems(1) = rs!Registro_Fecha
            itmX.SubItems(2) = rs!Cedula
            itmX.SubItems(3) = rs!Nombre
            itmX.SubItems(4) = rs!Registro_Usuario
            itmX.SubItems(5) = rs!Notas
            
          
        
     Case rbActas.Item(2).Value 'Asistencia
   
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Registro_Fecha)
            itmX.SubItems(1) = rs!Cedula
            itmX.SubItems(2) = rs!Nombre
            itmX.SubItems(3) = rs!Registro_Usuario
    
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


Private Sub cboActa_Click()
If vPaso Then Exit Sub

Call sbEstudio_Comite_Resolucion_Load

End Sub

Private Sub chkMoraCancelada_Click()
Call vGrid_SheetChanged(1, 1)
End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

lblOperacion.Caption = Operacion.OperacionConsulta

vPaso = True
    rbActas.Item(0).Value = True

    cboActa.Clear
    cboActa.AddItem "Estudio Crédito"
    cboActa.AddItem "Trámite Crédito"
    cboActa.Text = "Estudio Crédito"
    
    tcMain.Item(0).Selected = True
vPaso = False

End Sub


Private Sub sbCargaInicial()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSys_Consulta_Integrada_Creditos_Detalle " & lblOperacion.Caption
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    
    If rs!Plazo = 999 Then
        lblAmorticaCaption.Caption = "Recaudado"
    Else
        lblAmorticaCaption.Caption = "Amortizado"
    End If
    
    
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    txtCodigo.Text = rs!Codigo
    
    
    txtLineaDesc.Text = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
    txtDivisa.Text = rs!COD_DIVISA & ""
    txtGarantia.Text = rs!GarantiaDesc & ""
    
    txtDestinoCod.Text = rs!cod_destino & ""
    txtDestino.Text = rs!DestinoDesc
       
    txtOficinaCod.Text = rs!cod_oficina_r & ""
    txtOficina.Text = rs!OficinaDesc & ""
    
    txtRecursoCod.Text = rs!Cod_Grupo & ""
    txtRecurso.Text = rs!RecursoDesc & ""
    
    txtActividad.Text = rs!Cod_actividad & ""
    txtActividadDesc.Text = rs!ActividadDesc & ""
    
    txtCanal.Text = rs!Canal_Tipo & ""
    txtCanalDesc.Text = rs!CanalDesc & ""
    
    
    txtComite.Text = rs!id_Comite & ""
    txtComiteDesc.Text = rs!ComiteDesc
    
    txtColocador.Text = rs!ID_PROMOTOR & ""
    txtColocadorDesc.Text = rs!PromotorDesc
    
    txtAntiguedad.Text = rs!Antiguedad
    
    txtUsuario.Text = UCase(IIf(IsNull(rs!Userfor), "", rs!Userfor))
    
    txtFormalizacion.Text = IIf(IsNull(rs!FechaForp), "", Format(rs!FechaForp, "dd/mm/yyyy"))
    txtComprobante.Text = IIf(IsNull(rs!TDOCUMENTO), "", rs!TDOCUMENTO) & "-" & IIf(IsNull(rs!nDocumento), "", rs!nDocumento)
    
    txtPlazo.Text = CStr(rs!Plazo)
    txtInteresC.Text = Format(rs!interesv, "Standard") 'rs!int
    
    txtCuentaCliente.Text = rs!CUENTA_IBAN & "" 'CUENTA_CLIENTE
    
    'Etiquetas
    medMontoGirado.Caption = Format(rs!monto_girado, "Standard")
    medCuota.Caption = Format(rs!Cuota, "Standard")
    txtIntPagado.Text = Format(rs!interesc, "Standard")
    txtAmortizado.Text = Format(rs!Amortiza, "Standard")
    
    medMontoApr.Caption = Format(rs!monto_credito, "Standard")
    txtSaldo.Text = Format(rs!Saldo_Credito, "Standard")
        
    lblPoliza.Caption = Format(rs!Poliza_Cuota, "Standard")
        
    
    'Desembolso
    If IsNull(rs!documento_referido) Then
       txtSalidaTipo.Text = ""
       txtSalidaDesc.Text = ""
    Else
       txtSalidaTipo.Text = rs!Salida_Tipo & ".." & rs!documento_referido & ""
       txtSalidaDesc.Text = rs!Salida_Desc
    End If
    
    txtCuotasDeducidas.Text = IIf(IsNull(rs!cuotas_planilla), "", rs!cuotas_planilla)
    txtCuotasPagadas.Text = IIf(IsNull(rs!cuotas_directas), "", rs!cuotas_directas)
    txtCuotasAnuladas.Text = IIf(IsNull(rs!CUOTAS_ANULADAS), "", rs!CUOTAS_ANULADAS)
    
    
    lblEstadoPrestamo.Caption = rs!Estado_Desc
    
    If rs!MoraCuotas = 0 Then
            lblEstadoPrestamo.Caption = lblEstadoPrestamo.Caption & "¦ Al Día"
    Else
            lblEstadoPrestamo.Caption = lblEstadoPrestamo.Caption & "¦ Mora"
    End If
    
    txtCuotasMorosidad.Text = rs!MoraCuotas
    
    If IsNull(rs!TBP_PuntosAdd) Then
        lblTasa.Caption = "Tasa %"
        txtPtsAddTBP.Text = "N/A"
    Else
        lblTasa.Caption = "Tasa :TBP + " & rs!TBP_PuntosAdd & " pts"
        txtPtsAddTBP.Text = rs!TBP_PuntosAdd
    End If
    
    Select Case rs!Tasa_Mora_Tipo
        Case "PTS"
            txtPtsAddMora.Text = rs!tasa_mora_add & " pts"
            txtIntM.Text = Format(rs!tasa_mora_add + rs!interesv, "Standard")
        Case "POR"
            txtPtsAddMora.Text = rs!tasa_mora_add & "%"
            txtIntM = Format(rs!interesv + (rs!interesv * rs!tasa_mora_add / 100), "###.00")
        Case "N/A"
            txtPtsAddMora.Text = "0"
            txtIntM.Text = Format(0, "###.00")
        Case "TF"
            txtPtsAddMora.Text = "0"
            txtIntM.Text = Format(rs!tasa_mora_add, "###.00")
    End Select

    
    If IsNull(rs!Tasa_Piso) Then
       txtTasaPiso.Text = "N/A"
    Else
       If rs!Tasa_Piso = 0 Then
           txtTasaPiso.Text = "N/A"
       Else
           txtTasaPiso.Text = rs!Tasa_Piso
       End If
    End If
     
    
    If rs!LiqTasaX = 0 Then
        lblTasa.Caption = lblTasa.Caption & " + LiqPts"
        txtPtsAddLiq.Text = 0
    Else
        txtPtsAddLiq.Text = rs!Liq_Valor
    End If

    txtAnoPrimerAbono = Mid(rs!PriDeduc, 1, 4)
    txtMesPrimerAbono = Mid(rs!PriDeduc, 5, 6)
    
    txtAnoUltimoAbono = Mid(rs!FecUlt, 1, 4)
    txtMesUltimoAbono = Mid(rs!FecUlt, 5, 6)

    txtAnoFecProceso = Mid(GLOBALES.glngFechaCR, 1, 4)
    txtMesFecProceso = Mid(GLOBALES.glngFechaCR, 5, 6)

    
    Select Case rs!Estado
      Case "A"
        txtMesTerminacion.Text = Month(rs!Termina)
        txtAnioTerminacion.Text = Year(rs!Termina)
      Case "C"
        txtMesTerminacion.Text = Mid(rs!FecUlt, 5, 2)
        txtAnioTerminacion.Text = Mid(rs!FecUlt, 1, 4)
    End Select

    txtDiaPago.Text = IIf(rs!dia_pago = 32, "Ult.Dia.Mes.", "Todos los " & rs!dia_pago)
   txtFactor.Text = rs!Base_Calculo_Desc

End If
rs.Close

Call vGrid_SheetChanged(1, 1)

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub imgPlanPagos_Click()
Operacion.OperacionConsulta = lblOperacion.Caption
frmCR_PlanPagos.Show vbModal
End Sub


Private Sub rbActas_Click(Index As Integer)

If vPaso Then Exit Sub

Call sbEstudio_Comite_Resolucion_Load

End Sub

Private Sub sbBancos_Status_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


 txtBancos = ""

 strSQL = "select R.id_solicitud,R.codigo,R.cedula,R.monto_girado,C.descripcion,S.nombre" _
        & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
        & " inner join Catalogo C on R.codigo = C.codigo" _
        & " where R.id_solicitud = " & lblOperacion.Caption & " and R.estado in('A','C')"
 Call OpenRecordSet(rs, strSQL)
 
 If Not rs.EOF And Not rs.BOF Then
   
   txtBancos.Tag = rs!Codigo
   txtBancos = txtBancos & "Línea         : " & rs!Codigo & vbCrLf
   txtBancos = txtBancos & "Descripción   : " & rs!Descripcion & vbCrLf
   txtBancos = txtBancos & "Identificación: " & rs!Cedula & vbCrLf
   txtBancos = txtBancos & "Nombre        : " & rs!Nombre & vbCrLf
   txtBancos = txtBancos & "Monto a Girar : " & Format(rs!monto_girado, "Standard") & vbCrLf & vbCrLf & vbCrLf
   rs.Close
    
   'Remesa
   strSQL = "select Td.*, T.ESTADO, T.USUARIO , T.FECHA_INICIO , T.FECHA_CORTE, T.FECHA" _
          & "  from CRD_REMESAS_TES T inner join CRD_REMESAS_TES_DETALLE Td on T.COD_REMESA = Td.COD_REMESA" _
          & "  Where Td.Id_solicitud = " & lblOperacion.Caption
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
      txtBancos = txtBancos & ":::. REMESA DE PAGO .::" & vbCrLf & vbCrLf
      txtBancos = txtBancos & "Remesa Id      : " & rs!cod_remesa & vbCrLf
      txtBancos = txtBancos & "Estado         : "
      Select Case rs!Estado
        Case "A"
            txtBancos = txtBancos & "Abierta" & vbCrLf
        Case "C"
            txtBancos = txtBancos & "Cerrada" & vbCrLf
        Case "T"
            txtBancos = txtBancos & "Trasladada" & vbCrLf
      End Select
      txtBancos = txtBancos & "Fecha Creación : " & rs!fecha & vbCrLf
      txtBancos = txtBancos & "Usuario        : " & rs!Usuario & vbCrLf
      txtBancos = txtBancos & "Monto          : " & Format(rs!Monto, "Standard") & vbCrLf
      txtBancos = txtBancos & "Desembolsos Add: " & Format(rs!DESEMBOLSOS, "Standard") & vbCrLf
      txtBancos = txtBancos & "Tesoreria Id   : " & rs!NSolicitud & vbCrLf & vbCrLf & vbCrLf
   Else
      txtBancos = txtBancos & ">>> REMESA DE PAGO: NO SE LOCALIZO NINGUNA <<<" & vbCrLf & vbCrLf
   End If
   rs.Close
   
   'Bancos
   strSQL = "select T.NSOLICITUD,T.id_banco,T.ndocumento,T.BENEFICIARIO, T.FECHA_SOLICITUD, T.FECHA_EMISION , T.ESTADO " _
          & ", '[' + B.CTA  + '] ' + B.DESCRIPCION   as 'Cuenta_Desc' , Bg.DESCRIPCION as 'Banco_Desc', Td.DESCRIPCION as 'Tipo_Desc'" _
          & ", case when T.Estado in('P','S') then 'SOLICITADA' when T.Estado in('T','E', 'I') then 'EMITIDO' when T.Estado in('A','N') then 'ANULADA' end as 'ESTADO_DESC' " _
          & ", T.DOCUMENTO_BASE" _
          & "  from Tes_Transacciones T inner join TES_BANCOS B on T.ID_BANCO = B.ID_BANCO" _
          & "      inner join TES_BANCOS_GRUPOS Bg on B.COD_GRUPO = Bg.COD_GRUPO" _
          & "      inner join TES_TIPOS_DOC Td on T.tipo = Td.TIPO" _
          & " where T.op = " & lblOperacion.Caption & " and T.estado in('I','T','P', 'E')"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
      txtBancos = txtBancos & ":::. BANCOS .::" & vbCrLf & vbCrLf
      txtBancos = txtBancos & "Estado       : " & rs!Estado_Desc & vbCrLf
      txtBancos = txtBancos & "Solicitud    : " & rs!NSolicitud & vbCrLf
      txtBancos = txtBancos & "Documento    : " & rs!nDocumento & "     TF: " & rs!DOCUMENTO_BASE & vbCrLf
      txtBancos = txtBancos & "Beneficiario : " & rs!Beneficiario & vbCrLf
      txtBancos = txtBancos & "Banco        : " & rs!Banco_Desc & vbCrLf
      txtBancos = txtBancos & "Cuenta       : " & rs!Cuenta_Desc & vbCrLf
      txtBancos = txtBancos & "Tipo         : " & rs!Tipo_Desc & vbCrLf
      txtBancos = txtBancos & "Fec.Solicita : " & Format(rs!fecha_solicitud, "yyyy-MM-dd hh:mm:ss") & vbCrLf
      txtBancos = txtBancos & "Fec.Emite    : " & Format(rs!Fecha_Emision & "", "yyyy-MM-dd hh:mm:ss") & vbCrLf
      
   End If
   rs.Close

 End If


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 1
    Call sbEstudio_Comite_Resolucion_Load
  Case 3
    Call sbBancos_Status_Load
End Select

End Sub

Private Sub TimerInicio_Timer()
TimerInicio.Interval = 0
TimerInicio.Enabled = False

Call sbCargaInicial

End Sub

Private Sub tlbDetalle_ButtonClick(ByVal Button As MSComctlLib.Button)

If Button.Image = 4 Then
   Button.Image = 3
   vGrid.top = 2640
   vGrid.Height = 5895
Else
   Button.Image = 4
   vGrid.top = 5160
   vGrid.Height = 3375
End If

End Sub


Private Sub sbMovimientos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

With vGrid

 .ActiveSheet = 1
 .Sheet = 1
 .MaxRows = 0

strSQL = "exec spCrd_Movimientos_New " & lblOperacion.Caption & ", 1"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 
 .MaxRows = .MaxRows + 1
 .Row = .MaxRows
 
 For i = 1 To 16
    .col = i
    Select Case i
      Case 1 ' Proceso
            .Text = Format(rs!Proceso, "####-##")
      Case 2 ' Fecha Corte
            .Text = Format(rs!Fecha_Corte, "dd/mm/yyyy")
      Case 3 ' No.Cuota
            .Text = rs!Num_Cuota
      
      Case 4 ' Fecha Mov.
            .Text = Format(rs!fecha, "dd/mm/yyyy")
      
      Case 5 'Movimiento Total
            .Text = Format(rs!Total, "Standard")
      Case 6 ' Int.Cor
            .Text = Format(rs!IntCor, "Standard")
      Case 7 ' Int.Mor
            .Text = Format(rs!IntMor, "Standard")
      Case 8 ' Amortiza
            .Text = Format(rs!Principal, "Standard")
      Case 9 ' Poliza
            .Text = Format(rs!Poliza, "Standard")
      Case 10 ' Cargo
            .Text = Format(rs!Cargo, "Standard")
      Case 11 ' Saldo
            .Text = Format(rs!Saldo, "Standard")
      Case 12 ' Tipo Doc.
            .Text = Trim(rs!tcon2 & "")
      Case 13 ' Numero Doc.
            .Text = rs!nCon & ""
    
      Case 14 'Concepto
            .Text = rs!CONCEPTO & ""
      Case 15 'Usuario
            .Text = rs!Usuario & ""
      Case 16 'Cajas
            .Text = rs!Cajas & ""
    
    End Select
 
 Next i
 
 If chkMoraCancelada.Value = vbChecked And rs!Tipo = "O" Then
      'Borra la ultima linea porque no es morosidad
      .MaxRows = .MaxRows - 1
 End If
 
 
 rs.MoveNext
Loop
rs.Close

End With

End Sub



Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vGrid.ActiveSheet = NewSheet
vGrid.MaxRows = 0

If NewSheet = 1 Then
   chkMoraCancelada.Visible = True
Else
   chkMoraCancelada.Visible = False
End If

Select Case NewSheet
  Case 1 'Movimientos
    Call sbMovimientos
    
  Case 2 'Morosidad
    If GLOBALES.SysPlanPagos = 1 Then
         strSQL = "select SUBSTRING ( CONVERT(varchar(10), FECHA_PROCESO),1,4) + ' - ' + SUBSTRING ( CONVERT(varchar(10), FECHA_PROCESO),5,2)" _
                & ",INTCOR,INTMOR,PRINCIPAL,POLIZA,CARGOS,(INTCOR+INTMOR+PRINCIPAL+CARGOS+POLIZA)" _
                & " From CRD_OPERACION_TRANSAC" _
                & " where MORA_DIAS > 0 and ESTADO = 'A' and ID_SOLICITUD = " & lblOperacion.Caption & " order by FECHA_PROCESO"

    
    Else
         strSQL = "select SUBSTRING ( CONVERT(varchar(10), FECHAP),1,4) + ' - ' + SUBSTRING ( CONVERT(varchar(10), FECHAP),5,2)" _
                & ", INTC,INTM,AMORTIZA,0 as 'POLIZA',CARGO, (INTC + INTM + AMORTIZA + CARGO)" _
                & " from morosidad where estado = 'A' and id_solicitud = " & lblOperacion.Caption _
                & " order by Fechap desc"
    End If
    Call sbCargaGridFps7(vGrid, 7, strSQL, True, NewSheet, 0)
  
  Case 3 'Cierre
   strSQL = "select Anio, case when MES = 1 then 'Enero' when MES = 2 then 'Febrero' when MES = 3 then 'Marzo'" _
          & "   when MES = 4 then 'Abril' when MES = 5 then 'Mayo' when MES = 6 then 'Junio'" _
          & "   when MES = 7 then 'Julio' when MES = 8 then 'Agosto' when MES = 9 then 'Septiembre'" _
          & "   when MES = 10 then 'Octubre' when MES = 11 then 'Noviembre' when MES = 12 then 'Diciembre' end" _
          & "    ,SALDO_FINAL,TOTAL_DEBITOS, TOTAL_CREDITOS , SALDO_FINAL, Case when OPEX = 1 then 'Si' else 'No' end" _
          & "    , case when PROCESO = 'N' then 'Normal' when PROCESO = 'T' then 'Tra.Deuda' when PROCESO = 'J' then 'Cbr.Jud.' else 'Incobrable' end as 'Proceso'" _
          & "    , case when estado = 'A' then 'Activa' when estado = 'C' then 'Cancelada' when estado = 'N' then 'Anulada' else 'En Tramite' end as 'Estado'" _
          & "    , SUBSTRING ( CONVERT(varchar(10), PRIDEDUC),1,4) + ' - ' + SUBSTRING ( CONVERT(varchar(10), PRIDEDUC),5,2)" _
          & "    , SUBSTRING ( CONVERT(varchar(10), FECULT ),1,4) + ' - ' + SUBSTRING ( CONVERT(varchar(10), FECULT),5,2)" _
          & "    ,TASA, PLAZO, CUOTA" _
          & "  From ASE_PER_CERRADOS" _
          & "  Where ID_SOLICITUD = " & lblOperacion.Caption _
          & "  order by ANIO desc, MES desc"
    
    Call sbCargaGridFps7(vGrid, 14, strSQL, True, NewSheet, 0)
  
  Case 4 'Correcciones
   strSQL = "select FECHA,dbo.fxCrdMovimientoCorrectivo(MOVIMIENTO), USUARIO,DETALLE,NOTAS " _
          & " from credito_suBit where id_solicitud = " & lblOperacion.Caption _
          & " order by fecha desc"
    Call sbCargaGridFps7(vGrid, 5, strSQL, True, NewSheet, 0, True)
    
  Case 5 'Fiadores
   strSQL = "select S.cedula,S.nombre,E.descripcion as Estado,I.descripcion as Inst " _
          & " from fiadores F inner join Socios S on F.cedulaf = S.cedula" _
          & " inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
          & " inner join AFI_ESTADOS_PERSONA E on E.cod_estado = S.estadoActual" _
          & "  where F.estado = 'A' and F.id_solicitud = " & lblOperacion.Caption
   Call sbCargaGridFps7(vGrid, 4, strSQL, True, NewSheet, 0)
  
  Case 6 'Refundiciones
   strSQL = "select ID_SOLICITUD,CODIGO,0 AS INTCOR, 0 INTMOR,0 AS CARGO, MONTO  from REFUNDE_RETENCION" _
          & "   Where ID_SOLICITUDR = " & lblOperacion.Caption _
          & "  Union" _
          & "  select ID_SOLICITUD,CODIGO,INTCOR,INTMOR,isnull(CARGOS,0) as CARGO,MONTO from REFUNDICIONES" _
          & "   WHERE ID_SOLICITUDR = " & lblOperacion.Caption
   Call sbCargaGridFps7(vGrid, 6, strSQL, True, NewSheet, 0)
  
  Case 7 'Desembolsos
    strSQL = "SELECT CONCEPTO,MONTO,CUENTA_CONTA,RETENER,MODIFICA" _
           & " From DESEMBOLSOS where id_solicitud = " & lblOperacion.Caption
    Call sbCargaGridFps7(vGrid, 5, strSQL, True, NewSheet, 0)
  
  Case 8 'Tags
    strSQL = "select T.TAG_CODIGO,T.DESCRIPCION, O.REGISTRO_FECHA,O.REGISTRO_USUARIO , O.NOTAS" _
           & " from CRD_TAGS T inner join CRD_OPERACION_TAGS O on T.TAG_CODIGO = O.TAG_CODIGO" _
           & " Where O.ID_SOLICITUD = " & lblOperacion.Caption
   Call sbCargaGridFps7(vGrid, 5, strSQL, True, NewSheet, 0, True)
    
End Select

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
