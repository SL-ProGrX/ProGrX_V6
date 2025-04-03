VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmCO_PrincipalX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobro Administrativo y Judicial"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   HelpContextID   =   4001
   Icon            =   "frmCO_PrincipalX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   9060
   Begin TabDlg.SSTab ssTabPrincipal 
      Height          =   4455
      Left            =   0
      TabIndex        =   20
      Top             =   1680
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      ForeColor       =   12582912
      TabCaption(0)   =   "Estado de Cuenta"
      TabPicture(0)   =   "frmCO_PrincipalX.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lswAvisos"
      Tab(0).Control(1)=   "txtTotalMora"
      Tab(0).Control(2)=   "txtAmortizacionAtrasada"
      Tab(0).Control(3)=   "txtInteresesMoratorios"
      Tab(0).Control(4)=   "cmdEnviarACobroJudicial"
      Tab(0).Control(5)=   "txtCuotasAnuladas"
      Tab(0).Control(6)=   "txtCuotasDirectas"
      Tab(0).Control(7)=   "txtCuotasPlanilla"
      Tab(0).Control(8)=   "txtInteresActual"
      Tab(0).Control(9)=   "txtMontoRecalculo"
      Tab(0).Control(10)=   "txtPlazoRecalculo"
      Tab(0).Control(11)=   "txtCuota"
      Tab(0).Control(12)=   "txtInteresPorcentaje"
      Tab(0).Control(13)=   "txtPlazo"
      Tab(0).Control(14)=   "txtAmortizado"
      Tab(0).Control(15)=   "txtInteresPagado"
      Tab(0).Control(16)=   "txtSaldo"
      Tab(0).Control(17)=   "txtMonto"
      Tab(0).Control(18)=   "Label3(2)"
      Tab(0).Control(19)=   "Label3(0)"
      Tab(0).Control(20)=   "Label3(1)"
      Tab(0).Control(21)=   "Label3(3)"
      Tab(0).Control(22)=   "lblObservacionesCBR"
      Tab(0).Control(23)=   "Label1(14)"
      Tab(0).Control(24)=   "lblFechaCBR"
      Tab(0).Control(25)=   "Label1(13)"
      Tab(0).Control(26)=   "lblOpex"
      Tab(0).Control(27)=   "lblProceso"
      Tab(0).Control(28)=   "Label2(12)"
      Tab(0).Control(29)=   "Label2(11)"
      Tab(0).Control(30)=   "Label2(10)"
      Tab(0).Control(31)=   "Line1(0)"
      Tab(0).Control(32)=   "Label6(0)"
      Tab(0).Control(33)=   "Label5"
      Tab(0).Control(34)=   "Label4"
      Tab(0).Control(35)=   "Label2(9)"
      Tab(0).Control(36)=   "Label2(8)"
      Tab(0).Control(37)=   "Label2(7)"
      Tab(0).Control(38)=   "Label2(6)"
      Tab(0).Control(39)=   "Label2(5)"
      Tab(0).Control(40)=   "Label2(4)"
      Tab(0).Control(41)=   "Label2(3)"
      Tab(0).Control(42)=   "Label2(2)"
      Tab(0).Control(43)=   "Label2(1)"
      Tab(0).Control(44)=   "Label2(0)"
      Tab(0).Control(45)=   "Line2"
      Tab(0).Control(46)=   "Label1(12)"
      Tab(0).Control(47)=   "Label1(11)"
      Tab(0).ControlCount=   48
      TabCaption(1)   =   "Dirección"
      TabPicture(1)   =   "frmCO_PrincipalX.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label7(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label7(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ImageList1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtProvincia"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtCanton"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtDistrito"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtEMail"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtApartado"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lswTelefonos"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtDireccion"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Cuotas"
      TabPicture(2)   =   "frmCO_PrincipalX.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboTipoCuotas"
      Tab(2).Control(1)=   "cboVisualiza"
      Tab(2).Control(2)=   "lswAbonos"
      Tab(2).Control(3)=   "imgReporteCuotas"
      Tab(2).Control(4)=   "Label15"
      Tab(2).Control(5)=   "lblCuotas"
      Tab(2).Control(6)=   "Label6(1)"
      Tab(2).Control(7)=   "imgVisualiza"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Fiadores"
      TabPicture(3)   =   "frmCO_PrincipalX.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraTraspaso"
      Tab(3).Control(1)=   "txtDirFiadores"
      Tab(3).Control(2)=   "ssTabAuxFiador"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Reportes"
      TabPicture(4)   =   "frmCO_PrincipalX.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ssTabRep"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Antiguedad de Saldos"
      TabPicture(5)   =   "frmCO_PrincipalX.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cboRepAnt"
      Tab(5).Control(1)=   "cboAntiguedad"
      Tab(5).Control(2)=   "chkSaldo"
      Tab(5).Control(3)=   "optAntiguedad(1)"
      Tab(5).Control(4)=   "optAntiguedad(0)"
      Tab(5).Control(5)=   "fraAntiguedad"
      Tab(5).Control(6)=   "cmdIniciarLista"
      Tab(5).Control(7)=   "cmdCodigosAtrasados"
      Tab(5).Control(8)=   "cmdCatalogo"
      Tab(5).Control(9)=   "chkAntiguedadTodas"
      Tab(5).Control(10)=   "lswAntiguedad"
      Tab(5).Control(11)=   "Line4"
      Tab(5).Control(12)=   "Line3(0)"
      Tab(5).Control(13)=   "imgReporteAntiguedad"
      Tab(5).Control(14)=   "Label17(5)"
      Tab(5).ControlCount=   15
      TabCaption(6)   =   "OP.Generadas"
      TabPicture(6)   =   "frmCO_PrincipalX.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraReversionDeTraspaso"
      Tab(6).Control(1)=   "Frame3"
      Tab(6).Control(2)=   "lswOperacionesGeneradas"
      Tab(6).Control(3)=   "Label17(3)"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Opciones"
      TabPicture(7)   =   "frmCO_PrincipalX.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdDeduccionPlanilla"
      Tab(7).Control(1)=   "chkDeducirPlanilla"
      Tab(7).Control(2)=   "cmdCBRParametros"
      Tab(7).Control(3)=   "chkParMantieneTasaSocios"
      Tab(7).Control(4)=   "chkParMantieneTasaNSocios"
      Tab(7).Control(5)=   "txtParCodigo"
      Tab(7).Control(6)=   "Line8(0)"
      Tab(7).Control(7)=   "Label16(0)"
      Tab(7).Control(8)=   "Label16(1)"
      Tab(7).Control(9)=   "Line8(1)"
      Tab(7).Control(10)=   "Line9(0)"
      Tab(7).Control(11)=   "Line9(1)"
      Tab(7).Control(12)=   "Label28"
      Tab(7).ControlCount=   13
      Begin VB.CommandButton cmdDeduccionPlanilla 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   -68280
         TabIndex        =   153
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkDeducirPlanilla 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Deducir Cuotas por Planilla"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69600
         TabIndex        =   152
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton cmdCBRParametros 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   -68280
         TabIndex        =   151
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CheckBox chkParMantieneTasaSocios 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Mantener Tasa Original del Deudor a fiadores Socios (Traspaso Deudas)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   150
         Top             =   2400
         Width           =   5775
      End
      Begin VB.CheckBox chkParMantieneTasaNSocios 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Mantener Tasa Original del Deudor a fiadores NO Socios (Traspaso Deudas)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   149
         Top             =   2640
         Width           =   5775
      End
      Begin VB.TextBox txtParCodigo 
         Height          =   315
         Left            =   -69000
         TabIndex        =   148
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ComboBox cboRepAnt 
         Height          =   315
         Left            =   -68880
         Style           =   2  'Dropdown List
         TabIndex        =   144
         Top             =   1080
         Width           =   2775
      End
      Begin MSComctlLib.ListView lswAvisos 
         Height          =   1215
         Left            =   -69120
         TabIndex        =   132
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2143
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox cboAntiguedad 
         Height          =   315
         ItemData        =   "frmCO_PrincipalX.frx":09AA
         Left            =   -68880
         List            =   "frmCO_PrincipalX.frx":09B4
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CheckBox chkSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Mora Legal - Sin Intereses"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -70200
         TabIndex        =   130
         Top             =   720
         Width           =   2415
      End
      Begin VB.Frame fraReversionDeTraspaso 
         Height          =   1095
         Left            =   -73920
         TabIndex        =   118
         Top             =   3240
         Visible         =   0   'False
         Width           =   6615
         Begin VB.CommandButton cmdReversaTraspasoDeudas 
            Caption         =   "Reversar Traspaso"
            Height          =   375
            Left            =   4680
            TabIndex        =   123
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtTRAFD_MONTO 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   122
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtTRAFD_Plazo 
            Height          =   315
            Left            =   3840
            TabIndex        =   121
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtTRAFD_Int 
            Height          =   315
            Left            =   3840
            TabIndex        =   120
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtTRAFD_Cuota 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   119
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdCancelaReversionTraspaso 
            Caption         =   "&Cancelar Traspaso"
            Height          =   375
            Left            =   4680
            TabIndex        =   128
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label24 
            Caption         =   "Plazo"
            Height          =   255
            Left            =   3240
            TabIndex        =   127
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label25 
            Caption         =   "Interes"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   126
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label25 
            Caption         =   "Nuevo Monto"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   125
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label25 
            Caption         =   "Nueva Cuota"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   124
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -74880
         TabIndex        =   108
         Top             =   2520
         Width           =   8775
         Begin VB.Label lblOperacionActualDeudor 
            Height          =   285
            Left            =   1560
            TabIndex        =   129
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblSaldoActualDeudor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7320
            TabIndex        =   117
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblInteresActualDeudor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6240
            TabIndex        =   116
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblPlazoActualDeudor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5040
            TabIndex        =   115
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblMontoActualDeudor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   114
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Saldo"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   4
            Left            =   6840
            TabIndex        =   113
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Interés"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   3
            Left            =   5640
            TabIndex        =   112
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Plazo"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   2
            Left            =   4560
            TabIndex        =   111
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Monto"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   110
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "OP Actual Deudor"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.OptionButton optAntiguedad 
         Alignment       =   1  'Right Justify
         Caption         =   "Mora Financiera"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   2160
         Width           =   2775
      End
      Begin VB.OptionButton optAntiguedad 
         Alignment       =   1  'Right Justify
         Caption         =   "Mora Legal"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   1800
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.Frame fraTraspaso 
         Caption         =   "Traspaso de Deudas"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   -74880
         TabIndex        =   95
         Top             =   2640
         Width           =   8775
         Begin VB.CheckBox chkTraspaso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Traspasar la deuda a todos los fiadores por Igual"
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
            Height          =   315
            Left            =   4080
            TabIndex        =   143
            Top             =   960
            Value           =   1  'Checked
            Width           =   4575
         End
         Begin VB.CommandButton cmdTraspasoDeudas 
            Caption         =   "&Traspasar Deuda"
            Height          =   315
            Left            =   6960
            TabIndex        =   142
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtCodigoNuevo 
            Height          =   315
            Left            =   3480
            TabIndex        =   137
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtPorcentajeDeuda 
            Height          =   315
            Left            =   840
            TabIndex        =   100
            Text            =   "100"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtMontoFiador 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   99
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtPlazoFiador 
            Height          =   315
            Left            =   840
            TabIndex        =   98
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtInteresFiador 
            Height          =   315
            Left            =   2280
            TabIndex        =   97
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtCuotaFiador 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   96
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Código"
            Height          =   315
            Index           =   0
            Left            =   2880
            TabIndex        =   141
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblCedulaFiador 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   2880
            TabIndex        =   140
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblNombreFiador 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   4320
            TabIndex        =   139
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label lblCodigoNuevo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   4320
            TabIndex        =   138
            Top             =   600
            Width           =   4335
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFFFF&
            X1              =   2760
            X2              =   2760
            Y1              =   120
            Y2              =   1680
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            BorderWidth     =   2
            Index           =   0
            X1              =   2760
            X2              =   2760
            Y1              =   120
            Y2              =   1680
         End
         Begin VB.Label Label19 
            Caption         =   "Plazo"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Interes"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   104
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Monto"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "% Deuda"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Cuota"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   101
            Top             =   1320
            Width           =   615
         End
      End
      Begin VB.Frame fraAntiguedad 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   -70320
         TabIndex        =   92
         Top             =   3480
         Visible         =   0   'False
         Width           =   4215
         Begin MSComctlLib.ProgressBar prgBar 
            Height          =   215
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   370
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblAntiguedad 
            Caption         =   "..."
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   3975
         End
      End
      Begin VB.TextBox txtTotalMora 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67920
         MultiLine       =   -1  'True
         TabIndex        =   83
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtAmortizacionAtrasada 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67920
         MultiLine       =   -1  'True
         TabIndex        =   82
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtInteresesMoratorios 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67920
         MultiLine       =   -1  'True
         TabIndex        =   81
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton cmdIniciarLista 
         Caption         =   "&Iniciar"
         Height          =   495
         Left            =   -71760
         TabIndex        =   76
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCodigosAtrasados 
         Caption         =   "Códigos Atrasados"
         Height          =   495
         Left            =   -73320
         TabIndex        =   75
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmdCatalogo 
         Caption         =   "&Catalogo"
         Height          =   495
         Left            =   -74880
         TabIndex        =   74
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CheckBox chkAntiguedadTodas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Evaluar todos los Códigos"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -70200
         TabIndex        =   73
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtDirFiadores 
         ForeColor       =   &H00FF0000&
         Height          =   2175
         Left            =   -68520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboTipoCuotas 
         Height          =   315
         ItemData        =   "frmCO_PrincipalX.frx":09F7
         Left            =   -69720
         List            =   "frmCO_PrincipalX.frx":0A07
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboVisualiza 
         Height          =   315
         ItemData        =   "frmCO_PrincipalX.frx":0A31
         Left            =   -73680
         List            =   "frmCO_PrincipalX.frx":0A3B
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   975
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   840
         Width           =   7335
      End
      Begin MSComctlLib.ListView lswTelefonos 
         Height          =   1695
         Left            =   1320
         TabIndex        =   60
         Top             =   2640
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ext"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Contacto"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.TextBox txtApartado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   59
         Top             =   2280
         Width           =   7335
      End
      Begin VB.TextBox txtEMail 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   58
         Top             =   1920
         Width           =   7335
      End
      Begin VB.TextBox txtDistrito 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6960
         TabIndex        =   57
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtCanton 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
         TabIndex        =   56
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtProvincia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   55
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdEnviarACobroJudicial 
         Caption         =   "&Enviar a Cobro Judicial"
         Height          =   735
         Left            =   -67440
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txtCuotasAnuladas 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtCuotasDirectas 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtCuotasPlanilla 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtInteresActual 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69960
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtMontoRecalculo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71040
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtPlazoRecalculo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69960
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCuota 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   1440
         Width           =   1660
      End
      Begin VB.TextBox txtInteresPorcentaje 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1080
         Width           =   1660
      End
      Begin VB.TextBox txtPlazo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   720
         Width           =   1660
      End
      Begin VB.TextBox txtAmortizado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtInteresPagado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.ListView lswAbonos 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   62
         Top             =   1080
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lswOperacionesGeneradas 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   69
         Top             =   600
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3413
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Operación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Interes"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Plazo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Int.Atrasado"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswAntiguedad 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   71
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5318
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   360
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":0A6B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":0D87
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":10A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":13BF
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":16DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":19F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":1D13
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":202F
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":234B
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":174BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":2C62F
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCO_PrincipalX.frx":417A1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab ssTabAuxFiador 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   145
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         _Version        =   393216
         TabOrientation  =   2
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Fiadores"
         TabPicture(0)   =   "frmCO_PrincipalX.frx":56913
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lswFiadores"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Teléfonos"
         TabPicture(1)   =   "frmCO_PrincipalX.frx":5692F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lswTelefonosFiadores"
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView lswFiadores 
            Height          =   2100
            Left            =   330
            TabIndex        =   146
            Top             =   30
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3704
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cédula"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Empleado"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lswTelefonosFiadores 
            Height          =   2100
            Left            =   -74670
            TabIndex        =   147
            Top             =   0
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3704
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Tipo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Número"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Ext"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Contacto"
               Object.Width           =   6068
            EndProperty
         End
      End
      Begin TabDlg.SSTab ssTabRep 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   157
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7011
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Operación"
         TabPicture(0)   =   "frmCO_PrincipalX.frx":5694B
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraFechas"
         Tab(0).Control(1)=   "lswRepOp"
         Tab(0).Control(2)=   "imgReporteOperacion"
         Tab(0).Control(3)=   "lblRepOp"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Generales"
         TabPicture(1)   =   "frmCO_PrincipalX.frx":56967
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label1(15)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label1(4)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label1(2)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lblRepGen"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lblXDescribe"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "imgReporteGeneral"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label22(0)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label22(1)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label22(2)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label1(18)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label1(17)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "lswRepGen"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "cboDestino"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txtReporteX"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtDesde"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txtHasta"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "cboRep"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "chkLineas"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "cboCartera"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "cboGarantia"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).ControlCount=   20
         Begin VB.ComboBox cboGarantia 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   187
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cboCartera 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   960
            Width           =   3375
         End
         Begin VB.CheckBox chkLineas 
            Appearance      =   0  'Flat
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5400
            TabIndex        =   185
            Top             =   1575
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.ComboBox cboRep 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   171
            Top             =   270
            Width           =   3375
         End
         Begin VB.TextBox txtHasta 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6480
            TabIndex        =   170
            Text            =   "80"
            Top             =   3030
            Width           =   615
         End
         Begin VB.TextBox txtDesde 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5280
            TabIndex        =   169
            Text            =   "1"
            Top             =   3030
            Width           =   615
         End
         Begin VB.TextBox txtReporteX 
            Height          =   315
            Left            =   4560
            TabIndex        =   168
            ToolTipText     =   "Presione F4 Para Consultar"
            Top             =   1830
            Width           =   735
         End
         Begin VB.Frame fraFechas 
            Caption         =   "Fechas de Corte"
            ForeColor       =   &H00FF0000&
            Height          =   2415
            Left            =   -69840
            TabIndex        =   160
            Top             =   1080
            Visible         =   0   'False
            Width           =   3495
            Begin VB.ComboBox cboRepX 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   162
               Top             =   1080
               Width           =   2655
            End
            Begin VB.CommandButton cmdAceptarFechas 
               Height          =   615
               Left            =   2640
               Picture         =   "frmCO_PrincipalX.frx":56983
               Style           =   1  'Graphical
               TabIndex        =   161
               Top             =   1680
               Width           =   735
            End
            Begin MSComCtl2.DTPicker dtpFechaInicio 
               Height          =   315
               Left            =   720
               TabIndex        =   163
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   20381699
               CurrentDate     =   36431
            End
            Begin MSComCtl2.DTPicker dtpFechaCorte 
               Height          =   315
               Left            =   720
               TabIndex        =   164
               Top             =   720
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   20381699
               CurrentDate     =   36431
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Filtro"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   167
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Corte"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   166
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Inicio"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   165
               Top             =   360
               Width           =   615
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   3360
               X2              =   120
               Y1              =   1560
               Y2              =   1560
            End
         End
         Begin VB.ComboBox cboDestino 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   2160
            Width           =   3375
         End
         Begin MSComctlLib.ListView lswRepOp 
            Height          =   3135
            Left            =   -74880
            TabIndex        =   159
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   5530
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Reporte"
               Object.Width           =   8362
            EndProperty
         End
         Begin MSComctlLib.ListView lswRepGen 
            Height          =   3135
            Left            =   120
            TabIndex        =   172
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   5530
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Reporte"
               Object.Width           =   7304
            EndProperty
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Garantía"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   17
            Left            =   4560
            TabIndex        =   189
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cartera"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   18
            Left            =   4560
            TabIndex        =   188
            Top             =   960
            Width           =   735
         End
         Begin VB.Image imgReporteOperacion 
            Height          =   480
            Left            =   -66960
            Picture         =   "frmCO_PrincipalX.frx":5724D
            Stretch         =   -1  'True
            Top             =   2910
            Width           =   480
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Estado"
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
            Height          =   315
            Index           =   2
            Left            =   4560
            TabIndex        =   181
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde"
            Height          =   315
            Index           =   1
            Left            =   4680
            TabIndex        =   180
            Top             =   3030
            Width           =   615
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta"
            Height          =   315
            Index           =   0
            Left            =   5880
            TabIndex        =   179
            Top             =   3030
            Width           =   615
         End
         Begin VB.Image imgReporteGeneral 
            Height          =   480
            Left            =   8040
            Picture         =   "frmCO_PrincipalX.frx":57B17
            Stretch         =   -1  'True
            Top             =   2880
            Width           =   480
         End
         Begin VB.Label lblXDescribe 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5280
            TabIndex        =   178
            Top             =   1830
            Width           =   3375
         End
         Begin VB.Label lblRepOp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -74880
            TabIndex        =   177
            Top             =   120
            Width           =   4815
         End
         Begin VB.Label lblRepGen 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   120
            Width           =   4335
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Línea"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   175
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Destino"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   4560
            TabIndex        =   174
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuotas de Atraso"
            ForeColor       =   &H00FF0000&
            Height          =   795
            Index           =   15
            Left            =   4560
            TabIndex        =   173
            Top             =   2640
            Width           =   2655
         End
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -66840
         X2              =   -73680
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF8080&
         Caption         =   "Opciones de la Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   -74520
         TabIndex        =   156
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF8080&
         Caption         =   "Opciones del Módulo de Cobro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   -74520
         TabIndex        =   155
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -66960
         X2              =   -73800
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -74520
         X2              =   -70440
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74520
         X2              =   -70440
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label Label28 
         Caption         =   "Código x Omisión en Traspaso de Deudas"
         Height          =   255
         Left            =   -72840
         TabIndex        =   154
         Top             =   3000
         Width           =   4095
      End
      Begin VB.Label Label7 
         Caption         =   "Teléfonos"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   136
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Dirección"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   135
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos en Mora"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   -69240
         TabIndex        =   77
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos de Recálculos"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   38
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Información Deducciones "
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   -74880
         TabIndex        =   39
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Historial de Avisos"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   -69120
         TabIndex        =   133
         Top             =   360
         Width           =   3015
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   -70320
         X2              =   -66000
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   -70320
         X2              =   -66000
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label lblObservacionesCBR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   -71280
         TabIndex        =   91
         Top             =   3840
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Observaciones CBR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   -71280
         TabIndex        =   90
         Top             =   3600
         Width           =   3735
      End
      Begin VB.Label lblFechaCBR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -72720
         TabIndex        =   89
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fecha Envio CBR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   -74880
         TabIndex        =   88
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label lblOpex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -72720
         TabIndex        =   86
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblProceso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   85
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Total"
         Height          =   255
         Index           =   12
         Left            =   -69000
         TabIndex        =   80
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Principal"
         Height          =   255
         Index           =   11
         Left            =   -69000
         TabIndex        =   79
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Intereses"
         Height          =   255
         Index           =   10
         Left            =   -69000
         TabIndex        =   78
         Top             =   2280
         Width           =   735
      End
      Begin VB.Image imgBusqueda_Rapida 
         Height          =   255
         Index           =   0
         Left            =   2640
         Picture         =   "frmCO_PrincipalX.frx":583E1
         Stretch         =   -1  'True
         ToolTipText     =   "Busqueda Rápida"
         Top             =   -1560
         Width           =   255
      End
      Begin VB.Image imgReporteAntiguedad 
         Height          =   480
         Left            =   -66600
         Picture         =   "frmCO_PrincipalX.frx":586EB
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Líneas Disponibles"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   -74880
         TabIndex        =   72
         Top             =   460
         Width           =   4575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operaciones Generadas a Fiadores por Traspaso de Deudas"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   70
         Top             =   360
         Width           =   8775
      End
      Begin VB.Image imgReporteCuotas 
         Height          =   360
         Left            =   -66600
         Picture         =   "frmCO_PrincipalX.frx":58FB5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo de Cuotas"
         Height          =   255
         Left            =   -71040
         TabIndex        =   66
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   840
         Width           =   8775
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Abono"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   63
         Top             =   480
         Width           =   1095
      End
      Begin VB.Image imgVisualiza 
         Height          =   255
         Left            =   -67800
         Picture         =   "frmCO_PrincipalX.frx":5987F
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Apartado"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   54
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Distrito"
         Height          =   255
         Left            =   6000
         TabIndex        =   52
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Cantón"
         Height          =   255
         Left            =   3120
         TabIndex        =   51
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         Index           =   0
         X1              =   -75000
         X2              =   -66000
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label6 
         Caption         =   "Anuladas"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   45
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Directas"
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Cta Planilla"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Int % Actual"
         Height          =   255
         Index           =   9
         Left            =   -71760
         TabIndex        =   30
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Plazo Recálculo"
         Height          =   255
         Index           =   8
         Left            =   -71760
         TabIndex        =   29
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota"
         Height          =   255
         Index           =   7
         Left            =   -72000
         TabIndex        =   28
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Int % (Original)"
         Height          =   255
         Index           =   6
         Left            =   -72000
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Plazo"
         Height          =   255
         Index           =   5
         Left            =   -72000
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Monto"
         Height          =   255
         Index           =   4
         Left            =   -71760
         TabIndex        =   25
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Amortizado"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Int. Pagado"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Monto"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -75000
         X2              =   -66000
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OP. Ex-Socio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   -72720
         TabIndex        =   87
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   84
         Top             =   3600
         Width           =   2175
      End
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Nombre de la Persona"
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   18
      ToolTipText     =   "Descripción del código"
      Top             =   480
      Width           =   6015
   End
   Begin VB.TextBox txtCedula 
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      ToolTipText     =   "Cédula de la Persona"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      ToolTipText     =   "Código del Préstamo"
      Top             =   480
      Width           =   1815
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   182
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   688
      _CBWidth        =   9060
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "txtOpInfo"
      MinHeight1      =   315
      Width1          =   1305
      NewRow1         =   0   'False
      Child2          =   "txtOperacion"
      MinHeight2      =   315
      Width2          =   1620
      NewRow2         =   0   'False
      Child3          =   "tlbPrincipal"
      MinHeight3      =   330
      Width3          =   1980
      NewRow3         =   0   'False
      Begin VB.TextBox txtOperacion 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   30
         Width           =   1425
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   3150
         TabIndex        =   184
         Top             =   30
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   582
         ButtonWidth     =   2143
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Adjuntos"
               Key             =   "adjuntos"
               Object.ToolTipText     =   "Comprobantes para la Gestión"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Refrescar"
               Key             =   "refrescar"
               Object.ToolTipText     =   "Actualiza Estado Laboral Fiadores"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Reversar"
               Key             =   "reversar"
               Object.ToolTipText     =   "Reversa el Movimiento Realizado a la Operacion"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Reportes"
               Key             =   "reportes"
               ImageIndex      =   12
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   6
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repIngresos"
                     Text            =   "Ingresos Mora"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repEgresos"
                     Text            =   "Egresos Mora"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repAbonos"
                     Text            =   "Abonos del Mes"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repPlanilla"
                     Text            =   "Planilla (Comparativo)"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "x"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Archivo"
                     Text            =   "Archivo de Estudio"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "&Cerrar"
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cerrar Ventana"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtOpInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   183
         Text            =   "# Operación"
         Top             =   30
         Width           =   1110
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Cuota"
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   134
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblPagare 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7680
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblDocumento 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblGarantia 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblUltimoMovimiento 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblPrimerDeduccion 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pagaré (Letra)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   7680
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   6000
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Garantía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ult. Mov"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1era Deduc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblEstadoMoroso 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Cédula"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmCO_PrincipalX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type vTab
 Direccion As Integer   '1 y 0
 Fiadores As Integer    'Sirven Para Indicar si el Tab a sido
 Antiguedad As Integer  'Seleccionado por primera vez o no, para no repetir
 OPGeneradas As Integer 'Busquedas sobre una misma operacion
End Type
Dim vTabs As vTab, vOperacion As Boolean 'vOperacion es para almacenar el ultimo
                                         'Numero de Operacion Consultado
Dim mCurIntc As Currency, mCurIntm As Currency 'Para Alm. Interes Corriente y Moratorios Totales
                                              





Private Sub chkAntiguedadTodas_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Mueve los codigos disponibles a codigos seleccionados en la antiguedad de
'                Saldos, para generacion de estos
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim itmX As ListItem, lng As Long
 With lswAntiguedad
   For lng = 1 To lswAntiguedad.ListItems.Count
        If chkAntiguedadTodas.Value = 0 Then
          .ListItems.Item(lng).Checked = False
        Else
          .ListItems.Item(lng).Checked = True
        End If
   Next lng
 End With
End Sub


Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
 
  strSQL = "select cod_destino + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  catalogo_destinos"
  Call sbLlenaCbo(cboDestino, strSQL)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_destino) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtReporteX & "'"
  Call sbLlenaCbo(cboDestino, strSQL)

End If

End Sub

Private Sub chkTraspaso_Click()
If chkTraspaso.Value = vbChecked Then
 txtPorcentajeDeuda = 100
Else
 txtPorcentajeDeuda = ""
End If
End Sub

Private Sub cmdAceptarFechas_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprimir los Reportes Generales del esta ventana
'REFERENCIAS   : fxFechaServidor - (Devuelve la fecha del Servidor)
'OBSERVACIONES : Utiliza Variables Globales
'-------------------------------------------------------------------------------------------
Dim strRuta As String, strSQL As String, vSQLx As String

Me.MousePointer = vbHourglass


Select Case Mid(cboRepX.Text, 1, 2)
 Case "00" 'Todos
   vSQLx = ""
 Case "01" 'Socios
   vSQLx = " AND {SOCIOS.ESTADOACTUAL} = 'S'"
 Case "02" 'Opex
   vSQLx = " AND ({SOCIOS.ESTADOACTUAL} = 'A' OR {SOCIOS.ESTADOACTUAL} = 'P')"
 Case "03" 'No Socios
   vSQLx = " AND {SOCIOS.ESTADOACTUAL} = 'N'"
 Case "04" 'Ren.Interna
   vSQLx = " AND {SOCIOS.ESTADOACTUAL} = 'A'"
 Case "05" 'Ren.Patronal
   vSQLx = " AND {SOCIOS.ESTADOACTUAL} = 'P'"
End Select



With frmContenedor.Crt
  strRuta = App.Path + "\Credito\Reportes\"
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Cobro Administrativo y Judicial"
    .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"


Select Case lblRepOp.Tag
  Case "REVER" 'Casos con Reversion
    .ReportFileName = strRuta + "CasosConReversion.rpt"
    strSQL = "{REG_CREDITOS.FECHA_ENVIAPROCESO} >= Date(" & Year(dtpFechaInicio.Value) & "," _
            & Month(dtpFechaInicio.Value) & "," & Day(dtpFechaInicio.Value) & ")" _
            & " AND {REG_CREDITOS.FECHA_ENVIAPROCESO} <= Date(" _
            & Year(dtpFechaCorte.Value) & "," & Month(dtpFechaCorte.Value) & "," _
            & Day(dtpFechaCorte.Value) & ")" _
            & " AND {REG_CREDITOS.PROCESO} = 'N'"
    .SelectionFormula = strSQL & vSQLx
    
    .Formulas(2) = "SubTitulo='DE " & dtpFechaInicio.Value & " HASTA " & dtpFechaCorte.Value & " / FILTRO " & Mid(cboRepX.Text, 4, 30) & "'"
    
  Case "ENVCBR" 'Casos Enviados a Cobro Judicial
    .ReportFileName = strRuta + "CasosEnCobroJudicial.rpt"
    strSQL = "{REG_CREDITOS.FECHA_ENVIAPROCESO} >= Date(" & Year(dtpFechaInicio.Value) & "," _
            & Month(dtpFechaInicio.Value) & "," & Day(dtpFechaInicio.Value) & ")" _
            & " AND {REG_CREDITOS.FECHA_ENVIAPROCESO} <= Date(" _
            & Year(dtpFechaCorte.Value) & "," & Month(dtpFechaCorte.Value) & "," _
            & Day(dtpFechaCorte.Value) & ")" _
            & " AND {REG_CREDITOS.PROCESO} = 'J'"
    .SelectionFormula = strSQL & vSQLx
    .Formulas(2) = "SubTitulo=' DE " & dtpFechaInicio & " HASTA " & dtpFechaCorte & " / FILTRO " & Mid(cboRepX.Text, 4, 30) & "'"
  
  Case "TRADEUD" 'Casos Traspaso - Deudor
    .ReportFileName = strRuta + "CasosTraspasoDeudas.rpt"
    strSQL = "{REG_CREDITOS.FECHA_ENVIAPROCESO} >= Date(" & Year(dtpFechaInicio.Value) & "," _
            & Month(dtpFechaInicio.Value) & "," & Day(dtpFechaInicio.Value) & ")" _
            & " AND {REG_CREDITOS.FECHA_ENVIAPROCESO} <= Date(" _
            & Year(dtpFechaCorte.Value) & "," & Month(dtpFechaCorte.Value) & "," _
            & Day(dtpFechaCorte.Value) & ")" _
            & " AND {REG_CREDITOS.PROCESO} = 'T'"
    .SelectionFormula = strSQL & vSQLx
  
  Case "TRAFIA" 'Casos Traspaso - Fiadores
    .ReportFileName = strRuta + "CasosTraspasoFiadores.rpt"
    strSQL = "{REG_CREDITOS.FECHAFORP} >= Date(" & Year(dtpFechaInicio.Value) & "," _
            & Month(dtpFechaInicio.Value) & "," & Day(dtpFechaInicio.Value) & ")" _
            & " AND {REG_CREDITOS.FECHAFORP} <= Date(" _
            & Year(dtpFechaCorte.Value) & "," & Month(dtpFechaCorte.Value) & "," _
            & Day(dtpFechaCorte.Value) & ") AND IsNull ({REG_CREDITOS.REFERENCIA})=FALSE"
    .SelectionFormula = strSQL & vSQLx
    .Formulas(2) = "SubTitulo='TRASPASO DE DEUDAS / FILTRO " & Mid(cboRepX.Text, 4, 30) & "'"

End Select
 
 .PrintReport
End With

Me.MousePointer = vbDefault
fraFechas.Visible = False

End Sub

Private Sub cmdCancelaReversionTraspaso_Click()
 fraReversionDeTraspaso.Visible = False
End Sub

Private Sub cmdCatalogo_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Llenar Lsw de codigos disponibles con los codigos del catalogo de crédito
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, itmX As ListItem

Me.MousePointer = vbHourglass
rs.Open "select * from catalogo", glogon.Conection, adOpenForwardOnly

With lswAntiguedad
  .ListItems.Clear

  Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!Codigo, , 4)
    itmX.SubItems(1) = rs!Descripcion
   rs.MoveNext
  Loop
End With

rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub cmdCBRParametros_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update par_ahcr set CO_TASA_SOCIO = " & chkParMantieneTasaSocios.Value _
       & ",CO_TASA_NSOCIO = " & chkParMantieneTasaNSocios.Value _
       & ",CO_CODIGO = '" & txtParCodigo & "'"
glogon.Conection.Execute strSQL

Call Bitacora("Modifica", "Parametros de Tasas de Cobro (Traslados)")

MsgBox "Parámetros Actualizados...", vbInformation

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdCodigosAtrasados_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : LLenar el Lsw de códigos disponibles solo con aquellos códigos del catalogo
'                 de créditos que esten atrasados (morosos)
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, itmX As ListItem, strSQL As String

Me.MousePointer = vbHourglass

strSQL = "select A.codigo,A.descripcion from catalogo A inner join morosidad B " _
        & "On A.codigo = B.Codigo where B.estado = 'A' group by A.Codigo,A.descripcion"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

With lswAntiguedad
  .ListItems.Clear

  Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!Codigo, , 4)
    itmX.SubItems(1) = rs!Descripcion
   rs.MoveNext
  Loop
End With

rs.Close

Me.MousePointer = vbDefault

End Sub

Function fxValidaPasoCobroJudicial() As Boolean
If UCase(lblProceso.Caption) = "NORMAL" Then
  fxValidaPasoCobroJudicial = True
Else
  fxValidaPasoCobroJudicial = False
End If
End Function

Sub ReversaCobroJudicial()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Reversión de un envio a Cobro Judicial
'REFERENCIAS   : fxFechaServior - (Devuelve fecha del servidor)
'                Bitacora - (Registra el Movimiento efectuado)
'OBSERVACIONES : Ver Readecuacion con Cambio de Operacion (Para Ajustar Nuevos Montos)
'                Genera Asiento
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, strSQL As String, strCuentas As String
Dim vFecha As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

vFecha = fxFechaServidor

'Inicia Transacciones
glogon.Conection.BeginTrans


'Inidica que el proceso del operacion es Normal y no Cobro
strSQL = "update reg_creditos set proceso = 'N',opex = " & fxOpex(txtCedula) _
       & " where id_solicitud = " & txtOperacion
glogon.Conection.Execute strSQL

lblProceso.Caption = "NORMAL"


'Desbloque la Morosidad
strSQL = "update morosidad set estadoi = 'A' where estado = 'A' and estadoi = 'J' and id_solicitud = " & txtOperacion.Text
glogon.Conection.Execute strSQL


'Busca la cuenta contable en la que se encuentra registrada la operacion
If lblOpex.Caption = "NO" Then
 strCuentas = "select ctanamort as ctaPrincipal,ctacamort as ctaPrincipalCobro from catalogo"
 strCuentas = strCuentas & " where codigo = '" & txtCodigo & "'"
Else 'opex
 strCuentas = "select ctaoamort as ctaPrincipal,ctacamort as ctaPrincipalCobro from catalogo"
 strCuentas = strCuentas & " where codigo = '" & txtCodigo & "'"
End If 'opex

With rs
 .Open strCuentas, glogon.Conection, adOpenStatic
 strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber,tmp_fecha,tmp_estado_asiento)" _
        & " values('TRA','" & glogon.Usuario & "','RCBR" & txtOperacion & "','" & !ctaprincipal & "'," & CCur(txtSaldo) _
        & ",'D','" & Format(vFecha, "yyyy/mm/dd") & "','P')"
  
 glogon.Conection.Execute strSQL
  
 strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber,tmp_fecha,tmp_estado_asiento)" _
        & " values('TRA','" & glogon.Usuario & "','RCBR" & txtOperacion & "','" & !ctaprincipalcobro & "'," & CCur(txtSaldo) _
        & ",'H','" & Format(vFecha, "yyyy/mm/dd") & "','P')"
  
 glogon.Conection.Execute strSQL
  
 .Close
End With

Call Bitacora("Reversa", "Cobro Judicial a la Operación:" & txtOperacion)

'Cierra Transacciones
glogon.Conection.CommitTrans

Me.MousePointer = vbDefault

MsgBox "- La operación fue reversada a estado NORMAL " & vbCrLf & vbCrLf _
     & "- Se generó Asiento (RCBR" & txtOperacion & ")", vbInformation

Exit Sub

vError:
Me.MousePointer = vbDefault
glogon.Conection.RollbackTrans
MsgBox Err.Description, vbCritical
End Sub

Sub EjecuteCobroJudicial()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Ejecuta el Cobro Judicial a una Operación
'REFERENCIAS   : FxFechaServidor - (Devuelve la Fecha del servidor)
'                Bitacora - (Registra el movimiento realizado)
'OBSERVACIONES : Genera Asiento
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, strSQL As String, strCuentas As String
Dim strObservacion As String, vFecha As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

strObservacion = ""
vFecha = fxFechaServidor

'Busca la cuenta contable en la que se encuentra registrada la operacion
If lblOpex.Caption = "NO" Then
 strCuentas = "select ctanamort as ctaPrincipal,ctacamort as ctaPrincipalCobro from catalogo"
 strCuentas = strCuentas & " where codigo = '" & txtCodigo & "'"
Else 'opex
 strCuentas = "select ctaoamort as ctaPrincipal,ctacamort as ctaPrincipalCobro from catalogo"
 strCuentas = strCuentas & " where codigo = '" & txtCodigo & "'"
End If 'opex

'Aqui observacion
strObservacion = InputBox("Digite la Observación para este cobro judicial : ", "Observación del Cobro Judicial")
If Len(Trim(strObservacion)) = 0 Then strObservacion = "NADA"


'Inicia Transacciones
glogon.Conection.BeginTrans

'Actualiza reg_creditos campos : Fecha_enviaproceso,observacion_proceso,proceso
strSQL = "update reg_creditos set fecha_enviaproceso = '" & Format(vFecha, "yyyy/mm/dd") _
       & "',observacion_proceso = '" & strObservacion & "',proceso = 'J' where " _
       & "id_solicitud = " & Trim(txtOperacion)
glogon.Conection.Execute strSQL

'NUEVO : Actualiza ESTADOI en Morosidades para que no Acumele mas intereses moratorios
'SE actualiza con J - VERIFICAR EL PROCESO MENSUAL:
strSQL = "update morosidad set estadoi = 'J' where estado = 'A' and id_solicitud = " & txtOperacion.Text
glogon.Conection.Execute strSQL

lblProceso.Caption = "COBRO JUDICIAL"
lblFechaCBR.Caption = Format(vFecha, "dd/mm/yyyy")
lblObservacionesCBR.Caption = strObservacion


'Asiento
With rs
 .Open strCuentas, glogon.Conection, adOpenStatic
 
 strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber,tmp_fecha,tmp_estado_asiento)" _
        & " values('TRA','" & glogon.Usuario & "','CBR" & txtOperacion & "','" & !ctaprincipalcobro & "'," & CCur(txtSaldo) _
        & ",'D','" & Format(vFecha, "yyyy/mm/dd") & "','P')"
 
 glogon.Conection.Execute strSQL
 
 strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber,tmp_fecha,tmp_estado_asiento)" _
        & " values('TRA','" & glogon.Usuario & "','CBR" & txtOperacion & "','" & !ctaprincipal & "'," & CCur(txtSaldo) _
        & ",'H','" & Format(vFecha, "yyyy/mm/dd") & "','P')"
  
 glogon.Conection.Execute strSQL
  
 .Close
End With

Call Bitacora("Aplica", "Cobro Judicial a la Operación:" & txtOperacion)

'Cierra Transacciones
glogon.Conection.CommitTrans

MsgBox "- La operación fue enviada a Cobro Judicial" & vbCrLf & vbCrLf _
     & "- Se generó Asiento (CBR" & txtOperacion & ")", vbInformation

Me.MousePointer = vbDefault

Exit Sub

vError:
Me.MousePointer = vbDefault
glogon.Conection.RollbackTrans
MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdDeduccionPlanilla_Click()
Dim strSQL As String

On Error GoTo vError

If chkDeducirPlanilla.Value = 1 Then
 strSQL = "update reg_creditos set IND_DEDUCE_PLANILLA='S' where id_solicitud = " & txtOperacion
 Call Bitacora("Registra", "Indica la Deducción de Planilla de la OP: " & txtOperacion)
Else
 strSQL = "update reg_creditos set IND_DEDUCE_PLANILLA='N' where id_solicitud = " & txtOperacion
 Call Bitacora("Registra", "Indica la NO Deducción de Planilla de la OP: " & txtOperacion)
End If
glogon.Conection.Execute strSQL

MsgBox "Actualización Realizada...", vbInformation

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdEnviarACobroJudicial_Click()
Dim iRespuesta As Integer

If fxValidaPasoCobroJudicial Then
 iRespuesta = MsgBox("Esta seguro que desea enviar a cobro judicial esta Operación", vbYesNo)
 If iRespuesta = vbYes Then
    Call EjecuteCobroJudicial
 End If
Else 'Validacion
 MsgBox "No se puede ejecutar el cobro judicial verifique la información", vbCritical
End If 'Validacion

End Sub

Private Sub cmdIniciarLista_Click()
 lswAntiguedad.ListItems.Clear
End Sub
Function fxValidaTextosNumericos(txt As TextBox) As Boolean

fxValidaTextosNumericos = True

End Function


Function fxValidaDatosFiadorDeuda() As Boolean
Dim iPasa As Integer
If Len(lblCedulaFiador.Caption) > 0 And Val(txtPorcentajeDeuda.Text) > 0 And Val(txtCuotaFiador.Text) > 0 And Val(txtPorcentajeDeuda.Text) <= 100 Then
    fxValidaDatosFiadorDeuda = True
Else
    fxValidaDatosFiadorDeuda = False
    MsgBox "No se especificó ningún fiador o porcentaje de deuda...", vbCritical
End If

End Function

Function fxOpex(strCedula As String) As Integer
Dim rsX As New ADODB.Recordset
 
rsX.Source = "select estadoactual from socios where cedula = '" & strCedula & "'"
rsX.Open , glogon.Conection, adOpenStatic
 
If rsX!estadoactual = "S" Or rsX!estadoactual = "N" Then
 fxOpex = 0 'Socios y no Socios Cargan la misma Cuenta
Else
 fxOpex = 1 'Ren. Asociacion y Patrono cargan la misma Cuenta
End If
rsX.Close
End Function

Sub AsientoFiadores(curMonto As Currency, strCedula As String, vFecha As Date)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Crea Asientos de Traspaso de Deudas para los fiadores
'REFERENCIAS   : FxFechaServidor - (Devuelve la Fecha del Servidor)
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim rsA As New ADODB.Recordset, strSQL As String

If fxOpex(strCedula) = 0 Then
  strSQL = "select ctanamort as ctaAmortiza "
Else 'cuentas exsocios
  strSQL = "select ctaoamort as ctaAmortiza "
End If
strSQL = strSQL & "from catalogo where codigo = '" & txtCodigoNuevo & "'"
rsA.Open strSQL, glogon.Conection, adOpenStatic
If curMonto > 0 Then
    strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
        & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
        & "','" & txtOperacion & "-" & Format(Day(vFecha), "00") & "','" & rsA!ctaamortiza & "'," & curMonto & ",'D','" _
        & Format(vFecha, "yyyy/mm/dd") & "','P')"
    glogon.Conection.Execute strSQL
End If
rsA.Close

End Sub

Sub TraspasoPorFiador(strCedula As String, intPorcentaje As Integer)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Realizar el Traspaso de la Deuda por porcentaje a Cada Fiador
'REFERENCIAS   : FxFechaServidor - (Devuelve la Fecha del Servidor)
'                Bitacora - (Registra el movimiento realizado)
'                ConsultaOperacion - (Refresca Informacion de la Ventana)
'                AsientoFiadores - (Crea Asientos de Traspaso de Deudas por Fiador)
'OBSERVACIONES : Utiliza Variables Globales y de Módulo
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, intNumero As Integer, lngPriDeduc As Long
Dim curDeuda As Currency, strSQL As String, rs2 As New ADODB.Recordset
Dim strObservacion As String, curRepartido(3) As Currency, curUltimo(3) As Currency
Dim vFecha As Date

strObservacion = InputBox("Digite la Observación para este traspaso de deudas : ", "Observación del Cobro Judicial")
If Len(Trim(strObservacion)) = 0 Then strObservacion = "NADA"

Me.MousePointer = vbHourglass

On Error GoTo vError

vFecha = fxFechaServidor

curDeuda = mCurIntc + mCurIntm + CCur(txtSaldo)

If Mid(Trim(str(GLOBALES.glngFechaCR)), 5, 6) = 12 Then
 lngPriDeduc = (Val(Mid(Trim(str(GLOBALES.glngFechaCR)), 1, 4)) + 1) & "01"
Else
  lngPriDeduc = GLOBALES.glngFechaCR + 1
End If
intNumero = 0


'Inicia Transacciones
glogon.Conection.BeginTrans

 curDeuda = (curDeuda * intPorcentaje) / 100
   strSQL = "insert into reg_creditos(codigo,id_comite,cedula,estadosol" _
          & ",plazo,int,interesv,montoapr,prideduc,fechaforp,saldo,amortiza,interesc" _
          & ",cuota,referencia,userrec,userfor,garantia,firma_deudor" _
          & ",monto_girado,cuotas_planilla,cuotas_directas,cuotas_anuladas,Tesoreria,opex,FECULT) values" _
          & "('" & txtCodigoNuevo & "',1,'" & Trim(strCedula) & "','F'," & txtPlazoFiador _
          & "," & txtInteresFiador & "," & txtInteresFiador & "," & Format(curDeuda, "###########0.00") & "," & lngPriDeduc _
          & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & Format(curDeuda, "###########0.00") _
          & ",0,0," & Format(fxCalcula_Cuota(CLng(curDeuda), txtPlazoFiador, txtInteresFiador), "##########0.00") & "," & txtOperacion _
          & ",'" & glogon.Usuario & "','" & glogon.Usuario & "','F',1" _
          & ",0,0,0,0,'" & Format(vFecha, "yyyy/mm/dd") & "'," & fxOpex(strCedula) & "," & GLOBALES.glngFechaCR & ")"
    glogon.Conection.Execute strSQL
    Call AsientoFiadores(curDeuda, strCedula, vFecha)

'Hacer programa para abonar en morosidad y saber cuanto paga en intc,intm,amortiza

curRepartido(1) = 0
curRepartido(2) = 0
curRepartido(3) = 0

rs.Open "select * from morosidad where estado = 'A' and id_solicitud = " _
    & txtOperacion, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 If IIf(IsNull(rs!intc), 0, rs!intc) + IIf(IsNull(rs!intm), 0, rs!intm) + IIf(IsNull(rs!Amortiza), 0, rs!Amortiza) <= curDeuda Then
    curRepartido(1) = curRepartido(1) + IIf(IsNull(rs!intc), 0, rs!intc)
    curRepartido(2) = curRepartido(2) + IIf(IsNull(rs!intm), 0, rs!intm)
    curRepartido(3) = curRepartido(3) + IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
    curDeuda = curDeuda - IIf(IsNull(rs!intc), 0, rs!intc)
    curDeuda = curDeuda - IIf(IsNull(rs!intm), 0, rs!intm)
    curDeuda = curDeuda - IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
    strSQL = "Update morosidad set abintc = intc, abintm = intm, abamortiza = amortiza ," _
           & "estado = 'C',fecult = '" & Format(vFecha, "yyyy/mm/dd") & "',tcon = 4,ncon = " _
           & txtOperacion & " where id_moro = " & rs!id_moro
    glogon.Conection.Execute strSQL
 Else
    curUltimo(1) = 0
    curUltimo(2) = 0
    curUltimo(3) = 0
    
    If curDeuda >= IIf(IsNull(rs!intc), 0, rs!intc) Then
      curRepartido(1) = curRepartido(1) + IIf(IsNull(rs!intc), 0, rs!intc)
      curUltimo(1) = IIf(IsNull(rs!intc), 0, rs!intc)
      curDeuda = curDeuda - IIf(IsNull(rs!intc), 0, rs!intc)
    Else
     If curDeuda > 0 Then
      curRepartido(1) = curRepartido(1) + curDeuda
      curUltimo(1) = curDeuda
      curDeuda = 0
     End If
    End If
    
    If curDeuda >= IIf(IsNull(rs!intm), 0, rs!intm) Then
      curRepartido(2) = curRepartido(2) + IIf(IsNull(rs!intm), 0, rs!intm)
      curUltimo(2) = IIf(IsNull(rs!intm), 0, rs!intm)
      curDeuda = curDeuda - IIf(IsNull(rs!intm), 0, rs!intm)
    Else
     If curDeuda > 0 Then
      curRepartido(2) = curRepartido(2) + curDeuda
      curUltimo(2) = curDeuda
      curDeuda = 0
     End If
    End If
    
    If curDeuda >= IIf(IsNull(rs!Amortiza), 0, rs!Amortiza) Then
      curRepartido(3) = curRepartido(3) + IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
      curUltimo(3) = IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
      curDeuda = curDeuda - IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
    Else
     If curDeuda > 0 Then
      curRepartido(3) = curRepartido(3) + curDeuda
      curUltimo(3) = curDeuda
      curDeuda = 0
     End If
    End If
    
    strSQL = "Update morosidad set abintc = " & curUltimo(1) _
           & ", abintm = " & curUltimo(2) & ", abamortiza = " & curUltimo(3) & " ,estado = 'C', tcon = 4,ncon = " _
           & txtOperacion & ",fecult = '" & Format(vFecha, "yyyy/mm/dd") & "'" _
           & " where id_moro = " & rs!id_moro
    glogon.Conection.Execute strSQL
    
    'Inserta el Registro con la diferencia
    strSQL = "insert morosidad(id_solicitud,codigo,intc,intm,amortiza,fechap,fecult,estado,estadoi,fecap) values" _
           & "(" & rs!ID_SOLICITUD & ",'" & rs!Codigo & "'," & IIf(IsNull(rs!intc), 0, rs!intc) - curUltimo(1) _
           & "," & IIf(IsNull(rs!intm), 0, rs!intm) - curUltimo(2) & "," & IIf(IsNull(rs!Amortiza), 0, rs!Amortiza) - curUltimo(3) _
           & "," & rs!fechap & ",'" & Format(vFecha, "yyyy/mm/dd") & "','A','A'," & GLOBALES.glngFechaCR & ")"
    glogon.Conection.Execute strSQL
 End If
 
 rs.MoveNext
Loop 'Aplicacion a Cuotas Morosas
 
rs.Close

'ACTUALIZA REG_CREDITOS
If Abs((curDeuda + curRepartido(3)) - CCur(txtSaldo)) < 1 Then
'   strSQL = "update reg_creditos set saldo = saldo - " & curDeuda + curRepartido(3) & ", amortiza = amortiza + " & curDeuda + curRepartido(3) _
'       & ", interesc = interesc + " & curRepartido(1) + curRepartido(2) & ", estado = 'C', Proceso = 'T'" _
'       & ",fecha_enviaproceso = '" & Format(fxFechaServidor, "yyyy/mm/dd") & "'" _
'       & ",observacion_proceso = '" & strObservacion & "'" _
'       & " where id_solicitud = " & txtOperacion
   
'Cambiado por esta para que aparezca en el estado de cuenta
   strSQL = "update reg_creditos set saldo = saldo - " & curDeuda + curRepartido(3) & ", amortiza = amortiza + " & curDeuda + curRepartido(3) _
       & ", interesc = interesc + " & curRepartido(1) + curRepartido(2) & ", estado = 'A', Proceso = 'T'" _
       & ",fecha_enviaproceso = '" & Format(vFecha, "yyyy/mm/dd") & "'" _
       & ",observacion_proceso = '" & strObservacion & "'" _
       & " where id_solicitud = " & txtOperacion

Else
   strSQL = "update reg_creditos set saldo = saldo - " & curDeuda + curRepartido(3) & ", amortiza = amortiza + " & curDeuda + curRepartido(3) _
       & ", interesc = interesc + " & curRepartido(1) + curRepartido(2) & ", estado = 'A', Proceso = 'T'" _
       & ",fecha_enviaproceso = '" & Format(vFecha, "yyyy/mm/dd") & "'" _
       & ",observacion_proceso = '" & strObservacion & "'" _
       & " where id_solicitud = " & txtOperacion
End If
glogon.Conection.Execute strSQL


'INSERT EN CREDITOS DT POR LA DIFERENCIA EN EL SALDO

If curDeuda > 0 Then
    strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
           & "FECHAP,TCON,NCON,ESTADO) values('" & txtCodigo & "'," _
           & txtOperacion & ",0," & curDeuda _
           & ",0," & curDeuda & ",'" & Format(vFecha, "yyyy/mm/dd") _
           & "'," & GLOBALES.glngFechaCR & ",4,8888,'A')"
    glogon.Conection.Execute strSQL
End If


''Hacer Asiento Aqui *************************************************************

 If lblOpex.Caption = "NO" Then
  strSQL = "select ctanintc as ctaIntc, ctanintm as ctaIntm, ctanamort as ctaAmortiza"
 Else 'cuentas opex
  strSQL = "select ctaointc as ctaIntc, ctaointm as ctaIntm, ctaoamort as ctaAmortiza"
 End If
 strSQL = strSQL & " from catalogo where codigo = '" & txtCodigo & "'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 If (curDeuda + curRepartido(3)) > 0 Then
    strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
        & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
        & "','" & txtOperacion & "-" & Format(Day(vFecha), "00") & "','" & rs!ctaamortiza & "'," & Format(curDeuda + curRepartido(3), "########0.00") & ",'H','" _
        & Format(vFecha, "yyyy/mm/dd") & "','P')"
    glogon.Conection.Execute strSQL
  End If
  
 If curRepartido(1) > 0 Then
    strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
        & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
        & "','" & txtOperacion & "-" & Format(Day(vFecha), "00") & "','" & rs!ctaintc & "'," & curRepartido(1) & ",'H','" _
        & Format(vFecha, "yyyy/mm/dd") & "','P')"
    glogon.Conection.Execute strSQL
 End If
  
 If curRepartido(2) > 0 Then
    strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
        & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
        & "','" & txtOperacion & "-" & Format(Day(vFecha), "00") & "','" & rs!ctaintm & "'," & curRepartido(2) & ",'H','" _
        & Format(vFecha, "yyyy/mm/dd") & "','P')"
    glogon.Conection.Execute strSQL
 End If
 rs.Close

'BITACORA
Call Bitacora("Aplica", "Traspaso Deudas Por Porc. de la Operación:" & txtOperacion)

'Cierra Transacciones
glogon.Conection.CommitTrans

MsgBox "Traspaso de Deudas a Fiadores Realizado Satisfactoriamente...", vbInformation

Call ConsultaOperacion

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 glogon.Conection.RollbackTrans
 MsgBox Err.Description, vbCritical
End Sub

Sub TraspasoTotal()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Realiza el Traspaso de Deudas por Igual a Todos los Fiadores
'REFERENCIAS   : fxFechaServidor - (Devuelve la Fecha del Servidor)
'                ConsultaOperacion - (Refresca la informacion de la Ventana)
'                AsientoFiadores - (Crea el Asiento de Traspaso por cada Fiador)
'OBSERVACIONES : Divide el Monto de la Deuda (Saldo + Intereses) y los reparte a los
'                Fiadores. Utiliza Variables Globales y de Módulo
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, intNumero As Integer, lngPriDeduc As Long
Dim curDeuda As Currency, strSQL As String, rs2 As New ADODB.Recordset
Dim strObservacion As String, vFecha As Date


strObservacion = InputBox("Digite la Observación para este traspaso de deudas : ", "Observación del Cobro Judicial")
If Len(Trim(strObservacion)) = 0 Then strObservacion = "NADA"

Me.MousePointer = vbHourglass

On Error GoTo vError

vFecha = fxFechaServidor

'Se incluye la poliza PSD
curDeuda = mCurIntc + mCurIntm + CCur(txtSaldo)

If Mid(Trim(str(GLOBALES.glngFechaCR)), 5, 6) = 12 Then
 lngPriDeduc = (Val(Mid(Trim(str(GLOBALES.glngFechaCR)), 1, 4)) + 1) & "01"
Else
  lngPriDeduc = GLOBALES.glngFechaCR + 1
End If
intNumero = 0

'Inicia Transacciones
glogon.Conection.BeginTrans

With rs
 .Open "select * from fiadores where estado = 'A' and id_solicitud =" & txtOperacion, glogon.Conection, adOpenStatic
 intNumero = .RecordCount
 curDeuda = curDeuda / intNumero
 Do While Not .EOF
   strSQL = "insert into reg_creditos(codigo,id_comite,cedula,estadosol" _
          & ",plazo,int,interesv,montoapr,prideduc,fechaforp,saldo,amortiza,interesc" _
          & ",cuota,referencia,userrec,userfor,garantia,firma_deudor" _
          & ",monto_girado,cuotas_planilla,cuotas_directas,cuotas_anuladas,Tesoreria,opex,FECULT) values" _
          & "('" & txtCodigoNuevo & "',1,'" & Trim(!cedulaf) & "','F'," & txtPlazoFiador & "," _
          & txtInteresFiador & "," & txtInteresFiador & "," & Format(curDeuda, "###########0.00") & "," & lngPriDeduc _
          & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & Format(curDeuda, "###########0.00") _
          & ",0,0," & Format(fxCalcula_Cuota(CLng(curDeuda), txtPlazoFiador, txtInteresFiador), "##########0.00") & "," & txtOperacion _
          & ",'" & glogon.Usuario & "','" & glogon.Usuario & "','F',1" _
          & ",0,0,0,0,'" & Format(vFecha, "yyyy/mm/dd") & "'," & fxOpex(!cedulaf) & "," & GLOBALES.glngFechaCR & ")"
    glogon.Conection.Execute strSQL
    Call AsientoFiadores(curDeuda, !cedulaf, vFecha)
  .MoveNext
 Loop
 .Close
End With


'Cancela Mora Activa del Deudor
strSQL = "Update morosidad set abintc = intc, abintm = intm, abamortiza = amortiza ,estado = 'C', tcon = 4,ncon = " _
       & txtOperacion & ",fecult = '" & Format(vFecha, "yyyy/mm/dd") & "' where estado = 'A'" _
       & " and id_solicitud = " & txtOperacion
glogon.Conection.Execute strSQL


'Cambiado por esta para que aparezca en los reportes de estados de cuenta
strSQL = "update reg_creditos set saldo = saldo - " & CCur(txtSaldo) & ", amortiza = amortiza + " & CCur(txtSaldo) _
       & ", interesc = interesc + " & CCur(txtInteresesMoratorios) & ", estado = 'A', Proceso = 'T'" _
       & ",fecha_enviaproceso = '" & Format(vFecha, "yyyy/mm/dd") & "'" _
       & ",observacion_proceso = '" & strObservacion & "'" _
       & " where id_solicitud = " & txtOperacion
glogon.Conection.Execute strSQL

'INSERT EN CREDITOS DT POR LA DIFERENCIA EN EL SALDO

strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
       & "FECHAP,TCON,NCON,ESTADO) values('" & txtCodigo & "'," _
       & txtOperacion & ",0," & CCur(txtSaldo) - CCur(txtAmortizacionAtrasada) _
       & ",0," & CCur(txtSaldo) - CCur(txtAmortizacionAtrasada) & ",'" & Format(vFecha, "yyyy/mm/dd") _
       & "'," & GLOBALES.glngFechaCR & ",4,8888,'A')"
glogon.Conection.Execute strSQL


'Hacer Asiento Aqui *************************************************************

 If lblOpex.Caption = "NO" Then
  strSQL = "select ctanintc as ctaIntc, ctanintm as ctaIntm, ctanamort as ctaAmortiza"
 Else 'cuentas opex
  strSQL = "select ctaointc as ctaIntc, ctaointm as ctaIntm, ctaoamort as ctaAmortiza"
 End If
 strSQL = strSQL & " from catalogo where codigo = '" & txtCodigo & "'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 
 If CCur(txtSaldo) > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-" & Format(Day(vFecha), "00") & "','" & rs!ctaamortiza & "'," & Format(txtSaldo, "########0.00") & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  glogon.Conection.Execute strSQL
 End If
 
 If mCurIntc > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-" & Format(Day(vFecha), "00") & "','" & rs!ctaintc & "'," & mCurIntc & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  glogon.Conection.Execute strSQL
 End If
 
 If mCurIntm > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-" & Format(Day(vFecha), "00") & "','" & rs!ctaintm & "'," & mCurIntm & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  glogon.Conection.Execute strSQL
 End If
 rs.Close
 
'BITACORA
Call Bitacora("Aplica", "Traspaso de Deudas de la Operación:" & txtOperacion)

'Cierra Transacciones
glogon.Conection.CommitTrans
 
MsgBox "Traspaso de Deudas a Fiadores Realizado Satisfactoriamente..." _
       & vbCrLf & vbCrLf & " - Se Generó Asiento (" & txtOperacion & "-" _
       & Format(Day(Date), "00") & ")", vbInformation

Call ConsultaOperacion
 
Me.MousePointer = vbDefault
 
Exit Sub

vError:
 Me.MousePointer = vbDefault
 glogon.Conection.RollbackTrans
 MsgBox Err.Description, vbCritical
End Sub

Function fxValidaDatosTraspasoDeudas() As Boolean
Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
'Verificar el Saldo de la Operacion,Plazo,Interes,etc


fxValidaDatosTraspasoDeudas = True

If Not IsNumeric(txtSaldo) Then
   fxValidaDatosTraspasoDeudas = False
   MsgBox "Verifique el Saldo de la Operación", vbInformation
   Exit Function
End If
If Not IsNumeric(txtPlazoFiador) Then
   fxValidaDatosTraspasoDeudas = False
   MsgBox "Verifique el plazo suministrado para la nueva Operación", vbInformation
       Exit Function
End If
If Not IsNumeric(txtInteresFiador) Then
   fxValidaDatosTraspasoDeudas = False
   MsgBox "Verifique la tasa de interes para nueva Operación", vbInformation
   Exit Function
End If
If Not IsNumeric(txtPorcentajeDeuda) Then
   fxValidaDatosTraspasoDeudas = False
   MsgBox "Verifique la tasa de interes para nueva Operación", vbInformation
   Exit Function
End If

If (Val(txtSaldo) = 0) Or (Len(Trim(txtSaldo)) = 0) Then
 fxValidaDatosTraspasoDeudas = False
 MsgBox "No se puede Traspasar Nada porque el Saldo está en CERO", vbInformation
 Exit Function
End If

If (Val(txtPlazoFiador) <= 0) Or (Len(Trim(txtPlazoFiador)) = 0) Or (Val(txtPlazoFiador) > 100) Then
 fxValidaDatosTraspasoDeudas = False
 MsgBox "El plazo especificado para la(s) operacion(es) no es válido", vbInformation
 Exit Function
End If

If (Val(txtPorcentajeDeuda) <= 0) Or (Len(Trim(txtPorcentajeDeuda)) = 0) Or (Val(txtPorcentajeDeuda) > 100) Then
 fxValidaDatosTraspasoDeudas = False
 MsgBox "El porcentaje de traspaso especificado para la(s) operacion(es) no es válido", vbInformation
 Exit Function
End If

If (Val(txtInteresFiador) <= 0) Or (Len(Trim(txtInteresFiador)) = 0) Or (Val(txtInteresFiador) > 100) Then
 fxValidaDatosTraspasoDeudas = False
 MsgBox "El Interes especificado para la(s) operacion(es) no es válido", vbInformation
 Exit Function
End If

If (Len(txtCodigoNuevo) <= 0) Then
 fxValidaDatosTraspasoDeudas = False
 MsgBox "No se ha especificado ningun código para las operaciones de los fiadores", vbInformation
 Exit Function
Else
  txtCodigoNuevo = UCase(txtCodigoNuevo)
End If

rs.CursorLocation = adUseServer
rs.Open "select coalesce(count(*),0) as Existe from catalogo where codigo ='" _
    & txtCodigoNuevo & "'", glogon.Conection, adOpenStatic
If rs!existe = 0 Then
 fxValidaDatosTraspasoDeudas = False
 rs.Close
 MsgBox "El código especificado no existe en el catalogo de créditos...", vbInformation
 Exit Function
End If
rs.Close

'Eliminado ya que si se pueden tener varias operaciones con el mismo codigo
''''rs.CursorLocation = adUseServer
''''rs.Open "select cedulaf from fiadores where estado ='A' and id_solicitud = " & txtOperacion, glogon.Conection, adOpenStatic
''''Do While rs.EOF = False
''''  rs2.Open "select coalesce(count(*),0) as Existe from reg_Creditos where cedula = '" _
''''     & rs!cedulaf & "' and codigo = '" & txtCodigoNuevo & "' and estado = 'A'", glogon.Conection, adOpenStatic
''''  If rs2!existe > 0 Then
''''     rs2.Close
''''     rs.Close
''''     fxValidaDatosTraspasoDeudas = False
''''     MsgBox "Existe un fiador que ya posee un crédito con el código especificado...", vbInformation
''''     Exit Function
''''  End If
''''
''''  rs2.Close
''''  rs.MoveNext
''''Loop
''''rs.Close
End Function

Private Sub AsientoTraspasoFiadorDeudorF(curMonto As Currency, curIntc As Currency _
                , curIntm As Currency, strCedula As String, strCodigo As String, vFecha As Date)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Crea el Asiento de REVERSION de un Traspaso de Deudas para los Fiadores
'REFERENCIAS   : fxFechaServidor - (Devuelve Fecha del Servidor)
'OBSERVACIONES : Ver Reversiones de Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, strSQL As String

 If fxOpex(strCedula) = 0 Then
  strSQL = "select ctanintc as ctaIntc, ctanintm as ctaIntm, ctanamort as ctaAmortiza "
 Else 'cuentas opex
  strSQL = "select ctaointc as ctaIntc, ctaointm as ctaIntm, ctaoamort as ctaAmortiza "
 End If
 strSQL = strSQL & "from catalogo where codigo = '" & strCodigo & "'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 
 If curMonto > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-FD" & Format(Day(vFecha), "00") & "','" & rs!ctaamortiza & "'," & curMonto & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  glogon.Conection.Execute strSQL
 End If
 
 If curIntc > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-FD" & Format(Day(vFecha), "00") & "','" & rs!ctaintc & "'," & curIntc & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  glogon.Conection.Execute strSQL
 End If
 
 If curIntm > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-FD" & Format(Day(vFecha), "00") & "','" & rs!ctaintm & "'," & curIntm & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  glogon.Conection.Execute strSQL
 End If
 rs.Close
End Sub


Private Sub AsientoTraspasoFiadorDeudor(curMonto As Currency, strCedula As String, vFecha As Date)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Crea asiento de REVERSION de Traspaso de Deudas para el Deudor
'REFERENCIAS   : fxFechaServidor - (Devuelve la fecha del Sistema)
'OBSERVACIONES : Ver Reversion de Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim rsA As New ADODB.Recordset, strSQL As String

If fxOpex(strCedula) = 0 Then
  strSQL = "select ctanamort as ctaAmortiza "
Else 'cuentas exsocios
  strSQL = "select ctaoamort as ctaAmortiza "
End If
strSQL = strSQL & "from catalogo where codigo = '" & txtCodigo & "'"
rsA.Open strSQL, glogon.Conection, adOpenStatic

If curMonto > 0 Then
    strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
        & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
        & "','" & txtOperacion & "-FD" & Format(Day(vFecha), "00") & "','" & rsA!ctaamortiza & "'," & curMonto & ",'D','" _
        & Format(vFecha, "yyyy/mm/dd") & "','P')"
    glogon.Conection.Execute strSQL
End If
rsA.Close
End Sub

Private Sub cmdReversaTraspasoDeudas_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Reversar el Traspaso de Deudas
'REFERENCIAS   : AsientoTraspasoFiadorDeudorF -(Crea Lineas de Asiento de Reversion - Fiadores)
'                AsientoTraspasoFiadorDeudor -(Crea Lineas de Asiento de Reversion - Deudor)
'                fxFechaServidor -(Devuelve la Fecha del Servidor)
'                Bitacora - (Registra el movimiento realizado)
'OBSERVACIONES : Se Ejecutan los casos seleccionados, Utiliza variables globales
'-------------------------------------------------------------------------------------------

Dim itmX As ListItem, lng As Long, strSQL As String
Dim rs As New ADODB.Recordset, lngPriDeduc As Long
Dim xIntC As Currency, xIntM As Currency, lngUltimaOperacion As Long
Dim xAmortiza As Currency, vFecha As Date, vpaso As Boolean

xIntC = 0
xIntM = 0
xAmortiza = 0
vpaso = False

On Error GoTo vError

If CCur(txtTRAFD_MONTO) = 0 Then Exit Sub

'Verifica que exista una opcion marcada
With lswOperacionesGeneradas.ListItems
 For lng = 1 To .Count
  If .Item(lng).Checked Then
     vpaso = True
  End If
 Next lng
End With


If Not vpaso Then
  MsgBox "No se ha marcado ningun (deuda) de fiador, para reversión verifique...?", vbExclamation
  Exit Sub
End If

vFecha = fxFechaServidor

'Cancelar Operaciones de los fiadores Marcados
If Mid(Trim(str(GLOBALES.glngFechaCR)), 5, 6) = 12 Then
 lngPriDeduc = (Val(Mid(Trim(str(GLOBALES.glngFechaCR)), 1, 4)) + 1) & "01"
Else
  lngPriDeduc = GLOBALES.glngFechaCR + 1
End If

'Inicia Transacciones
glogon.Conection.BeginTrans

 With lswOperacionesGeneradas
   For lng = 1 To .ListItems.Count
    Set itmX = .FindItem(lng, lvwTag)
     If itmX Is Nothing Then  'No lo encontro.
      'nada
     Else
      .ListItems.Item(lng).Selected = lng
      If .SelectedItem.Checked Then
        rs.Open "select sum(intc) as Intc, sum(intm) as Intm,sum(amortiza) as Amortiza from morosidad " _
            & "where estado = 'A' and id_solicitud = " & .SelectedItem.Text, glogon.Conection, adOpenStatic
         xIntC = IIf(IsNull(rs!intc), 0, rs!intc)
         xIntM = IIf(IsNull(rs!intm), 0, rs!intm)
         xAmortiza = IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
        rs.Close
        strSQL = "Update morosidad set abintc = intc, abintm = intm, abamortiza = amortiza ,estado = 'C', tcon = 4,ncon = " _
               & .SelectedItem.Text & ",fecult = '" & Format(vFecha, "yyyy/mm/dd") & "' where estado = 'A'" _
               & " and id_solicitud = " & .SelectedItem.Text
        glogon.Conection.Execute strSQL
        
        'Deberia de realizar registro en creditos_dt
        
        strSQL = "update reg_creditos set saldo = saldo - " & CCur(.SelectedItem.SubItems(4)) & ", amortiza = amortiza + " & CCur(.SelectedItem.SubItems(4)) _
               & ", interesc = interesc + " & CCur(.SelectedItem.SubItems(8)) & ", estado = 'C', Proceso = 'N'" _
               & ",fecha_enviaproceso = '" & Format(vFecha, "yyyy/mm/dd") & "'" _
               & ",observacion_proceso = 'REVERSION DE TRASPASO DE DEUDAS'" _
               & " where id_solicitud = " & .SelectedItem.Text
        glogon.Conection.Execute strSQL
        
        'Inserta en creditos DT Cancelación (CCur(.SelectedItem.SubItems(4)) - xamortiza)
        xAmortiza = CCur(.SelectedItem.SubItems(4)) - xAmortiza
        
        strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
               & "FECHAP,TCON,NCON,ESTADO) values('" & .SelectedItem.SubItems(1) & "'," _
               & .SelectedItem.Text & ",0," & xAmortiza _
               & ",0," & xAmortiza & ",'" & Format(vFecha, "yyyy/mm/dd") _
               & "'," & GLOBALES.glngFechaCR & ",4,8888,'A')"
        glogon.Conection.Execute strSQL
        
        'ASIENTO
        Call AsientoTraspasoFiadorDeudorF(CCur(.SelectedItem.SubItems(4)), xIntC, xIntM, Trim(.SelectedItem.SubItems(2)), .SelectedItem.SubItems(1), vFecha)
      
      End If
     End If
   Next lng
 End With

If Val(lblOperacionActualDeudor) = 0 Then
  'Insertar Nueva Operacion
   strSQL = "insert into reg_creditos(codigo,id_comite,cedula,estadosol" _
          & ",plazo,int,interesv,montoapr,prideduc,fechaforp,saldo,amortiza,interesc" _
          & ",cuota,referencia,userrec,userfor,garantia,firma_deudor" _
          & ",monto_girado,cuotas_planilla,cuotas_directas,cuotas_anuladas,Tesoreria,opex,OBSERVACION_PROCESO,FECULT) values" _
          & "('" & txtCodigo & "',1,'" & Trim(txtCedula) & "','F'," & txtTRAFD_Plazo & "," _
          & txtTRAFD_Int & "," & txtTRAFD_Int & "," & CCur(txtTRAFD_MONTO) & "," & lngPriDeduc _
          & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & CCur(txtTRAFD_MONTO) & ",0,0," _
          & Format(fxCalcula_Cuota(CLng(CCur(txtTRAFD_MONTO)), txtTRAFD_Plazo, txtTRAFD_Int), "##########0.00") & "," & txtOperacion _
          & ",'" & glogon.Usuario & "','" & glogon.Usuario & "','F',1" _
          & ",0,0,0,0,'" & Format(vFecha, "yyyy/mm/dd") & "'," & fxOpex(txtCedula) & ",'REVERSION DE TRASPASO FIADOR - DEUDOR'," _
          & GLOBALES.glngFechaCR & ")"
    glogon.Conection.Execute strSQL
  
  'Recupera la nuevo operacion
   lngUltimaOperacion = fxUltimaOperacion(txtCedula)
  
  'Hereda Fiadores Operacion Anterior
  With rs
   .CursorLocation = adUseServer
   .Source = "select * from fiadores where id_solicitud = " & txtOperacion.Text
   .Open , glogon.Conection, adOpenStatic
   Do While Not .EOF
    strSQL = "insert fiadores(id_solicitud,codigo,cedulaf,nombre,firma,estado) values(" _
           & lngUltimaOperacion & ",'" & !Codigo & "','" & !cedulaf _
           & "','" & !Nombre & "','" & !firma & "','" & !Estado & "')"
    glogon.Conection.Execute strSQL
    .MoveNext
   Loop
   .Close
  End With
  
  'Aqui el Asiento
   Call AsientoTraspasoFiadorDeudor(CCur(txtTRAFD_MONTO), txtCedula, vFecha)

Else

  'Actualizar Operacion
  strSQL = "update reg_creditos set " _
        & "montoapr = " & CCur(txtTRAFD_MONTO) & "," _
        & "saldo = saldo + " & CCur(txtTRAFD_MONTO) - CCur(lblSaldoActualDeudor.Caption) & "," _
        & "plazo = " & CCur(txtTRAFD_Plazo) & "," _
        & "interesv = " & CCur(txtTRAFD_Int) & "," _
        & "cuota = " & CCur(txtTRAFD_Cuota) & " where id_solicitud = " & lblOperacionActualDeudor.Caption
  glogon.Conection.Execute strSQL
  'Aqui el Asiento
  Call AsientoTraspasoFiadorDeudor(CCur(txtTRAFD_MONTO) - CCur(lblSaldoActualDeudor.Caption), txtCedula, vFecha)

End If

'BITACORA
Call Bitacora("Reversa", "Traspaso de Deudas de la Operación:" & txtOperacion)

'Cierra Transacciones
glogon.Conection.CommitTrans

MsgBox "- Reversión de Traspaso Realizada Satisfactoriamente..." _
       & vbCrLf & vbCrLf & "- Se Generó Asiento ([OPERACION]" _
       & "-FD" & Format(Day(vFecha), "00") & ")", vbInformation

fraReversionDeTraspaso.Visible = False
Call OperacionesGeneradas

Exit Sub

vError:
 glogon.Conection.RollbackTrans
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub cmdTraspasoDeudas_Click()
Dim iRespuesta As Integer
'1. Verificar si se realiza el pase a todos o solo a uno
'2. Aplicar porcentaje a cada uno (Saldo + Intereses en Mora)
'NO SE PUEDEN HACER MOVIMIENTOS DE TRASPASOS SI EL CREDITO SE ENCUENTRA EN CBR
If lblProceso.Caption = "COBRO JUDICIAL" Then
 MsgBox "No se puede realizar traspaso de deudas porque es ya se encuentra en Cobro Judicial", vbInformation
Else
If fxValidaDatosTraspasoDeudas Then
    If chkTraspaso.Value = vbChecked Then
     iRespuesta = MsgBox("Esta Seguro de Traspasar la Deuda a Todos los Fiadores", vbYesNo)
     If iRespuesta = vbYes Then
      'Pasa a todos los fiadores, Aplica Cargo PSD al Traslado
      Call TraspasoTotal
     End If
    
    Else 'Por persona y porcentaje
    
     iRespuesta = MsgBox("Esta Seguro de Traspasar el " & Me.txtPorcentajeDeuda & "% de la deuda al fiador seleccionado", vbYesNo)
     If iRespuesta = vbYes And fxValidaDatosFiadorDeuda Then
        'NO Aplica Cargo PSD al traslado.
        Call TraspasoPorFiador(lblCedulaFiador.Caption, txtPorcentajeDeuda)
     End If
    
    End If 'CHK

End If 'Validacion
End If
End Sub

Private Sub Form_Activate()
 vModulo = 4
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
 vOperacion = False 'Inicializar
 dtpFechaInicio.Value = fxFechaServidor
 dtpFechaCorte.Value = dtpFechaInicio
 cboVisualiza.Text = "Abonos Ordinarios"
 cboAntiguedad.Text = "Con Base a la Fecha de la Cuota"
 cboTipoCuotas.Text = "Todas"
 Call LimpiaDatos
End Sub


Sub LlenaAbonos()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Llena Lsw de Cuotas con el Listado de Abonos Ordinarios y Extraordinarios
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, itmX As ListItem, strSQL As String
Dim curTotales(4) As Currency

curTotales(1) = 0
curTotales(2) = 0
curTotales(3) = 0
curTotales(4) = 0


With lswAbonos
 Select Case cboTipoCuotas.Text
  Case "Todas"
    strSQL = "select * from creditos_dt where id_solicitud = " & txtOperacion
  Case "Activas"
    strSQL = "select * from creditos_dt where id_solicitud = " & txtOperacion & " and estado = 'A'"
  Case "Anuladas"
    strSQL = "select * from creditos_dt where id_solicitud = " & txtOperacion & " and estado = 'N'"
  Case Else
    strSQL = "select * from creditos_dt where id_solicitud = " & txtOperacion & " and estado = 'A'"
 End Select
 strSQL = strSQL & " order by fechas"
 rs.CursorLocation = adUseServer
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While rs.EOF = False
  Set itmX = .ListItems.Add(.ListItems.Count + 1, , Format(IIf(IsNull(rs!fechap), "", rs!fechap), "####-##"), , 3)
   itmX.SubItems(1) = Format(IIf(IsNull(rs!Fechas), Date, rs!Fechas), "dd/mm/yyyy")
   itmX.SubItems(2) = Format(IIf(IsNull(rs!cuota), 0, rs!cuota), "###,###,###,##0.00")
   itmX.SubItems(3) = Format(IIf(IsNull(rs!Abono), 0, rs!Abono), "###,###,###,##0.00")
   itmX.SubItems(4) = Format(IIf(IsNull(rs!intcp), 0, rs!intcp), "###,###,###,##0.00")
   itmX.SubItems(5) = Format(IIf(IsNull(rs!Amortiza), 0, rs!Amortiza), "###,###,###,##0.00")
   itmX.SubItems(6) = fxTipoComprobante(IIf(IsNull(rs!tcon), 0, rs!tcon))
   itmX.SubItems(7) = IIf(IsNull(rs!nCon), 0, rs!nCon)
   itmX.SubItems(8) = fxEstadoCuota(IIf(IsNull(rs!Estado), "", rs!Estado))
   itmX.Tag = itmX.Index
   curTotales(1) = curTotales(1) + IIf(IsNull(rs!cuota), 0, rs!cuota)
   curTotales(2) = curTotales(2) + IIf(IsNull(rs!Abono), 0, rs!Abono)
   curTotales(3) = curTotales(3) + IIf(IsNull(rs!intcp), 0, rs!intcp)
   curTotales(4) = curTotales(4) + IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
   
  rs.MoveNext
 Loop
 rs.Close
 
  Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
   itmX.SubItems(2) = "---------------"
   itmX.SubItems(3) = "---------------"
   itmX.SubItems(4) = "---------------"
   itmX.SubItems(5) = "---------------"
   
 
  Set itmX = .ListItems.Add(.ListItems.Count + 1, , "TOTALES")
   itmX.SubItems(2) = Format(curTotales(1), "###,###,###,##0.00")
   itmX.SubItems(3) = Format(curTotales(2), "###,###,###,##0.00")
   itmX.SubItems(4) = Format(curTotales(3), "###,###,###,##0.00")
   itmX.SubItems(5) = Format(curTotales(4), "###,###,###,##0.00")
 

End With
End Sub

Sub LlenaCuotasMorosas()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Llena Lsw de Cuotas con el Listado de Cuotas en Mora (Canceladas, Activas,etc)
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, itmX As ListItem, strSQL As String
Dim curTotales(8) As Currency

curTotales(1) = 0
curTotales(2) = 0
curTotales(3) = 0
curTotales(4) = 0
curTotales(5) = 0
curTotales(6) = 0
curTotales(7) = 0
curTotales(8) = 0


With lswAbonos
 Select Case cboTipoCuotas.Text
  Case "Todas"
    strSQL = "select * from morosidad where id_solicitud = " & txtOperacion
  Case "Activas"
    strSQL = "select * from morosidad where id_solicitud = " & txtOperacion & " and estado = 'A'"
  Case "Anuladas"
    strSQL = "select * from morosidad where id_solicitud = " & txtOperacion & " and estado = 'N'"
  Case "Canceladas"
    strSQL = "select * from morosidad where id_solicitud = " & txtOperacion & " and estado = 'C'"
  Case Else
    strSQL = "select * from morosidad where id_solicitud = " & txtOperacion & " and estado = 'A'"
 End Select
 strSQL = strSQL & " order by fechap"
 rs.CursorLocation = adUseServer
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While rs.EOF = False
  Set itmX = .ListItems.Add(.ListItems.Count + 1, , Format(IIf(IsNull(rs!fechap), "", rs!fechap), "####-##"), , 3)
   itmX.SubItems(1) = Format(IIf(IsNull(rs!fecult), Date, rs!fecult), "dd/mm/yyyy")
   itmX.SubItems(2) = Format(IIf(IsNull(rs!intc), 0, rs!intc), "###,###,###,##0.00")
   itmX.SubItems(3) = Format(IIf(IsNull(rs!intm), 0, rs!intm), "###,###,###,##0.00")
   itmX.SubItems(4) = Format(IIf(IsNull(rs!Amortiza), 0, rs!Amortiza), "###,###,###,##0.00")
   itmX.SubItems(5) = Format(IIf(IsNull(rs!Amortiza), 0, rs!Amortiza) + IIf(IsNull(rs!intm), 0, rs!intm) + IIf(IsNull(rs!intc), 0, rs!intc), "###,###,###,##0.00")
   itmX.SubItems(6) = Format(IIf(IsNull(rs!abintc), 0, rs!abintc), "###,###,###,##0.00")
   itmX.SubItems(7) = Format(IIf(IsNull(rs!abintm), 0, rs!abintm), "###,###,###,##0.00")
   itmX.SubItems(8) = Format(IIf(IsNull(rs!abAmortiza), 0, rs!abAmortiza), "###,###,###,##0.00")
   itmX.SubItems(9) = Format(IIf(IsNull(rs!abAmortiza), 0, rs!abAmortiza) + IIf(IsNull(rs!abintm), 0, rs!abintm) + IIf(IsNull(rs!abintc), 0, rs!abintc), "###,###,###,##0.00")
   itmX.SubItems(10) = fxTipoComprobante(IIf(IsNull(rs!tcon), 0, rs!tcon))
   itmX.SubItems(11) = IIf(IsNull(rs!nCon), 0, rs!nCon)
   itmX.SubItems(12) = fxEstadoCuota(IIf(IsNull(rs!Estado), "", rs!Estado))
   itmX.SubItems(13) = fxEstadoCuota(IIf(IsNull(rs!Estadoi), "", rs!Estadoi))
   itmX.Tag = itmX.Index

   
   curTotales(1) = curTotales(1) + IIf(IsNull(rs!intc), 0, rs!intc)
   curTotales(2) = curTotales(2) + IIf(IsNull(rs!intm), 0, rs!intm)
   curTotales(3) = curTotales(3) + IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
   curTotales(4) = curTotales(4) + IIf(IsNull(rs!Amortiza), 0, rs!Amortiza) + IIf(IsNull(rs!intm), 0, rs!intm) + IIf(IsNull(rs!intc), 0, rs!intc)
   curTotales(5) = curTotales(5) + IIf(IsNull(rs!abintc), 0, rs!abintc)
   curTotales(6) = curTotales(6) + IIf(IsNull(rs!abintm), 0, rs!abintm)
   curTotales(7) = curTotales(7) + IIf(IsNull(rs!abAmortiza), 0, rs!abAmortiza)
   curTotales(8) = curTotales(8) + IIf(IsNull(rs!abAmortiza), 0, rs!abAmortiza) + IIf(IsNull(rs!abintm), 0, rs!abintm) + IIf(IsNull(rs!abintc), 0, rs!abintc)
   
  rs.MoveNext
 Loop
 rs.Close
 
   Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
   itmX.SubItems(2) = "---------------"
   itmX.SubItems(3) = "---------------"
   itmX.SubItems(4) = "---------------"
   itmX.SubItems(5) = "---------------"
   itmX.SubItems(6) = "---------------"
   itmX.SubItems(7) = "---------------"
   itmX.SubItems(8) = "---------------"
   itmX.SubItems(9) = "---------------"
    
  Set itmX = .ListItems.Add(.ListItems.Count + 1, , "TOTALES")
   itmX.SubItems(2) = Format(curTotales(1), "###,###,###,##0.00")
   itmX.SubItems(3) = Format(curTotales(2), "###,###,###,##0.00")
   itmX.SubItems(4) = Format(curTotales(3), "###,###,###,##0.00")
   itmX.SubItems(5) = Format(curTotales(4), "###,###,###,##0.00")
   itmX.SubItems(6) = Format(curTotales(5), "###,###,###,##0.00")
   itmX.SubItems(7) = Format(curTotales(6), "###,###,###,##0.00")
   itmX.SubItems(8) = Format(curTotales(7), "###,###,###,##0.00")
   itmX.SubItems(9) = Format(curTotales(8), "###,###,###,##0.00")

 
End With
End Sub

Private Sub sbBusqueda(Index As Integer)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Realizar Busquedas Rapidas de la ventana
'REFERENCIAS   : ConsultaOperacion - (Refresca la informacion de la ventana)
'                CambiaDatos       - (Limpia ciertos campos de la ventana)
'                fxDescribeCodigo  - (Devuelve la descripcion de el código de crédito)
'                frmBusquedas      - (Formulario de Busquedas Rápidas
'OBSERVACIONES : Bloque de Busquedas Rápidas
'-------------------------------------------------------------------------------------------

gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Index
  Case 1 'txtOperacion
    gBusquedas.Consulta = "select id_solicitud as Operacion,codigo,cedula,montoapr,saldo from reg_creditos"
    gBusquedas.Orden = "id_solicitud"
    gBusquedas.Columna = "id_solicitud"
    gBusquedas.Filtro = " and estadosol = 'F'"
    frmBusquedas.Show vbModal
    txtOperacion = gBusquedas.Resultado
    If Len(Trim(txtOperacion)) > 0 Then
      Call ConsultaOperacion
    End If
  Case 2 'txtCodigo
   If Len(Trim(txtCedula)) > 0 Then
        gBusquedas.Consulta = "select id_solicitud as Operacion,codigo,cedula,proceso,estado from reg_creditos"
        gBusquedas.Orden = "id_solicitud"
        gBusquedas.Columna = "id_solicitud"
        gBusquedas.Filtro = " and estadosol = 'F' and cedula ='" & txtCedula & "'"
        frmBusquedas.Show vbModal
        txtOperacion = gBusquedas.Resultado
        If Len(Trim(txtOperacion)) > 0 Then
          Call ConsultaOperacion
        End If
    Else
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
    End If
  
  Case 3 'txtCedula
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  Case 4 'txtDescripcion
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  Case 5 'txtNombre
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
  Case 6 'txtCodigoNuevo
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        frmBusquedas.Show vbModal
        txtCodigoNuevo = gBusquedas.Resultado
        If Len(Trim(txtCodigoNuevo)) > 0 Then
          lblCodigoNuevo.Caption = fxDescribeCodigo(Trim(txtCodigoNuevo))
        End If
  Case 7 'txtReporteX
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        frmBusquedas.Show vbModal
        txtReporteX = gBusquedas.Resultado
        If Len(Trim(txtReporteX)) > 0 Then
          lblXDescribe.Caption = fxDescribeCodigo(Trim(txtReporteX))
        End If
End Select

End Sub

Function fxSaldo(lngSolicitud As Long)
Dim rsX As New ADODB.Recordset
With rsX
 .Open "select saldo from reg_creditos where id_solicitud = " & lngSolicitud, glogon.Conection, adOpenStatic
 If .EOF And .BOF Then
  fxSaldo = 0
 Else
  fxSaldo = !Saldo
 End If
 .Close
End With
End Function

Sub AntiguedadSaldos(strCodigo As String)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Procesa Antiguedad de Saldos de los Códigos Seleccionados
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Existen varios tipos de selecciones, segun el tipo de antiguedad que se
'                desee, como : Mora Legal, Mora Financiera y Mora Legal sin intereses
'                ANTIGUEDAD DE SALDOS POR NUMERO DE CUOTAS
'-------------------------------------------------------------------------------------------

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "insert antiguedad_saldos(codigo) values('" & strCodigo & "')"

glogon.Conection.Execute strSQL

Select Case True
 Case optAntiguedad(0).Value 'Mora Legal
    strSQL = "select A.ID_SOLICITUD,A.CODIGO,A.INTC,A.INTM,A.CUOTA,B.SALDO " _
            & " from Vista_Morosidad A INNER JOIN Reg_creditos B ON A.id_solicitud = B.id_solicitud " _
            & " inner join Socios S on B.cedula = S.cedula" _
            & " Where A.codigo = '" & strCodigo & "'"
 Case optAntiguedad(1).Value 'Mora Financiera
    strSQL = "select V.ID_SOLICITUD,V.CODIGO,V.INTC,V.INTM,V.CUOTA,V.AMORTIZA AS SALDO " _
            & " from Vista_Morosidad V inner join reg_creditos R on V.id_solicitud = R.id_solicitud" _
            & " inner join Socios S on R.cedula = S.cedula" _
            & " Where V.codigo = '" & strCodigo & "'"
 End Select

'Filtro
Select Case Mid(cboRepAnt.Text, 1, 2)
    Case "00" 'Todos
    Case "01" 'Socios
      strSQL = strSQL & " AND S.estadoActual = 'S'"
    Case "02" 'Opex
      strSQL = strSQL & " AND S.estadoActual in('A','P')"
    Case "03" 'No Socios
      strSQL = strSQL & " AND S.estadoactual = 'N'"
    Case "04" 'Ren.Interna
      strSQL = strSQL & " AND S.estadoActual = 'A'"
    Case "05" 'Ren.Patronal
      strSQL = strSQL & " AND S.estadoActual = 'P'"
End Select


lblAntiguedad.Caption = "Procesando Antiguedad Código : " & Trim(strCodigo)
lblAntiguedad.Refresh

prgBar.Value = 1

rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
 strSQL = "update antiguedad_saldos set "
 
If chkSaldo.Value = vbUnchecked Then
 
 Select Case rs!cuota
   Case 1
    strSQL = strSQL + " rng1_monto = rng1_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng1_casos = rng1_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 2
    strSQL = strSQL + " rng2_monto = rng2_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng2_casos = rng2_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 3
    strSQL = strSQL + " rng3_monto = rng3_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng3_casos = rng3_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 4, 5, 6
    strSQL = strSQL + " rng4_monto = rng4_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng4_casos = rng4_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 7, 8, 9, 10, 11, 12
    strSQL = strSQL + " rng5_monto = rng5_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng5_casos = rng5_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case Else 'es mayor
    strSQL = strSQL + " rng6_monto = rng6_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng6_casos = rng6_casos + 1 where codigo = '" & rs!Codigo & "'"
 End Select

Else

 Select Case rs!cuota
   Case 1
    strSQL = strSQL + " rng1_monto = rng1_monto + " & rs!Saldo _
        & ", rng1_casos = rng1_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 2
    strSQL = strSQL + " rng2_monto = rng2_monto + " & rs!Saldo _
        & ", rng2_casos = rng2_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 3
    strSQL = strSQL + " rng3_monto = rng3_monto + " & rs!Saldo _
        & ", rng3_casos = rng3_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 4, 5, 6
    strSQL = strSQL + " rng4_monto = rng4_monto + " & rs!Saldo _
        & ", rng4_casos = rng4_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 7, 8, 9, 10, 11, 12
    strSQL = strSQL + " rng5_monto = rng5_monto + " & rs!Saldo _
        & ", rng5_casos = rng5_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case Else 'es mayor
    strSQL = strSQL + " rng6_monto = rng6_monto + " & rs!Saldo _
        & ", rng6_casos = rng6_casos + 1 where codigo = '" & rs!Codigo & "'"
 End Select
 
 End If
 
 rs.MoveNext
 glogon.Conection.Execute strSQL
 prgBar.Value = prgBar.Value + 1
Loop

rs.Close

'Limpia Movimientos en Cero
strSQL = "delete antiguedad_saldos where (rng1_monto + rng2_monto + rng3_monto + rng4_monto" _
       & " + rng5_monto + rng6_monto) = 0"
glogon.Conection.Execute strSQL



End Sub


Sub AntiguedadSaldosFecha(strCodigo As String)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Procesa Antiguedad de Saldos de los Códigos Seleccionados
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Existen varios tipos de selecciones, segun el tipo de antiguedad que se
'                desee, como : Mora Legal, Mora Financiera y Mora Legal sin intereses
'                ANTIGUEDAD DE SALDOS POR FECHAS DE LAS CUOTAS
'-------------------------------------------------------------------------------------------

Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date


vFecha = Format(fxFechaServidor, "yyyy/mm/dd")

strSQL = "insert antiguedad_saldos(codigo) values('" & strCodigo & "')"

glogon.Conection.Execute strSQL

strSQL = "select M.ID_SOLICITUD,M.CODIGO,M.INTC,M.INTM,M.AMORTIZA AS SALDO,M.FECULT" _
        & " from Morosidad M inner join reg_creditos R on M.id_solicitud = R.id_solicitud" _
        & " inner join Socios S on R.cedula = S.cedula" _
        & " Where M.codigo = '" & strCodigo & "' and M.estado = 'A'"
'Filtro
Select Case Mid(cboRepAnt.Text, 1, 2)
    Case "00" 'Todos
    Case "01" 'Socios
      strSQL = strSQL & " AND S.estadoActual = 'S'"
    Case "02" 'Opex
      strSQL = strSQL & " AND S.estadoActual in('A','P')"
    Case "03" 'No Socios
      strSQL = strSQL & " AND S.estadoactual = 'N'"
    Case "04" 'Ren.Interna
      strSQL = strSQL & " AND S.estadoActual = 'A'"
    Case "05" 'Ren.Patronal
      strSQL = strSQL & " AND S.estadoActual = 'P'"
End Select


lblAntiguedad.Caption = "Procesando Antiguedad Código : " & Trim(strCodigo)
lblAntiguedad.Refresh

prgBar.Value = 1

rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
 strSQL = "update antiguedad_saldos set "
 
If chkSaldo.Value = 0 Then
  
 
 Select Case DateDiff("m", Format(rs!fecult, "yyyy/mm/dd"), vFecha) + 1
   Case 1
    strSQL = strSQL + " rng1_monto = rng1_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng1_casos = rng1_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 2
    strSQL = strSQL + " rng2_monto = rng2_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng2_casos = rng2_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 3
    strSQL = strSQL + " rng3_monto = rng3_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng3_casos = rng3_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 4, 5, 6
    strSQL = strSQL + " rng4_monto = rng4_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng4_casos = rng4_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 7, 8, 9, 10, 11, 12
    strSQL = strSQL + " rng5_monto = rng5_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng5_casos = rng5_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case Else 'es mayor
    strSQL = strSQL + " rng6_monto = rng6_monto + " & rs!Saldo + rs!intc + rs!intm _
        & ", rng6_casos = rng6_casos + 1 where codigo = '" & rs!Codigo & "'"
 End Select

Else

 Select Case DateDiff("m", vFecha, Format(rs!fecult, "yyyy/mm/dd")) + 1
   Case 1
    strSQL = strSQL + " rng1_monto = rng1_monto + " & rs!Saldo _
        & ", rng1_casos = rng1_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 2
    strSQL = strSQL + " rng2_monto = rng2_monto + " & rs!Saldo _
        & ", rng2_casos = rng2_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 3
    strSQL = strSQL + " rng3_monto = rng3_monto + " & rs!Saldo _
        & ", rng3_casos = rng3_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 4, 5, 6
    strSQL = strSQL + " rng4_monto = rng4_monto + " & rs!Saldo _
        & ", rng4_casos = rng4_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case 7, 8, 9, 10, 11, 12
    strSQL = strSQL + " rng5_monto = rng5_monto + " & rs!Saldo _
        & ", rng5_casos = rng5_casos + 1 where codigo = '" & rs!Codigo & "'"
   Case Else 'es mayor
    strSQL = strSQL + " rng6_monto = rng6_monto + " & rs!Saldo _
        & ", rng6_casos = rng6_casos + 1 where codigo = '" & rs!Codigo & "'"
 End Select
 
 End If
 
 rs.MoveNext
 glogon.Conection.Execute strSQL
 prgBar.Value = prgBar.Value + 1
Loop
rs.Close

'Limpia Movimientos en Cero
strSQL = "delete antiguedad_saldos where (rng1_monto + rng2_monto + rng3_monto + rng4_monto" _
       & " + rng5_monto + rng6_monto) = 0"
glogon.Conection.Execute strSQL

End Sub

Private Sub imgReporteAntiguedad_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprimir Antiguedad de Saldos
'REFERENCIAS   : AntiguedadSaldos - (Procesa Antiguedad de Saldos)
'                fxFechaServidor - (Devuelve la Fecha del Servidor)
'OBSERVACIONES : Se Debe de Borrar la Tabla de Salida antes de procesar A.S.
'                Utiliza Variables globales
'-------------------------------------------------------------------------------------------
Dim itmX As ListItem, lng As Long, intPasa As Integer
Dim strSubTitulo As String, iRespuesta As Integer

If cboAntiguedad.Text <> "Con Base al Número de Cuotas" _
   And optAntiguedad.Item(0).Value Then
   
   MsgBox "No se puede aplicar Mora Legal con base a las fechas de las cuotas", vbCritical
   
   Exit Sub
End If

Me.MousePointer = vbHourglass
intPasa = 0
On Error GoTo vError
strSubTitulo = "EMITIDO ANTERIORMENTE"

iRespuesta = MsgBox("Desea Generar un Listado de Antiguedad Actualizado a la fecha", vbYesNo)

If iRespuesta = vbYes Then
    
    strSubTitulo = InputBox("Digite el SubTitulo Para la Antiguedad de Saldos : ", "SubTitulo del Reporte de Antiguedad de Saldos")
    If cboAntiguedad.Text = "Con Base al Número de Cuotas" Then
        strSubTitulo = "[Basado en Cuotas] [" & strSubTitulo & "]"
    Else
        strSubTitulo = "[Basado en Fechas] [" & strSubTitulo & "]"
    End If
    fraAntiguedad.Visible = True
    fraAntiguedad.Refresh
    
    glogon.Conection.Execute "delete antiguedad_saldos"

    With lswAntiguedad
       For lng = 1 To .ListItems.Count
        If .ListItems.Item(lng).Checked = True Then
           If cboAntiguedad.Text = "Con Base al Número de Cuotas" Then
              Call AntiguedadSaldos(.ListItems.Item(lng).Text)
           Else
              Call AntiguedadSaldosFecha(.ListItems.Item(lng).Text)
           End If
         intPasa = 1
        End If
       Next lng
     End With
    
    fraAntiguedad.Visible = False
Else
 intPasa = 1
End If 'Regenerar informacion

If intPasa = 1 Then
    With frmContenedor.Crt
        .Reset
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .WindowTitle = "Reportes del Módulo de Cobro Administrativo y Judicial"
        .ReportFileName = App.Path + "\Credito\Reportes\AntiguedadSaldos.rpt"
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
       Select Case True
          Case optAntiguedad(0).Value
               .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS - MORA LEGAL" & IIf((chkSaldo.Value = 0), "'", " - Sin Int.'")
          Case optAntiguedad(1).Value
               .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS - MORA FINANCIERA'"
       End Select
        .Formulas(2) = "Subtitulo='" & UCase(strSubTitulo) & " / FILTRO " & Mid(cboRepAnt.Text, 4, 30) & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .PrintReport
    End With
Else
 MsgBox "No se Procesó ningun código de préstamo...", vbInformation
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
End Sub


Private Sub imgReporteGeneral_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprime Reportes Generales de Cobro
'REFERENCIAS   : fxFechaServidor - (Devuelve la Fecha del Servidor)
'OBSERVACIONES : Utiliza variables globales
'-------------------------------------------------------------------------------------------
Dim strSQL As String, vSubTitulo As String

Me.MousePointer = vbHourglass

Select Case Mid(cboRep.Text, 1, 2)
 Case "00" 'Todos
   strSQL = ""
 Case "01" 'Socios
   strSQL = " AND {SOCIOS.ESTADOACTUAL} = 'S'"
 Case "02" 'Opex
   strSQL = " AND ({SOCIOS.ESTADOACTUAL} = 'A' OR {SOCIOS.ESTADOACTUAL} = 'P')"
 Case "03" 'No Socios
   strSQL = " AND {SOCIOS.ESTADOACTUAL} = 'N'"
 Case "04" 'Ren.Interna
   strSQL = " AND {SOCIOS.ESTADOACTUAL} = 'A'"
 Case "05" 'Ren.Patronal
   strSQL = " AND {SOCIOS.ESTADOACTUAL} = 'P'"
End Select


If cboDestino.Text <> "TODOS" Then
  strSQL = strSQL & " AND {REG_CREDITOS.COD_DESTINO} = '" & fxCodigoCbo(cboDestino) & "'"
End If

If cboGarantia.Text <> "TODOS" Then
  strSQL = strSQL & " AND {REG_CREDITOS.GARANTIA} = '" & Mid(cboGarantia.Text, 1, 1) & "'"
End If

vSubTitulo = "Cartera : " & cboCartera.Text & " / Estado : " & Mid(cboRep.Text, 4, 30) _
                 & " / Garantía : " & cboGarantia.Text & " / Destino : " & cboDestino.Text

If chkLineas.Value = vbChecked Then
  vSubTitulo = vSubTitulo & " / Línea : Todas"
Else
  strSQL = strSQL & " AND {REG_CREDITOS.CODIGO} = '" & txtReporteX.Text & "'"
  vSubTitulo = vSubTitulo & " / Línea : " & txtReporteX.Text
End If

If cboCartera.Text <> "(Todas las Carteras)" Then
    strSQL = strSQL & " AND {CBR_CLASIFICACION_DETALLE.COD_CLASIFICACION} = '" & fxCodigoCbo(cboCartera) & "'"
End If

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro Administrativo y Judicial"
     
     
  Select Case lblRepGen.Tag
   Case "GENDET" 'Reporte General - Detallado
    .ReportFileName = App.Path & "\Credito\Reportes\CbrListadoDetallado.rpt"
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Titulo='REPORTE GENERAL DETALLADO DE MOROSIDAD'"
    .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
    .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
    .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
    .SelectionFormula = .SelectionFormula & strSQL

   Case "GENRSM" 'Reporte General - Resumen
    .ReportFileName = App.Path & "\Credito\Reportes\CbrListadoResumen.rpt"
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
    .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Cuota='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
    .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
    .SelectionFormula = .SelectionFormula & strSQL
   
   
   Case "GENAGD"    ' General - Detallado Agrupado"
    .ReportFileName = App.Path & "\Credito\Reportes\CbrListadoDetalladoAgr.rpt"
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Titulo='REPORTE GENERAL DETALLADO DE MOROSIDAD - ESP.AGRUPADO'"
    .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
    .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
    .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
    .SelectionFormula = .SelectionFormula & strSQL
    
   Case "GENRGD"     'General - Resumen Agrupado"
    .ReportFileName = App.Path & "\Credito\Reportes\CbrListadoResumenAgr.rpt"
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
    .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Cuota='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
    .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
    .SelectionFormula = .SelectionFormula & strSQL
   
   
   
   Case "ESPCON" 'Convenios
    .ReportFileName = App.Path & "\Credito\Reportes\CbrListadoXConvenios.rpt"
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
    .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & " / FILTRO " & Mid(cboRep.Text, 4, 30) & "'"
    .Formulas(3) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
    .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta & " And {VISTA_MOROSIDAD.CODIGO}='" & txtReporteX & "'"
    .SelectionFormula = .SelectionFormula & strSQL
  
   Case "MORCAR" 'Resumen Comparativo
    .ReportFileName = App.Path & "\Credito\Reportes\CbrComparativo.rpt"
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
    .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    Select Case Mid(cboRep.Text, 1, 2)
     Case "00" 'Todos
        .StoredProcParam(0) = "S"
        .StoredProcParam(1) = "A"
        .StoredProcParam(2) = "P"
        .StoredProcParam(3) = "N"
     Case "01" 'Socios
        .StoredProcParam(0) = "S"
        .StoredProcParam(1) = "S"
        .StoredProcParam(2) = "S"
        .StoredProcParam(3) = "S"
     Case "02" 'Opex
        .StoredProcParam(0) = "A"
        .StoredProcParam(1) = "A"
        .StoredProcParam(2) = "P"
        .StoredProcParam(3) = "P"
     Case "03" 'No Socios
        .StoredProcParam(0) = "N"
        .StoredProcParam(1) = "N"
        .StoredProcParam(2) = "N"
        .StoredProcParam(3) = "N"
    End Select

   strSQL = ""
   If cboGarantia.Text <> "TODOS" Then
      strSQL = "{spCBRComparativo;1.garantia} = '" & Mid(cboGarantia.Text, 1, 1) & "'"
   End If
   If chkLineas.Value = vbUnchecked Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{spCBRComparativo;1.codigo} = '" & txtReporteX.Text & "'"
   End If

   If cboCartera.Text <> "(Todas las Carteras)" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{spCBRComparativo;1.cod_clasificacion} = '" & fxCodigoCbo(cboCartera) & "'"
   End If


   .SelectionFormula = strSQL
     
  
   Case "MORGAR" 'Reporte Mora x Garantia
        .ReportFileName = App.Path & "\Credito\Reportes\CbrListadoXGarantia.rpt"
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = "{REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'" & strSQL
  
    Case "MORGAG"  'Mora x Garantía - Agrupado
        .ReportFileName = App.Path & "\Credito\Reportes\CbrListadoXGarantiaAgr.rpt"
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = "{REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'" & strSQL
  
  
  End Select
    
    .PrintReport
End With

Me.MousePointer = vbDefault


End Sub

Private Sub imgReporteOperacion_Click()

'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprime Reportes sobre la operacion u operaciones
'REFERENCIAS   : fxFechaServidor - (Devuelve la Fecha del Servidor)
'OBSERVACIONES : Utiliza variables globales
'-------------------------------------------------------------------------------------------
Dim strRuta As String, strSQL As String, vMes As Integer
Dim rs As New ADODB.Recordset

strSQL = "select * from par_ahcr"
rs.Open strSQL, glogon.Conection, adOpenStatic
vMes = Mid(GLOBALES.glngFechaCR, 5, 2)
If rs!cr_apl = 0 Then
 If vMes = 1 Then
   vMes = 12
 Else
   vMes = vMes - 1
 End If
End If
rs.Close

Me.MousePointer = vbHourglass

With frmContenedor.Crt
  strRuta = App.Path + "\Credito\Reportes\"
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Cobro Administrativo y Judicial"


Select Case lblRepOp.Tag
  Case "ULTEC" 'Ultimo Estado
    Select Case lblProceso.Caption
     Case "TRASPASO DEUDAS"
          .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
          .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
          .ReportFileName = strRuta + "BoletaTraspaso.rpt"
     Case "COBRO JUDICIAL"
          .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
          .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
          .ReportFileName = strRuta + "BoletaCobroJudicial.rpt"
    End Select
    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
  
  Case "ECBR" 'EC CBR Resumen

  Case "ETSBR" 'Etiquetas Sobres
         .ReportFileName = strRuta + "Sobres.rpt"
         .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
  
  Case "PRIAVI" 'Primer Aviso
     .ReportFileName = strRuta + "Carta1.rpt"
     .Formulas(0) = "MesProceso = '" & Format(vMes, "00") & "'"
     .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA} >= 1 and {REG_CREDITOS.CEDULA} = '" & Trim(txtCedula) & "'"
     .SubreportToChange = "Fiadores"
     .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA} >= 1 And {REG_CREDITOS.CEDULA} = '" & Trim(txtCedula) & "'"
  
  Case "SEGAVI" 'Segundo Aviso
     .ReportFileName = strRuta + "Carta2.rpt"
     .Formulas(0) = "MesProceso = '" & Format(vMes, "00") & "'"
     .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA} >= 1 and {REG_CREDITOS.CEDULA} = '" & Trim(txtCedula) & "'"
     .SubreportToChange = "Fiadores"
     .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA} >= 1 And {REG_CREDITOS.CEDULA} = '" & Trim(txtCedula) & "'"
     
  Case "NOTMOV" 'Notificacion del Movimiento
     .ReportFileName = strRuta + "Notificacion.rpt"
     .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
  
  Case "REVER" 'Boleta Reversion
    fraFechas.Visible = True
  Case "ENVCBR" 'Casos Enviados a Cobro Judicial
    fraFechas.Visible = True
  Case "TRADEUD" 'Casos Traspaso - Deudor
    fraFechas.Visible = True
  Case "TRAFIA" 'Casos Traspaso - Fiadores
    fraFechas.Visible = True
  
  Case "TRAREV" 'Boleta de Reversion de Traspaso de Deudas
          .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
          .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
          .Formulas(2) = "subtitulo='BOLETA DE REVERSION DE TRASPASO'"
          .ReportFileName = strRuta + "BoletaTraspasoReversion.rpt"
          .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
End Select

 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub imgVisualiza_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Llena informacion de las cuotas de la operacion
'REFERENCIAS   : LLenaAbonos - (Carga Abonos Ordinarios y Extraordinarios)
'                LlenaCuotasMorosas - (Carga Cuotas en Mora)
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------

Me.MousePointer = vbHourglass

lblCuotas.Caption = cboVisualiza.Text & " - " & cboTipoCuotas.Text
lblCuotas.Refresh

With lswAbonos
  .ListItems.Clear
  .ColumnHeaders.Clear
    Select Case cboVisualiza.Text
      Case "Abonos Ordinarios"
        .ColumnHeaders.Add , , "Proceso", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Cuota", 1200, 1
        .ColumnHeaders.Add , , "Abono", 1200, 1
        .ColumnHeaders.Add , , "Intereses", 1200, 1
        .ColumnHeaders.Add , , "Amortización", 1200, 1
        .ColumnHeaders.Add , , "T.Comp", 1000
        .ColumnHeaders.Add , , "# Comp", 1200
        .ColumnHeaders.Add , , "Estado", 1200
        
        Call LlenaAbonos
      
      Case "Abonos a Cuotas Morosas"
        .ColumnHeaders.Add , , "Proceso", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Int.Cor", 1200, 1
        .ColumnHeaders.Add , , "Int.Mor", 1200, 1
        .ColumnHeaders.Add , , "Amortización", 1200, 1
        .ColumnHeaders.Add , , "T.Atrasado", 1200, 1
        .ColumnHeaders.Add , , "Ab.Int.Cor", 1200, 1
        .ColumnHeaders.Add , , "Ab.Int.Mor", 1200, 1
        .ColumnHeaders.Add , , "Ab.Amortiza", 1200, 1
        .ColumnHeaders.Add , , "T.Abono", 1200, 1
        .ColumnHeaders.Add , , "T.Comp.", 1000
        .ColumnHeaders.Add , , "# Comp.", 1200
        .ColumnHeaders.Add , , "Estado", 1000
        .ColumnHeaders.Add , , "Est.Orig.", 1000
        Call LlenaCuotasMorosas
    
    End Select
End With

Me.MousePointer = vbDefault

End Sub


Private Sub lswAbonos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 Set Conlsw.lswX = lswAbonos
 Conlsw.Abre
End If
End Sub


Private Function fxPlazoRestante(vOperacion As Long) As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vResultado As Long, vPrimerDeduc As Long
Dim vAnio As Integer, vMes As Byte

strSQL = "select Plazo,PriDeduc from reg_creditos where id_solicitud = " & vOperacion
rs.Open strSQL, glogon.Conection, adOpenStatic

vPrimerDeduc = fxPrimerDeduccion

If Not rs.EOF And Not rs.BOF Then
 vAnio = Mid(CStr(rs!prideduc), 1, 4)
 vMes = Mid(CStr(rs!prideduc), 5, 2)
 i = 0
 vResultado = CLng(vAnio & Format(vMes, "00"))
 Do While vPrimerDeduc >= vResultado
   If vMes = 12 Then
      vMes = 1
      vAnio = vAnio + 1
   Else
      vMes = vMes + 1
   End If
   i = i + 1
   vResultado = CLng(vAnio & Format(vMes, "00"))
 Loop
 i = rs!Plazo - i
 
 i = i + 1 'Plazo Restante + 1, para Cuadrar Fecha de Conclusion
 
 If i <= 0 Then i = 1
 fxPlazoRestante = i

Else
 fxPlazoRestante = 1
End If
rs.Close

End Function

Private Sub lswFiadores_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Despliega datos del fiador Seleccionado, y lo activa para posible traspaso
'                de deudas
'REFERENCIAS   : fxProvincia - (Devuelve el número o descripcion de las provincias)
'                Telefonos   - (Carga los número telefonicos del fiador)
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, strProvincia As String
Me.MousePointer = vbHourglass
txtDirFiadores = ""
lswTelefonosFiadores.ListItems.Clear

If lswFiadores.SelectedItem.Text <> "" Then
 lblCedulaFiador.Caption = lswFiadores.SelectedItem.Text
 lblNombreFiador.Caption = lswFiadores.SelectedItem.ListSubItems(1).Text
 txtInteresFiador = txtInteresActual
 txtPlazoFiador = fxPlazoRestante(txtOperacion)

With lswFiadores
 rs.Open "select provincia,canton,distrito,direccion from socios where cedula = '" _
    & .SelectedItem.Text & "'", glogon.Conection, adOpenStatic
 If Not rs.EOF And Not rs.BOF Then
  strProvincia = fxProvincia(IIf(IsNull(rs!provincia), 11, rs!provincia))
  txtDirFiadores = "PROVINCIA : " & strProvincia & vbCrLf _
     & "CANTON : " & fxCanton(IIf(IsNull(rs!provincia), 11, rs!provincia), IIf(IsNull(rs!canton), 0, rs!canton)) & vbCrLf _
     & "DISTRITO : " & IIf(IsNull(rs!Distrito), "", rs!Distrito) & vbCrLf _
     & "DIR : " & IIf(IsNull(rs!Direccion), "", rs!Direccion)
 End If
 rs.Close

 Call Telefonos(lswTelefonosFiadores, .SelectedItem.Text)
 
End With

End If 'selected
Me.MousePointer = vbDefault
End Sub

Sub OperacionesGeneradas()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Carga las operaciones de los fiadores que se generaron por un traspaso de
'                deudas.
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, itmX As ListItem
Dim colReferencias() As Long, i As Integer, iTotal As Integer

Label17(3).Refresh
lblMontoActualDeudor.Caption = 0
lblPlazoActualDeudor.Caption = 0
lblInteresActualDeudor.Caption = 0
lblSaldoActualDeudor.Caption = 0
lblOperacionActualDeudor.Caption = 0

With lswOperacionesGeneradas
 .ListItems.Clear
 rs.CursorLocation = adUseServer
 rs.Open "select * from reg_creditos where estado in('A','C') and referencia = " & txtOperacion, glogon.Conection, adOpenStatic
 ReDim colReferencias(rs.RecordCount) As Long
 iTotal = rs.RecordCount
 i = 1
 Do While rs.EOF = False
  If Trim(rs!Cedula) = Trim(txtCedula) Then
    lblMontoActualDeudor.Caption = Format(rs!montoapr, "###,###,###,##0.00")
    lblPlazoActualDeudor.Caption = rs!Plazo
    lblInteresActualDeudor.Caption = rs!interesv
    lblSaldoActualDeudor.Caption = Format(rs!Saldo, "###,###,###,##0.00")
    lblOperacionActualDeudor.Caption = rs!ID_SOLICITUD
  Else
    Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!ID_SOLICITUD, , 3)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = Format(rs!montoapr, "###,###,###,##0.00")
     itmX.SubItems(4) = Format(rs!Saldo, "###,###,###,##0.00")
     itmX.SubItems(5) = Format(rs!cuota, "###,###,###,##0.00")
     itmX.SubItems(6) = rs!Int
     itmX.SubItems(7) = rs!Plazo
     itmX.SubItems(8) = fxTotalInteresEnMora(rs!ID_SOLICITUD)
     colReferencias(i) = rs!ID_SOLICITUD
     itmX.Tag = itmX.Index
     i = i + 1
  End If
    rs.MoveNext
 Loop
 rs.Close
 
End With
End Sub


Private Sub lswOperacionesGeneradas_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Carga Montos para una posible Reversion de Traspaso de deudas
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ver Traspaso de Deudas y Reversion
'-------------------------------------------------------------------------------------------

'Volver a Calcular los montos de las Operaciones
Dim itmX As ListItem, lng As Long

txtTRAFD_MONTO = 0

 With lswOperacionesGeneradas
   For lng = 1 To .ListItems.Count
     If .ListItems.Item(lng).Checked Then
        'Saldo + Intereses Atrasados de los Fiadores Marcados
        txtTRAFD_MONTO = CCur(txtTRAFD_MONTO) + CCur(.ListItems.Item(lng).SubItems(4)) + CCur(.ListItems.Item(lng).SubItems(8))
     End If
   Next lng
 End With
 
 txtTRAFD_MONTO = Format(txtTRAFD_MONTO, "Standard")
 txtTRAFD_Int = txtInteresActual
 txtTRAFD_Plazo = fxPlazoRestante(txtOperacion)
  
End Sub

Private Sub lswOperacionesGeneradas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 Set Conlsw.lswX = lswOperacionesGeneradas
 Conlsw.Abre
End If
End Sub

Private Sub lswRepGen_ItemClick(ByVal Item As MSComctlLib.ListItem)

lblRepGen.Caption = Item.Text
lblRepGen.Tag = Item.Key

End Sub

Private Sub lswRepOp_ItemClick(ByVal Item As MSComctlLib.ListItem)

lblRepOp.Caption = Item.Text
lblRepOp.Tag = Item.Key

End Sub

Private Sub ssTabPrincipal_Click(PreviousTab As Integer)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Actualiza informacion de los Tabs
'REFERENCIAS   : Fiadores    - (Carga lsw de fiadores de la operacion)
'                OperacionesGenerada - (Carga lsw de Operaciones Generadas)
'                Telefonos   - (Carga los número telefonicos del deudor)
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------


Select Case ssTabPrincipal.Tab
 Case 1
  If vOperacion Then
    If vTabs.Direccion = 0 Then
        vTabs.Direccion = 1
        Call Telefonos(lswTelefonos, txtCedula)
    End If
  Else
   ssTabPrincipal.Tab = PreviousTab
  End If
 
 Case 2
  
  If Not vOperacion Then
   ssTabPrincipal.Tab = PreviousTab
  Else
   Call imgVisualiza_Click
  End If
  
 Case 3
  If vOperacion Then
    If vTabs.Fiadores = 0 Then
        vTabs.Fiadores = 1
        txtPorcentajeDeuda = 100
        Call Fiadores
    End If
  Else
   ssTabPrincipal.Tab = PreviousTab
  End If
 Case 6
  If vOperacion Then
    If vTabs.OPGeneradas = 0 Then
        vTabs.OPGeneradas = 1
        Call OperacionesGeneradas
    End If
  Else
   ssTabPrincipal.Tab = PreviousTab
  End If

 Case 7
   Call sbOpciones

End Select

End Sub

Private Sub sbOpciones()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select CO_TASA_SOCIO,CO_TASA_NSOCIO,CO_CODIGO from par_AhCr"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  chkParMantieneTasaSocios.Value = rs!co_Tasa_Socio
  chkParMantieneTasaNSocios.Value = rs!co_Tasa_NSocio
  txtParCodigo = rs!co_codigo
End If
rs.Close

End Sub

Private Sub sbAdjuntos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

i = MsgBox("Desea Imprimir el Estado de Cuenta de los Fiadores?", vbYesNo)

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro Administrativo y Judicial"
    
      .ReportFileName = App.Path & "\Credito\Reportes\DetalleMorosidadOperacion.rpt"
      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .Formulas(1) = "SubTitulo='CUOTAS ATRASADAS ACTIVAS'"
      .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
      .Formulas(3) = "Titulo='DETALLE DE CUOTAS MOROSAS DE LA OPERACION'"
      .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
      .SubreportToChange = "Morosidad"
      .SelectionFormula = "{MOROSIDAD.ID_SOLICITUD} = {?Pm-REG_CREDITOS.ID_SOLICITUD} AND {MOROSIDAD.ESTADO} ='A'"
    .PrintReport
End With

Me.MousePointer = vbDefault

'Llamar el Estado de Cuenta
Call sbEstadoCuenta(txtCedula)

'Estado de Cuentas de los Fiadores
If i = vbYes Then
    strSQL = "select cedulaf from fiadores where estado = 'A' and id_solicitud = " & txtOperacion
    rs.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rs.EOF
     Call sbEstadoCuenta(rs!cedulaf)
     rs.MoveNext
    Loop
    rs.Close
End If

Exit Sub

vError:

End Sub

Private Sub sbActualizaEstadoLaboralFiadores()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "update fiadores set interno = 1" _
       & " where fia_consec in(select F.fia_consec" _
       & " from fiadores F inner join Socios S on F.cedulaF = S.cedula" _
       & " inner join reg_creditos R on F.id_solicitud = R.id_solicitud and R.estado = 'A'" _
       & " where F.interno = 0 and S.estadoactual in('A','S'))"
glogon.Conection.Execute strSQL

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer

Select Case Button.Key
 Case "adjuntos"
    Call sbAdjuntos
 
 Case "refrescar"
    Call sbActualizaEstadoLaboralFiadores
 
 Case "reversar"
  iRespuesta = MsgBox("Esta seguro que desea Reversar (CBR JUD- TRASPASO) esta Operación", vbYesNo)
  If iRespuesta = vbYes Then
    Select Case lblProceso.Caption
     Case "TRASPASO DEUDAS"
      ssTabPrincipal.Tab = 6
      fraReversionDeTraspaso.Visible = True
      txtTRAFD_MONTO = 0
      txtTRAFD_Int = 0
      txtTRAFD_Plazo = 0
      txtTRAFD_Cuota = 0
     Case "COBRO JUDICIAL"
      Call ReversaCobroJudicial
     Case Else
      MsgBox "No hay Nada que Reversar verifique la información", vbInformation
    End Select
  End If
 Case "cerrar"
  Unload Me

End Select

End Sub


Private Sub sbCbrArchivoEstudio()
Dim strSQL As String, rs As New ADODB.Recordset
Dim fn, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

fn = FreeFile

Open "C:\ArchivoEstudio.txt" For Output As #fn

strSQL = "OPERACION,CODIGO,CEDULA,NOMBRE,GARANTIA,MONTO,SALDO,Mor.INTC,Mor.INTM,Mor.AMORTIZA" _
       & ",Mor.FINANCIERA,Mor.LEGAL,PRI-DED,Mor.CUOTAS,COMITE,Ult.Mov.,ESTADO,Fec.APORTE,AHORROS" _
       & ",APORTES,LIQUIDEZ,Mor.Prs.FINANCIERA,PLANILLA,PLAZO,INTERES,Mor.Prs.LEGAL,ESTADO_LABORAL"
Print #fn, strSQL
Print #fn, ""


strSQL = "select R.id_solicitud,R.codigo,R.Cedula,S.nombre,R.garantia,R.montoapr" _
       & ",R.saldo,V.intc,V.intm,V.amortiza,(V.intc+V.intm+V.amortiza) as Financiera" _
       & ",(V.intc+V.intm+R.saldo) as Legal,R.prideduc,V.cuota,C.descripcion as comite" _
       & ",R.fecult,S.estadoactual,A.fecAporte,A.ahorro+A.capitaliza as Ahorros" _
       & ",A.aporte,coalesce(P.porc_liquidez,0) as Liquidez,dbo.fxCBRMoraPersona(R.cedula,'F') as MoraPersona" _
       & ",R.ind_deduce_planilla,R.plazo,R.interesv,dbo.fxCBRMoraPersona(R.cedula,'L') as MoraPersonaLegal" _
       & ",coalesce(S.estadolaboral,0) as EstadoLaboral" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join vista_morosidad V on R.id_Solicitud = V.id_solicitud" _
       & " inner join ahorro_consolidado A on S.cedula = A.cedula" _
       & " inner join comites C on R.id_comite = C.id_comite" _
       & " inner join catalogo X on R.codigo = X.codigo and X.retencion = 'N'" _
       & " left join pra_principal P on R.id_solicitud = P.id_solicitud"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 strSQL = ""
 For i = 0 To rs.Fields.Count - 1
    strSQL = strSQL & rs.Fields(i).Value & ","
 Next i
 Print #fn, strSQL
 rs.MoveNext
Loop
rs.Close

Close #fn

Me.MousePointer = vbDefault
MsgBox "Se Creó el Archivo : C:\ArchivoEstudio.txt", vbInformation

Exit Sub

vError:
Me.MousePointer = vbDefault
MsgBox Err.Description, vbCritical

End Sub


Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vFecha As Long

If UCase(ButtonMenu.Key) = "ARCHIVO" Then
  Call sbCbrArchivoEstudio
  Exit Sub
End If

vFecha = GLOBALES.glngFechaCR
On Error Resume Next
vFecha = CLng(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Reportes de Mora"))

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro Administrativo y Judicial"

    Select Case UCase(ButtonMenu.Key)
      Case "REPINGRESOS"
        .ReportFileName = App.Path + "\Credito\Reportes\IngresosAMora.rpt"
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='FECHA PROCESO : " & Format(vFecha, "####-##") & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = "{MOROSIDAD.FECHAP}=" & vFecha
      Case "REPEGRESOS"
      
      Case "REPABONOS"
      
      Case "REPPLANILLA"
        .ReportFileName = App.Path + "\Credito\Reportes\CbrPlanillaComparativa.rpt"
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='FECHA PROCESO : " & Format(vFecha, "####-##") & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = "{vCbrPlanillaComparativa.proceso}=" & vFecha & " and ({vCbrPlanillaComparativa.Enviado} - {vCbrPlanillaComparativa.Recibido} > 10)"
            
    End Select
    .PrintReport
End With

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtNombre = fxNombre(txtCedula)
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtDescripcion = fxDescribeCodigo(txtCodigo)
End Sub

Private Sub txtCodigoNuevo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(6)
End Sub

Private Sub txtCodigoNuevo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  txtCodigoNuevo = UCase(txtCodigoNuevo)
  lblCodigoNuevo.Caption = fxDescribeCodigo(Trim(txtCodigoNuevo))
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(5)
End Sub

Private Sub txtOperacion_Change()
 Call CambiaDatos
End Sub

Sub CambiaDatos()
 vOperacion = False
 ssTabPrincipal.TabEnabled(0) = False
 ssTabPrincipal.TabEnabled(1) = False
 ssTabPrincipal.TabEnabled(2) = False
 ssTabPrincipal.TabEnabled(3) = False
 ssTabPrincipal.TabEnabled(6) = False
 ssTabPrincipal.TabEnabled(7) = False
 
 
 vTabs.Antiguedad = 0
 vTabs.Direccion = 0
 vTabs.Fiadores = 0
 vTabs.OPGeneradas = 0
'Tab Estado de Cuenta y General
 txtCodigo = ""
 txtNombre = ""
 txtDescripcion = ""
 txtCedula = ""
 lblEstado.Caption = ""
 lblEstadoMoroso.Caption = ""
 lblPrimerDeduccion.Caption = ""
 lblUltimoMovimiento.Caption = ""
 lblGarantia.Caption = ""
 lblDocumento.Caption = ""
 lblPagare.Caption = ""
 txtMonto = ""
 txtPlazo = ""
 txtSaldo = ""
 txtAmortizado = ""
 txtInteresPorcentaje = ""
 txtCuota = ""
 txtInteresPagado = ""
 txtCuotasPlanilla = ""
 txtCuotasDirectas = ""
 txtCuotasAnuladas = ""
 txtPlazoRecalculo = ""
 txtInteresActual = ""
 txtMontoRecalculo = ""
 txtInteresesMoratorios = ""
 txtAmortizacionAtrasada = ""
 txtTotalMora = ""
 lblProceso.Caption = ""
 lblOpex.Caption = ""
 lblFechaCBR.Caption = "__/__/____"
 lblObservacionesCBR.Caption = ""
 lswTelefonos.ListItems.Clear
'Tab Cuotas
 lswAbonos.ListItems.Clear
 lblCuotas.Caption = ""

'Tab Fiadores
 txtDirFiadores = ""
 lswFiadores.ListItems.Clear
 lswTelefonosFiadores.ListItems.Clear
 lblCedulaFiador.Caption = ""
 lblNombreFiador.Caption = ""
 txtCodigoNuevo = ""
 lblCodigoNuevo.Caption = ""
 txtPorcentajeDeuda = "100"
 txtMontoFiador = ""
 txtPlazoFiador = ""
 txtInteresFiador = ""
 txtCuotaFiador = ""
 chkTraspaso.Value = 1
 lswOperacionesGeneradas.ListItems.Clear
 ssTabPrincipal.Tab = 4
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call ConsultaOperacion
End Sub


Sub HistorialAvisos(lngOperacion As Long)
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListItem

strSQL = "select * from cobro_avisos where id_solicitud = " _
       & lngOperacion & " order by fecha_aviso"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Set itmX = lswAvisos.ListItems.Add(, , Format(rs!fecha_aviso, "dd/mm/yyyy"))
      Select Case rs!tipo_aviso
        Case 1
            itmX.SubItems(1) = "Primer Aviso"
        Case 2
            itmX.SubItems(1) = "Segundo Aviso"
        Case Else
            itmX.SubItems(1) = "Otro Aviso"
      End Select
  rs.MoveNext
Loop
rs.Close

End Sub

Sub ConsultaOperacion()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Actualizar la informacion de la ventana segun la operacion seleccionada
'REFERENCIAS   : MorosidadActiva - (Carga Datos de Mora Activa de la Operacion)
'                fxDescribeCodigo - (Devuelve la descripcion de el código del crédito)
'                DatosBoleta - (Carga los datos personales)
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, strSQL As String, blnContinua As Boolean

On Error GoTo Fin
Me.MousePointer = vbHourglass

blnContinua = True

With rs
 .Open "select * from reg_creditos where id_solicitud = " & txtOperacion, glogon.Conection, adOpenStatic
 
 If .EOF And .BOF Then
  blnContinua = False
  vOperacion = False
  ssTabPrincipal.TabEnabled(0) = False
  ssTabPrincipal.TabEnabled(1) = False
  ssTabPrincipal.TabEnabled(2) = False
  ssTabPrincipal.TabEnabled(3) = False
  ssTabPrincipal.TabEnabled(6) = False
  ssTabPrincipal.TabEnabled(7) = False
  ssTabPrincipal.Tab = 4
  MsgBox "No se encontró número de solicitud...", vbInformation
 
 Else
    
    vOperacion = True
    ssTabPrincipal.TabEnabled(0) = True
    ssTabPrincipal.TabEnabled(1) = True
    ssTabPrincipal.TabEnabled(2) = True
    ssTabPrincipal.TabEnabled(3) = True
    ssTabPrincipal.TabEnabled(6) = True
    ssTabPrincipal.TabEnabled(7) = True
    
    ssTabPrincipal.Tab = 0
    txtCodigo = !Codigo
    txtDescripcion = fxDescribeCodigo(!Codigo)
    Call DatosBoleta(!Cedula)
    txtCedula = !Cedula
    lblEstado.Caption = fxDescribeEstado(IIf(IsNull(!Estado), "N", !Estado))
    
    Select Case UCase(IIf(IsNull(!proceso), "N", !proceso))
     Case "N"
      lblProceso = "NORMAL"
     Case "T"
      lblProceso = "TRASPASO DEUDAS"
     Case "J"
      lblProceso = "COBRO JUDICIAL"
     Case Else
      lblProceso = "NORMAL"
    End Select
    
    If IIf(IsNull(!opex), 0, !opex) = 1 Then
       lblOpex.Caption = "SI"
    Else
       lblOpex.Caption = "NO"
    End If
    
    If Not IsNull(!ind_deduce_planilla) Then
        chkDeducirPlanilla.Value = IIf((!ind_deduce_planilla = "S"), vbChecked, vbUnchecked)
    End If
    
    lblFechaCBR.Caption = IIf(IsNull(!fecha_enviaproceso), ("__/__/____"), Format(!fecha_enviaproceso, "dd/mm/yyyy"))
    lblObservacionesCBR.Caption = IIf(IsNull(!observacion_proceso), "NADA", !observacion_proceso)
 
    Call MorosidadActiva
    
    lblPrimerDeduccion.Caption = Format(IIf(IsNull(!prideduc), "", !prideduc), "####-##")
    lblUltimoMovimiento.Caption = Format(IIf(IsNull(!fecult), "", !fecult), "####-##")
    lblGarantia.Caption = fxGarantia(!garantia)
    lblDocumento.Caption = IIf(IsNull(!Tdocumento), "", !Tdocumento) & "-" & IIf(IsNull(!ndocumento), "", !ndocumento)
    lblPagare.Caption = "NA"
    txtMonto = Format(IIf(IsNull(!montoapr), "", !montoapr), "###,###,###,##0.00")
    txtPlazo = IIf(IsNull(!Plazo), "", !Plazo)
    txtSaldo = Format(IIf(IsNull(!Saldo), "", !Saldo), "###,###,###,##0.00")
    txtAmortizado = Format(IIf(IsNull(!Amortiza), "", !Amortiza), "###,###,###,##0.00")
    txtInteresPorcentaje = IIf(IsNull(!Int), "", !Int)
    txtCuota = Format(IIf(IsNull(!cuota), "", !cuota), "###,###,###,##0.00")
    txtInteresPagado = Format(IIf(IsNull(!interesc), "", !interesc), "###,###,###,##0.00")
    txtCuotasPlanilla = IIf(IsNull(!cuotas_planilla), "", !cuotas_planilla)
    txtCuotasDirectas = IIf(IsNull(!cuotas_directas), "", !cuotas_directas)
    txtCuotasAnuladas = IIf(IsNull(!CUOTAS_ANULADAS), "", !CUOTAS_ANULADAS)
    txtPlazoRecalculo = IIf(IsNull(!plazo_recalculo), "", !plazo_recalculo)
    txtInteresActual = IIf(IsNull(!interesv), "", !interesv)
    txtMontoRecalculo = Format(IIf(IsNull(!monto_recalculo), "", !monto_recalculo), "###,###,###,##0.00")
 
    Call HistorialAvisos(!ID_SOLICITUD)
 
 End If
 .Close
End With
Fin:
Me.MousePointer = vbDefault

End Sub


Sub MorosidadActiva()
Dim rsMorosidad As New ADODB.Recordset, strSQL As String

With rsMorosidad
 strSQL = "select coalesce(sum(intc),0) as intc,coalesce(sum(intm),0) as intm," _
        & "coalesce(sum(amortiza),0) as amortiza from morosidad" _
        & " where estado = 'A' and id_solicitud = " & txtOperacion
 .Open strSQL, glogon.Conection, adOpenStatic
    mCurIntc = IIf(IsNull(!intc), 0, !intc)
    mCurIntm = IIf(IsNull(!intm), 0, !intm)
    txtInteresesMoratorios = Format(IIf(IsNull(!intc), 0, !intc) + IIf(IsNull(!intm), 0, !intm), "###,###,###,##0.00")
    txtAmortizacionAtrasada = Format(IIf(IsNull(!Amortiza), "0", !Amortiza), "###,###,###,##0.00")
    txtTotalMora = Format((CCur(txtInteresesMoratorios) + CCur(txtAmortizacionAtrasada)), "###,###,###,##0.00")
 lblEstadoMoroso = IIf((CCur(txtTotalMora) = 0), "AL DIA", "MOROSO")
 .Close
End With
End Sub

Sub DatosBoleta(strCedula As String)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Carga los datos Personales
'REFERENCIAS   : fxProvincia - (Devuelve el número o descripcion de las provincias)
'                fxCanton   - (Devuelve la descripcion del canton)
'                fxUnidad   - (Devuelve la descripcion de la unidad Prog. o de Trabajo)
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim rsDatosBoleta As New ADODB.Recordset
With rsDatosBoleta
.Open "select * from socios where cedula = '" & Trim(strCedula) & "'", glogon.Conection, adOpenStatic
If .EOF And .BOF Then
 MsgBox "No se encontró registro de la persona con cédula = " & strCedula, vbCritical
Else
 txtNombre = IIf(IsNull(!Nombre), "", !Nombre)
 txtProvincia = fxProvincia(IIf(IsNull(!provincia), 11, !provincia))
 txtCanton = fxCanton(IIf(IsNull(!provincia), 0, !provincia), IIf(IsNull(!canton), 0, !canton))
 txtDistrito = IIf(IsNull(!Distrito), "", !Distrito)
 txtDireccion = IIf(IsNull(!Direccion), "", !Direccion)
 txtEMail = IIf(IsNull(!Af_Email), "", !Af_Email)
 txtApartado = IIf(IsNull(!Apto), "", !Apto)
End If
.Close
End With
End Sub

Sub LimpiaDatos()
Dim strSQL As String, rs As New ADODB.Recordset

 'tabs inactivos
 ssTabPrincipal.TabEnabled(0) = False
 ssTabPrincipal.TabEnabled(1) = False
 ssTabPrincipal.TabEnabled(2) = False
 ssTabPrincipal.TabEnabled(3) = False
 ssTabPrincipal.TabEnabled(6) = False
 ssTabPrincipal.TabEnabled(7) = False

mCurIntc = 0
mCurIntm = 0

 lswAvisos.ListItems.Clear

 'fin
 ssTabPrincipal.Tab = 0
 vTabs.Antiguedad = 0
 vTabs.Direccion = 0
 vTabs.Fiadores = 0
 vTabs.OPGeneradas = 0
'Tab Estado de Cuenta y General
 txtOperacion = ""
 txtCodigo = ""
 txtNombre = ""
 txtDescripcion = ""
 txtCedula = ""
 lblEstado.Caption = ""
 lblEstadoMoroso.Caption = ""
 lblPrimerDeduccion.Caption = ""
 lblUltimoMovimiento.Caption = ""
 lblGarantia.Caption = ""
 lblDocumento.Caption = ""
 lblPagare.Caption = ""
 txtMonto = ""
 txtPlazo = ""
 txtSaldo = ""
 txtAmortizado = ""
 txtInteresPorcentaje = ""
 txtCuota = ""
 txtInteresPagado = ""
 txtCuotasPlanilla = ""
 txtCuotasDirectas = ""
 txtCuotasAnuladas = ""
 txtPlazoRecalculo = ""
 txtInteresActual = ""
 txtMontoRecalculo = ""
 txtInteresesMoratorios = ""
 txtAmortizacionAtrasada = ""
 txtTotalMora = ""
 lblProceso.Caption = ""
 lblOpex.Caption = ""
 lblFechaCBR.Caption = "__/__/____"
 lblObservacionesCBR.Caption = ""
 
 'Tab Direccion
 
 txtProvincia = ""
 txtCanton = ""
 txtDistrito = ""
 txtDireccion = ""
 txtEMail = ""
 txtApartado = ""
 lswTelefonos.ListItems.Clear
'Tab Cuotas
 lswAbonos.ListItems.Clear
 lblCuotas.Caption = ""
'Tab Fiadores
 txtDirFiadores = ""
 lswFiadores.ListItems.Clear
 lswTelefonosFiadores.ListItems.Clear
 lblCedulaFiador.Caption = ""
 lblNombreFiador.Caption = ""
 txtCodigoNuevo = ""
 lblCodigoNuevo.Caption = ""
 txtPorcentajeDeuda = 100
 txtMontoFiador = ""
 txtPlazoFiador = ""
 txtInteresFiador = ""
 txtCuotaFiador = ""
 chkTraspaso.Value = 1
 
'Tab Reportes
 cboRep.Clear
 cboRep.AddItem "00 - Todos"
 cboRep.AddItem "01 - Socios"
 cboRep.AddItem "02 - Ex.Socios"
 cboRep.AddItem "03 - No Socios"
 cboRep.AddItem "04 - Ren.Interna"
 cboRep.AddItem "05 - Ren.Patronal"
 cboRep.Text = "00 - Todos"
 
 cboRepX.Clear
 cboRepX.AddItem "00 - Todos"
 cboRepX.AddItem "01 - Socios"
 cboRepX.AddItem "02 - Ex.Socios"
 cboRepX.AddItem "03 - No Socios"
 cboRepX.AddItem "04 - Ren.Interna"
 cboRepX.AddItem "05 - Ren.Patronal"
 cboRepX.Text = "00 - Todos"
 
 
 cboRepAnt.Clear
 cboRepAnt.AddItem "00 - Todos"
 cboRepAnt.AddItem "01 - Socios"
 cboRepAnt.AddItem "02 - Ex.Socios"
 cboRepAnt.AddItem "03 - No Socios"
 cboRepAnt.AddItem "04 - Ren.Interna"
 cboRepAnt.AddItem "05 - Ren.Patronal"
 cboRepAnt.Text = "00 - Todos"
 
 
 'Tab Reportes
 
 With lswRepOp.ListItems
   .Clear
   .Add , "ULTEC", "Ultimo Estado"
   .Add , "ECBR", "Estado de Cuenta Cobro"
   .Add , "ETSBR", "Equitetado de Sobres"
   .Add , "PRIAVI", "Carta - Primer Aviso"
   .Add , "SEGAVI", "Carta - Segundo Aviso"
   .Add , "NOTMOV", "Notificación de Movimiento Realizado"
   .Add , "REVER", "Estado de Reversión"
   .Add , "ENVCBR", "Casos en Cobro Judicial"
   .Add , "TRADEUD", "Traspaso de Deudas"
   .Add , "TRAFIA", "Traspasos Fiadores con Operaciones"
   .Add , "TRAREV", "Boleta de Reversion de Traspaso de Deudas"
 End With
 
 With lswRepGen.ListItems
   .Clear
   .Add , "GENDET", "General - Detallado"
   .Add , "GENRSM", "General - Resumen"
   
   .Add , "GENAGD", "General - Detallado Agrupado"
   .Add , "GENRGD", "General - Resumen Agrupado"
   
   .Add , "ESPCON", "Especial Convenios"
   .Add , "MORCAR", "Comparativo - Resumen"
   .Add , "MORGAR", "Mora x Garantía"
   .Add , "MORGAG", "Mora x Garantía - Agrupado"
 End With
 
 cboDestino.Clear
 cboDestino.AddItem "TODOS"
 cboDestino.Text = "TODOS"
 
 
 cboGarantia.Clear
 cboGarantia.AddItem "A - Sobre Ahorros"
 cboGarantia.AddItem "F - Fiduciaria"
 cboGarantia.AddItem "H - Hipotecaria"
 cboGarantia.AddItem "X - Acciones"
 cboGarantia.AddItem "Y - Fondos de Inversion"
 cboGarantia.AddItem "N - Sin Garantía"
 cboGarantia.AddItem "TODOS"
 cboGarantia.Text = "TODOS"
 
 cboCartera.Clear
 cboCartera.AddItem "(Todas las Carteras)"
 strSQL = "select rtrim(cod_clasificacion) + ' - ' + descripcion as ItemX" _
        & " from CBR_CLASIFICACION_CARTERA order by cod_clasificacion"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
  cboCartera.AddItem rs!itemx
  rs.MoveNext
 Loop
 cboCartera.Text = "(Todas las Carteras)"
 rs.Close
 
'Tab Antiguedad
lswAntiguedad.ListItems.Clear

'Tab OP Generadas

lswOperacionesGeneradas.ListItems.Clear
ssTabPrincipal.Tab = 4

End Sub

Sub Telefonos(lsw As ListView, strCedula As String)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Carga lsw con los datos de los número de teléfonos de la persona
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, itmX As ListItem
Me.MousePointer = vbHourglass
rs.Open "select * from telefonos where cedula = '" & Trim(strCedula) & "'", glogon.Conection, adOpenForwardOnly

If rs.EOF And rs.BOF Then
 'nada
Else
 Do While rs.EOF = False
  Set itmX = lsw.ListItems.Add(lsw.ListItems.Count + 1, , fxTipoTelefono(rs!Tipo), , 1)
   itmX.SubItems(1) = IIf(IsNull(rs!Numero), "", rs!Numero)
   itmX.SubItems(2) = IIf(IsNull(rs!Ext), "", rs!Ext)
   itmX.SubItems(3) = IIf(IsNull(rs!contacto), "", rs!contacto)
  rs.MoveNext
 Loop
End If

rs.Close
Me.MousePointer = vbDefault
End Sub


Sub Fiadores()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Carga lsw con los datos de los número de teléfonos de la persona
'REFERENCIAS   : fxEmpleadoPatrono - (Devuelve 1 = si es empleado y 0 si no)
'                fxNombre - (Devuelve el nombre de la persona)
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem

Me.MousePointer = vbHourglass

strSQL = "select F.cedulaf,F.interno,S.nombre" _
       & " from fiadores F inner join Socios S on F.cedulaf = S.cedula" _
       & " where F.id_solicitud = " & Trim(txtOperacion) & " and F.estado = 'A'"

rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 
Do While Not rs.EOF
   Set itmX = lswFiadores.ListItems.Add(lswFiadores.ListItems.Count + 1, , (rs!cedulaf), , 5)
    itmX.SubItems(1) = rs!Nombre & ""
    itmX.SubItems(2) = IIf((rs!interno = 0), "NO", "SI") ' IIf(fxEmpleadoPatrono(rs!cedulaf) = 0, "NO", "SI")
    itmX.Tag = lswFiadores.ListItems.Count
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub txtPorcentajeDeuda_Change()
Dim curDeuda As Currency

If Len(Trim(txtInteresesMoratorios)) = 0 Then
    txtInteresesMoratorios = 0
End If

If Len(Trim(txtSaldo)) = 0 Then
    txtSaldo = 0
End If

curDeuda = CCur(txtInteresesMoratorios) + CCur(txtSaldo)

If Len(Trim(txtPorcentajeDeuda)) = 0 Then
  txtPorcentajeDeuda = 0
End If

txtMontoFiador = Format((curDeuda * (CCur(txtPorcentajeDeuda) / 100)), "###,###,###,##0.00")

End Sub

Private Sub txtInteresFiador_Change()
On Error Resume Next
If CCur(IIf((txtInteresFiador = ""), 0, txtInteresFiador)) > 0 And CCur(IIf((txtPlazoFiador = ""), 0, txtPlazoFiador)) > 0 _
    And CCur(IIf((txtMontoFiador = ""), 0, txtMontoFiador)) > 0 Then
 txtCuotaFiador = fxCalcula_Cuota(CCur(txtMontoFiador), CCur(txtPlazoFiador), CCur(txtInteresFiador))
End If
End Sub

Private Sub txtmontofiador_Change()
Dim x As Integer
If CCur(IIf((txtInteresFiador = ""), 0, txtInteresFiador)) > 0 And CCur(IIf((txtPlazoFiador = ""), 0, txtPlazoFiador)) > 0 _
    And CCur(IIf((txtMontoFiador = ""), 0, txtMontoFiador)) > 0 Then
 txtCuotaFiador = fxCalcula_Cuota(CCur(txtMontoFiador), CCur(txtPlazoFiador), CCur(txtInteresFiador))
End If
End Sub

Private Sub txtPlazoFiador_Change()
On Error Resume Next
If CCur(IIf((txtInteresFiador = ""), 0, txtInteresFiador)) > 0 And CCur(IIf((txtPlazoFiador = ""), 0, txtPlazoFiador)) > 0 _
    And CCur(IIf((txtMontoFiador = ""), 0, txtMontoFiador)) > 0 Then
 txtCuotaFiador = fxCalcula_Cuota(CCur(txtMontoFiador), CCur(txtPlazoFiador), CCur(txtInteresFiador))
End If
End Sub


Private Sub txtReporteX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(7)
If KeyCode = vbKeyReturn Then cboDestino.SetFocus
End Sub

Private Sub txtReporteX_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  lblXDescribe.Caption = fxDescribeCodigo(Trim(txtReporteX))
End If
End Sub

Private Sub sbLlenaCbo(cboX As ComboBox, strSQL As String, Optional vTodos As Boolean = True)
Dim rs As New ADODB.Recordset

cboX.Clear

rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 cboX.AddItem rs!itmX
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboX.Text = rs!itmX
End If
rs.Close

If vTodos Then
    cboX.AddItem "TODOS"
    cboX.Text = "TODOS"
End If

End Sub


Private Sub txtReporteX_LostFocus()
 If Len(Trim(txtReporteX)) > 0 Then lblXDescribe.Caption = fxDescribeCodigo(Trim(txtReporteX))
 Call chkLineas_Click
End Sub


Private Sub txtTRAFD_Int_Change()

If CCur(IIf((txtTRAFD_Int = ""), 0, txtTRAFD_Int)) > 0 And CCur(IIf((txtTRAFD_Plazo = ""), 0, txtTRAFD_Plazo)) > 0 _
    And CCur(IIf((txtTRAFD_MONTO = ""), 0, txtTRAFD_MONTO)) > 0 Then
 txtTRAFD_Cuota = fxCalcula_Cuota(CCur(txtTRAFD_MONTO), CCur(txtTRAFD_Plazo), CCur(txtTRAFD_Int))
 txtTRAFD_Cuota = Format(txtTRAFD_Cuota, "###,###,###,##0.00")
End If
End Sub

Private Sub txtTRAFD_Monto_Change()
Dim x As Integer
If CCur(IIf((txtTRAFD_Int = ""), 0, txtTRAFD_Int)) > 0 And CCur(IIf((txtTRAFD_Plazo = ""), 0, txtTRAFD_Plazo)) > 0 _
    And CCur(IIf((txtTRAFD_MONTO = ""), 0, txtTRAFD_MONTO)) > 0 Then
 txtTRAFD_Cuota = fxCalcula_Cuota(CCur(txtTRAFD_MONTO), CCur(txtTRAFD_Plazo), CCur(txtTRAFD_Int))
 txtTRAFD_Cuota = Format(txtTRAFD_Cuota, "###,###,###,##0.00")
End If
End Sub

Private Sub txtTRAFD_Plazo_Change()
If CCur(IIf((txtTRAFD_Int = ""), 0, txtTRAFD_Int)) > 0 And CCur(IIf((txtTRAFD_Plazo = ""), 0, txtTRAFD_Plazo)) > 0 _
    And CCur(IIf((txtTRAFD_MONTO = ""), 0, txtTRAFD_MONTO)) > 0 Then
 txtTRAFD_Cuota = fxCalcula_Cuota(CCur(txtTRAFD_MONTO), CCur(txtTRAFD_Plazo), CCur(txtTRAFD_Int))
 txtTRAFD_Cuota = Format(txtTRAFD_Cuota, "###,###,###,##0.00")
End If
End Sub


