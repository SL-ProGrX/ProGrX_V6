VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_APA_Operaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Pagarés - Operaciones"
   ClientHeight    =   8220
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   12690
   Icon            =   "frmCR_APA_Operaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   240
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":13916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":1A178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":209DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":2723C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":2DA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":34300
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":3AB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":413C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":47C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":4E488
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":4E4E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":500E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":5694B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":5D1AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":63A0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Operaciones.frx":6A271
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Acreedores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstDatos 
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
         Height          =   870
         Left            =   120
         TabIndex        =   58
         Top             =   7080
         Width           =   2895
      End
      Begin MSComctlLib.ListView lswAcreedores 
         Height          =   2625
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4630
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lswContactos 
         Height          =   2385
         Left            =   120
         TabIndex        =   57
         Top             =   3960
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4207
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   0
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   135
         Top             =   6600
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Localiza"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   134
         Top             =   3480
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Contactos"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   133
         Top             =   120
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Acreedores"
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
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   360
         Left            =   120
         Picture         =   "frmCR_APA_Operaciones.frx":6A39B
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   120
         Picture         =   "frmCR_APA_Operaciones.frx":6DBCF
         Top             =   6600
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   120
         Picture         =   "frmCR_APA_Operaciones.frx":6DD08
         Top             =   3480
         Width           =   360
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7335
      Left            =   3360
      TabIndex        =   25
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Operaciones"
      TabPicture(0)   =   "frmCR_APA_Operaciones.frx":6DEA6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(16)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tlbBusquedaOperaciones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tlbListaOperaciones"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vGridOperaciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtOperacionBusqueda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboEstadoOpeBusq"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkFiltroOperaciones"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "+ Operaciones"
      TabPicture(1)   =   "frmCR_APA_Operaciones.frx":6DEC2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTipoCambio"
      Tab(1).Control(1)=   "cboMoneda"
      Tab(1).Control(2)=   "cboTipoPago"
      Tab(1).Control(3)=   "cboOficina"
      Tab(1).Control(4)=   "txtResponsabilidad"
      Tab(1).Control(5)=   "txtComisionAdmin"
      Tab(1).Control(6)=   "txtCuota"
      Tab(1).Control(7)=   "txtCod_Acreedor"
      Tab(1).Control(8)=   "txtDescripcionAcreedor"
      Tab(1).Control(9)=   "txtOperacion"
      Tab(1).Control(10)=   "txtMonto"
      Tab(1).Control(11)=   "txtTasa"
      Tab(1).Control(12)=   "txtPlazo"
      Tab(1).Control(13)=   "txtSaldo"
      Tab(1).Control(14)=   "cboTipo"
      Tab(1).Control(15)=   "txtDia_Pago"
      Tab(1).Control(16)=   "txtNotas"
      Tab(1).Control(17)=   "tlbIncluirOperaciones"
      Tab(1).Control(18)=   "fraOptComision"
      Tab(1).Control(19)=   "fraOptResponsabilidad"
      Tab(1).Control(20)=   "dtpPrimerPago"
      Tab(1).Control(21)=   "dtpFormalizacion"
      Tab(1).Control(22)=   "Label2(54)"
      Tab(1).Control(23)=   "Label2(53)"
      Tab(1).Control(24)=   "Label2(52)"
      Tab(1).Control(25)=   "Label2(51)"
      Tab(1).Control(26)=   "Label2(44)"
      Tab(1).Control(27)=   "Label2(7)"
      Tab(1).Control(28)=   "Line14"
      Tab(1).Control(29)=   "lblFec_Actualizacion"
      Tab(1).Control(30)=   "lblEstado"
      Tab(1).Control(31)=   "Label2(6)"
      Tab(1).Control(32)=   "Label2(40)"
      Tab(1).Control(33)=   "Label2(43)"
      Tab(1).Control(34)=   "Label2(41)"
      Tab(1).Control(35)=   "Line2"
      Tab(1).Control(36)=   "Label2(0)"
      Tab(1).Control(37)=   "Label2(1)"
      Tab(1).Control(38)=   "Label2(2)"
      Tab(1).Control(39)=   "Label2(3)"
      Tab(1).Control(40)=   "Label2(4)"
      Tab(1).Control(41)=   "Label2(5)"
      Tab(1).Control(42)=   "Label2(8)"
      Tab(1).Control(43)=   "Label2(9)"
      Tab(1).Control(44)=   "Label2(11)"
      Tab(1).Control(45)=   "Label2(12)"
      Tab(1).Control(46)=   "Label2(13)"
      Tab(1).Control(47)=   "Label2(14)"
      Tab(1).Control(48)=   "Label2(10)"
      Tab(1).ControlCount=   49
      TabCaption(2)   =   "Pagos"
      TabPicture(2)   =   "frmCR_APA_Operaciones.frx":6DEDE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraAutorizados"
      Tab(2).Control(1)=   "chkFiltroPagos"
      Tab(2).Control(2)=   "cboEstadoPagoBusq"
      Tab(2).Control(3)=   "vGridPagos"
      Tab(2).Control(4)=   "tlbListaPagos"
      Tab(2).Control(5)=   "tlbBusquedaPagos"
      Tab(2).Control(6)=   "tlbPagosOtros"
      Tab(2).Control(7)=   "dtpFecPagosDesde"
      Tab(2).Control(8)=   "dtpFecPagosHasta"
      Tab(2).Control(9)=   "Label2(21)"
      Tab(2).Control(10)=   "Label2(18)"
      Tab(2).Control(11)=   "Label2(20)"
      Tab(2).Control(12)=   "Line7"
      Tab(2).Control(13)=   "Line6"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "+ Pagos"
      TabPicture(3)   =   "frmCR_APA_Operaciones.frx":6DEFA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraPago"
      Tab(3).Control(1)=   "tlbIncluirPagos"
      Tab(3).Control(2)=   "lblEstadoPago"
      Tab(3).Control(3)=   "lblFechaPago"
      Tab(3).Control(4)=   "lblUsuarioPago"
      Tab(3).Control(5)=   "lblTesoreriaFecha"
      Tab(3).Control(6)=   "lblTesoreriaUsuario"
      Tab(3).Control(7)=   "lblTesoreriaSolicitud"
      Tab(3).Control(8)=   "lblUltimoSaldo"
      Tab(3).Control(9)=   "lblUltimaTasa"
      Tab(3).Control(10)=   "lblUltimaCuota"
      Tab(3).Control(11)=   "Label3"
      Tab(3).Control(12)=   "Label2(35)"
      Tab(3).Control(13)=   "Label2(31)"
      Tab(3).Control(14)=   "Label2(27)"
      Tab(3).Control(15)=   "Line10"
      Tab(3).Control(16)=   "Line9"
      Tab(3).Control(17)=   "Label2(45)"
      Tab(3).Control(18)=   "Label2(46)"
      Tab(3).Control(19)=   "Label2(49)"
      Tab(3).Control(20)=   "Label2(39)"
      Tab(3).Control(21)=   "Label2(37)"
      Tab(3).Control(22)=   "Label2(36)"
      Tab(3).Control(23)=   "Label2(34)"
      Tab(3).Control(24)=   "Label2(33)"
      Tab(3).Control(25)=   "Label2(32)"
      Tab(3).Control(26)=   "Line4"
      Tab(3).Control(27)=   "Line8"
      Tab(3).ControlCount=   28
      Begin VB.TextBox txtTipoCambio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68280
         TabIndex        =   122
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox cboMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox cboTipoPago 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   3960
         Width           =   2295
      End
      Begin MSComctlLib.ImageCombo cboOficina 
         Height          =   345
         Left            =   -73080
         TabIndex        =   118
         Top             =   4320
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraAutorizados 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   107
         Top             =   7560
         Visible         =   0   'False
         Width           =   7575
         Begin VB.TextBox txtCedulaAutorizado 
            Height          =   375
            Left            =   1920
            TabIndex        =   110
            Top             =   1440
            Width           =   1575
         End
         Begin MSComctlLib.Toolbar tlbAutorizados 
            Height          =   990
            Left            =   240
            TabIndex        =   108
            Top             =   1080
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   1746
            ButtonWidth     =   1693
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgMenu"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Aplicar"
                  Key             =   "Aplicar"
                  ImageIndex      =   16
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Salir"
                  Key             =   "Salir"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00FFFFFF&
            X1              =   360
            X2              =   7080
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line12 
            BorderColor     =   &H80000004&
            X1              =   0
            X2              =   7440
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Autorizado"
            Height          =   255
            Left            =   3120
            TabIndex        =   114
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   3480
            TabIndex        =   113
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblNombreAutorizado 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3480
            TabIndex        =   112
            Top             =   1440
            Width           =   3615
         End
         Begin VB.Label Label4 
            Caption         =   "Cédula"
            Height          =   255
            Left            =   1920
            TabIndex        =   111
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.TextBox txtResponsabilidad 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68280
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtComisionAdmin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Frame fraPago 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   -74640
         TabIndex        =   60
         Top             =   840
         Width           =   7095
         Begin VB.TextBox txtComisionPago 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   5
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtDocumentoPago 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   38
            TabIndex        =   0
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtInteresesPago 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   3
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtCargosPago 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   4
            Top             =   1920
            Width           =   1575
         End
         Begin VB.ComboBox cboFormaPago 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtMontoPago 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   1
            Top             =   840
            Width           =   1575
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaPago 
            Height          =   330
            Left            =   1680
            TabIndex        =   129
            Top             =   1200
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin XtremeSuiteControls.DateTimePicker dtpVencePago 
            Height          =   330
            Left            =   5160
            TabIndex        =   130
            Top             =   1200
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin VB.Label lblSaldoPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   90
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblTasaPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   89
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblAmortizacionPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   87
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblNPago 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   86
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Fec Movimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   48
            Left            =   240
            TabIndex        =   73
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Comisiones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   47
            Left            =   240
            TabIndex        =   72
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   240
            TabIndex        =   71
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Intereses"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   240
            TabIndex        =   70
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Otros Cargos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   240
            TabIndex        =   69
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   3840
            TabIndex        =   68
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Tasa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   3840
            TabIndex        =   67
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Amortización"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   3840
            TabIndex        =   66
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Forma Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   3840
            TabIndex        =   64
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Número"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   3840
            TabIndex        =   63
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Cuota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   62
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Fec Vence"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   3840
            TabIndex        =   61
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.TextBox txtCuota 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68280
         TabIndex        =   14
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CheckBox chkFiltroOperaciones 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   55
         Top             =   6360
         Width           =   1695
      End
      Begin VB.CheckBox chkFiltroPagos 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -74160
         TabIndex        =   54
         Top             =   5880
         Width           =   255
      End
      Begin VB.ComboBox cboEstadoPagoBusq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69960
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   6240
         Width           =   1575
      End
      Begin VB.ComboBox cboEstadoOpeBusq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   6720
         Width           =   1575
      End
      Begin VB.TextBox txtOperacionBusqueda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   44
         Top             =   6720
         Width           =   2295
      End
      Begin VB.TextBox txtCod_Acreedor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcionAcreedor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtOperacion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtMonto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtTasa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68280
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtPlazo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtSaldo 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68280
         TabIndex        =   23
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtDia_Pago 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68280
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtNotas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   -73080
         MaxLength       =   495
         TabIndex        =   17
         Top             =   4680
         Width           =   6375
      End
      Begin FPSpreadADO.fpSpread vGridOperaciones 
         Height          =   5175
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   8655
         _Version        =   524288
         _ExtentX        =   15266
         _ExtentY        =   9128
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   491
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_Operaciones.frx":6DF16
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbListaOperaciones 
         Height          =   312
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   5268
         _ExtentX        =   9287
         _ExtentY        =   556
         ButtonWidth     =   1640
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Agregar Acreedor"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar Acreedor"
               ImageIndex      =   2
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Editar"
               Key             =   "Editar"
               Object.ToolTipText     =   "Editar Acreedor"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ver"
               Key             =   "Ver"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Cerrar"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbIncluirOperaciones 
         Height          =   330
         Left            =   -74760
         TabIndex        =   22
         Top             =   480
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         ButtonWidth     =   1826
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Key             =   "Guardar"
               Object.ToolTipText     =   "Guardar los cambios"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridPagos 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   29
         Top             =   1080
         Width           =   7815
         _Version        =   524288
         _ExtentX        =   13785
         _ExtentY        =   8281
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   491
         MaxRows         =   501
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_Operaciones.frx":6E613
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbListaPagos 
         Height          =   330
         Left            =   -74760
         TabIndex        =   30
         Top             =   480
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         ButtonWidth     =   1640
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Agregar Acreedor"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar Acreedor"
               ImageIndex      =   2
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ver"
               Key             =   "Ver"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBusquedaOperaciones 
         Height          =   315
         Left            =   6720
         TabIndex        =   48
         Top             =   6720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         ButtonWidth     =   1640
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBusquedaPagos 
         Height          =   315
         Left            =   -68280
         TabIndex        =   52
         Top             =   6240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         ButtonWidth     =   1640
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbIncluirPagos 
         Height          =   330
         Left            =   -74760
         TabIndex        =   6
         Top             =   360
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   556
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Key             =   "Guardar"
               Object.ToolTipText     =   "Guardar los cambios"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprime Boleta del Traslado"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "CancelarOperacion"
               Object.ToolTipText     =   "Contactos del Acreedor"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbPagosOtros 
         Height          =   312
         Left            =   -69120
         TabIndex        =   109
         Top             =   480
         Width           =   1788
         _ExtentX        =   3149
         _ExtentY        =   556
         ButtonWidth     =   2858
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Girar a Nombre"
               Key             =   "NGiro"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraOptComision 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -73320
         TabIndex        =   102
         Top             =   5760
         Width           =   2535
         Begin VB.OptionButton OptComisionOperacion 
            Caption         =   "Operación"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptComisionGarantias 
            Caption         =   "Garantías"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   19
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame fraOptResponsabilidad 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   375
         Left            =   -70800
         TabIndex        =   103
         Top             =   5640
         Width           =   4695
         Begin VB.OptionButton OptResponOperacion 
            Caption         =   "Operación"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptResponGarantias 
            Caption         =   "Garantías"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   21
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Respon X Saldos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   50
            Left            =   600
            TabIndex        =   104
            Top             =   120
            Width           =   1695
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpPrimerPago 
         Height          =   330
         Left            =   -73080
         TabIndex        =   125
         Top             =   3600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.DateTimePicker dtpFormalizacion 
         Height          =   330
         Left            =   -68280
         TabIndex        =   126
         Top             =   3600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.DateTimePicker dtpFecPagosDesde 
         Height          =   330
         Left            =   -73560
         TabIndex        =   127
         Top             =   6240
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
      Begin XtremeSuiteControls.DateTimePicker dtpFecPagosHasta 
         Height          =   330
         Left            =   -72120
         TabIndex        =   128
         Top             =   6240
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
      Begin VB.Label Label2 
         Caption         =   "Tipo Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   -69720
         TabIndex        =   124
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   -74400
         TabIndex        =   123
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   -74400
         TabIndex        =   119
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   -74400
         TabIndex        =   117
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   -66600
         TabIndex        =   116
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -71400
         TabIndex        =   115
         Top             =   3240
         Width           =   255
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000005&
         X1              =   -74400
         X2              =   -67680
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblFec_Actualizacion 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69360
         TabIndex        =   106
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label lblEstado 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73080
         TabIndex        =   105
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Responsabilidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -69720
         TabIndex        =   101
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Comisión X Sal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   -74400
         TabIndex        =   100
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label lblEstadoPago 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68760
         TabIndex        =   99
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label lblFechaPago 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71520
         TabIndex        =   98
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label lblUsuarioPago 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73920
         TabIndex        =   97
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label lblTesoreriaFecha 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68760
         TabIndex        =   96
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label lblTesoreriaUsuario 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71520
         TabIndex        =   95
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label lblTesoreriaSolicitud 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73920
         TabIndex        =   94
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label lblUltimoSaldo 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68760
         TabIndex        =   93
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblUltimaTasa 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71520
         TabIndex        =   92
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblUltimaCuota 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73920
         TabIndex        =   91
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -70080
         TabIndex        =   88
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   -74640
         TabIndex        =   85
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tesorería"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   -74640
         TabIndex        =   84
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Ültima Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   -74640
         TabIndex        =   83
         Top             =   3600
         Width           =   975
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74640
         X2              =   -67200
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74640
         X2              =   -67200
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label2 
         Caption         =   "Tasa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   -72120
         TabIndex        =   82
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   -69600
         TabIndex        =   81
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   -74640
         TabIndex        =   80
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   39
         Left            =   -69600
         TabIndex        =   79
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fec Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   -72240
         TabIndex        =   78
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   -74640
         TabIndex        =   77
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   -69600
         TabIndex        =   76
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   -72240
         TabIndex        =   75
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Solicitud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   -74640
         TabIndex        =   74
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74640
         X2              =   -67320
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label2 
         Caption         =   "Comisión Admin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   -74400
         TabIndex        =   65
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   -69720
         TabIndex        =   59
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74640
         X2              =   -67320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   -74160
         TabIndex        =   53
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   -74760
         TabIndex        =   49
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   20
         Left            =   -70560
         TabIndex        =   50
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74640
         X2              =   -67080
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   17
         Left            =   4320
         TabIndex        =   46
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74760
         X2              =   -67080
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   7800
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000004&
         X1              =   -74640
         X2              =   -67560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   45
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Acreedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74400
         TabIndex        =   43
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   42
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   41
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tasa Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -69720
         TabIndex        =   40
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Plazo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74400
         TabIndex        =   39
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -69000
         TabIndex        =   38
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74400
         TabIndex        =   37
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fec Formalización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -69720
         TabIndex        =   36
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fec 1º Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74400
         TabIndex        =   34
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Día de Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -69720
         TabIndex        =   33
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   13
         Left            =   -74400
         TabIndex        =   32
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -74400
         TabIndex        =   31
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Actualización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -70800
         TabIndex        =   35
         Top             =   6360
         Width           =   1695
      End
   End
   Begin XtremeSuiteControls.PushButton btnRenumerar 
      Height          =   495
      Left            =   3360
      TabIndex        =   131
      ToolTipText     =   "Renumerar el No. Operación"
      Top             =   240
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   79
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
      Appearance      =   17
      Picture         =   "frmCR_APA_Operaciones.frx":6ECD6
   End
   Begin XtremeShortcutBar.ShortcutCaption lblOperacionActiva 
      Height          =   495
      Left            =   3360
      TabIndex        =   132
      Top             =   240
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Operación Activa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
Attribute VB_Name = "frmCR_APA_Operaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCod_Acreedor As String
Private mCod_Contacto As String
Private mOperacion As String
Private mNPago As Integer

Dim strSQL As String
Dim i As Integer
Dim vEdita As Boolean
Dim vEditaTodo As Boolean
Dim vDetalle As Boolean
Dim vCambios As Boolean
Dim gTipoCambio As Currency, gDivisa As String, gVariacion As Integer, gDivisaDesc As String
Dim gDivisaLocal As Integer, gVariacionAsiento As Integer, strMovimiento As String

Private Sub sbCargarListaOperaciones()
' Carga Lista de operaciones
    Dim strSQL As String
    
On Error GoTo error
    'Consulta la lista de las Operaciones
    strSQL = "select OPERACION, MONTO, CUOTA, SALDO, case when ESTADO = 'A' then 'Activa' " _
            & " when ESTADO = 'C' then  'Cancelado' else '' end as ESTADO " _
            & " from CRD_APA_OPERACIONES where COD_ACREEDOR = '" & Trim(mCod_Acreedor) & "'"
        If chkFiltroOperaciones.Value = 1 Then
        
            If txtOperacionBusqueda.Text <> Empty Then
                strSQL = strSQL & " and OPERACION like ('%" & Trim(txtOperacionBusqueda) & "%')"
            End If
            
            If cboEstadoOpeBusq <> "Todos" Then
                strSQL = strSQL & " and ESTADO = '" & Mid(cboEstadoOpeBusq, 1, 1) & "'"
            End If
            
        End If
        
    Call sbCargaGridCheckIni(vGridOperaciones, 5, strSQL)
    vGridOperaciones.MaxRows = vGridOperaciones.MaxRows - 1
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub

Private Sub sbCargarListaPagos()
    Dim strSQL As String
    
On Error GoTo error
    'Consulta la lista de los pagos
    strSQL = "select NPAGO, convert(varchar(20),PAGO_FECHA,103), DOCUMENTO, MONTO, " _
        & " case when ESTADO = 'P' then 'Pendiente' " _
        & " when ESTADO = 'C' then  'Cancelado' when ESTADO = 'E' then  'Ejecutado'" _
        & " when ESTADO = 'D' then  'Detallado' else '' end as ESTADO " _
        & " from CRD_APA_PAGOS where COD_ACREEDOR = '" & Trim(mCod_Acreedor) _
        & "' and OPERACION = '" & Trim(mOperacion) & "'"
        
        'Filtros
        If chkFiltroPagos.Value = 1 Then

            strSQL = strSQL & " and PAGO_FECHA >= '" & Format(dtpFecPagosDesde, "yyyymmdd 00:00:00") & "'" _
                    & " and PAGO_FECHA <= '" & Format(dtpFecPagosHasta, "yyyymmdd 23:59:59") & "'"
            
            If cboEstadoPagoBusq <> "Todos" Then
                strSQL = strSQL & " and ESTADO = '" & Mid(cboEstadoPagoBusq, 1, 1) & "'"
            End If

        End If
        
       strSQL = strSQL & " order by NPAGO desc "
        
    Call sbCargaGridCheckIni(vGridPagos, 5, strSQL)
    vGridPagos.MaxRows = vGridPagos.MaxRows - 1
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Function fxNombreAcreedor(ByRef mAcreedor As String) As String
    'Funcion que consulta el nombre del acreedor
    Dim strSQL As String
    Dim rs As New ADODB.Recordset

On Error GoTo error
    
    strSQL = "select DESCRIPCION from CRD_APA_ACREEDORES where COD_ACREEDOR = '" & mAcreedor & "'"
    
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then fxNombreAcreedor = rs.Fields(0)
    
    rs.Close
    Exit Function
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Function

Private Sub sbCargarUltimoPago(ByRef Cod_Acreedor As String, ByRef Operacion As String, Optional nPago As Integer)
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
On Error GoTo error
    
    If nPago = Empty Then
    
        strSQL = "select SALDO, TASA, CUOTA from CRD_APA_OPERACIONES where COD_ACREEDOR = '" & Cod_Acreedor & "'" _
                    & " and OPERACION = '" & Trim(mOperacion) & "'"
            
        Call OpenRecordSet(rs, strSQL)
            
        lblUltimaCuota = Format(IIf(IsNull(rs!Cuota), Empty, rs!Cuota), "Standard")
        lblUltimaTasa = Format(IIf(IsNull(rs!Tasa), Empty, rs!Tasa), "Standard")
        lblUltimoSaldo = Format(IIf(IsNull(rs!Saldo), Empty, rs!Saldo), "Standard")
    
    
    Else
        If nPago > 1 Then
        
        
            strSQL = "select MONTO, DETALLE_TASA, DETALLE_SALDO from CRD_APA_PAGOS where COD_ACREEDOR = '" & Cod_Acreedor & "'" _
                    & " and OPERACION = '" & Trim(mOperacion) & "'"
            
            If nPago <> Empty Then
                
                nPago = nPago - 1
                strSQL = strSQL & " and NPAGO = " & Trim(nPago)
                
            End If
            
            Call OpenRecordSet(rs, strSQL)
            
            lblUltimaCuota = Format(IIf(IsNull(rs!Monto), Empty, rs!Monto), "Standard")
            lblUltimaTasa = Format(IIf(IsNull(rs!Detalle_Tasa), Empty, rs!Detalle_Tasa), "Standard")
            lblUltimoSaldo = Format(IIf(IsNull(rs!Detalle_Saldo), Empty, rs!Detalle_Saldo), "Standard")
            
        Else
        
            strSQL = "select CUOTA_ORIGINAL, TASA_ORIGINAL, MONTO  from CRD_APA_OPERACIONES where COD_ACREEDOR = '" & Cod_Acreedor & "'" _
                    & " and OPERACION = '" & Trim(Operacion) & "'"
            
            Call OpenRecordSet(rs, strSQL)
            
            If Not rs.EOF Then
                lblUltimaCuota = Format(IIf(IsNull(rs!CUOTA_ORIGINAL), Empty, rs!CUOTA_ORIGINAL), "Standard")
                lblUltimaTasa = Format(IIf(IsNull(rs!TASA_ORIGINAL), Empty, rs!TASA_ORIGINAL), "Standard")
                lblUltimoSaldo = Format(IIf(IsNull(rs!Monto), Empty, rs!Monto), "Standard")
            End If
        End If
    End If
    rs.Close
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub sbLlenarListaAcreedores()
' Carga lista de acreedores
On Error GoTo error

    Dim vItem As MSComctlLib.ListItem
    Dim vLvw As MSComctlLib.ListView
    Dim vKey As String
    Dim rs As New ADODB.Recordset
    
    Me.lswAcreedores.ColumnHeaders.Clear
    Me.lswAcreedores.ListItems.Clear
    
    Set vLvw = Me.lswAcreedores
    vLvw.ColumnHeaders.Add , , "Descripción", 2400
    
    strSQL = "select COD_ACREEDOR, DESCRIPCION, ESTADO  from dbo.CRD_APA_ACREEDORES " & _
             " order by DESCRIPCION "
    Call OpenRecordSet(rs, strSQL)

    While Not rs.EOF
        
        vKey = Trim(rs.Fields("COD_ACREEDOR")) & "(CA)"
        
        Set vItem = lswAcreedores.ListItems.Add(, vKey, Trim(rs.Fields!Descripcion))
        
        rs.MoveNext
    Wend

    rs.Close
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbLlenarListaContactos()

On Error GoTo error


        Dim vItem As MSComctlLib.ListItem
        Dim vLvw As MSComctlLib.ListView
        Dim vKey As String
        Dim rs As New ADODB.Recordset
        
        Me.lswContactos.ColumnHeaders.Clear
        Me.lswContactos.ListItems.Clear
        
        Set vLvw = Me.lswContactos
        vLvw.ColumnHeaders.Add , , "Descripción", 2400
        
        strSQL = "select COD_CONTACTO, NOMBRE from dbo.CRD_APA_CONTACTOS where COD_ACREEDOR = '" & mCod_Acreedor & _
                 "' order by COD_CONTACTO "
        Call OpenRecordSet(rs, strSQL)
    
        While Not rs.EOF
            
            vKey = Trim(rs.Fields("COD_CONTACTO")) & "(CC)"
            
            Set vItem = lswContactos.ListItems.Add(, vKey, Trim(rs.Fields!Nombre))
            
            rs.MoveNext
        Wend
    rs.Close
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub sbCalcularSaldos()

On Error GoTo error

    If lblUltimoSaldo = Empty Then lblUltimoSaldo = 0
    If lblAmortizacionPago = Empty Then lblAmortizacionPago = 0
    If txtMontoPago = Empty Then txtMontoPago = 0
    If txtInteresesPago = Empty Then txtInteresesPago = 0
    If txtCargosPago = Empty Then txtCargosPago = 0
    If txtComisionPago = Empty Then txtComisionPago = 0
    
    lblAmortizacionPago = CDbl(txtMontoPago) - (CDbl(txtInteresesPago) + CDbl(txtCargosPago) + CDbl(txtComisionPago))
    lblAmortizacionPago = Format(lblAmortizacionPago, "Standard")
    
    If CDbl(lblUltimoSaldo) <> 0 Then
        lblTasaPago = CDbl(txtInteresesPago) / CDbl(lblUltimoSaldo) * 12 * 100
        lblTasaPago = Format(lblTasaPago, "Standard")
    Else
        lblTasaPago = Format(0, "Standard")
    End If
    
    lblSaldoPago = CDbl(lblUltimoSaldo) - CDbl(lblAmortizacionPago)
    lblSaldoPago = Format(lblSaldoPago, "Standard")
    
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbLlenarListaTelefonos()

On Error GoTo error

        Dim rs As New ADODB.Recordset
        
        Me.lstDatos.Clear
        
        If mCod_Contacto <> Empty Then
            
            strSQL = "select isnull(TEL_CEL,'') as TEL_CEL , isnull(TEL_TRABAJO,'') as TEL_TRABAJO, " & _
                     " isnull(TEL_FAX,'') as TEL_FAX, isnull(EMAIL,'') as EMAIL " & _
                     " from dbo.CRD_APA_CONTACTOS where COD_CONTACTO =  " & mCod_Contacto & _
                     " and COD_ACREEDOR = '" & mCod_Acreedor & "'"
                     
            Call OpenRecordSet(rs, strSQL)
        
            If Not rs.EOF Then
                If Trim(rs.Fields(0)) <> "" Then
                    lstDatos.AddItem ("Celular:")
                    lstDatos.AddItem (Trim(rs.Fields!TEL_CEL))
                End If
                
                If Trim(rs.Fields(1)) <> "" Then
                    lstDatos.AddItem ("Trabajo:")
                    lstDatos.AddItem (Trim(rs.Fields!TEL_TRABAJO))
                End If
                
                If Trim(rs.Fields(2)) <> "" Then
                    lstDatos.AddItem ("Fax:")
                    lstDatos.AddItem (rs.Fields!TEL_FAX)
                End If
                
                If Trim(rs.Fields(3)) <> "" Then
                    lstDatos.AddItem ("E-mail:")
                    lstDatos.AddItem (Trim(rs.Fields!EMAIL))
                End If
            End If
            rs.Close
        End If
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub btnRenumerar_Click()
  frmCR_APA_OperacionRenumera.Show vbModal
End Sub

Private Sub cboFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFechaPago.SetFocus
End Sub

Private Sub cboMoneda_Change()
  vCambios = True
End Sub

Private Sub cboMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCambio.SetFocus
End Sub

Private Sub cboMoneda_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset
Dim Cod_Moneda As String

 Cod_Moneda = Trim(cboMoneda.Text)
 strSQL = "Select dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",'" & Mid(Cod_Moneda, 1, 3) & "' ,dbo.MyGetdate(),'V') as 'TipoCambio'"
 Call OpenRecordSet(rs, strSQL)
 
 txtTipoCambio = rs!TipoCambio
 rs.Close
 
End Sub

Private Sub cboOficina_Change()
    vCambios = True
End Sub

Private Sub cboTipo_Click()
    vCambios = True
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDia_Pago.SetFocus
End Sub


Private Sub cboTipoPago_Click()
vCambios = True
End Sub

Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub chkFiltroOperaciones_Click()
    Call sbCargarListaOperaciones
End Sub

Private Sub chkFiltroPagos_Click()
    Call sbCargarListaPagos
End Sub

Private Sub dtpFechaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVencePago.SetFocus
End Sub

Private Sub dtpFormalizacion_Change()
    vCambios = True
End Sub

Private Sub dtpFormalizacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoPago.SetFocus
End Sub

Private Sub dtpPrimerPago_Change()
    vCambios = True
End Sub

Private Sub dtpPrimerPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFormalizacion.SetFocus
End Sub

Private Sub dtpVencePago_Change()
    vCambios = True
End Sub

Private Sub dtpVencePago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtInteresesPago.SetFocus
End Sub

Private Sub Form_Activate()

    vModulo = 14 'Modulo de Credito
    
End Sub

Private Sub Form_Load()
    
    vModulo = 14 'Modulo de Credito
    
    
    If GLOBALES.gEnlace = 0 Then
        Call sbgCntParametros
    End If
    
    '' Carga nombre de la ternimal
    If Len(glogon.Maquina) = 0 Then
        Call sbMaquina
    End If
    
    mCod_Acreedor = Empty
    ssTab.Tab = 0
    Call ssTab_Click(0)
    Call sbCargaCombos
    Call sbLlenarListaAcreedores
    Call sbCargarListaOperaciones
    dtpFecPagosDesde.Value = fxFechaServidor()
    dtpFecPagosHasta.Value = dtpFecPagosDesde.Value
    Call sbCargaOficinas
    'Call sbCargaBanco

End Sub

Private Sub sbGuardarOperacion()
Dim strSQL As String, rs As New ADODB.Recordset, Responsabilidad_Base As String
Dim Comision_Base As String
Dim vCod_Oficina As String
Dim Fecha_Prox_Pago As String

If OptComisionOperacion.Value Then
    Comision_Base = "O"
Else
    Comision_Base = "G"
End If
If OptResponOperacion.Value Then
    Responsabilidad_Base = "O"
Else
    Responsabilidad_Base = "G"
End If

If Mid(cboTipo, 1, 1) = "C" Then
   vCod_Oficina = DeCodificaPrimaryKey(cboOficina.SelectedItem.Key, 1, "(id)")
Else
   vCod_Oficina = Empty
End If

On Error GoTo vError

If vEdita Then

    If vEditaTodo Then
    
        strSQL = "update crd_apa_operaciones set porc_responsabilidad = '" & Trim(txtResponsabilidad) _
               & "',tipo = '" & Mid(cboTipo, 1, 1) & "', FECHA_PRIMER_PAGO = '" & Format(dtpPrimerPago.Value, "yyyy/mm/dd") _
               & "',notas = '" & Trim(txtNotas) & "', FECHA_FORMALIZA = '" & Format(dtpFormalizacion.Value, "yyyy/mm/dd") _
               & "',monto = " & CDec(txtMonto) _
               & ",tasa = " & CDec(txtTasa) _
               & ",tasa_original = " & CDec(txtTasa) _
               & ",plazo = " & CInt(txtPlazo) _
               & ",plazo_original = " & CInt(txtPlazo) _
               & ",cuota = " & CDec(txtCuota) _
               & ",cuota_original = " & CDec(txtCuota) _
               & ",dia_de_pago = " & Trim(txtDia_Pago) _
               & ",RESPONSABILIDAD_BASE = " & pc(Responsabilidad_Base) _
               & ",COMISION_BASE = " & pc(Comision_Base) _
               & ",COMISION_ADMIN = " & CDec(txtComisionAdmin) _
               & ",COD_OFICINA = '" & Trim(vCod_Oficina) _
               & "',PERIOCIDAD_PAGO = '" & Mid(cboTipoPago, 1, 1) _
               & "',COD_DIVISA='" & Mid(cboMoneda.Text, 1, 3) _
               & "',TIPO_CAMBIO=" & txtTipoCambio _
               & "  where operacion = '" & Trim(txtOperacion) & "'" _
               & "  and cod_acreedor = '" & Trim(txtCod_Acreedor) & "'"
               
        Call ConectionExecute(strSQL)
        
        Call Bitacora("MODIFICA", "APA Operación: " & Trim(txtOperacion) & " Acreedor:" & Trim(txtCod_Acreedor))
    
    Else
    
         strSQL = "update crd_apa_operaciones set porc_responsabilidad = " & Trim(txtResponsabilidad) _
               & ",notas = '" & Trim(txtNotas) & "', FECHA_PRIMER_PAGO = '" & Format(dtpPrimerPago.Value, "yyyy/mm/dd") _
               & "',plazo = " & CInt(txtPlazo) & ", FECHA_FORMALIZA = '" & Format(dtpFormalizacion.Value, "yyyy/mm/dd") _
               & "',dia_de_pago = " & Trim(txtDia_Pago) _
               & ",RESPONSABILIDAD_BASE = " & pc(Responsabilidad_Base) _
               & ",COMISION_BASE = " & pc(Comision_Base) _
               & ",COMISION_ADMIN = " & CDec(txtComisionAdmin) _
               & ",COD_OFICINA = '" & Trim(vCod_Oficina) _
               & ",PERIOCIDAD_PAGO = " & pc(cboTipoPago) _
               & "',COD_DIVISA='" & Mid(cboMoneda.Text, 1, 3) _
               & "',TIPO_CAMBIO=" & txtTipoCambio _
               & "' where operacion = '" & Trim(txtOperacion) & "'" _
               & " and cod_acreedor = '" & Trim(txtCod_Acreedor) & "'"
               
        Call ConectionExecute(strSQL)
        
        Call Bitacora("MODIFICA", "APA Operación: " & Trim(txtOperacion) & " Acreedor:" & Trim(txtCod_Acreedor))
    End If
    
Else
   strSQL = "insert crd_apa_operaciones(COD_ACREEDOR, OPERACION, PORC_RESPONSABILIDAD, TIPO, NOTAS, MONTO, SALDO, TASA, TASA_ORIGINAL, PLAZO, PLAZO_ORIGINAL " _
          & ",CUOTA,CUOTA_ORIGINAL,FECHA_FORMALIZA,FECHA_PRIMER_PAGO,DIA_DE_PAGO,COMISION_ADMIN,ESTADO,RESPONSABILIDAD_BASE,COMISION_BASE,COD_OFICINA " _
          & ",PERIOCIDAD_PAGO,FECHA_PROX_PAGO,COD_DIVISA,TIPO_CAMBIO)" _
          & " values('" & Trim(txtCod_Acreedor) & "','" & Trim(txtOperacion) & "','" & txtResponsabilidad & "','" & Mid(cboTipo, 1, 1) & "','" & Trim(txtNotas.Text) & "'," & CCur(txtMonto) & "," & CCur(txtMonto) & "," & CCur(txtTasa) & "," & CCur(txtTasa) _
          & "," & CInt(txtPlazo) & "," & CInt(txtPlazo) & "," & CCur(txtCuota.Text) & "," & CCur(txtCuota.Text) & ",'" & Format(dtpFormalizacion.Value, "yyyy/mm/dd") _
          & "','" & Format(dtpPrimerPago.Value, "yyyy/mm/dd") & "'," & Trim(txtDia_Pago) & "," & CCur(txtComisionAdmin) & ",'A'" _
          & "," & pc(Responsabilidad_Base) & "," & pc(Comision_Base) & ",'" & Trim(vCod_Oficina) & "', '" & Mid(cboTipoPago, 1, 1) & "','" & Format(dtpPrimerPago.Value, "yyyy/mm/dd") _
          & "','" & Mid(cboMoneda.Text, 1, 3) & "'," & txtTipoCambio & ")"
                    

    Call ConectionExecute(strSQL)
    
    Call Bitacora("REGISTRA", "APA Operación: " & Trim(txtOperacion) & " Acreedor:" & Trim(txtCod_Acreedor))
     
     
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation


'Call sbToolBar(tlbPrincipal, "activo")
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub sbGuardarPago()
Dim strSQL As String, rs As New ADODB.Recordset
Dim nPago As Integer, FPago As String, vFecha As Date
Dim NumeroPago As String

On Error GoTo vError

vFecha = fxFechaServidor

    Select Case cboFormaPago.ItemData(cboFormaPago.ListIndex)
    Case 1
        FPago = "DC"
    Case 2
        FPago = "CK"
    End Select

    Call sbAgregarPago(mCod_Acreedor, _
                            mOperacion, _
                            glogon.Usuario, _
                            Format(vFecha, "yyyymmdd hh:mm:ss"), _
                            Format(dtpVencePago.Value, "yyyymmdd"), _
                            "E", _
                            txtMontoPago, _
                            txtInteresesPago, _
                            txtCargosPago, _
                            lblAmortizacionPago, _
                            lblTasaPago, _
                            lblSaldoPago, _
                            Format(vFecha, "yyyymmdd hh:mm:ss"), _
                            glogon.Usuario, _
                            txtComisionPago, _
                            txtDocumentoPago, _
                            FPago)
                            
    
    strSQL = "select max(NPAGO) from CRD_APA_PAGOS where COD_ACREEDOR =" & pc(Trim(mCod_Acreedor)) _
            & " and OPERACION = " & pc(Trim(mOperacion))
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        NumeroPago = rs.Fields(0)
    End If
    rs.Close
                            
    Call Bitacora("REGISTRA", "APA Pago: " & NumeroPago & " Operación: " & Trim(mOperacion) & " Acreedor:" & Trim(txtCod_Acreedor))
     
    MsgBox "Información guardada satisfactoriamente...", vbInformation
    
    'Call sbToolBar(tlbPrincipal, "activo")
    Call RefrescaTags(Me)
    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbAgregarPago(ByVal Cod_Acreedor As String, _
                                ByVal Operacion As String, _
                                ByVal Pago_Usuario As String, _
                                ByVal PAGO_FECHA As String, _
                                ByVal Fecha_Pago As String, _
                                ByVal Estado As String, _
                                ByVal Monto As Double, _
                                ByVal Detalle_Intereses As Double, _
                                ByVal Detalle_Cargos As Double, _
                                ByVal Detalle_Amortiza As Double, _
                                ByVal Detalle_Tasa As Double, _
                                ByVal Detalle_Saldo As Double, _
                                ByVal Detalle_Fecha As String, _
                                ByVal Detalle_Usuario As String, _
                                ByVal Detalle_Comision As Double, _
                                ByVal Documento As String, _
                                ByVal Forma_Pago As String)
                                    
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCRDAPA_PAGOS_A " & pcc(Cod_Acreedor) _
                                & pcc(Operacion) _
                                & pcc(Pago_Usuario) _
                                & pcc(PAGO_FECHA) _
                                & pcc(Fecha_Pago) _
                                & pcc(Estado) _
                                & CDec(Monto) & "," _
                                & CDec(Detalle_Intereses) & "," _
                                & CDec(Detalle_Cargos) & "," _
                                & CDec(Detalle_Amortiza) & "," _
                                & CDec(Detalle_Tasa) & "," _
                                & CDec(Detalle_Saldo) & "," _
                                & pcc(Detalle_Fecha) _
                                & pcc(Detalle_Usuario) _
                                & CDec(Detalle_Comision) & "," _
                                & pcc(Documento) _
                                & pc(Forma_Pago)
                                        
Call ConectionExecute(strSQL)
Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al agregar el pago a la operación  " & mOperacion & " Error " & Err.Description
End Sub

Private Sub lswAcreedores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo error

    Call sbListaMarcarSoloUno(lswAcreedores, Item)
    mCod_Acreedor = Empty
    mOperacion = Empty
    mNPago = Empty
    If Item.Checked = True Then
        mCod_Acreedor = DeCodificaPrimaryKey(Item.Key, 1, "(CA)")
    End If
    Call sbLimpiarFiltrosOperaciones
    Call sbLlenarListaContactos
    Call sbCargarListaOperaciones
    Call sbLimpiarControlesPagos
    Call sbLimpiarControlesOperaciones
    
    ssTab.Tab = 0
    lstDatos.Clear
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
 

Private Sub lswContactos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo error

    Call sbListaMarcarSoloUno(lswContactos, Item)
    mCod_Contacto = Empty
    If Item.Checked = True Then
        mCod_Contacto = DeCodificaPrimaryKey(Item.Key, 1, "(CC)")
    End If
    Call sbLlenarListaTelefonos
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub OptComisionGarantias_Click()
    vCambios = True
End Sub

Private Sub OptComisionGarantias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then OptResponOperacion.SetFocus
End Sub

Private Sub OptComisionOperacion_Click()
    vCambios = True
End Sub

Private Sub OptComisionOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then OptComisionGarantias.SetFocus
End Sub

Private Sub OptResponGarantias_Click()
    vCambios = True
End Sub

Private Sub OptResponGarantias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCod_Acreedor.SetFocus
End Sub

Private Sub OptResponOperacion_Click()
    vCambios = True
End Sub

Private Sub OptResponOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then OptResponGarantias.SetFocus
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
On Error GoTo error

    Select Case ssTab.Tab
    Case 0, 2
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = False
        ssTab.TabEnabled(2) = True
        ssTab.TabEnabled(3) = False
        
        If ssTab.Tab = 2 Then
            mOperacion = fxGridValorMarcado(vGridOperaciones, 2)
            lblOperacionActiva.Caption = mOperacion
            
            Call sbCargarListaPagos
            
        End If
        
    Case 1
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = True
        ssTab.TabEnabled(2) = True
        ssTab.TabEnabled(3) = False
    Case 3
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = False
        ssTab.TabEnabled(2) = True
        ssTab.TabEnabled(3) = True
    End Select
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlbAutorizados_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
On Error GoTo error
    
    Select Case UCase(Button.Key)
    Case "APLICAR"
        
        ' Valida si ya se realizó el pase del pago a tesoreria
        strSQL = "select count(*) from CRD_APA_PAGOS where TESORERIA_FECHA is not null " _
                        & " and COD_ACREEDOR = " & pc(mCod_Acreedor) _
                        & " and OPERACION = " & pc(mOperacion) _
                        & " and NPAGO = " & mNPago
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) > 0 Then
            MsgBox "No se puede cambiar el Autorizado, ya se realizó el traslado a tesoreria"
            Exit Sub
        End If
        rs.Close
        
        If txtCedulaAutorizado.Text <> Empty Then
        
            strSQL = "select count(*) from CRD_APA_AUTORIZADOSCK where COD_ACREEDOR = " & pc(mCod_Acreedor) _
                   & " and CEDULA =" & pc(Trim(txtCedulaAutorizado))
            Call OpenRecordSet(rs, strSQL)
            If rs.Fields(0) = 0 Then
                MsgBox "No existe ese número de cédula asignado a ese acreedor"
                Exit Sub
            End If
            
            strSQL = "update CRD_APA_PAGOS set CEDULA_AUTORIZADO = " & pc(txtCedulaAutorizado) _
                        & " where COD_ACREEDOR = " & pc(mCod_Acreedor) _
                        & "       and OPERACION = " & pc(mOperacion) _
                        & "       and NPAGO = " & mNPago
            Call ConectionExecute(strSQL)
            
        Else
        
            strSQL = "update CRD_APA_PAGOS set CEDULA_AUTORIZADO = null " _
                        & " where COD_ACREEDOR = " & pc(mCod_Acreedor) _
                        & "       and OPERACION = " & pc(mOperacion) _
                        & "       and NPAGO = " & mNPago
            Call ConectionExecute(strSQL)
            
        End If
        
        fraAutorizados.Visible = False
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation
        
    Case "SALIR"
        fraAutorizados.Visible = False
    End Select
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlbBusquedaOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
    Select Case UCase(Button.Key)
    Case "BUSCAR"
        chkFiltroOperaciones.Value = 1
        Call sbCargarListaOperaciones
    End Select
    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlbBusquedaPagos_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
    
    Select Case UCase(Button.Key)
    Case "BUSCAR"
        chkFiltroPagos.Value = 1
        Call sbCargarListaPagos
    End Select
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlbIncluirOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error

    Select Case UCase(Button.Key)
        Case "GUARDAR"
        
            If vCambios Then
                If fxValida Then
                    Call sbGuardarOperacion
                Else
                    Exit Sub
                End If
            End If
            Call sbCargarListaOperaciones
            Call sbLimpiarControlesOperaciones
            ssTab.Tab = 0
            
    End Select
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxValida() As Boolean
    Dim strSQL As String, rs As New ADODB.Recordset
    Dim vMensaje As String

On Error GoTo error
    
    vMensaje = ""
    fxValida = True
    
    If txtCod_Acreedor.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el código del acreedor"
    If txtOperacion.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el número de operación"
    If txtResponsabilidad.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el porcentaje de responsabilidad"
    If txtMonto.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto del crédito"
    If txtTasa.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar la tasa del crédito"
    If txtPlazo.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el plazo del crédito"
    If txtCuota.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar la cuota del crédito"
    If txtDia_Pago.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el día de pago"
    If txtComisionAdmin.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar la comisión por administración"
    
    If vEdita = False Then
    
        'Verifica que exista ninguna operación con ese código
        strSQL = "select isnull(count(*),0) as Existe from CRD_APA_OPERACIONES" _
               & " where COD_ACREEDOR = '" & Trim(txtCod_Acreedor.Text) & "'" _
               & " and OPERACION = '" & Trim(txtOperacion.Text) & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe > 0 Then
           vMensaje = vMensaje & vbCrLf & " Ya Existe una operación con ese código en este acreedor"
        End If
        rs.Close
        
    End If
    
    If Len(vMensaje) > 0 Then
      fxValida = False
      MsgBox vMensaje, vbCritical
    End If
    Exit Function
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Function

Private Function fxNumPagosporOperacion() As Integer
    Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo error

    'Verifica si la operación ya tiene un pago asociado
    strSQL = "select isnull(count(*),0) as Existe from CRD_APA_PAGOS" _
            & " where COD_ACREEDOR = '" & Trim(txtCod_Acreedor.Text) & "'" _
            & " and OPERACION = '" & Trim(txtOperacion.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
        fxNumPagosporOperacion = rs!Existe
    Else
        fxNumPagosporOperacion = 0
    End If
    rs.Close
    Exit Function
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
        
End Function

Private Function fxValidaPago() As Boolean
    Dim strSQL As String, rs As New ADODB.Recordset
    Dim vMensaje As String
    
On Error GoTo error
    
    vMensaje = ""
    fxValidaPago = True
    

        If txtInteresesPago.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto de los intereses"
        If txtCargosPago.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto de los cargos"
        If lblAmortizacionPago.Caption = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto de la amortización"
        If lblTasaPago.Caption = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto de ls tasa"
        If lblSaldoPago.Caption = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto del saldo del crédito"
        If txtMontoPago.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto del pago"
        If txtComisionPago.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el monto de la comisión"

        If lblAmortizacionPago < 0 Then
            vMensaje = vMensaje & vbCrLf & "La suma de los cargos es mayor que la cuota"
        End If
        
    If Len(vMensaje) > 0 Then
      fxValidaPago = False
      MsgBox vMensaje, vbCritical
    End If
    Exit Function
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub tlbIncluirPagos_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
     
    Select Case UCase(Button.Key)
    Case "GUARDAR"
    
        If vCambios = True Then
            If fxValidaPago Then
                Call sbGuardarPago
            Else
                Exit Sub
            End If
        End If
        Call sbCargarListaPagos
        Call sbLimpiarControlesPagos
        ssTab.Tab = 2
        
    Case "CANCELAROPERACION"
    
        Call sbLimpiarControlesPagos
        ssTab.Tab = 2
        
    End Select
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlbListaOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
    Dim rs As New ADODB.Recordset, strSQL As String

    Select Case UCase(Button.Key)
        Case "NUEVO"
        
            ssTab.Tab = 1
            Call sbLimpiarControlesOperaciones
            vEdita = False
            vCambios = False
            Call sbBloquearControlesOperaciones("")
            If mCod_Acreedor <> Empty Then
                txtCod_Acreedor = Trim(mCod_Acreedor)
                txtDescripcionAcreedor = fxNombreAcreedor(mCod_Acreedor)
                txtOperacion.SetFocus
            Else
                txtCod_Acreedor.SetFocus
            End If
            
        Case "EDITAR"
        
            mOperacion = fxGridValorMarcado(vGridOperaciones, 2)
            lblOperacionActiva.Caption = mOperacion
            
            If mOperacion = Empty Then
                MsgBox "Seleccione la operación"
                Exit Sub
            End If
            Call sbConsultaOperacion(mOperacion, mCod_Acreedor)
            vEdita = True
            vEditaTodo = True
            vCambios = False
            '' Valida si la operacion tiene pagos asociados
            If fxNumPagosporOperacion = 0 Then
                Call sbBloquearControlesOperaciones("editartodo")
            Else
                vEditaTodo = False
                Call sbBloquearControlesOperaciones("editar")
            End If
            ssTab.Tab = 1
            
        Case "VER"
        
            mOperacion = fxGridValorMarcado(vGridOperaciones, 2)
            lblOperacionActiva.Caption = mOperacion
            
            If mOperacion = Empty Then
                MsgBox "Seleccione la operación"
                Exit Sub
            End If
            Call sbConsultaOperacion(mOperacion, mCod_Acreedor)
            Call sbBloquearControlesOperaciones("ver")
            vEdita = False
            vCambios = False
            ssTab.Tab = 1
            
        Case "CERRAR"
        
            mOperacion = fxGridValorMarcado(vGridOperaciones, 2)
            lblOperacionActiva.Caption = mOperacion
            
            If mOperacion = Empty Then
                MsgBox "Seleccione la operación"
                Exit Sub
            End If
            
            strSQL = "select SALDO, ESTADO from CRD_APA_OPERACIONES where COD_ACREEDOR = '" & Trim(mCod_Acreedor) & "'" _
              & " and OPERACION = '" & Trim(mOperacion) & "'"
            Call OpenRecordSet(rs, strSQL)
            
            If Not rs.EOF Then
                If rs.Fields(0) > 0 Then
                    MsgBox "No es posible cerrar la operación seleccionada porque tiene saldo mayor a cero"
                    Exit Sub
                End If
                If rs.Fields(1) <> "A" Then
                    MsgBox "Solo es posible cerrar operaciones activas"
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            
            rs.Close
            
            strSQL = "update CRD_APA_OPERACIONES set ESTADO = 'C' where COD_ACREEDOR = '" & Trim(mCod_Acreedor) & "'" _
              & " and OPERACION = '" & Trim(mOperacion) & "'"
              
            Call OpenRecordSet(rs, strSQL)
            
            Call Bitacora("CERRAR", "APA Operación: " & Trim(mOperacion) & " Acreedor:" & Trim(mCod_Acreedor))
            
            Call sbCargarListaOperaciones
            
            MsgBox "Información guardada satisfactoriamente...", vbInformation
    
    End Select
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsultaOperacion(Operacion As String, Acreedor As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer


On Error GoTo vError

    Me.MousePointer = vbHourglass

    strSQL = "select O.COD_ACREEDOR, O.OPERACION, O.PORC_RESPONSABILIDAD, O.TIPO, O.NOTAS, O.MONTO, O.TASA, O.PLAZO, O.CUOTA " _
              & ",O.SALDO, O.FECHA_FORMALIZA, O.FECHA_ACTUALIZA,O.PERIOCIDAD_PAGO,A.DESCRIPCION AS 'ACREEDOR'" _
              & ",O.FECHA_PRIMER_PAGO, O.DIA_DE_PAGO, O.ESTADO, O.COD_DIVISA, O.TIPO_CAMBIO" _
              & ",O.COMISION_ADMIN, O.RESPONSABILIDAD_BASE, O.COMISION_BASE,O.PERIOCIDAD_PAGO" _
              & " from CRD_APA_OPERACIONES O inner join CRD_APA_ACREEDORES A on O.cod_Acreedor = A.cod_acreedor" _
              & " where O.COD_ACREEDOR = '" & Trim(Acreedor) & "' and O.OPERACION = '" & Trim(Operacion) & "'"
    Call OpenRecordSet(rs, strSQL)

    If Not rs.BOF And Not rs.EOF Then
    '  Call sbToolBar(tlbPrincipal, "activo")
  
        vEdita = True
        txtCod_Acreedor.Text = Trim(rs!Cod_Acreedor)
        
        Select Case Trim(rs!Tipo)
            Case "M"
                cboTipo.Text = "Multiple"
            Case "U"
                cboTipo.Text = "Una a Una"
            Case "C"
                cboTipo.Text = "Capital de trabajo"
        End Select
        
        txtOperacion.Text = Trim(rs!Operacion)
        txtResponsabilidad.Text = Trim(rs!PORC_RESPONSABILIDAD)
        txtNotas.Text = Trim(IsNull(rs!Notas))
        txtMonto.Text = Format(rs!Monto, "Standard")
        txtTasa.Text = Format(rs!Tasa, "Standard")
        txtPlazo.Text = rs!Plazo
        txtCuota.Text = Format(rs!Cuota, "Standard")
        txtSaldo.Text = Format(rs!Saldo, "Standard")
        dtpFormalizacion.Value = rs!Fecha_Formaliza
        dtpPrimerPago.Value = rs!Fecha_Primer_Pago
        lblFec_Actualizacion.Caption = Format(IIf(IsNull(rs!FECHA_ACTUALIZA), Empty, Trim(rs!FECHA_ACTUALIZA)), "dd/mm/yyyy hh:mm")
        txtDia_Pago.Text = rs!dia_de_pago
        txtTipoCambio.Text = rs!TIPO_CAMBIO
        
        Select Case rs!Estado
            Case "A"
                lblEstado = "Activa"
            Case "C"
                lblEstado = "Cancelada"
        End Select
        
        txtComisionAdmin = rs!COMISION_ADMIN
        
        Select Case rs!Responsabilidad_Base
            Case "O"
                OptResponOperacion.Value = True
            Case "G"
                OptResponGarantias.Value = True
        End Select

        Select Case rs!Comision_Base
            Case "O"
                OptComisionOperacion.Value = True
            Case "G"
                OptComisionGarantias.Value = True
        End Select
        
        Select Case Trim(rs!PERIOCIDAD_PAGO)
            Case "M"
                cboTipoPago.Text = "Mensual"
            Case "T"
                cboTipoPago.Text = "Trimestral"
            Case "S"
                cboTipoPago.Text = "Semestral"
            Case "A"
                cboTipoPago.Text = "Anual"
        End Select
        
        Select Case Trim(rs!cod_divisa)
            Case "COL"
                cboMoneda.Text = "Colones"
            Case "DOL"
                cboMoneda.Text = "Dolares"
            Case "EUR"
                cboMoneda.Text = "Euros"
        End Select
        
        txtDescripcionAcreedor.Text = rs!Acreedor
    Else
        MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsultaPagos(Operacion As String, Acreedor As String, Pago As Integer)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select COD_ACREEDOR, OPERACION, NPAGO, DOCUMENTO, PAGO_USUARIO, PAGO_FECHA, FECHA_PAGO, ESTADO" _
          & ",MONTO,TESORERIA_SOLICITUD, TESORERIA_USUARIO, TESORERIA_FECHA, DETALLE_INTERESES," _
          & " DETALLE_CARGOS, DETALLE_AMORTIZA, DETALLE_TASA, DETALLE_SALDO, DETALLE_FECHA," _
          & " DETALLE_USUARIO, FORMA_PAGO, DETALLE_COMISION,ESTADO from CRD_APA_PAGOS " _
          & " where COD_ACREEDOR = '" & Trim(Acreedor) & "'" _
          & " and OPERACION = '" & Trim(Operacion) & "'" _
          & " and NPAGO = '" & Trim(Pago) & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
'  Call sbToolBar(tlbPrincipal, "activo")
  
vDetalle = True

lblNPago = rs!nPago
txtDocumentoPago = IIf(IsNull(rs!Documento), Empty, rs!Documento)

Select Case rs!Forma_Pago
    Case "DC"
        cboFormaPago.Text = "Débito Cuenta"
    Case "CK"
        cboFormaPago.Text = "Solicitud Cheque"
End Select

Select Case IIf(IsNull(rs!Estado), Empty, rs!Estado)
    Case "P"
        lblEstadoPago = "Pendiente"
    Case "C"
        lblEstadoPago = "Cancelado"
    Case "E"
        lblEstadoPago = "Ejecutado"
    Case "D"
        lblEstadoPago = "Detallado"
    Case ""
        lblEstadoPago = Empty
End Select

txtMontoPago = Format(CDec(rs!Monto), "Standard")
dtpVencePago = IIf(IsNull(rs!Fecha_Pago), fxFechaServidor, rs!Fecha_Pago)
txtInteresesPago = IIf(IsNull(rs!Detalle_Intereses), Empty, Format(rs!Detalle_Intereses, "Standard"))
txtCargosPago = IIf(IsNull(rs!Detalle_Cargos), Empty, Format(rs!Detalle_Cargos, "Standard"))
lblAmortizacionPago = IIf(IsNull(rs!Detalle_Amortiza), Empty, Format(rs!Detalle_Amortiza, "Standard"))
lblTasaPago = IIf(IsNull(rs!Detalle_Tasa), Empty, Format(rs!Detalle_Tasa, "Standard"))
lblSaldoPago = IIf(IsNull(rs!Detalle_Saldo), Empty, Format(rs!Detalle_Saldo, "Standard"))
txtComisionPago = IIf(IsNull(rs!Detalle_Comision), Empty, Format(rs!Detalle_Comision, "Standard"))
lblUsuarioPago = IIf(IsNull(rs!Pago_Usuario), Empty, Trim(rs!Pago_Usuario))
lblFechaPago = Format(rs!PAGO_FECHA, "dd/mm/yyyy hh:mm")
lblTesoreriaSolicitud = IIf(IsNull(rs!Tesoreria_Solicitud), Empty, Trim(rs!Tesoreria_Solicitud))
lblTesoreriaUsuario = IIf(IsNull(rs!Tesoreria_Usuario), Empty, Trim(rs!Tesoreria_Usuario))
lblTesoreriaFecha = Format(IIf(IsNull(rs!Tesoreria_Fecha), Empty, Trim(rs!Tesoreria_Fecha)), "dd/mm/yyyy hh:mm")
   
Else
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaCombos()
On Error GoTo error

    cboTipo.Clear
    cboTipo.AddItem "Multiple"
    cboTipo.AddItem "Una a Una"
    cboTipo.AddItem "Capital de trabajo"
    cboTipo.Text = "Multiple"
    
    cboTipoPago.Clear
    cboTipoPago.AddItem "Mensual"
    cboTipoPago.AddItem "Trimestral"
    cboTipoPago.AddItem "Semestral"
    cboTipoPago.AddItem "Anual"
    cboTipoPago.Text = "Mensual"
        
    cboFormaPago.Clear
    cboFormaPago.AddItem "Débito Cuenta"
    cboFormaPago.ItemData(cboFormaPago.NewIndex) = 1
    cboFormaPago.AddItem "Solicitud Cheque"
    cboFormaPago.ItemData(cboFormaPago.NewIndex) = 2
    cboFormaPago.Text = "Débito Cuenta"
    
    cboEstadoOpeBusq.Clear
    cboEstadoOpeBusq.AddItem "Activa"
    cboEstadoOpeBusq.AddItem "Cancelada"
    cboEstadoOpeBusq.AddItem "Todos"
    cboEstadoOpeBusq.Text = "Todos"
 
    cboEstadoPagoBusq.Clear
    cboEstadoPagoBusq.AddItem "Pendiente"
    cboEstadoPagoBusq.AddItem "Cancelado"
    cboEstadoPagoBusq.AddItem "Ejecutado"
    cboEstadoPagoBusq.AddItem "Detallo"
    cboEstadoPagoBusq.AddItem "Todos"
    cboEstadoPagoBusq.Text = "Todos"
    
    cboMoneda.Clear
    cboMoneda.AddItem "Colones"
    cboMoneda.AddItem "Dolares"
    cboMoneda.AddItem "Euros"
    cboMoneda.Text = "Colones"
    
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbApaControlDivisas(iCodigo As String, iContabilidad As Integer)
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = " SELECT isnull(TC_VENTA,1) as TC_VENTA ,COD_DIVISA," _
      & " FROM CNTX_DIVISAS " _
      & " where COD_DIVISA = " & iCodigo & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
  
    gTipoCambio = rs!tc_venta
    gDivisa = rs!cod_divisa
    rs.Close
    If gDivisa = "DOL" Or gDivisa = "EUR" And gTipoCambio = 1 Then
       strSQL = "SELECT D.TC_VENTA, D.VARIACION,X.Descripcion  from CNTX_DIVISAS_TIPO_CAMBIO D inner join  " _
               & " CNTX_DIVISAS X on D.COD_DIVISA = X.COD_DIVISA where  D.COD_CONTABILIDAD = " & iContabilidad & " " _
                & " and D.cod_divisa = '" & gDivisa & "' order by corte desc"
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF Or Not rs.BOF Then
          gTipoCambio = rs!tc_venta
          gVariacion = rs!VARIACION
          gDivisaDesc = LCase(rs!Descripcion)
       End If
    End If
End If

txtTipoCambio.Text = gTipoCambio
If Val(txtTipoCambio) = 1 Then
    txtTipoCambio.Locked = True
Else
    txtTipoCambio.Locked = False
End If

End Sub


Private Sub tlbListaPagos_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
    Select Case UCase(Button.Key)
        Case "NUEVO"
        
            If mCod_Acreedor <> Empty And mOperacion <> Empty Then
                ssTab.Tab = 3
                Call sbLimpiarControlesPagos
                vCambios = False
                Call sbCargarUltimoPago(mCod_Acreedor, mOperacion)
                Call sbBloquearControlesPagos("NUEVO")
                txtDocumentoPago.SetFocus
            Else
                MsgBox "Selecciones el acreedor y la operación que desea agregar el pago"
            End If
 
            
        Case "VER"
        
            lblNPago = fxGridValorMarcado(vGridPagos, 2)
            
            If mCod_Acreedor <> Empty And mOperacion <> Empty And lblNPago.Caption <> Empty Then
                Call sbBloquearControlesPagos("VER")
                Call sbConsultaPagos(mOperacion, mCod_Acreedor, lblNPago)
                Call sbCargarUltimoPago(mCod_Acreedor, mOperacion, lblNPago)
                vCambios = False
                ssTab.Tab = 3
                
            Else
                MsgBox "Selecciones el acreedor, la operación y el pago que desea detallar"
            End If
            

        
    End Select
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub







Private Sub tlbPagosOtros_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "NGIRO"
    
        If fxGridValorMarcado(vGridPagos, 2) = Empty Then
            MsgBox "Debe seleccionar el pago que desea agregar el autorizado"
            Exit Sub
        Else
            mNPago = fxGridValorMarcado(vGridPagos, 2)
        End If
        
        Call sbCargarAutorizado
        fraAutorizados.Top = tlbListaPagos.Top
        fraAutorizados.Visible = True
        txtCedulaAutorizado.SetFocus
    End Select
End Sub

Private Sub sbCargarAutorizado()
    Dim rs As New ADODB.Recordset, strSQL As String
    
     strSQL = "select isnull(P.CEDULA_AUTORIZADO,''), isnull(A.NOMBRE,'') as NOMBRE from CRD_APA_PAGOS P " _
            & " left join CRD_APA_AUTORIZADOSCK A on P.CEDULA_AUTORIZADO = A.CEDULA where P.COD_ACREEDOR = " & pc(mCod_Acreedor) _
            & " and P.OPERACION =" & pc(mOperacion) _
            & " and P.NPAGO =" & mNPago
        
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        txtCedulaAutorizado = rs.Fields(0)
        lblNombreAutorizado = rs.Fields(1)
    End If

End Sub

Private Sub sbCargaOficinas()
    Dim rs As New ADODB.Recordset, strSQL As String
    
    strSQL = "Select COD_OFICINA, DESCRIPCION" _
            & " from SIF_OFICINAS"
        
    Call OpenRecordSet(rs, strSQL)
    
    Do While Not rs.EOF
        cboOficina.ComboItems.Add , rs.Fields("COD_OFICINA") & "(id)", UCase(Trim(rs.Fields("DESCRIPCION")))
        rs.MoveNext
    Loop
    
    rs.Close
    
End Sub


Private Sub txtCargosPago_Change()
    vCambios = True
End Sub

Private Sub txtCargosPago_GotFocus()
    Call sbMarcarTXT(txtCargosPago)
End Sub

Private Sub txtCargosPago_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionPago.SetFocus
End Sub

Private Sub txtCargosPago_LostFocus()
    If txtCargosPago.Text <> Empty Then
        If Not IsNumeric(txtCargosPago) Then
            MsgBox "El campo solo permite valores numéricos"
            txtCargosPago.SetFocus
        End If
        txtCargosPago = Format(txtCargosPago, "Standard")
    End If
    Call sbCalcularSaldos
End Sub



Private Sub txtCedulaAutorizado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "NOMBRE"
        gBusquedas.Orden = "NOMBRE"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select CEDULA,NOMBRE from CRD_APA_AUTORIZADOSCK "
        gBusquedas.Filtro = " and COD_ACREEDOR = " & pc(mCod_Acreedor)
        frmBusquedas.Show vbModal
        txtCedulaAutorizado = gBusquedas.Resultado
        lblNombreAutorizado = gBusquedas.Resultado2
    
    End If
End Sub

Private Sub txtCedulaAutorizado_LostFocus()
    Dim rs As New ADODB.Recordset, strSQL As String
    
     strSQL = "select CEDULA, isnull(NOMBRE,'') as NOMBRE from CRD_APA_AUTORIZADOSCK " _
            & " where COD_ACREEDOR = " & pc(mCod_Acreedor) _
            & " and CEDULA =" & pc(Trim(txtCedulaAutorizado))
            
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        lblNombreAutorizado = rs.Fields(1)
    Else
        lblNombreAutorizado.Caption = Empty
    End If
    
End Sub

Private Sub txtCod_Acreedor_Change()
    vCambios = True
End Sub

Private Sub txtCod_Acreedor_GotFocus()
    Call sbMarcarTXT(txtCod_Acreedor)
End Sub

Private Sub txtCod_Acreedor_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOperacion.SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "Cod_Acreedor"
        gBusquedas.Orden = "Cod_Acreedor"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
        frmBusquedas.Show vbModal
        txtCod_Acreedor = gBusquedas.Resultado
        txtDescripcionAcreedor = gBusquedas.Resultado2
        txtOperacion.SetFocus
    
    End If
    
    
    
    
End Sub

Private Sub txtCod_Acreedor_LostFocus()
    Call sbNombreAcreedor
End Sub

Private Sub sbNombreAcreedor()
    Dim strSQL As String, rs As New ADODB.Recordset
    
    If txtCod_Acreedor.Text <> Empty Then
        strSQL = "select DESCRIPCION from CRD_APA_ACREEDORES where COD_ACREEDOR = " & pc(Trim(txtCod_Acreedor))
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            txtDescripcionAcreedor = rs.Fields(0)
        Else
            txtDescripcionAcreedor = Empty
            MsgBox "No existe un acreedor con ese código"
            txtCod_Acreedor.SetFocus
        End If
    End If
    
End Sub



Private Sub txtComisionAdmin_Change()
    vCambios = True
End Sub

Private Sub txtComisionAdmin_GotFocus()
    Call sbMarcarTXT(txtComisionAdmin)
End Sub

Private Sub txtComisionAdmin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtResponsabilidad.SetFocus
End Sub

Private Sub txtComisionAdmin_LostFocus()
    If txtComisionAdmin.Text <> Empty Then
        If Not IsNumeric(txtComisionAdmin) Then
            MsgBox "El campo solo permite valores numéricos"
            txtComisionAdmin.SetFocus
        End If
    End If
End Sub

Private Sub txtComisionPago_Change()
    vCambios = True
End Sub

Private Sub txtComisionPago_GotFocus()
    Call sbMarcarTXT(txtComisionPago)
End Sub

Private Sub txtComisionPago_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumentoPago.SetFocus
End Sub

Private Sub txtComisionPago_LostFocus()
    If txtComisionPago.Text <> Empty Then
        txtComisionPago = Format(txtComisionPago, "Standard")
        If Not IsNumeric(txtComisionPago) Then
            MsgBox "El campo solo permite valores numéricos"
            txtComisionPago.SetFocus
        End If
    End If
    Call sbCalcularSaldos
End Sub

Private Sub txtCuota_Change()
    vCambios = True
End Sub

Private Sub txtCuota_GotFocus()
    Call sbMarcarTXT(txtCuota)
End Sub

Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionAdmin.SetFocus
End Sub

Private Sub txtCuota_LostFocus()
    If txtCuota.Text <> Empty Then
        txtCuota = Format(txtCuota, "Standard")
        If Not IsNumeric(txtCuota) Then
            MsgBox "El campo solo permite valores numéricos"
            txtCuota.SetFocus
        End If
    End If
End Sub


Private Sub txtDia_Pago_Change()
    vCambios = True
End Sub

Private Sub txtDia_Pago_GotFocus()
    Call sbMarcarTXT(txtDia_Pago)
End Sub

Private Sub txtDia_Pago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtDia_Pago_LostFocus()
    If txtDia_Pago.Text <> Empty Then
        If Not IsNumeric(txtDia_Pago) Then
            MsgBox "El campo solo permite valores numéricos"
            txtDia_Pago.SetFocus
            Exit Sub
        End If
        If txtDia_Pago < 1 Or txtDia_Pago > 31 Then
            MsgBox "El campo solo permite valores entre 1-31"
            txtDia_Pago.SetFocus
            Exit Sub
        End If
    End If

End Sub



Private Sub txtDocumentoPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoPago.SetFocus
End Sub

Private Sub txtInteresesPago_Change()
    vCambios = True
End Sub

Private Sub txtInteresesPago_GotFocus()
    Call sbMarcarTXT(txtInteresesPago)
End Sub

Private Sub txtInteresesPago_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargosPago.SetFocus
End Sub

Private Sub txtInteresesPago_LostFocus()
    If txtInteresesPago.Text <> Empty Then
        If Not IsNumeric(txtInteresesPago) Then
            MsgBox "El campo solo permite valores numéricos"
            txtInteresesPago.SetFocus
        End If
        txtInteresesPago = Format(txtInteresesPago, "Standard")
    End If
    Call sbCalcularSaldos
End Sub

Private Sub txtMonto_Change()
    vCambios = True
End Sub

Private Sub txtMonto_GotFocus()
    Call sbMarcarTXT(txtMonto)
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTasa.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
    If txtMonto.Text <> Empty Then
        If Not IsNumeric(txtMonto) Then
            MsgBox "El campo solo permite valores numéricos"
            txtMonto.SetFocus
        End If
        txtMonto = Format(txtMonto, "Standard")
    End If
End Sub

Private Sub txtMontoPago_Change()
    vCambios = True
End Sub

Private Sub txtMontoPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboFormaPago.SetFocus
End Sub

Private Sub txtMontoPago_LostFocus()
    If txtMontoPago.Text <> Empty Then
        If Not IsNumeric(txtMontoPago) Then
            MsgBox "El campo solo permite valores numéricos"
            txtMontoPago.SetFocus
            Exit Sub
        End If
        txtMontoPago = Format(txtMontoPago, "Standard")
    End If
    Call sbCalcularSaldos
End Sub

Private Sub txtNotas_Change()
    vCambios = True
End Sub

Private Sub txtNotas_GotFocus()
    Call sbMarcarTXT(txtNotas)
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then OptComisionOperacion.SetFocus
End Sub

Private Sub txtOperacion_Change()
    vCambios = True
End Sub

Private Sub txtOperacion_GotFocus()
    Call sbMarcarTXT(txtOperacion)
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
End Sub

Private Sub txtPlazo_Change()
    vCambios = True
End Sub

Private Sub txtPlazo_GotFocus()
    Call sbMarcarTXT(txtPlazo)
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuota.SetFocus
End Sub

Private Sub txtPlazo_LostFocus()
    
    If txtPlazo.Text <> Empty Then
        If Not IsNumeric(txtPlazo) Then
            MsgBox "El campo solo permite valores numéricos"
            txtPlazo.SetFocus
        End If
    End If
    
End Sub

Private Sub txtResponsabilidad_Change()
    vCambios = True
End Sub

Private Sub txtResponsabilidad_GotFocus()
    Call sbMarcarTXT(txtResponsabilidad)
End Sub

Private Sub txtResponsabilidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpPrimerPago.SetFocus
End Sub

Private Sub txtResponsabilidad_LostFocus()
    If txtResponsabilidad.Text <> Empty Then
        If Not IsNumeric(txtResponsabilidad) Then
            MsgBox "El campo solo permite valores numéricos"
            txtResponsabilidad.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtSaldo_Change()
    vCambios = True
End Sub

Private Sub txtSaldo_GotFocus()
    Call sbMarcarTXT(txtSaldo)
End Sub

Private Sub txtSaldo_LostFocus()
    txtSaldo = Format(txtSaldo, "Standard")
End Sub



Private Sub txtTasa_Change()
    vCambios = True
End Sub



Private Sub txtTasa_GotFocus()
    Call sbMarcarTXT(txtTasa)
End Sub

Private Sub txtTasa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboMoneda.SetFocus
End Sub

Private Sub txtTasa_LostFocus()
    
    If txtTasa.Text <> Empty Then
        txtTasa = Format(txtTasa, "Standard")
        If Not IsNumeric(txtTasa) Then
            MsgBox "El campo solo permite valores numéricos"
            txtTasa.SetFocus
        End If
    End If
    
End Sub



Private Sub txtTipoCambio_Change()
  vCambios = True
End Sub

Private Sub txtTipoCambio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub vGridOperaciones_Click(ByVal Col As Long, ByVal Row As Long)
    Call sbGridMarcarSoloUno(vGridOperaciones, Row)
    lblOperacionActiva.Caption = fxGridValorMarcado(vGridOperaciones, 2)
End Sub

Private Sub sbLimpiarControlesOperaciones()
On Error GoTo error
    txtCod_Acreedor.Text = Empty
    txtDescripcionAcreedor.Text = Empty
    txtOperacion.Text = Empty
    txtResponsabilidad.Text = Empty
    txtNotas.Text = Empty
    txtMonto.Text = Empty
    txtTasa.Text = Empty
    txtPlazo.Text = Empty
    txtCuota.Text = Empty
    txtSaldo.Text = Empty
    dtpFormalizacion.Value = fxFechaServidor()
    dtpPrimerPago.Value = fxFechaServidor()
    lblFec_Actualizacion.Caption = Empty
    txtDia_Pago.Text = Empty
    lblEstado.Caption = Empty
    lblOperacionActiva.Caption = Empty
    txtComisionAdmin.Text = Empty
    cboTipo.Text = "Multiple"
    cboTipoPago.Text = "Mensual"
    OptComisionOperacion.Value = True
    OptResponOperacion.Value = True
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbLimpiarControlesPagos()

On Error GoTo error
    lblNPago.Caption = Empty
    txtMontoPago.Text = Empty
    dtpVencePago.Value = fxFechaServidor()
    dtpFechaPago.Value = dtpVencePago.Value
    txtInteresesPago.Text = Empty
    txtCargosPago.Text = Empty
    lblAmortizacionPago.Caption = Empty
    lblTasaPago.Caption = Empty
    lblSaldoPago.Caption = Empty
    lblTesoreriaSolicitud.Caption = Empty
    lblTesoreriaUsuario.Caption = Empty
    lblTesoreriaFecha.Caption = Empty
    lblUsuarioPago.Caption = Empty
    lblFechaPago.Caption = Empty
    lblEstadoPago.Caption = Empty
    txtDocumentoPago.Text = Empty
    cboFormaPago.Text = "Débito Cuenta"
    txtComisionPago.Text = Empty
    lblSaldoPago.Caption = Empty
    
    lblUltimaTasa.Caption = Empty
    lblUltimaCuota.Caption = Empty
    lblUltimoSaldo.Caption = Empty
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub



Private Sub sbBloquearControlesOperaciones(ByVal Modo As String)
On Error GoTo error

    txtCod_Acreedor.Locked = False
    txtOperacion.Locked = False
    txtResponsabilidad.Locked = False
    txtNotas.Locked = False
    txtMonto.Locked = False
    txtTasa.Locked = False
    txtPlazo.Locked = False
    txtCuota.Locked = False
    txtSaldo.Locked = False
    txtComisionAdmin.Locked = False
    dtpFormalizacion.Enabled = True
    dtpPrimerPago.Enabled = True
    txtDia_Pago.Locked = False
    cboTipo.Locked = False
    cboTipoPago.Locked = False
    tlbIncluirOperaciones.Enabled = True
    fraOptComision.Enabled = True
    fraOptResponsabilidad.Enabled = True
    Select Case UCase(Modo)
    Case "EDITARTODO"
        txtCod_Acreedor.Locked = True
        txtOperacion.Locked = True
        
    Case "EDITAR"
        txtCod_Acreedor.Locked = True
        txtOperacion.Locked = True
        txtMonto.Locked = True
        txtTasa.Locked = True
        txtPlazo.Locked = True
        txtCuota.Locked = True
        txtSaldo.Locked = True
        dtpFormalizacion.Enabled = False
        dtpPrimerPago.Enabled = False
        cboTipo.Locked = True
        cboTipoPago.Locked = True
      
    Case "VER"
        txtCod_Acreedor.Locked = True
        txtOperacion.Locked = True
        txtResponsabilidad.Locked = True
        txtNotas.Locked = True
        txtMonto.Locked = True
        txtTasa.Locked = True
        txtPlazo.Locked = True
        txtCuota.Locked = True
        txtSaldo.Locked = True
        dtpFormalizacion.Enabled = False
        dtpPrimerPago.Enabled = False
        txtDia_Pago.Locked = True
        cboTipo.Locked = True
        cboTipoPago.Locked = True
        txtComisionAdmin.Locked = True
        tlbIncluirOperaciones.Enabled = False
        fraOptComision.Enabled = False
        fraOptResponsabilidad.Enabled = False
    End Select
    Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbBloquearControlesPagos(ByVal Modo As String)
    fraPago.Enabled = True
    Select Case UCase(Modo)
    Case "NUEVO"
    Case "VER"
        fraPago.Enabled = False
    Case "DETALLE"
        fraPago.Enabled = False
    End Select
End Sub


Private Sub sbLimpiarFiltrosOperaciones()
    chkFiltroOperaciones.Value = vbUnchecked
    txtOperacionBusqueda.Text = Empty
    cboEstadoOpeBusq.Text = "Todos"
End Sub

Private Sub vGridPagos_Click(ByVal Col As Long, ByVal Row As Long)
    Call sbGridMarcarSoloUno(vGridPagos, Row)
End Sub

'Private Sub sbCargaBanco()
'    Dim strSQL As String
'    Dim rs As New ADODB.Recordset
'
'    strSQL = "Select ID_BANCO, DESCRIPCION" _
'            & " from TES_BANCOS"
'
'    Call OpenRecordSet(rs, strSQL)
'
'    Do While Not rs.EOF
'        cboBanco.ComboItems.Add , rs.Fields("ID_BANCO") & "(id)", UCase(Trim(rs.Fields("DESCRIPCION")))
'        rs.MoveNext
'    Loop
'
'    rs.Close
'End Sub



