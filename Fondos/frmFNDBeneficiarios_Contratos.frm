VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmFNDBeneficiarios_Contratos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beneficiarios Contratos"
   ClientHeight    =   7110
   ClientLeft      =   2280
   ClientTop       =   3645
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10080
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5652
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   9852
      _Version        =   1441792
      _ExtentX        =   17378
      _ExtentY        =   9970
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
      Color           =   2048
      ItemCount       =   2
      Item(0).Caption =   "Listado"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   28
      Item(1).Control(0)=   "txtObservacion"
      Item(1).Control(1)=   "txtApellido1"
      Item(1).Control(2)=   "txtApellido2"
      Item(1).Control(3)=   "txtNombre"
      Item(1).Control(4)=   "txtDireccion"
      Item(1).Control(5)=   "txtApartadoPostal"
      Item(1).Control(6)=   "txtTelefono2"
      Item(1).Control(7)=   "txtTelefono1"
      Item(1).Control(8)=   "txtPorcentaje"
      Item(1).Control(9)=   "txtCedula"
      Item(1).Control(10)=   "dtpFechaNacimiento"
      Item(1).Control(11)=   "Label2"
      Item(1).Control(12)=   "Label3"
      Item(1).Control(13)=   "Label4(1)"
      Item(1).Control(14)=   "Label7(1)"
      Item(1).Control(15)=   "Label16"
      Item(1).Control(16)=   "Label15(0)"
      Item(1).Control(17)=   "Label14"
      Item(1).Control(18)=   "Label8"
      Item(1).Control(19)=   "Label4(0)"
      Item(1).Control(20)=   "Lbl5"
      Item(1).Control(21)=   "Lbl4"
      Item(1).Control(22)=   "Lbl3(0)"
      Item(1).Control(23)=   "Lbl1"
      Item(1).Control(24)=   "cboParentesco"
      Item(1).Control(25)=   "Label15(1)"
      Item(1).Control(26)=   "txtEmail"
      Item(1).Control(27)=   "txtCodigo"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5052
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   9732
         _Version        =   1441792
         _ExtentX        =   17166
         _ExtentY        =   8911
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
      End
      Begin XtremeSuiteControls.ComboBox cboParentesco 
         Height          =   312
         Left            =   -68320
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441792
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
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   -68320
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441792
         _ExtentX        =   4043
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaNacimiento 
         Height          =   315
         Left            =   -65080
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   792
         Left            =   -68320
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   7452
         _Version        =   1441792
         _ExtentX        =   13144
         _ExtentY        =   1397
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   912
         Left            =   -68320
         TabIndex        =   11
         Top             =   4080
         Visible         =   0   'False
         Width           =   7452
         _Version        =   1441792
         _ExtentX        =   13144
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   -68320
         TabIndex        =   12
         Top             =   3720
         Visible         =   0   'False
         Width           =   7452
         _Version        =   1441792
         _ExtentX        =   13144
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApartadoPostal 
         Height          =   312
         Left            =   -62920
         TabIndex        =   13
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441792
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   312
         Left            =   -68320
         TabIndex        =   14
         Top             =   2400
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono2 
         Height          =   312
         Left            =   -65680
         TabIndex        =   15
         Top             =   2400
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   -63760
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441792
         _ExtentX        =   5101
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
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   312
         Left            =   -66040
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441792
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   312
         Left            =   -68320
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441792
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
         Height          =   312
         Left            =   -61720
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
         Text            =   "0"
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   -61720
         TabIndex        =   20
         Top             =   5160
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
         Transparent     =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Apellido 1"
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
         Left            =   -68320
         TabIndex        =   34
         Top             =   1020
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Apellido 2"
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
         Left            =   -66040
         TabIndex        =   33
         Top             =   1020
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Nombre"
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
         Index           =   1
         Left            =   -63760
         TabIndex        =   32
         Top             =   1020
         Visible         =   0   'False
         Width           =   2892
      End
      Begin VB.Label Label7 
         Caption         =   "Teléfono 1"
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
         Left            =   -69640
         TabIndex        =   31
         Top             =   2400
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label16 
         Caption         =   "Dirección"
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
         Left            =   -69640
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label15 
         Caption         =   "Email"
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
         Left            =   -69640
         TabIndex        =   29
         Top             =   3720
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label14 
         Caption         =   "Apto. Postal"
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
         Left            =   -64000
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label8 
         Caption         =   "Observación"
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
         Left            =   -69640
         TabIndex        =   27
         Top             =   4080
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfono 2"
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
         Left            =   -66640
         TabIndex        =   26
         Top             =   2400
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Lbl5 
         Caption         =   "Porcentaje de Beneficio"
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
         Left            =   -63640
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Lbl4 
         Caption         =   "Fec. Nac."
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
         Left            =   -65920
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Lbl3 
         Caption         =   "Parentesco"
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
         Left            =   -69640
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Lbl1 
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
         Height          =   252
         Left            =   -69640
         TabIndex        =   22
         Top             =   636
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Linea Id:"
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
         Left            =   -63160
         TabIndex        =   21
         Top             =   5160
         Visible         =   0   'False
         Width           =   1092
      End
   End
   Begin VB.Label lblOperadora 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Operadora:"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   6132
   End
   Begin VB.Label lblPlan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Plan:"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   3852
   End
   Begin VB.Label lblContrato 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contrato:"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   840
      Width           =   2292
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos del contrato:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1764
   End
End
Attribute VB_Name = "frmFNDBeneficiarios_Contratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vCedulaPrincipal As String
Dim vPaso As Boolean



Private Sub sbCargaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vPaso = True
lsw.ListItems.Clear

strSQL = "Select B.*, isnull(P.Descripcion,'') as 'Parentesco_Desc'" _
        & " from FND_CONTRATOS_BENEFICIARIOS B left join SYS_PARENTESCOS P on B.Parentesco = P.cod_Parentesco" _
        & " Where Cedula='" & vCedulaPrincipal _
        & "' and cod_Operadora = " & lblOperadora.Tag _
        & "  and cod_Plan = '" & lblPlan.Tag & "'" _
        & "  and cod_Contrato = " & lblContrato.Tag
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!consec)
    itmX.SubItems(1) = rs!cedulaBn
    itmX.SubItems(2) = rs!Nombre
    itmX.SubItems(3) = Format(rs!FechaNac, "dd/mm/yyyy")
    itmX.SubItems(4) = rs!parentesco_Desc
    itmX.SubItems(5) = rs!Porcentaje & " %"
   rs.MoveNext
Loop
     
rs.Close
vPaso = False

End Sub


Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18

On Error GoTo vError

vCedulaPrincipal = GLOBALES.gCedulaActual
lblOperadora.Caption = "Operadora:  " & fxOperadora
lblPlan.Caption = "Plan:  " & Trim(GLOBALES.gTag2)
lblContrato.Caption = "Contrato:  " & Trim(GLOBALES.gTag3)

lblOperadora.Tag = GLOBALES.gTag
lblContrato.Tag = GLOBALES.gTag3
lblPlan.Tag = GLOBALES.gTag2

 strSQL = "select rtrim(cod_Parentesco) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from sys_Parentescos where activo = 1"
 Call sbCbo_Llena_New(cboParentesco, strSQL, False, True)
 
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Identificación", 1400
    .Add , , "Nombre", 3400
    .Add , , "Fec.Nac.", 1400, vbCenter
    .Add , , "Parentesco", 1400
    .Add , , "Porcentaje", 1200, vbCenter
 End With


 vEdita = True
 Call sbToolBarIconos(tlbPrincipal)
 Call sbToolBar(tlbPrincipal, "nuevo")
 Call sbLimpiaPantalla(0)


 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla(Optional Index As Integer = 1)
Dim strSQL As String, rs As New ADODB.Recordset

If Trim(lblOperadora.Tag) = "" Then Exit Sub

tcMain.Item(Index).Selected = True

Select Case Index
    Case 0 'Lista
        Call sbCargaLsw
    Case 1 'Caso
        vCodigo = 0
        txtCodigo = ""
        
        strSQL = "select isnull(count(*),0) + 1 as Consec from FND_CONTRATOS_BENEFICIARIOS where cedula = '" & vCedulaPrincipal & "' and cod_plan = '" & lblPlan.Tag & "'" _
                 & " and cod_contrato = " & lblContrato.Tag & " and cod_operadora  = " & lblOperadora.Tag & ""
        Call OpenRecordSet(rs, strSQL)
          txtCedula = Trim(vCedulaPrincipal) & "-" & Format(rs!consec, "00")
        rs.Close
        
        txtApellido1 = ""
        txtApellido2 = ""
        txtNombre = ""
        
        dtpFechaNacimiento.MaxDate = fxFechaServidor
        dtpFechaNacimiento.Value = dtpFechaNacimiento.MaxDate
        
        txtPorcentaje.Text = "0"
        
        txtObservacion.Text = ""
        txtDireccion.Text = ""
        txtApartadoPostal.Text = ""
        txtEmail.Text = ""
        txtTelefono1.Text = ""
        txtTelefono2.Text = ""
        
End Select


End Sub


Private Sub Form_Unload(Cancel As Integer)
GLOBALES.gTag = ""
GLOBALES.gTag2 = ""
GLOBALES.gTag3 = ""
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
   
If vPaso Then Exit Sub

Call sbConsulta(Item.Text)

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 0 Then

    Call sbCargaLsw

End If
End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCedula.SetFocus
      Call sbToolBar(tlbPrincipal, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCedula.SetFocus
      Call sbToolBar(tlbPrincipal, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlbPrincipal, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlbPrincipal, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select consec,cedulaBN,nombre from FND_CONTRATOS_BENEFICIARIOS"
       gBusquedas.Filtro = " and cedula = '" & vCedulaPrincipal & "'"
       frmBusquedas.Show vbModal
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       Call sbConsulta(txtCodigo)
        
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vApellido1 As String, vApellido2 As String, vNombre1 As String, vNombre2 As String
Dim vEspacio As Integer, i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select B.*, isnull(P.Descripcion,'') as 'Parentesco_Desc', isnull(P.Cod_Parentesco,'') as 'cod_Parentesco'" _
        & " from FND_CONTRATOS_BENEFICIARIOS B left join SYS_PARENTESCOS P on B.Parentesco = P.cod_Parentesco" _
        & " Where consec = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlbPrincipal, "activo")
  
  tcMain.Item(1).Selected = True
  
  vEdita = True
  vCodigo = rs!consec
  txtCodigo.Text = CStr(rs!consec)
  
  
   vEspacio = 1
   For i = 1 To Len(Trim(rs!Nombre))
     If Mid(Trim(rs!Nombre), i, 1) <> " " Then
        Select Case vEspacio
         Case 1
          vApellido1 = vApellido1 & Mid(Trim(rs!Nombre), i, 1)
         Case 2
          vApellido2 = vApellido2 & Mid(Trim(rs!Nombre), i, 1)
         Case 3
          vNombre1 = vNombre1 & Mid(Trim(rs!Nombre), i, 1)
         Case Is >= 4
          vNombre2 = vNombre2 & Mid(Trim(rs!Nombre), i, 1)
        End Select
     Else
        vEspacio = vEspacio + 1
     End If
   Next i
   txtApellido1 = vApellido1
   txtApellido2 = vApellido2
   txtNombre = vNombre1 & " " & vNombre2
   
   txtPorcentaje = Format(rs!Porcentaje, "###.00")
   
   txtCedula = Trim(rs!cedulaBn)
   txtObservacion = Trim(rs!Notas & "")
       
    txtDireccion = Trim(rs!Direccion & "")
    txtApartadoPostal = Trim(rs!apto_postal & "")
    txtEmail = Trim(rs!Email & "")
    
    txtTelefono1 = Trim(rs!telefono1 & "")
    txtTelefono2 = Trim(rs!telefono2 & "")

    Call sbCboAsignaDato(cboParentesco, rs!parentesco_Desc, True, rs!cod_Parentesco)

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

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Verifica que exista ningun otro Beneficiario con la misma cedula juridica
strSQL = "select isnull(count(*),0) as 'Existe'" _
       & " from FND_CONTRATOS_BENEFICIARIOS" _
       & " where cedula = '" & vCedulaPrincipal _
       & "' and cedulaBN = '" & Trim(txtCedula) _
       & "' and consec <> " & vCodigo _
       & "  and cod_Operadora = " & lblOperadora.Tag _
       & "  and cod_Plan = '" & lblPlan.Tag & "'" _
       & "  and cod_Contrato = " & lblContrato.Tag
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   vMensaje = vMensaje & vbCrLf & " - Ya Existe ya un Beneficiario registrado con la mismo número de identificación ..."
End If
rs.Close

If cboParentesco.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se ha seleccionado ningún parentesco..."

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Beneficiario no es válido ..."
If txtApellido1 = "" Then vMensaje = vMensaje & vbCrLf & " - txtApellido 1 del Beneficiario no es válido ..."
If txtApellido2 = "" Then vMensaje = vMensaje & vbCrLf & " - txtApellido 2 del Beneficiario no es válido ..."

'Verificar que el porcentaje no supere el 100 %

If Not IsNumeric(txtPorcentaje) Then
   vMensaje = vMensaje & vbCrLf & " - El porcentaje no es válido ..."
Else
    strSQL = "select isnull(sum(porcentaje),0) as 'Porcentaje'" _
           & " from FND_CONTRATOS_BENEFICIARIOS" _
           & " where cedula = '" & vCedulaPrincipal & "' and consec <> " & vCodigo _
           & "  and cod_Operadora = " & lblOperadora.Tag _
           & "  and cod_Plan = '" & lblPlan.Tag & "'" _
           & "  and cod_Contrato = " & lblContrato.Tag

    Call OpenRecordSet(rs, strSQL)
    If CCur(txtPorcentaje) + rs!Porcentaje > 100 Then
       vMensaje = vMensaje & vbCrLf & " - El porcentaje sobre pasa el total del 100% del total de los beneficiarios ..."
    End If
    rs.Close
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update FND_CONTRATOS_BENEFICIARIOS set nombre = '" & UCase(Trim(txtApellido1)) & " " & UCase(Trim(txtApellido2)) & " " & UCase(Trim(txtNombre)) _
         & "',cedulaBN = '" & txtCedula & "', parentesco = '" & cboParentesco.ItemData(cboParentesco.ListIndex) _
         & "',notas = '" & txtObservacion & "',direccion = '" & txtDireccion & "',apto_postal = '" & txtApartadoPostal _
         & "',email = '" & txtEmail & "',telefono1 = '" & txtTelefono1 & "',telefono2 = '" & txtTelefono2 _
         & "',fechaNac = '" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") & "',porcentaje = " & CCur(txtPorcentaje) _
         & " where consec = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Beneficiario de Plan: Op." & lblOperadora.Tag & "..Pln:" & lblPlan.Tag & "..Cnt:" & lblContrato.Tag & "..Id: " & vCodigo)

Else
   strSQL = "insert FND_CONTRATOS_BENEFICIARIOS(cedula,cedulaBN,Nombre,parentesco,fechaNac,porcentaje,direccion,notas,telefono1" _
          & ",telefono2,email,apto_postal,cod_operadora,cod_plan,cod_contrato) values('" & vCedulaPrincipal & "','" & txtCedula _
          & "','" & UCase(Trim(txtApellido1)) & " " & UCase(Trim(txtApellido2)) & " " & UCase(Trim(txtNombre)) _
          & "','" & cboParentesco.ItemData(cboParentesco.ListIndex) & "','" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") _
          & "'," & CCur(txtPorcentaje) & ",'" & txtDireccion & "','" & txtObservacion & "','" & txtTelefono1 _
          & "','" & txtTelefono2 & "','" & txtEmail & "','" & txtApartadoPostal & "'," & lblOperadora.Tag & ",'" & lblPlan.Tag & "'," & lblContrato.Tag & ")"
   Call ConectionExecute(strSQL)
    
   strSQL = "select isnull(max(consec),0) as ultimo from FND_CONTRATOS_BENEFICIARIOS" _
          & " where cedula = '" & vCedulaPrincipal & "'"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo.Text = CStr(rs!ultimo)
     vCodigo = txtCodigo
   rs.Close
    
    
  Call Bitacora("Registra", "Beneficiario de Plan: Op." & lblOperadora.Tag & "..Pln:" & lblPlan.Tag & "..Cnt:" & lblContrato.Tag & "..Id: " & vCodigo)
  
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbToolBar(tlbPrincipal, "activo")
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete FND_CONTRATOS_BENEFICIARIOS where consec = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Beneficiario de Plan: Op." & lblOperadora.Tag & "..Pln:" & lblPlan.Tag & "..Cnt:" & lblContrato.Tag & "..Id: " & vCodigo)
  
  Call sbLimpiaPantalla
  Call sbToolBar(tlbPrincipal, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido1.SetFocus
End Sub

Private Sub txtApellido1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido2.SetFocus
End Sub

Private Sub txtApellido2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboParentesco.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select consec,cedulabn,nombre from FND_CONTRATOS_BENEFICIARIOS"
  gBusquedas.Filtro = " and cedula = '" & vCedulaPrincipal & "'"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub cboParentesco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFechaNacimiento.SetFocus
End Sub

Private Sub dtpFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPorcentaje.SetFocus
End Sub

Private Sub txtPorcentaje_GotFocus()
On Error GoTo vError
 txtPorcentaje = CCur(txtPorcentaje)
vError:
End Sub

Private Sub txtPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub

Private Sub txtPorcentaje_LostFocus()
On Error GoTo vError
 txtPorcentaje = Format(CCur(txtPorcentaje), "###.00")
vError:
End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartadoPostal.SetFocus
End Sub

Private Sub txtApartadoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub

Private Function fxOperadora() As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select descripcion from fnd_operadoras where cod_operadora = " & GLOBALES.gTag & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
   fxOperadora = rs!Descripcion
Else
   fxOperadora = ""
End If

End Function




