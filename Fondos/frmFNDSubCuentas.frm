VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDSubCuentas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SubCuentas"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
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
      TabIndex        =   1
      Top             =   1320
      Width           =   9852
      _Version        =   1572864
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
      Color           =   32
      ItemCount       =   2
      SelectedItem    =   1
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
      Item(1).Control(8)=   "txtCedula"
      Item(1).Control(9)=   "dtpFechaNacimiento"
      Item(1).Control(10)=   "Label2"
      Item(1).Control(11)=   "Label3"
      Item(1).Control(12)=   "Label4(1)"
      Item(1).Control(13)=   "Label7(1)"
      Item(1).Control(14)=   "Label16"
      Item(1).Control(15)=   "Label15(0)"
      Item(1).Control(16)=   "Label14"
      Item(1).Control(17)=   "Label8"
      Item(1).Control(18)=   "Label4(0)"
      Item(1).Control(19)=   "Lbl5"
      Item(1).Control(20)=   "Lbl4"
      Item(1).Control(21)=   "Lbl3(0)"
      Item(1).Control(22)=   "Lbl1"
      Item(1).Control(23)=   "cboParentesco"
      Item(1).Control(24)=   "Label15(1)"
      Item(1).Control(25)=   "txtEmail"
      Item(1).Control(26)=   "txtCodigo"
      Item(1).Control(27)=   "txtCuota"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5052
         Left            =   -69880
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1572864
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
         Left            =   1680
         TabIndex        =   3
         Top             =   1800
         Width           =   1572
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
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   2292
         _Version        =   1572864
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
         Height          =   312
         Left            =   4320
         TabIndex        =   5
         Top             =   1800
         Width           =   1332
         _Version        =   1572864
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
         Left            =   1680
         TabIndex        =   6
         Top             =   2880
         Width           =   7452
         _Version        =   1572864
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
         Left            =   1680
         TabIndex        =   7
         Top             =   4080
         Width           =   7452
         _Version        =   1572864
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
         Left            =   1680
         TabIndex        =   8
         Top             =   3720
         Width           =   7452
         _Version        =   1572864
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
         Left            =   7080
         TabIndex        =   9
         Top             =   2400
         Width           =   2052
         _Version        =   1572864
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
         Left            =   1680
         TabIndex        =   10
         Top             =   2400
         Width           =   1572
         _Version        =   1572864
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
         Left            =   4320
         TabIndex        =   11
         Top             =   2400
         Width           =   1572
         _Version        =   1572864
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
         Left            =   6240
         TabIndex        =   12
         Top             =   1320
         Width           =   2892
         _Version        =   1572864
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
         Left            =   3960
         TabIndex        =   13
         Top             =   1320
         Width           =   2292
         _Version        =   1572864
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
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   2292
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   8280
         TabIndex        =   16
         Top             =   5160
         Width           =   852
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtCuota 
         Height          =   312
         Left            =   7200
         TabIndex        =   15
         Top             =   1800
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
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
         Left            =   6840
         TabIndex        =   30
         Top             =   5160
         Width           =   1092
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
         Left            =   360
         TabIndex        =   29
         Top             =   636
         Width           =   1212
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
         Left            =   360
         TabIndex        =   28
         Top             =   1800
         Width           =   972
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
         Left            =   3360
         TabIndex        =   27
         Top             =   1800
         Width           =   852
      End
      Begin VB.Label Lbl5 
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
         Left            =   6000
         TabIndex        =   26
         Top             =   1800
         Width           =   1212
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
         Left            =   3360
         TabIndex        =   25
         Top             =   2400
         Width           =   972
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
         Left            =   360
         TabIndex        =   24
         Top             =   4080
         Width           =   1452
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
         Left            =   6000
         TabIndex        =   23
         Top             =   2400
         Width           =   1092
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
         Left            =   360
         TabIndex        =   22
         Top             =   3720
         Width           =   1092
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
         Left            =   360
         TabIndex        =   21
         Top             =   2760
         Width           =   1332
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
         Left            =   360
         TabIndex        =   20
         Top             =   2400
         Width           =   972
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
         Left            =   6240
         TabIndex        =   19
         Top             =   1020
         Width           =   2892
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
         Left            =   3960
         TabIndex        =   18
         Top             =   1020
         Width           =   2412
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
         Left            =   1680
         TabIndex        =   17
         Top             =   1020
         Width           =   2292
      End
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
      TabIndex        =   34
      Top             =   480
      Width           =   1764
   End
   Begin VB.Label lblContrato 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
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
      TabIndex        =   33
      Top             =   840
      Width           =   3132
   End
   Begin VB.Label lblPlan 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
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
      TabIndex        =   32
      Top             =   840
      Width           =   4692
   End
   Begin VB.Label lblOperadora 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
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
      TabIndex        =   31
      Top             =   480
      Width           =   6972
   End
End
Attribute VB_Name = "frmFNDSubCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vCedulaPrincipal As String
Dim vPaso As Boolean

Private Sub sbActualizaCuotaContrato()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spFnd_SubCuentas_Maestro_Update " & gFondos.Operadora _
       & ", '" & gFondos.Plan & "', " & gFondos.Contrato & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vPaso = True
lsw.ListItems.Clear

strSQL = "Select B.*, isnull(P.Descripcion,'') as 'Parentesco_Desc'" _
        & " from FND_SubCUENTAS B left join SYS_PARENTESCOS P on B.Parentesco = P.cod_Parentesco" _
        & " Where " _
        & "  cod_Operadora = " & lblOperadora.Tag _
        & "  and cod_Plan = '" & lblPlan.Tag & "'" _
        & "  and cod_Contrato = " & lblContrato.Tag
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!IdX)
    itmX.SubItems(1) = rs!Cedula
    itmX.SubItems(2) = rs!Nombre
    itmX.SubItems(3) = Format(rs!FechaNac, "dd/mm/yyyy")
    itmX.SubItems(4) = rs!Parentesco_Desc
    itmX.SubItems(5) = Format(rs!Cuota, "Standard")
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

vCedulaPrincipal = gFondos.Cedula
lblPlan.Caption = "Plan:  " & gFondos.Plan
lblContrato.Caption = "Contrato:  " & gFondos.Contrato

lblOperadora.Tag = gFondos.Operadora
lblContrato.Tag = gFondos.Contrato
lblPlan.Tag = gFondos.Plan

lblOperadora.Caption = "Operadora:  " & fxOperadora

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
    .Add , , "Mensualidad", 1200, vbRightJustify
 End With


 vEdita = True
 Call sbToolBarIconos(tlbPrincipal)
 Call sbToolBar(tlbPrincipal, "nuevo")
 Call sbLimpiaPantalla(0)


 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 
 If gFondos.SubCuenta > 0 Then
    Call sbConsulta(CLng(gFondos.SubCuenta))
 End If
 
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
        
        strSQL = "select isnull(count(*),0) + 1 as IDX from FND_SubCUENTAS where cod_plan = '" & lblPlan.Tag & "'" _
                 & " and cod_contrato = " & lblContrato.Tag & " and cod_operadora  = " & lblOperadora.Tag & ""
        Call OpenRecordSet(rs, strSQL)
          txtCedula = Trim(vCedulaPrincipal) & "-" & Format(rs!IdX, "00")
        rs.Close
        
        txtApellido1 = ""
        txtApellido2 = ""
        txtNombre = ""
        
        dtpFechaNacimiento.MaxDate = fxFechaServidor
        dtpFechaNacimiento.Value = dtpFechaNacimiento.MaxDate
        
        txtCuota.Text = "0"
        
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
       gBusquedas.Consulta = "select IDX,CEDULA,nombre from FND_SubCUENTAS"
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
        & " from FND_SubCUENTAS B left join SYS_PARENTESCOS P on B.Parentesco = P.cod_Parentesco" _
        & " Where IDX = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlbPrincipal, "activo")
  
  tcMain.Item(1).Selected = True
  
  vEdita = True
  vCodigo = rs!IdX
  txtCodigo.Text = CStr(rs!IdX)
  
  
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
   
   txtCuota = Format(rs!Cuota, "Standard")
   
   txtCedula = Trim(rs!Cedula)
   txtObservacion = Trim(rs!Notas & "")
       
    txtDireccion = Trim(rs!direccion & "")
    txtApartadoPostal = Trim(rs!apto_postal & "")
    txtEmail = Trim(rs!Email & "")
    
    txtTelefono1 = Trim(rs!telefono1 & "")
    txtTelefono2 = Trim(rs!telefono2 & "")

    Call sbCboAsignaDato(cboParentesco, rs!Parentesco_Desc, True, rs!Cod_Parentesco)

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
       & " from FND_SubCUENTAS" _
       & " where cedula = '" & vCedulaPrincipal _
       & "' and CEDULA = '" & Trim(txtCedula) _
       & "' and IDX <> " & vCodigo _
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

'If Not IsNumeric(txtCuota) Then
'   vMensaje = vMensaje & vbCrLf & " - El porcentaje no es válido ..."
'Else
'    strSQL = "select isnull(sum(porcentaje),0) as 'Porcentaje'" _
'           & " from FND_SubCUENTAS" _
'           & " where cedula = '" & vCedulaPrincipal & "' and IDX <> " & vCodigo _
'           & "  and cod_Operadora = " & lblOperadora.Tag _
'           & "  and cod_Plan = '" & lblPlan.Tag & "'" _
'           & "  and cod_Contrato = " & lblContrato.Tag
'
'    Call OpenRecordSet(rs, strSQL)
'    If CCur(txtCuota) + rs!Porcentaje > 100 Then
'       vMensaje = vMensaje & vbCrLf & " - El porcentaje sobre pasa el total del 100% del total de los beneficiarios ..."
'    End If
'    rs.Close
'End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update FND_SubCUENTAS set nombre = '" & UCase(Trim(txtApellido1)) & " " & UCase(Trim(txtApellido2)) & " " & UCase(Trim(txtNombre)) _
         & "',CEDULA = '" & txtCedula & "', parentesco = '" & cboParentesco.ItemData(cboParentesco.ListIndex) _
         & "',notas = '" & txtObservacion & "',direccion = '" & txtDireccion & "',apto_postal = '" & txtApartadoPostal _
         & "',email = '" & txtEmail & "',telefono1 = '" & txtTelefono1 & "',telefono2 = '" & txtTelefono2 _
         & "',fechaNac = '" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") & "',cuota = " & CCur(txtCuota) _
         & " where IDX = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Sub-Cuenta de Plan: Op." & lblOperadora.Tag & "..Pln:" & lblPlan.Tag & "..Cnt:" & lblContrato.Tag & "..Id: " & vCodigo)

Else
   strSQL = "select isnull(max(IDX),0) + 1 as ultimo from FND_SubCUENTAS" _
          & " where cod_operadora = " & lblOperadora.Tag & " and cod_plan = '" & lblPlan.Tag & "' and cod_Contrato = " & lblContrato.Tag
   Call OpenRecordSet(rs, strSQL)
     txtCodigo.Text = CStr(rs!ultimo)
     vCodigo = txtCodigo
   rs.Close
       
   
   strSQL = "insert FND_SubCUENTAS(IdX,cedula,Nombre,parentesco,fechaNac,cuota, APORTES, RENDIMIENTO, direccion,notas,telefono1" _
          & ",telefono2,email,apto_postal,cod_operadora,cod_plan,cod_contrato,cod_beneficiario,estado)" _
          & " values(" & vCodigo & ",'" & txtCedula.Text _
          & "','" & UCase(Trim(txtApellido1)) & " " & UCase(Trim(txtApellido2)) & " " & UCase(Trim(txtNombre)) _
          & "','" & cboParentesco.ItemData(cboParentesco.ListIndex) & "','" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") _
          & "'," & CCur(txtCuota) & ", 0, 0, '" & txtDireccion & "','" & txtObservacion & "','" & txtTelefono1 _
          & "','" & txtTelefono2 & "','" & txtEmail & "','" & txtApartadoPostal _
          & "'," & lblOperadora.Tag & ",'" & lblPlan.Tag & "'," & lblContrato.Tag & ",0,'A')"
   Call ConectionExecute(strSQL)
    

    
  Call Bitacora("Registra", "Sub-Cuenta de Plan: Op." & lblOperadora.Tag & "..Pln:" & lblPlan.Tag & "..Cnt:" & lblContrato.Tag & "..Id: " & vCodigo)
  
End If

'Actualiza Contrato
Call sbActualizaCuotaContrato

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
  strSQL = "delete FND_SubCUENTAS where IDX = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Sub-Cuenta de Plan: Op." & lblOperadora.Tag & "..Pln:" & lblPlan.Tag & "..Cnt:" & lblContrato.Tag & "..Id: " & vCodigo)
  
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
  gBusquedas.Consulta = "select IDX,CEDULA,nombre from FND_SubCUENTAS"
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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuota.SetFocus
End Sub

Private Sub txtCuota_GotFocus()
On Error GoTo vError
 txtCuota = CCur(txtCuota)
vError:
End Sub

Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub

Private Sub txtCuota_LostFocus()
On Error GoTo vError
 txtCuota = Format(CCur(txtCuota), "Standard")
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

strSQL = "select descripcion from fnd_operadoras where cod_operadora = " & gFondos.Operadora & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
   fxOperadora = rs!Descripcion
Else
   fxOperadora = ""
End If

End Function
