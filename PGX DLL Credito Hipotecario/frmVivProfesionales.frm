VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmVivProfesionales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información Profesionales"
   ClientHeight    =   7680
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10200
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6372
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   11239
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
      Item(0).Caption =   "Registro"
      Item(0).ControlCount=   31
      Item(0).Control(0)=   "Label3(0)"
      Item(0).Control(1)=   "Label3(1)"
      Item(0).Control(2)=   "Label3(2)"
      Item(0).Control(3)=   "Label3(3)"
      Item(0).Control(4)=   "Label3(4)"
      Item(0).Control(5)=   "Label3(5)"
      Item(0).Control(6)=   "Label3(6)"
      Item(0).Control(7)=   "Label3(7)"
      Item(0).Control(8)=   "Label3(8)"
      Item(0).Control(9)=   "Label3(9)"
      Item(0).Control(10)=   "Label3(10)"
      Item(0).Control(11)=   "txtTelefono1"
      Item(0).Control(12)=   "txtTelefonoExt"
      Item(0).Control(13)=   "txtFax"
      Item(0).Control(14)=   "txtFaxExt"
      Item(0).Control(15)=   "txtApartadoPostal"
      Item(0).Control(16)=   "txtEmail"
      Item(0).Control(17)=   "txtDireccion"
      Item(0).Control(18)=   "txtObservacion"
      Item(0).Control(19)=   "cboEstado"
      Item(0).Control(20)=   "lswCuentas"
      Item(0).Control(21)=   "btnCuentas"
      Item(0).Control(22)=   "cboBancos"
      Item(0).Control(23)=   "Label3(11)"
      Item(0).Control(24)=   "Label3(13)"
      Item(0).Control(25)=   "cboTipoPago"
      Item(0).Control(26)=   "chkDesembolsos"
      Item(0).Control(27)=   "cboTipoRol"
      Item(0).Control(28)=   "txtNombre"
      Item(0).Control(29)=   "cboTipoId"
      Item(0).Control(30)=   "txtIdentificacion"
      Item(1).Caption =   "Contactos"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "chkVinculado"
      Item(1).Control(2)=   "btnVincular"
      Item(1).Control(3)=   "Label3(14)"
      Item(1).Control(4)=   "txtEmpresaId"
      Item(1).Control(5)=   "txtEmpresaNombre"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4932
         Left            =   -69880
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1441793
         _ExtentX        =   17166
         _ExtentY        =   8700
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswCuentas 
         Height          =   1572
         Left            =   1320
         TabIndex        =   4
         Top             =   4596
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   2773
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
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
         Height          =   312
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   1692
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkDesembolsos 
         Height          =   252
         Left            =   1320
         TabIndex        =   6
         Top             =   3840
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica desembolsos?"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   312
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
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
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefonoExt 
         Height          =   312
         Left            =   2520
         TabIndex        =   8
         Top             =   1320
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFax 
         Height          =   312
         Left            =   4200
         TabIndex        =   9
         Top             =   1320
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFaxExt 
         Height          =   312
         Left            =   5400
         TabIndex        =   10
         Top             =   1320
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
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
      End
      Begin XtremeSuiteControls.FlatEdit txtApartadoPostal 
         Height          =   312
         Left            =   7920
         TabIndex        =   11
         Top             =   1320
         Width           =   1932
         _Version        =   1441793
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   1320
         TabIndex        =   12
         Top             =   1800
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
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
      End
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   672
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   1185
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
         Height          =   792
         Left            =   1320
         TabIndex        =   14
         Top             =   2880
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   7920
         TabIndex        =   15
         Top             =   480
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      End
      Begin XtremeSuiteControls.ComboBox cboTipoRol 
         Height          =   312
         Left            =   4200
         TabIndex        =   16
         Top             =   480
         Width           =   1692
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCuentas 
         Height          =   372
         Left            =   8160
         TabIndex        =   17
         Top             =   4200
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cuentas Bancarias"
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
      End
      Begin XtremeSuiteControls.ComboBox cboBancos 
         Height          =   312
         Left            =   3240
         TabIndex        =   18
         Top             =   4236
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8493
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
      Begin XtremeSuiteControls.ComboBox cboTipoPago 
         Height          =   312
         Left            =   8160
         TabIndex        =   19
         Top             =   3840
         Width           =   1692
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   4200
         TabIndex        =   34
         Top             =   960
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      End
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   312
         Left            =   1320
         TabIndex        =   35
         Top             =   480
         Width           =   1692
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkVinculado 
         Height          =   252
         Left            =   -69520
         TabIndex        =   38
         Top             =   5520
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Vinculado con Empresa?"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnVincular 
         Height          =   372
         Left            =   -61480
         TabIndex        =   39
         Top             =   5880
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Vincular"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtEmpresaNombre 
         Height          =   312
         Left            =   -67240
         TabIndex        =   43
         Top             =   5880
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEmpresaId 
         Height          =   312
         Left            =   -68200
         TabIndex        =   42
         Top             =   5880
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   14
         Left            =   -69760
         TabIndex        =   40
         Top             =   5880
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Empresa:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   960
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   2
         Left            =   3240
         TabIndex        =   30
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Nombre"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   3
         Left            =   6360
         TabIndex        =   29
         Top             =   480
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   4
         Left            =   3240
         TabIndex        =   28
         Top             =   480
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Rol"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   6
         Left            =   3240
         TabIndex        =   26
         Top             =   1320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fax"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   7
         Left            =   6360
         TabIndex        =   25
         Top             =   1320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Apto. Postal"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "E Mail"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Dirección"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Observacion"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   11
         Left            =   1320
         TabIndex        =   21
         Top             =   4200
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta/Desembolso"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   13
         Left            =   5040
         TabIndex        =   20
         Top             =   3840
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Emitir"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1680
      TabIndex        =   32
      Top             =   480
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   37
      Top             =   7428
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Usuario de Modificación"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha Modificacion"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnSuspender 
      Height          =   315
      Left            =   8880
      TabIndex        =   44
      Top             =   480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Suspender"
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
   End
   Begin XtremeSuiteControls.Label lblSuspendido 
      Height          =   252
      Left            =   3840
      TabIndex        =   41
      Top             =   480
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8700
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "..."
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   33
      Top             =   480
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Persona Id.:"
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
End
Attribute VB_Name = "frmVivProfesionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean, vPaso As Boolean

Private Sub btnSuspender_Click()

If vCodigo = 0 Then Exit Sub

Dim frm As Form


Call sbFormsCall("frmVivEstadoProfesionales", , , , False, Me)
Call sbFormActivo("frmVivEstadoProfesionales", frm)

frm.sbConsulta_Externa_IdContacto (vCodigo)


End Sub

Private Sub btnVincular_Click()
Dim strSQL As String

If vCodigo = 0 Then Exit Sub
If Not IsNumeric(txtEmpresaId.Text) Then Exit Sub

On Error GoTo vError

If chkVinculado.Value = xtpChecked Then
        strSQL = "update ViviendaContactos set idEmpresa = " & txtEmpresaId.Text _
               & " where idContacto = " & vCodigo
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Hipotecario, Persona [Id:" & vCodigo & "] vinculado con [Id:" & txtEmpresaId.Text & "]")
        
        MsgBox "Persona vinculada con la empresa: " & txtEmpresaNombre.Text, vbInformation
Else
        strSQL = "update ViviendaContactos set idEmpresa = Null" _
               & " where idContacto = " & vCodigo
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Hipotecario, Persona [Id:" & vCodigo & "] desvinculado de alguna empresa!")
        
        MsgBox "Persona desvinculada de cualquier empresa!", vbInformation
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboBancos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And btnCuentas.Enabled Then btnCuentas.SetFocus
End Sub

Private Sub sbCuentas_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lswCuentas.ListItems.Clear
If vCodigo > 0 Then
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtIdentificacion.Text) & "'"
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswCuentas.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!Activa = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           itmX.SubItems(8) = rs!Registro_Usuario & ""
     
       rs.MoveNext
    Loop
    rs.Close
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCuentas_Click()


GLOBALES.gTag = Trim(txtIdentificacion.Text)
GLOBALES.gTag2 = "AFI"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdentificacion.SetFocus
End Sub


Private Sub cboTipoId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoRol.SetFocus

End Sub

Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBancos.SetFocus
End Sub

Private Sub chkVinculado_Click()

If chkVinculado.Value = xtpChecked Then
    btnVincular.Caption = "Vincular"
    txtEmpresaId.Enabled = True
    txtEmpresaNombre.Enabled = True
    If tcMain.SelectedItem = 1 Then
        txtEmpresaId.SetFocus
    End If
Else
    btnVincular.Caption = "Desvincular"
    txtEmpresaId.Enabled = False
    txtEmpresaNombre.Enabled = False
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo = "" Or Not IsNumeric(txtCodigo) Then txtCodigo.Text = "0"

If vScroll Then
    strSQL = "select Top 1 idContacto from ViviendaContactos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where idContacto > " & txtCodigo & " order by idContacto asc"
    Else
       strSQL = strSQL & " where idContacto < " & txtCodigo & " order by idContacto desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!IdContacto)
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
Dim strSQL As String

vModulo = 3

On Error GoTo vError
 
vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
 
lswCuentas.ColumnHeaders.Add 1, , "Cuenta", 2500
lswCuentas.ColumnHeaders.Add 2, , "Banco", 3500
lswCuentas.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 8, , "Fecha", 2500
lswCuentas.ColumnHeaders.Add 9, , "Usuario", 2500

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Persona Id", 1200
lsw.ColumnHeaders.Add , , "Identificación", 1800, vbCenter
lsw.ColumnHeaders.Add , , "Nombre", 4000
lsw.ColumnHeaders.Add , , "Estado", 1500, vbCenter
lsw.ColumnHeaders.Add , , "Tipo", 1500, vbCenter
lsw.ColumnHeaders.Add , , "Desembolsos?", 1500, , vbCenter
lsw.ColumnHeaders.Add , , "Suspendido?", 1300, vbCenter



strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)
 
 
cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"
 
cboTipoId.Clear
cboTipoId.AddItem "Física"
cboTipoId.AddItem "Jurídica"
cboTipoId.AddItem "Física"
 
cboTipoRol.Clear
cboTipoRol.AddItem "Abogado"
cboTipoRol.AddItem "Ingeniero"
cboTipoRol.AddItem "Contacto"
cboTipoRol.AddItem "Contacto"



cboTipoPago.Clear
cboTipoPago.AddItem fxTipoDocumento("CK")
cboTipoPago.AddItem fxTipoDocumento("TE")

 
 vEdita = True
 
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String, rs As New ADODB.Recordset


tcMain.Item(0).Selected = True

vCodigo = 0
txtCodigo = ""

lblSuspendido.Caption = ""

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""

cboEstado.Text = "Activo"
cboTipoId.Text = "Física"
cboTipoRol.Text = "Contacto"

chkDesembolsos.Value = vbUnchecked

cboTipoPago.Text = fxTipoDocumento("TE")

txtNombre.Text = ""
txtObservacion.Text = ""

txtIdentificacion.Text = ""

txtDireccion.Text = ""
txtApartadoPostal.Text = ""
txtEmail.Text = ""

txtTelefono1.Text = ""
txtTelefonoExt.Text = ""
txtFax.Text = ""
txtFaxExt.Text = ""

lswCuentas.ListItems.Clear

chkVinculado.Value = xtpUnchecked
txtEmpresaId.Text = ""
txtEmpresaNombre.Text = ""

txtCodigo.Enabled = True

End Sub

Public Sub sbConsulta_Externa_IdPersona(pIdentificacion As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim pContacto As Long

strSQL = "select idContacto from ViviendaContactos where Identificacion = '" & pIdentificacion & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
 pContacto = 0
Else
 pContacto = rs!IdContacto
End If
rs.Close

If pContacto > 0 Then
    Call sbConsulta(pContacto)
End If

End Sub


Public Sub sbConsulta_Externa_IdContacto(pContacto As Long)

Call sbConsulta(pContacto)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count > 0 Then
    Call sbConsulta(Item.Text)
End If
End Sub

Private Sub sbListado_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

    lsw.ListItems.Clear



    strSQL = "Select IdContacto, Identificacion, Nombre,PagaHonorarios" _
           & ", Case when Estado = 'A' then 'Activo' when estado = 'I' then 'Inactivo' end as 'Estado'" _
           & ", case TipoProfesional when 'I' then 'Ingeniero' when 'A' then 'Abodado' else 'Contacto' end as 'Tipo'" _
           & ", dbo.fxCrd_Viv_Profesional_Suspendido(IdContacto) as 'Suspendido'" _
           & " from ViviendaContactos where idEmpresa = " & vCodigo _
           & " order by nombre"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
        Set itmX = lsw.ListItems.Add(, , rs!IdContacto)
            itmX.SubItems(1) = rs!Identificacion & ""
            itmX.SubItems(2) = rs!Nombre
            itmX.SubItems(3) = rs!Estado
            itmX.SubItems(4) = rs!Tipo
            itmX.SubItems(5) = IIf(rs!PagaHonorarios = 1, "Sí", "No")
            itmX.SubItems(6) = IIf(rs!Suspendido = 1, "Sí", "No")
        rs.MoveNext
     Loop
     rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
    Call sbListado_Load
End If

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      cboTipoId.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
         gBusquedas.Columna = "nombre"
         gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select IdConctacto,Identificacion,nombre from ViviendaContactos"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(pPersonaId As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select P.*,B.descripcion as 'Banco', dbo.fxCrd_Viv_Profesional_Suspendido(P.IdContacto) as 'SuspendeActual'" _
       & ", isnull(E.IdContacto,0) as 'EmpresaId', isnull(E.Nombre,'') as 'EmpresaNombre'" _
       & " from ViviendaContactos P inner join Tes_Bancos B on P.cod_banco = B.id_banco" _
       & " left join ViviendaContactos E on P.idEmpresa = E.idContacto" _
       & " where P.idContacto = " & pPersonaId
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  tcMain.Item(0).Selected = True
  
  vEdita = True
  
  vCodigo = rs!IdContacto
  txtCodigo.Text = CStr(rs!IdContacto)
  
  If rs!SuspendeActual = 1 Then
      lblSuspendido.Caption = "Suspendido desde: " & Format(rs!SuspensionInicio, "dd/MM/yyyy")
  Else
      lblSuspendido.Caption = ""
  End If
  
    txtNombre.Text = rs!Nombre & ""
    txtObservacion.Text = rs!observacion & ""
    
    txtIdentificacion.Text = rs!Identificacion & ""
    
    If rs!Estado = "A" Then
      cboEstado.Text = "Activo"
    Else
      cboEstado.Text = "Inactivo"
    End If
    
    cboTipoPago.Text = fxTipoDocumento(rs!Emite)
    
    Select Case rs!TipoContacto 'Id
      Case "F"
        cboTipoId.Text = "Física"
      Case "J"
        cboTipoId.Text = "Jurídica"
      Case Else
        cboTipoId.Text = "Física"
    End Select
    
    
    Select Case rs!TipoProfesional
      Case "I"
        cboTipoRol.Text = "Ingeniero"
      Case "A"
        cboTipoRol.Text = "Abogado"
      Case "C"
        cboTipoRol.Text = "Contacto"
      Case Else
        cboTipoRol.Text = "Contacto"
    End Select
    
    chkDesembolsos.Value = rs!PagaHonorarios
    
    txtDireccion.Text = rs!Direccion & ""
    txtApartadoPostal.Text = rs!APTO_POSTAL & ""
    txtEmail.Text = rs!Email & ""
    
    txtTelefono1.Text = rs!telefono & ""
    txtTelefonoExt.Text = rs!telefono_ext & ""
    txtFax.Text = rs!fax & ""
    txtFaxExt.Text = rs!fax_ext & ""
    
    Call sbCboAsignaDato(cboBancos, Trim(rs!Banco), True, rs!cod_banco)

    If rs!EmpresaId > 0 Then
        chkVinculado.Value = xtpChecked
        txtEmpresaId.Text = rs!EmpresaId
        txtEmpresaNombre.Text = rs!EmpresaNombre
    Else
        chkVinculado.Value = xtpUnchecked
        txtEmpresaId.Text = ""
        txtEmpresaNombre.Text = ""
    End If
    Call chkVinculado_Click
    
   StatusBarX.Panels(1).Text = rs!RegistroUsuario & ""
   StatusBarX.Panels(2).Text = rs!RegistroFecha & ""

   Call sbCuentas_Load
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

If cboBancos.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se especificó una Cuenta Bancaria para Desembolsos ..."

If Mid(cboTipoPago.Text, 1, 1) = "T" _
   And lswCuentas.ListItems.Count = 0 Then vMensaje = vMensaje & vbCrLf & " - No se especificó la cuenta para las transferencias..."

If txtNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del contacto no es válido ..."
If txtIdentificacion.Text = "" Then vMensaje = vMensaje & vbCrLf & " - El número de identificación de la persona no es válido..."


strSQL = "select count(*) as 'Existe' from ViviendaContactos where idContacto <> " & vCodigo _
       & " and identificacion = '" & txtIdentificacion.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
    vMensaje = vMensaje & vbCrLf & " - El número de identificación de la persona ya esta siendo utilizado por otro contacto!"
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
  strSQL = "update ViviendaContactos set nombre = '" & Trim(txtNombre.Text) _
         & "',TipoContacto = '" & Mid(cboTipoId.Text, 1, 1) & "',Identificacion = '" & txtIdentificacion.Text _
         & "',Observacion = '" & txtObservacion.Text & "',Estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "',Emite = '" & fxTipoDocumento(cboTipoPago.Text) _
         & "',Direccion = '" & txtDireccion.Text & "',APTO_POSTAL = '" & txtApartadoPostal.Text _
         & "',Email = '" & txtEmail.Text & "',Telefono = '" & txtTelefono1.Text _
         & "',Telefono_ext = '" & txtTelefonoExt.Text & "',fax = '" & txtFax.Text _
         & "',Fax_ext = '" & txtFaxExt.Text & "',cod_banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & ",PagaHonorarios = " & chkDesembolsos.Value & ",TipoProfesional = '" & Mid(cboTipoRol.Text, 1, 1) _
         & "' where idContacto = " & vCodigo
         
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Credito Hipotecario, Contacto Id: " & vCodigo)

Else
   
   
   strSQL = "select isnull(max(idContacto),0) as ultimo from ViviendaContactos"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo.Text = CStr(rs!Ultimo + 1)
     vCodigo = txtCodigo.Text
   rs.Close
       
   strSQL = "insert ViviendaContactos(IdContacto, IdEmpresa, TipoContacto, Identificacion, Nombre, TipoProfesional, Telefono, Telefono_Ext" _
          & ", Fax, Fax_Ext, Email, Direccion, Apto_Postal, Estado, PagaHonorarios, Emite, Cod_Banco , RegistroFecha, RegistroUsuario)" _
          & " values(" & vCodigo & ",Null,'" & Mid(cboTipoId.Text, 1, 1) & "','" & txtIdentificacion.Text & "','" & txtNombre.Text _
          & "','" & Mid(cboTipoRol.Text, 1, 1) & "','" & txtTelefono1.Text & "','" & txtTelefonoExt.Text & "','" & txtFax.Text _
          & "','" & txtFaxExt.Text & "','" & txtEmail.Text & "','" & txtDireccion.Text & "','" & txtApartadoPostal.Text _
          & "','" & Mid(cboEstado.Text, 1, 1) & "'," & chkDesembolsos.Value & ",'" & fxTipoDocumento(cboTipoPago.Text) & "'," & cboBancos.ItemData(cboBancos.ListIndex) _
          & ", dbo.mygetdate(),'" & glogon.Usuario & "')"
 
   Call ConectionExecute(strSQL)
    

   Call Bitacora("Registra", "Credito Hipotecario, Contacto Id: " & vCodigo)
    
   txtCodigo.Enabled = True
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbToolBar(tlb, "activo")
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
  strSQL = "delete ViviendaContactos where idContacto = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Credito Hipotecario, Contacto Id: " & vCodigo)
  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
Private Sub txtApartadoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub


Private Sub txtEmpresaId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Identificación"
  gBusquedas.Col2Name = "Id Persona"
  gBusquedas.Columna = "Identificacion"
  gBusquedas.Orden = "Identificacion"
  gBusquedas.Consulta = "select Identificacion,idContacto,nombre from ViviendaContactos"
  gBusquedas.Filtro = " and TipoContacto = 'J' and IdContacto <> " & vCodigo
  frmBusquedas.Show vbModal
  txtEmpresaId.Text = gBusquedas.Resultado2
  txtEmpresaNombre.Text = gBusquedas.Resultado3
End If
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Identificación"
  gBusquedas.Col2Name = "Id Persona"
  gBusquedas.Columna = "Identificacion"
  gBusquedas.Orden = "Identificacion"
  gBusquedas.Consulta = "select Identificacion,idContacto,nombre from ViviendaContactos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado2
  If txtCodigo.Text <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado2))
End If

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    tcMain.Item(0).Selected = True
    If vEdita Then
      txtNombre.SetFocus
    Else
      cboTipoId.SetFocus
    End If
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Persona Id"
  gBusquedas.Col2Name = "Identificación"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "idContacto"
  gBusquedas.Orden = "idContacto"
  gBusquedas.Consulta = "select idContacto,Identificacion,Nombre from ViviendaContactos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub


Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub


Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFaxExt.SetFocus
End Sub

Private Sub txtFaxExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartadoPostal.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Identificación"
  gBusquedas.Col2Name = "Id Persona"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select Identificacion,idContacto,nombre from ViviendaContactos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado2
  If txtCodigo.Text <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado2))
End If

End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoPago.SetFocus
End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefonoExt.SetFocus
End Sub

Private Sub txtTelefonoExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFax.SetFocus
End Sub



