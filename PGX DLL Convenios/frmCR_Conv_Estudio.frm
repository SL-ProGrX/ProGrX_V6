VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Begin VB.Form frmCR_Conv_Estudio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convenios de Estudios"
   ClientHeight    =   7956
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9864
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7956
   ScaleWidth      =   9864
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      MaxLength       =   38
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      DataField       =   "e"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   9615
      _ExtentX        =   16955
      _ExtentY        =   12298
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmCR_Conv_Estudio.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblBtnColor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(16)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(15)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label18(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line2(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(7)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(10)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line2(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label12(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label12(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label12(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cboDistrito"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cboCanton"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtTelFax"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtWebSite"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtEMail2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtRazon"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtTelefono"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtCelular"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtTelefono2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtDireccion"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtAptoPostal"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtEMail"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtNotas"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cboProvincia"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtCedJur"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtContacto_01"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtContacto_02"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtCuentaDepositos"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cboTipoPago"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cboBanco"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "chkActiva"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "Carreras"
      TabPicture(1)   =   "frmCR_Conv_Estudio.frx":0121
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lswCarreras"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contratos"
      TabPicture(2)   =   "frmCR_Conv_Estudio.frx":01DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CheckBox chkActiva 
         Alignment       =   1  'Right Justify
         Caption         =   "Activa ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   50
         Top             =   820
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.ComboBox cboBanco 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCR_Conv_Estudio.frx":02F7
         Left            =   2760
         List            =   "frmCR_Conv_Estudio.frx":02F9
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   4560
         Width           =   6615
      End
      Begin VB.ComboBox cboTipoPago 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCR_Conv_Estudio.frx":02FB
         Left            =   2760
         List            =   "frmCR_Conv_Estudio.frx":0305
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox txtCuentaDepositos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         MaxLength       =   30
         TabIndex        =   44
         ToolTipText     =   "Ultima Compra realizada con este proveedor"
         Top             =   4920
         Width           =   3015
      End
      Begin MSComctlLib.ListView lswCarreras 
         Height          =   6135
         Left            =   -74640
         TabIndex        =   42
         Top             =   480
         Width           =   9015
         _ExtentX        =   15896
         _ExtentY        =   10816
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Carrera"
            Object.Width           =   9596
         EndProperty
      End
      Begin VB.TextBox txtContacto_02 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   41
         Top             =   5940
         Width           =   8055
      End
      Begin VB.TextBox txtContacto_01 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   40
         Top             =   5580
         Width           =   8055
      End
      Begin VB.TextBox txtCedJur 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   39
         Top             =   420
         Width           =   2055
      End
      Begin VB.ComboBox cboProvincia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCR_Conv_Estudio.frx":0321
         Left            =   1440
         List            =   "frmCR_Conv_Estudio.frx":033A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Provincia"
         Top             =   3060
         Width           =   1935
      End
      Begin VB.TextBox txtNotas 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1440
         TabIndex        =   15
         ToolTipText     =   "Notas sobre la persona"
         Top             =   6300
         Width           =   8055
      End
      Begin VB.TextBox txtEMail 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   1500
         Width           =   4935
      End
      Begin VB.TextBox txtAptoPostal 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   13
         Top             =   2220
         Width           =   4935
      End
      Begin VB.TextBox txtDireccion 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Dirección Exacta"
         Top             =   3060
         Width           =   6015
      End
      Begin VB.TextBox txtTelefono2 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   1500
         Width           =   2055
      End
      Begin VB.TextBox txtCelular 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   1860
         Width           =   2055
      End
      Begin VB.TextBox txtTelefono 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox txtRazon 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4560
         TabIndex        =   8
         Top             =   420
         Width           =   4935
      End
      Begin VB.TextBox txtEMail2 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   7
         Top             =   1860
         Width           =   4935
      End
      Begin VB.TextBox txtWebSite 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   1140
         Width           =   4935
      End
      Begin VB.TextBox txtTelFax 
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   2220
         Width           =   2055
      End
      Begin VB.ComboBox cboCanton 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCR_Conv_Estudio.frx":0383
         Left            =   1440
         List            =   "frmCR_Conv_Estudio.frx":0385
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3420
         Width           =   1935
      End
      Begin VB.ComboBox cboDistrito 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCR_Conv_Estudio.frx":0387
         Left            =   1440
         List            =   "frmCR_Conv_Estudio.frx":0389
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3780
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   49
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Cuenta Depósitos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   48
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   47
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   9480
         X2              =   240
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Forma de Pago:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   32
         Top             =   6300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Contacto 2.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   31
         Top             =   5940
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Contacto 1.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   30
         Top             =   5580
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   9480
         X2              =   240
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   29
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Apto. Postal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   28
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   27
         Top             =   3780
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cantón"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   26
         Top             =   3420
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   25
         Top             =   3060
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   9480
         X2              =   240
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   23
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   22
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Ced. Juridica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   21
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Razón Social"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3600
         TabIndex        =   20
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3600
         TabIndex        =   19
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Web Site"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3600
         TabIndex        =   18
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   360
         TabIndex        =   17
         Top             =   2220
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos Adicionales:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   34
         Top             =   5220
         Width           =   1695
      End
      Begin VB.Label lblBtnColor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2700
         Width           =   1695
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9000
      TabIndex        =   35
      Top             =   480
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   396
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   9864
      _ExtentX        =   17399
      _ExtentY        =   699
      BandCount       =   2
      _CBWidth        =   9864
      _CBHeight       =   396
      _Version        =   "6.7.9816"
      Child1          =   "tlb"
      MinWidth1       =   1800
      MinHeight1      =   336
      Width1          =   1800
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   336
      Width2          =   1104
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   37
         Top             =   30
         Width           =   8475
         _ExtentX        =   14944
         _ExtentY        =   572
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
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   120
      TabIndex        =   38
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmCR_Conv_Estudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vCantonMascara As String, vDistritoMascara As String, vFechaActual As Date
Dim vScroll As Boolean, vTipoJuridica As Integer, vPaso As Boolean



Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoPago.SetFocus
End Sub

Private Sub cboCanton_Click()
Dim vSQL As String

If vPaso Then Exit Sub

vSQL = " where Provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) _
     & " and Canton = '" & Format(cboCanton.ItemData(cboCanton.ListIndex), vCantonMascara) & "' order by descripcion"

Call sbCargaCbo(cboDistrito, "distritos", vSQL)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "
End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub cboProvincia_Click()
Dim vSQL As String

If vPaso Then Exit Sub

vSQL = " where provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) & " order by descripcion"

vPaso = True
 Call sbCargaCbo(cboCanton, "cantones", vSQL)
vPaso = False
Call cboCanton_Click

End Sub

Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDepositos.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_institucion from CRD_ESTUDIO_INSTITUCION"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_institucion > '" & txtCodigo.Text & "' order by cod_institucion asc"
    Else
       strSQL = strSQL & " where cod_institucion < '" & txtCodigo.Text & "' order by cod_institucion desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_Institucion
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

End Sub

Private Sub Form_Activate()
vModulo = 16
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 vModulo = 16

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
ssTab.Tab = 0

vEdita = False
vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")

'Mascara del canton
vCantonMascara = "0"
strSQL = "select MAX(LEN(canton)) as Caracteres from CANTONES"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
 vCantonMascara = SIFGlobal.fxStringRelleno(vCantonMascara, "D", "0", rs!Caracteres)
End If
rs.Close

'Mascara del distrito
vDistritoMascara = "0"
strSQL = "select MAX(LEN(distrito)) as Caracteres from Distritos"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
 vDistritoMascara = SIFGlobal.fxStringRelleno(vDistritoMascara, "D", "0", rs!Caracteres)
End If
rs.Close


'Carga Bancos
vPaso = True
strSQL = "select ID_Banco as Idx, rtrim(Descripcion) as ItmX from Tes_Bancos where estado = 'A'"
    Call sbLlenaCbo(cboBanco, strSQL, False, True)
vPaso = False

'Carga combo de Provincias
vPaso = True
 Call sbCargaCbo(cboProvincia, "provincias")
vPaso = False

 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String

vCodigo = ""
txtCodigo = ""

chkActiva.Value = vbChecked

txtCedJur.Text = ""
txtDescripcion = ""
txtTelefono.Text = ""
txtTelefono2.Text = ""
txtCelular.Text = ""
txtRazon.Text = ""
txtWebSite.Text = ""
txtEMail.Text = ""
txtEMail2.Text = ""
txtAptoPostal.Text = ""

txtDireccion = ""

strSQL = "select id_banco as Idx,descripcion as ItmX from Tes_Bancos"
Call sbLlenaCbo(cboBanco, strSQL, False, True)

cboTipoPago.Text = "Cheques"
txtCuentaDepositos.Text = ""

txtContacto_01.Text = ""
txtContacto_01.Text = ""
txtNotas.Text = ""


ssTab.Tab = 0
ssTab.TabEnabled(1) = False
ssTab.TabEnabled(2) = False

End Sub



Private Sub sbConsultaContratoDetalle()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

''Carga Pagadores
'lswP.ListItems.Clear
'strSQL = " select P.descripcion, C.*" _
'     & " from CRD_ESTUDIO_INSTITUCION P inner join CRD_ESTUDIO_INSTITUCION_Contratos_Pagadores C on P.cod_institucion = C.cod_institucion_pagador" _
'     & " where C.cod_contrato = '" & txtCodigo.Text & "' and C.cod_institucion = '" & vCodigo & "'"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
'    Set itmX = lswP.ListItems.Add(, , rs!cod_institucion_Pagador)
'        itmX.SubItems(1) = rs!descripcion
'        itmX.SubItems(2) = rs!Registro_Usuario & "..." & rs!Registro_Fecha
'    rs.MoveNext
'Loop
'rs.Close
'
'
''Carga Cargos de Suscripción
'lswC.ListItems.Clear
'strSQL = " select C.descripcion,S.*" _
'       & " from CxC_Cargos C inner join CRD_ESTUDIO_INSTITUCION_Contratos_Suscripciones S on C.cod_cargo = S.cod_cargo" _
'       & " where S.cod_contrato = '" & txtCodigo.Text & "' and S.cod_institucion = '" & vCodigo & "'"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
'    Set itmX = lswC.ListItems.Add(, , rs!cod_cargo)
'        itmX.SubItems(1) = rs!Descripcion
'
'        Select Case rs!Tipo
'           Case "P"
'             itmX.SubItems(2) = "Porcentual"
'           Case "M"
'             itmX.SubItems(2) = "Monto"
'        End Select
'
'
'        Select Case rs!Frecuencia_Tipo
'          Case "O"
'            itmX.SubItems(4) = "Operación"
'          Case "D"
'            itmX.SubItems(4) = "Días"
'        End Select
'
'        itmX.SubItems(3) = Format(rs!Valor, "Standard")
'        itmX.SubItems(5) = rs!Frecuencia_dias
'        itmX.SubItems(6) = Format(rs!Recaudado, "Standard")
'        itmX.SubItems(7) = Format(rs!Pago_Ultimo, "dd/mm/yyyy")
'        itmX.SubItems(8) = Format(rs!Pago_Proximo, "dd/mm/yyyy")
'        itmX.SubItems(9) = rs!Modifica
'        itmX.Checked = True
'    rs.MoveNext
'Loop
'rs.Close


vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub lswCarreras_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert CRD_ESTUDIO_INST_CARRERAS(cod_institucion,cod_carrera,registro_fecha,registro_usuario)" _
          & " values('" & vCodigo & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete CRD_ESTUDIO_INST_CARRERAS where cod_institucion = '" & vCodigo & "' and cod_carrera = '" & Item.Tag & "'"

End If
   
Call ConectionExecute(strSQL)

Call Bitacora(IIf(Item.Checked = True, "Aplica", "Elimina"), "Con.Est.Ins. Asignación de Carrera: " & Item.Tag & " a Inst: " & vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass


vPaso = True
Select Case ssTab.Tab
   Case 1 'Carreas Asignadas

       vPaso = True
       lswCarreras.ListItems.Clear
       
       strSQL = "select Car.COD_CARRERA,Car.DESCRIPCION,Asg.COD_INSTITUCION " _
              & " from CRD_ESTUDIO_CARRERAS Car left join CRD_ESTUDIO_INST_CARRERAS Asg on Car.COD_CARRERA = Asg.COD_CARRERA" _
              & " and Asg.COD_INSTITUCION = '" & vCodigo & "'" _
              & " Where Car.ACTIVO = 1" _
              & " order by Asg.COD_INSTITUCION desc,Car.COD_CARRERA"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswCarreras.ListItems.Add(, , rs!Descripcion)
             itmX.Tag = rs!Cod_Carrera
             itmX.Checked = IIf(IsNull(rs!cod_Institucion), vbUnchecked, vbChecked)
          rs.MoveNext
       Loop
       rs.Close
       vPaso = False

   Case 2 'Contratos


       vPaso = True
       
'       lswContratos.ListItems.Clear
'
'       strSQL = " select P.descripcion,C.*" _
'              & " from CxC_Contratos P inner join CRD_ESTUDIO_INSTITUCION_Contratos C on P.cod_contrato = C.cod_contrato" _
'              & " where C.cod_institucion = '" & vCodigo & "' order by C.Activo desc,C.Registro_Fecha desc"
'       Call OpenRecordSet(rs, strSQL)
'       Do While Not rs.EOF
'         Set itmX = lswContratos.ListItems.Add(, , rs!Contrato_Num)
'             itmX.SubItems(1) = rs!Cod_Contrato
'             itmX.SubItems(2) = rs!Descripcion
'             itmX.SubItems(3) = IIf(rs!ACTIVO = 1, "Sí", "No")
'
'             itmX.SubItems(4) = rs!Plazo
'             itmX.SubItems(5) = Format(rs!Tasa_Corriente, "Standard")
'             itmX.SubItems(6) = Format(rs!Tasa_Mora, "Standard")
'
'             If rs!Contrato_Tipo = "D" Then
'                 itmX.SubItems(7) = Format(rs!Contrato_Vence, "dd/mm/yyyy")
'             Else
'                 itmX.SubItems(7) = "Indefinido"
'             End If
'
'             itmX.SubItems(8) = rs!Registro_Usuario & "..." & rs!Registro_Fecha
'
'          rs.MoveNext
'       Loop
'       rs.Close

       vPaso = False

End Select

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_institucion,descripcion from CRD_ESTUDIO_INSTITUCION"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String


On Error GoTo vError

If Not fxSIFValidaCadena(txtCodigo.Text) Then
   Exit Sub
End If


Me.MousePointer = vbHourglass

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & " from CRD_ESTUDIO_INSTITUCION P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and convert(int,P.Canton) = convert(int,Dist.Canton) and P.distrito = Dist.distrito" _
       & " where P.cod_institucion = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!cod_Institucion
  txtCodigo = rs!cod_Institucion

  txtDescripcion = rs!Descripcion & ""
  
  chkActiva.Value = rs!activa
  
  txtCedJur.Text = rs!cedula_juridica & ""
  
  txtTelefono.Text = rs!telefono1 & ""
  txtTelefono2.Text = rs!telefono2 & ""
  txtCelular.Text = rs!celular & ""
  txtTelFax.Text = rs!Fax & ""
   
  txtRazon.Text = rs!Razon_Social & ""
  txtWebSite.Text = rs!WebSite & ""
  txtEMail.Text = rs!Email_01 & ""
  txtEMail2.Text = rs!Email_02 & ""
  txtAptoPostal.Text = rs!apto_postal & ""

  txtContacto_01.Text = rs!Contacto_01 & ""
  txtContacto_02.Text = rs!Contacto_02 & ""
  
  txtNotas = rs!NOTAS & ""

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")

  cboDistrito.ToolTipText = Trim(rs!distrito) & ""
  txtDireccion.Text = rs!Direccion

  'Datos de Cuentas Bancarias
  txtCuentaDepositos.Text = Trim(rs!Cuenta_Bancaria & "")
    If rs!Tipo_Emision = "CK" Then
      cboTipoPago.Text = "Cheques"
    Else
      cboTipoPago.Text = "Transferencia"
    End If
    
    If Not IsNull(rs!cod_banco) Then
     If rs!cod_banco > 0 Then
        Call sbCboAsignaDato(cboBanco, Trim(fxSIFCCodigos("D", rs!cod_banco, "Bancos")))
     End If
    End If
    

  ssTab.Tab = 0
  ssTab.TabEnabled(1) = True
  ssTab.TabEnabled(2) = True

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbFichaCliente(pcod_institucion As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

'Procedimiento
'1. Buscar si existe en la tabla CRD_ESTUDIO_INSTITUCION, si existe --> salir
'2.1. Si no existe buscar si existe en Socios, si no existe --> salir
'2.2. Si no existe buscar si existe en CxP_Proveedores, si no existe --> salir
'3. Si existe en (Socios o CxP_Proveedores) cargar los datos encontrados en pantalla


' Punto 1
strSQL = "select isnull(count(*),0) as Existe from CRD_ESTUDIO_INSTITUCION where cod_institucion = '" & pcod_institucion & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  Me.MousePointer = vbDefault
  Exit Sub
End If
rs.Close

'Punto 2.1
strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ",Tid.Descripcion as TipoIdDesc,Tid.Tipo_Personeria" _
       & ",dbo.fxAFITelefono(P.cod_institucion,1) as 'TelHab',dbo.fxAFITelefono(P.cod_institucion,2) as 'TelTra', dbo.fxAFITelefono(P.cod_institucion,3) as 'TelCell'" _
       & " from socios P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and P.Canton = Dist.Canton and P.distrito = Dist.distrito" _
       & " left join AFI_TIPOS_IDS Tid on P.tipo_id = Tid.tipo_id" _
       & " where P.cod_institucion = '" & pcod_institucion & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vCodigo = rs!cod_Institucion
  txtCodigo = rs!cod_Institucion

  txtDescripcion = rs!Descripcion & ""
  
  txtRazon.Text = rs!Razon_Social & ""
  txtWebSite.Text = ""
  txtEMail.Text = rs!AF_Email & ""
  txtEMail2.Text = ""
  txtAptoPostal.Text = Trim(rs!apto & "")

  txtTelefono.Text = rs!TelHab
  txtTelefono2.Text = rs!TelTra
  txtCelular.Text = rs!TelCell

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
     
  cboDistrito.ToolTipText = Trim(rs!distrito) & ""
  txtDireccion.Text = rs!Direccion & ""

  rs.Close
  Exit Sub
End If
rs.Close


Me.MousePointer = vbDefault
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

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - descripcion no es válido ..."

'Verifica que exista ningun otro proveedor con la misma cedula juridica
strSQL = "select isnull(count(*),0) as Existe from CRD_ESTUDIO_INSTITUCION" _
       & " where cod_institucion not in('" & vCodigo & "') and Cedula_Juridica = '" _
       & Trim(txtCedJur.Text) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   vMensaje = vMensaje & vbCrLf & " - Existe ya una institución registrada con la Misma Cedula Juridica ..."
End If
rs.Close


'Si Existe Enlace con SIF / Realizar esta verificacion
If cboBanco.ListIndex = -1 Then vMensaje = vMensaje & vbCrLf & " - No se especificó un Banco para Desembolsos ..."


If Mid(cboTipoPago.Text, 1, 1) = "T" _
   And Len(Trim(txtCuentaDepositos)) = 0 Then vMensaje = vMensaje & vbCrLf & " - No se especificó la cuenta para depositos de las transferencias..."



If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEstadoCivil As String

On Error GoTo vError


If vEdita Then
  strSQL = "update CRD_ESTUDIO_INSTITUCION set descripcion = '" & Trim(txtDescripcion.Text) & "',Razon_Social = '" & Trim(txtRazon.Text) _
         & "',telefono1 = '" & txtTelefono.Text & "',telefono2 = '" & txtTelefono2.Text & "',celular = '" & txtCelular.Text & "',Fax = '" & txtTelFax.Text & "',WebSite = '" _
         & txtWebSite.Text & "',apto_postal = '" & txtAptoPostal & "',email_01 = '" & txtEMail & "', email_02 = '" & txtEMail2.Text & "',direccion = '" & txtDireccion _
         & "',distrito = '" & Format(cboDistrito.ItemData(cboDistrito.ListIndex), vDistritoMascara) & "',canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "',provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) & ",cod_Banco = " & cboBanco.ItemData(cboBanco.ListIndex) _
         & ",Tipo_Emision = '" & IIf((cboTipoPago.Text = "Cheques"), "CK", "TE") & "',Cuenta_Bancaria = '" & txtCuentaDepositos.Text _
         & "',notas = '" & txtNotas & "',Contacto_01 = '" & txtContacto_01.Text & "',Contacto_02 = '" & txtContacto_02.Text & "', Activa = " & chkActiva.Value _
         & ",cedula_Juridica = '" & Trim(txtCedJur.Text) & "' where cod_institucion = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Modifica", "Conv.Est.Inst.: " & vCodigo)

Else
  vCodigo = txtCodigo

   strSQL = "insert into CRD_ESTUDIO_INSTITUCION(cod_institucion,descripcion,cedula_Juridica,razon_social,activa,celular,telefono1,telefono2,fax,Contacto_01,Contacto_02" _
          & ",apto_postal,email_01,email_02,webSite,notas,direccion,distrito,provincia,canton,cod_Banco,Tipo_Emision,Cuenta_Bancaria,Registro_Fecha,Registro_Usuario)" _
          & " values('" & vCodigo & "','" & txtDescripcion.Text & "','" & txtCedJur.Text & "','" & txtRazon.Text & "'," & chkActiva.Value & ",'" & txtCelular.Text _
          & "','" & txtTelefono.Text & "','" & txtTelefono2.Text & "','" & txtTelFax.Text & "','" & txtContacto_01.Text & "','" & txtContacto_02.Text _
          & "','" & txtAptoPostal.Text & "','" & txtEMail.Text & "','" & txtEMail2.Text & "','" & txtWebSite.Text _
          & "','" & txtNotas.Text & "','" & txtDireccion.Text & "','" & Format(cboDistrito.ItemData(cboDistrito.ListIndex), vDistritoMascara) & "'," _
          & cboProvincia.ItemData(cboProvincia.ListIndex) & ",'" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & IIf((cboTipoPago.Text = "Cheques"), "CK", "TE") _
          & "','" & Trim(txtCuentaDepositos.Text) & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Conv.Est.Inst.: " & vCodigo)

End If

ssTab.TabEnabled(1) = True
ssTab.TabEnabled(2) = True

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CRD_ESTUDIO_INSTITUCION where cod_institucion = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Conv.Est.Inst.: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCedJur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRazon.SetFocus
End Sub

Private Sub txtCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelFax.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_institucion"
  gBusquedas.Orden = "cod_institucion"
  gBusquedas.Consulta = "select cod_institucion,descripcion from CRD_ESTUDIO_INSTITUCION"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
  Call sbConsulta(txtCodigo.Text)
'  txtDescripcion.SetFocus
End Sub




Private Sub txtContacto_01_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContacto_02.SetFocus
End Sub

Private Sub txtContacto_02_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtCuentaDepositos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContacto_01.SetFocus
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBanco.SetFocus

End Sub

Private Sub txtEMail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedJur.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_institucion,descripcion from CRD_ESTUDIO_INSTITUCION"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtRazon_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCelular.SetFocus
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEMail2.SetFocus
End Sub


Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus

End Sub

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEMail.SetFocus
End Sub


