VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCC_Documentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Documentos"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   HelpContextID   =   9006
   Icon            =   "CC_Documentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab ssTab 
      Height          =   6135
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   1
      TabsPerRow      =   7
      TabHeight       =   520
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Documentos"
      TabPicture(0)   =   "CC_Documentos.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "txtDocumento"
      Tab(0).Control(2)=   "cboTipo"
      Tab(0).Control(3)=   "lswAsiento"
      Tab(0).Control(4)=   "FlatScrollBar"
      Tab(0).Control(5)=   "imgAnular"
      Tab(0).Control(6)=   "Label2(0)"
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(8)=   "Label2(1)"
      Tab(0).Control(9)=   "imgReImpresion"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Reportes"
      TabPicture(1)   =   "CC_Documentos.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraReportes"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Depósitos"
      TabPicture(2)   =   "CC_Documentos.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "txtNumDp"
      Tab(2).Control(2)=   "lswDP"
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(4)=   "Label11(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Usuarios"
      TabPicture(3)   =   "CC_Documentos.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ssTabAux"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Traslados"
      TabPicture(4)   =   "CC_Documentos.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraTraspaso"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Configuración"
      TabPicture(5)   =   "CC_Documentos.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraConfiguracion"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Cuentas Autorizadas"
      TabPicture(6)   =   "CC_Documentos.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "vGrid"
      Tab(6).ControlCount=   1
      Begin TabDlg.SSTab ssTabAux 
         Height          =   5415
         Left            =   -74760
         TabIndex        =   119
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   9551
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ReImprimen"
         TabPicture(0)   =   "CC_Documentos.frx":03CE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label12(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lswUsReImprimen"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cuentas Autorizadas"
         TabPicture(1)   =   "CC_Documentos.frx":03EA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkUsCtaTodos"
         Tab(1).Control(1)=   "lswUsCtaAsignadas"
         Tab(1).Control(2)=   "lswUsCtaAuto"
         Tab(1).Control(3)=   "lblUSCtaActual"
         Tab(1).Control(4)=   "Label12(1)"
         Tab(1).Control(5)=   "Label12(0)"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Códigos de Formalizaciones"
         TabPicture(2)   =   "CC_Documentos.frx":0406
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lswUsCodigos"
         Tab(2).Control(1)=   "Label12(2)"
         Tab(2).ControlCount=   2
         Begin VB.CheckBox chkUsCtaTodos 
            Alignment       =   1  'Right Justify
            Caption         =   "Todas"
            Height          =   255
            Left            =   -68400
            TabIndex        =   125
            Top             =   5040
            Width           =   1095
         End
         Begin MSComctlLib.ListView lswUsCtaAsignadas 
            Height          =   2535
            Left            =   -72840
            TabIndex        =   122
            Top             =   2400
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   4471
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cuenta"
               Object.Width           =   2893
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripcion"
               Object.Width           =   6068
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Creditos"
               Object.Width           =   1658
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Retencion"
               Object.Width           =   1658
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Fondos"
               Object.Width           =   1658
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Patrimonio"
               Object.Width           =   1658
            EndProperty
         End
         Begin MSComctlLib.ListView lswUsCtaAuto 
            Height          =   1695
            Left            =   -72840
            TabIndex        =   121
            Top             =   600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   2990
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Usuario"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   6068
            EndProperty
         End
         Begin MSComctlLib.ListView lswUsReImprimen 
            Height          =   4575
            Left            =   2160
            TabIndex        =   120
            Top             =   600
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   8070
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Usuario"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   6068
            EndProperty
         End
         Begin MSComctlLib.ListView lswUsCodigos 
            Height          =   4575
            Left            =   -72840
            TabIndex        =   128
            Top             =   600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   8070
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   2187
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripcion"
               Object.Width           =   6068
            EndProperty
         End
         Begin VB.Label Label12 
            Caption         =   "Seleccione a los usuarios Autorizados a Re-Imprimir Documentos (Deben Indicar Su Usuario y Clave de Acceso a la Base de Datos)"
            Height          =   1335
            Index           =   3
            Left            =   240
            TabIndex        =   131
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   $"CC_Documentos.frx":0422
            Height          =   1455
            Index           =   2
            Left            =   -74760
            TabIndex        =   127
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblUSCtaActual 
            Caption         =   "..."
            Height          =   255
            Left            =   -72840
            TabIndex        =   126
            Top             =   5040
            Width           =   4335
         End
         Begin VB.Label Label12 
            Caption         =   "Lista de Cuentas para Asignar al Usuario Seleccionado"
            Height          =   855
            Index           =   1
            Left            =   -74760
            TabIndex        =   124
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "Lista de Usuarios Disponibles (Seleccione el Usuario al que desea Asignar Cuentas)"
            Height          =   855
            Index           =   0
            Left            =   -74760
            TabIndex        =   123
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cambio de # de Depósito"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   108
         Top             =   4320
         Width           =   7935
         Begin VB.CommandButton cmdDepCambio 
            Caption         =   "&Cambiar # Depósito"
            Height          =   375
            Left            =   5400
            TabIndex        =   115
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDepNuevo 
            Height          =   315
            Left            =   1320
            TabIndex        =   114
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txtDepActual 
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   113
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txtDepDoc 
            Height          =   315
            Left            =   5400
            TabIndex        =   112
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cboDepTipo 
            Height          =   315
            ItemData        =   "CC_Documentos.frx":04B3
            Left            =   1320
            List            =   "CC_Documentos.frx":04B5
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "# Dep. Nuevo"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   117
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "# Dep. Actual"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   116
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "# Documento"
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   111
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   110
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtNumDp 
         Height          =   315
         Left            =   -73680
         TabIndex        =   89
         Top             =   840
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   3015
         Left            =   -74880
         TabIndex        =   51
         Top             =   840
         Width           =   8415
         Begin VB.TextBox txtBeneficiario 
            Height          =   315
            Left            =   1080
            TabIndex        =   64
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtFechaGenera 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3600
            TabIndex        =   62
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtEstado 
            Height          =   315
            Left            =   6480
            TabIndex        =   61
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtTipo 
            Height          =   315
            Left            =   3600
            TabIndex        =   60
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtUS_Genera 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtUS_Anula 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6480
            MultiLine       =   -1  'True
            TabIndex        =   58
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtUS_Traspasa 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   57
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtFechaAnula 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   6480
            TabIndex        =   56
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtFechaTraspasa 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3600
            TabIndex        =   55
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtPago 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   6480
            TabIndex        =   54
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtConcepto 
            Height          =   495
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   53
            ToolTipText     =   "Concepto del Recibo"
            Top             =   1680
            Width           =   7215
         End
         Begin VB.TextBox txtDetalle 
            Height          =   735
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   52
            ToolTipText     =   "Detalle de la nota"
            Top             =   2200
            Width           =   7215
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   7200
            Top             =   1560
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
                  Picture         =   "CC_Documentos.frx":04B7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CC_Documentos.frx":07D7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CC_Documentos.frx":0AF7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CC_Documentos.frx":0E17
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            Caption         =   "Estado"
            Height          =   255
            Left            =   5640
            TabIndex        =   77
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Monto"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Genera"
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   74
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   73
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblUS 
            Caption         =   "US.Genera"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblUS 
            Caption         =   "US.Anula"
            Height          =   255
            Index           =   1
            Left            =   5640
            TabIndex        =   71
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblUS 
            Caption         =   "US.Traspasa"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Anula"
            Height          =   255
            Index           =   2
            Left            =   5640
            TabIndex        =   69
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Traslada"
            Height          =   255
            Index           =   3
            Left            =   2880
            TabIndex        =   68
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Pago"
            Height          =   255
            Left            =   5640
            TabIndex        =   67
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblUS 
            Caption         =   "Concepto"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   66
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblUS 
            Caption         =   "Detalle"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   65
            Top             =   2160
            Width           =   855
         End
      End
      Begin VB.TextBox txtDocumento 
         Height          =   315
         Left            =   -71280
         TabIndex        =   49
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "CC_Documentos.frx":1133
         Left            =   -74040
         List            =   "CC_Documentos.frx":1135
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame fraTraspaso 
         Caption         =   "Traslado de Asientos a ContaXpress"
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
         Height          =   4215
         Left            =   -73560
         TabIndex        =   39
         Top             =   960
         Width           =   5775
         Begin VB.CheckBox chkAS_Depositos 
            Appearance      =   0  'Flat
            Caption         =   "Depósitos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   45
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAS_Recibos 
            Appearance      =   0  'Flat
            Caption         =   "Recibos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   44
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAS_ND 
            Appearance      =   0  'Flat
            Caption         =   "Notas de Débito"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   43
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAS_NC 
            Appearance      =   0  'Flat
            Caption         =   "Notas de Crédito"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   42
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton cmdAS_Aceptar 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   41
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CheckBox chkAsientoResumen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Crear Asiento Tipo Diario"
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
            Height          =   255
            Left            =   1680
            TabIndex        =   40
            Top             =   3480
            Width           =   2175
         End
         Begin MSComctlLib.ProgressBar prgBar 
            Height          =   135
            Left            =   120
            TabIndex        =   46
            Top             =   3240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSComCtl2.DTPicker dtpAsientoInicio 
            Height          =   300
            Left            =   3000
            TabIndex        =   135
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   178520067
            CurrentDate     =   36462
         End
         Begin MSComCtl2.DTPicker dtpAsientoCorte 
            Height          =   300
            Left            =   3000
            TabIndex        =   136
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   178520067
            CurrentDate     =   36462
         End
         Begin VB.Label Label8 
            Caption         =   "Corte"
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
            Left            =   2040
            TabIndex        =   138
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Inicio"
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
            Left            =   2040
            TabIndex        =   137
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblEstatus 
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   3000
            Width           =   5415
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   5520
            Y1              =   3360
            Y2              =   3360
         End
      End
      Begin VB.Frame fraConfiguracion 
         Caption         =   "Configuración"
         ForeColor       =   &H00FF0000&
         Height          =   5535
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   7935
         Begin VB.CheckBox chkUtilizaCtaAutorizadas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Utilizar Cuentas Autorizadas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5520
            TabIndex        =   118
            ToolTipText     =   "Formato del Recibo tipo Boucher"
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtConDPTA 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   4200
            Width           =   3255
         End
         Begin VB.TextBox txtConDPCta 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   3840
            Width           =   3255
         End
         Begin VB.TextBox txtConNDTA 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   3120
            Width           =   3255
         End
         Begin VB.TextBox txtConNDCta 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   2760
            Width           =   3255
         End
         Begin VB.TextBox txtConNCTA 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   2040
            Width           =   3255
         End
         Begin VB.TextBox txtConNCCta 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox txtConRETA 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   101
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox txtConRECta 
            Height          =   315
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   100
            Top             =   600
            Width           =   3255
         End
         Begin VB.CommandButton cmdGuardaConfiguracion 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   6720
            TabIndex        =   27
            Top             =   5040
            Width           =   975
         End
         Begin VB.CheckBox chkUtilizaRecibo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Utilizar Documentos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtTA_RE 
            Height          =   315
            Left            =   3000
            TabIndex        =   25
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtTA_NC 
            Height          =   315
            Left            =   3000
            TabIndex        =   24
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtTA_ND 
            Height          =   315
            Left            =   3000
            TabIndex        =   23
            Top             =   3120
            Width           =   1575
         End
         Begin VB.TextBox txtTA_DP 
            Height          =   315
            Left            =   3000
            TabIndex        =   22
            Top             =   4200
            Width           =   1575
         End
         Begin VB.TextBox txtID_DP 
            Height          =   315
            Left            =   3000
            TabIndex        =   21
            Top             =   4560
            Width           =   1575
         End
         Begin VB.TextBox txtID_ND 
            Height          =   315
            Left            =   3000
            TabIndex        =   20
            Top             =   3480
            Width           =   1575
         End
         Begin VB.TextBox txtID_NC 
            Height          =   315
            Left            =   3000
            TabIndex        =   19
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox txtID_RE 
            Height          =   315
            Left            =   3000
            TabIndex        =   18
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox chkReciboFlat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Utilizar Formato de Recibo voucher"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            ToolTipText     =   "Formato del Recibo tipo Boucher"
            Top             =   240
            Width           =   2895
         End
         Begin MSMask.MaskEdBox medNC 
            Height          =   315
            Left            =   3000
            TabIndex        =   28
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medRE 
            Height          =   315
            Left            =   3000
            TabIndex        =   29
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medND 
            Height          =   315
            Left            =   3000
            TabIndex        =   30
            Top             =   2760
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medDP 
            Height          =   315
            Left            =   3000
            TabIndex        =   31
            Top             =   3840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo Asiento"
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
            Height          =   315
            Index           =   15
            Left            =   1440
            TabIndex        =   99
            Top             =   4200
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuenta Contable"
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
            Height          =   315
            Index           =   14
            Left            =   1440
            TabIndex        =   98
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Consecutivo"
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
            Height          =   315
            Index           =   13
            Left            =   1440
            TabIndex        =   97
            Top             =   4560
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo Asiento"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   12
            Left            =   1440
            TabIndex        =   96
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuenta Contable"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   11
            Left            =   1440
            TabIndex        =   95
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Consecutivo"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   10
            Left            =   1440
            TabIndex        =   94
            Top             =   3480
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo Asiento"
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
            Height          =   315
            Index           =   9
            Left            =   1440
            TabIndex        =   93
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuenta Contable"
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
            Height          =   315
            Index           =   8
            Left            =   1440
            TabIndex        =   92
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Consecutivo"
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
            Height          =   315
            Index           =   7
            Left            =   1440
            TabIndex        =   91
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo Asiento"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   38
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuenta Contable"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   37
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Recibos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nota Crédito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nota Débito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   34
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Depósitos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   7800
            Y1              =   4920
            Y2              =   4920
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Consecutivo"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   1440
            TabIndex        =   32
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000C&
            BorderWidth     =   2
            X1              =   120
            X2              =   7800
            Y1              =   4920
            Y2              =   4920
         End
      End
      Begin VB.Frame fraReportes 
         ForeColor       =   &H8000000D&
         Height          =   5535
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   6255
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   132
            Top             =   1200
            Width           =   3135
         End
         Begin VB.ComboBox cboRepFiltro 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   840
            Width           =   3135
         End
         Begin VB.OptionButton optReportes 
            Caption         =   "Especial para Cierres"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   86
            Top             =   4560
            Width           =   1935
         End
         Begin VB.ComboBox cboRepTipo 
            Height          =   315
            ItemData        =   "CC_Documentos.frx":1137
            Left            =   3000
            List            =   "CC_Documentos.frx":1139
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   480
            Width           =   3135
         End
         Begin VB.OptionButton optReportes 
            Caption         =   "Reporte General Documentos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   2760
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton optReportes 
            Caption         =   "Documentos Emitidos"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   3120
            Width           =   2055
         End
         Begin VB.OptionButton optReportes 
            Caption         =   "Documentos Anulados"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   3480
            Width           =   2055
         End
         Begin VB.OptionButton optReportes 
            Caption         =   "Pendientes Traslado Asientos a Contabilidad"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   3840
            Width           =   3855
         End
         Begin VB.CheckBox chkTodasLasFechas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Todas las Fechas "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   480
            TabIndex        =   7
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox chkFechaAnulacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Utilizar como fecha Base Anulación"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3000
            TabIndex        =   6
            Top             =   1560
            Width           =   3135
         End
         Begin VB.CheckBox chkFechaEmision 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Utilizar como fecha Base Emisión"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3000
            TabIndex        =   5
            Top             =   1800
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.CheckBox chkFechaTraspaso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Utilizar como fecha Base Traspaso"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3000
            TabIndex        =   4
            Top             =   2040
            Width           =   3135
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Reporte"
            Height          =   375
            Left            =   5040
            TabIndex        =   3
            Top             =   5040
            Width           =   975
         End
         Begin VB.OptionButton optReportes 
            Caption         =   "Traslados de Asientos Generados a Contabilidad"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   2
            Top             =   4200
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker dtpDesde 
            Height          =   300
            Left            =   840
            TabIndex        =   8
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   189857795
            CurrentDate     =   36462
         End
         Begin MSComCtl2.DTPicker dtpHasta 
            Height          =   300
            Left            =   840
            TabIndex        =   13
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   189857795
            CurrentDate     =   36462
         End
         Begin VB.Label Label8 
            Caption         =   "Bancos"
            Height          =   255
            Index           =   8
            Left            =   2280
            TabIndex        =   133
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Filtro"
            Height          =   255
            Index           =   7
            Left            =   2280
            TabIndex        =   130
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   84
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reportes"
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
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   2400
            Width           =   6015
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Parámetros"
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
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   120
            Width           =   6015
         End
         Begin VB.Label Label8 
            Caption         =   "Desde"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Hasta"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   615
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   6000
            Y1              =   4920
            Y2              =   4920
         End
      End
      Begin MSComctlLib.ListView lswAsiento 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   50
         Top             =   4200
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Debe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Haber"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   255
         Left            =   -69360
         TabIndex        =   78
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1572865
      End
      Begin MSComctlLib.ListView lswDP 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   88
         Top             =   1200
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5318
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Concepto"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Deposito"
            Object.Width           =   2540
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   134
         Top             =   480
         Width           =   8415
         _Version        =   524288
         _ExtentX        =   14843
         _ExtentY        =   9340
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
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "CC_Documentos.frx":113B
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin VB.Image imgAnular 
         Height          =   255
         Left            =   -68400
         Picture         =   "CC_Documentos.frx":1861
         Stretch         =   -1  'True
         ToolTipText     =   "Anular (Recibos)"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "# Depósito"
         Height          =   255
         Left            =   -74760
         TabIndex        =   90
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consulta de Depósitos"
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
         Height          =   285
         Index           =   2
         Left            =   -74760
         TabIndex        =   87
         Top             =   480
         Width           =   7935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# Doc"
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
         Left            =   -72120
         TabIndex        =   81
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asiento del Documento"
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
         TabIndex        =   80
         Top             =   3960
         Width           =   8415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         Left            =   -74880
         TabIndex        =   79
         Top             =   480
         Width           =   855
      End
      Begin VB.Image imgReImpresion 
         Height          =   255
         Left            =   -68760
         Picture         =   "CC_Documentos.frx":1956
         Stretch         =   -1  'True
         ToolTipText     =   "Presione Aqui para Reimprimir el Doc."
         Top             =   480
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmCC_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean


Private Sub sbLimpiaDatos()
 txtBeneficiario = ""
 txtEstado = ""
 txtFechaAnula = ""
 txtFechaGenera = ""
 txtFechaTraspasa = ""
 txtConcepto = ""
 txtMonto = ""
 txtPago = ""
 txtTipo = ""
 txtUS_Anula = ""
 txtUS_Traspasa = ""
 txtUS_Genera = ""
 txtDetalle = ""
 lswAsiento.ListItems.Clear
End Sub


Private Sub cboDepTipo_Change()
txtDepActual = ""
txtDepDoc = ""
txtDepNuevo = ""
End Sub

Private Sub cboRepFiltro_Click()
On Error GoTo vError

If Mid(cboRepFiltro.Text, 1, 2) = "05" Then
 cboBanco.Enabled = True
Else
 cboBanco.Enabled = False
End If

vError:

End Sub

Private Sub chkUsCtaTodos_Click()
Dim strSQL As String, lng As Long

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "delete ase_usr_CtaAuto where usuario = '" & lblUSCtaActual.Caption & "'"
glogon.Conection.Execute strSQL

With lswUsCtaAsignadas
  For lng = 1 To .ListItems.Count
    
    .ListItems.Item(lng).Checked = chkUsCtaTodos.Value
     
    If .ListItems.Item(lng).Checked = True Then
       strSQL = "insert ase_usr_CtaAuto(cod_cuenta,usuario) values('" & .ListItems.Item(lng).Text _
              & "','" & lblUSCtaActual.Caption & "')"
       glogon.Conection.Execute strSQL
    End If
  Next lng
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub cmdAS_Aceptar_Click()
Dim iRespuesta As Integer

iRespuesta = MsgBox("Esta seguro de realizar el traspaso a contabilidad", vbYesNo)

If iRespuesta = vbYes Then
 
 If chkAsientoResumen.Value = vbChecked Then
    Me.MousePointer = vbHourglass
        If chkAS_Depositos.Value = vbChecked Then Call sbAsientoTipoDiario("DP")
        If chkAS_NC.Value = vbChecked Then Call sbAsientoTipoDiario("NC")
        If chkAS_ND.Value = vbChecked Then Call sbAsientoTipoDiario("ND")
        If chkAS_Recibos.Value = vbChecked Then Call sbAsientoTipoDiario("RE")
    
    MsgBox "Se realizó el pase de asientos a contabilidad ", vbInformation
    Me.MousePointer = vbDefault

 Else
     Call sbAsientoIndividual
 End If
End If 'Respuesta
End Sub


Private Sub chkTodasLasFechas_Click()
If chkTodasLasFechas.Value = vbChecked Then
 dtpDesde.Enabled = False
 dtpHasta.Enabled = False
 chkFechaAnulacion.Value = 0
 chkFechaAnulacion.Enabled = False
 chkFechaEmision.Value = 0
 chkFechaEmision.Enabled = False
 chkFechaTraspaso.Value = 0
 chkFechaTraspaso.Enabled = False
Else
 dtpDesde.Enabled = True
 dtpHasta.Enabled = True
 chkFechaAnulacion.Enabled = True
 chkFechaEmision.Enabled = True
 chkFechaTraspaso.Enabled = True
End If
End Sub


Private Function fxFechaReportes(vTipo As Integer) As String
If vTipo = 1 Then
' fxFechaReportes = Year(dtpDesde.Value) & "," & Month(dtpDesde.Value) & "," & Day(dtpDesde.Value)
  fxFechaReportes = "'" & Format(dtpDesde.Value, "yyyy,mm,dd") & " 00:00:00'"
  fxFechaReportes = " in DateTime(" & Year(dtpDesde.Value) & "," & Month(dtpDesde.Value) & "," & Day(dtpDesde.Value)
Else
' fxFechaReportes = Year(dtpHasta.Value) & "," & Month(dtpHasta.Value) & "," & Day(dtpHasta.Value)
  fxFechaReportes = "'" & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
End If


fxFechaReportes = " in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"



End Function



Private Sub cmdDepCambio_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update ase_documentos set dp = '" & txtDepNuevo _
       & "' where tipo = '" & fxTipoASEDoc(cboDepTipo.Text) _
       & "' and id_documento = " & txtDepDoc
glogon.Conection.Execute strSQL


MsgBox "# de depósito actualizado satisfactoriamente...", vbInformation
Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdGuardaConfiguracion_Click()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vPasa As Boolean
On Error GoTo CapturaError

'Validacion
'vPasa = fxValidaTipoAsiento(txtTA_RE)
'If vPasa Then vPasa = fxValidaTipoAsiento(txtTA_NC)
'If vPasa Then vPasa = fxValidaTipoAsiento(txtTA_ND)
'If vPasa Then vPasa = fxValidaTipoAsiento(txtTA_DP)
If vPasa Then vPasa = fxgCntCuentaValida(medRE)
If vPasa Then vPasa = fxgCntCuentaValida(medNC)
If vPasa Then vPasa = fxgCntCuentaValida(medND)
If vPasa Then vPasa = fxgCntCuentaValida(medDP)

If Not vPasa Then
  MsgBox "Datos inválidos verifiquelos...", vbCritical
  Exit Sub
End If

'Guardar Configuracion
strSQL = "update ase_consecutivos set " _
       & "cs_nota_credito = " & txtID_NC _
       & ",cs_nota_debito = " & txtID_ND _
       & ",cs_deposito = " & txtID_DP _
       & ",cs_recibo = " & txtID_RE _
       & ",cs_utilizar_recibo = '" & IIf((chkUtilizaRecibo.Value = vbChecked), "S", "N") _
       & "',cs_utilizar_reciboFlat = '" & IIf((chkReciboFlat.Value = vbChecked), "S", "N") _
       & "',cs_utilizar_CuentasAuto = '" & IIf((chkUtilizaCtaAutorizadas.Value = vbChecked), "S", "N") _
       & "',cs_nc_cuenta = '" & medNC _
       & "',cs_nd_cuenta = '" & medND _
       & "',cs_dp_cuenta = '" & medDP _
       & "',cs_re_cuenta = '" & medRE _
       & "',cs_nc_asiento = '" & txtTA_NC _
       & "',cs_nd_asiento = '" & txtTA_ND _
       & "',cs_dp_asiento = '" & txtTA_DP _
       & "',cs_re_asiento = '" & txtTA_RE & "'"
glogon.Conection.Execute strSQL


MsgBox "Parámetros Guardados Satisfactoriamente", vbInformation


Exit Sub
CapturaError:
 MsgBox Err.Description, vbCritical
End Sub

Private Function fxUsuarioNombre(vUsuario As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select descripcion from usuarios where nombre = '" & vUsuario & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
 fxUsuarioNombre = "[SIN DESCRIPCION]"
Else
 fxUsuarioNombre = "[" & UCase(Trim(rs!Descripcion)) & "]"

End If
rs.Close
End Function


Sub CreaDetalleAsiento(strTipo As String, strCaso As String, strCuenta As String, vFecha As Date, curMonto As Currency, DH As String, intLinea As Integer)
Dim strSQL As String, strNumero_Asiento As String

If UCase(DH) <> "D" Then 'dc - dh
  DH = "C"
End If

strNumero_Asiento = strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")

strSQL = "insert asientos_detalle(tipo_asiento,num_asiento,num_linea,num_cuenta,tipo_movimiento,monto,detalle,referencia,num_documento)" _
    & " values('AS','" & strNumero_Asiento & "'," & intLinea & "," & fxNumeroCuenta(strCuenta) & ",'" & DH & "'," _
    & curMonto & ",'" & strCaso & "','TRASPASO-CC','" & strTipo & "')"

glogon.Conection.Execute strSQL

End Sub

Private Sub sbBuscaConfiguracion(vObj As Object, vTipo As String)

If UCase(vTipo) = "C" Then
    frmCntX_ConsultaCuentas.Show vbModal
    vObj.Text = gCuenta
Else
    gBusquedas.Columna = "tipo_asiento"
    gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_tipos_asiento"
    gBusquedas.Orden = "tipo_asiento"
    gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
    frmBusquedas.Show vbModal
    vObj.Text = gBusquedas.Resultado
End If

End Sub

Private Sub cmdImprimir_Click()
Dim vTipo As String, vUsuario As String


If (chkTodasLasFechas.Value + chkFechaAnulacion.Value + chkFechaEmision.Value _
    + chkFechaTraspaso.Value) = 0 Then
  MsgBox "No se ha especificado ninguna fecha como parámetro de busqueda...", vbInformation
  Exit Sub
End If

If Mid(cboRepFiltro.Text, 1, 2) = "04" Then
    vUsuario = InputBox("Digite el usuario que desea visualizar", "Control de Documentos / Usr.Específico")
    If Len(Trim(vUsuario)) = 0 Then vUsuario = ""
End If


Me.MousePointer = vbHourglass

vTipo = fxTipoASEDoc(cboRepTipo.Text)

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes de Control de Documentos"
    
    .Connect = glogon.ConectRPT
   
   If Mid(cboRepFiltro.Text, 1, 2) = "02" Then
    .ReportFileName = SIFGlobal.fxSIFPathReportes("SIFDocumentoControlUsr.rpt")
   Else
    .ReportFileName = SIFGlobal.fxSIFPathReportes("SIFDocumentoControl.rpt")
   End If
    
   
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "usuario='Usuario : " & glogon.Usuario & "'"
    .Formulas(2) = "fecha='Fecha : " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  
  Select Case True
   Case optReportes(0).Value  'Reporte General
     .Formulas(3) = "SUBTITULO='REPORTE GENERAL - " & UCase(cboRepTipo.Text) & "'"
       If chkFechaAnulacion.Value = vbChecked Then
          .SelectionFormula = "CDATE({ASE_DOCUMENTOS.FECHA_ANULACION}) " & fxFechaReportes(1) _
                    & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .Formulas(4) = "fecha_anulacion = 'Fecha Anulación entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        
        End If
        
        If chkFechaEmision.Value = vbChecked Then
          .SelectionFormula = "CDATE({ASE_DOCUMENTOS.FECHA}) " & fxFechaReportes(1) _
                    & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .Formulas(4) = "fecha_emision = 'Fecha Emisión entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        End If
        
        If chkFechaTraspaso.Value = vbChecked Then
          .SelectionFormula = "CDATE({ASE_DOCUMENTOS.FECHA_TRASPASO}) " & fxFechaReportes(1) _
                    & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .Formulas(4) = "fecha_traspaso = 'Fecha Traspaso entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        End If
        
   Case optReportes(1).Value  'Emitidos
     .SelectionFormula = "{ASE_DOCUMENTOS.ESTADO} = 'I'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboRepTipo.Text) & " - EMITIDOS'"
     
   Case optReportes(2).Value  'Anulados
     .SelectionFormula = "{ASE_DOCUMENTOS.ESTADO} = 'A'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboRepTipo.Text) & " - ANULADOS'"
   
   Case optReportes(3).Value  'Pendientes de Traspaso
     .SelectionFormula = "{ASE_DOCUMENTOS.TRASPASO} = 'P'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboRepTipo.Text) & " - PENDIENTES TRSP.'"
   
   Case optReportes(4).Value  'Traspasos Generados
     .SelectionFormula = "{ASE_DOCUMENTOS.GENERADOS} = 'G'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboRepTipo.Text) & " - TRSP. GENERADOS'"
  
   Case optReportes(5).Value  'Especial de Cierres
    'Salir del Procedimiento Despues del Print
    .ReportFileName = SIFGlobal.fxSIFPathReportes("SIFDocumentoEspecialCierre.rpt")
    .Formulas(3) = "SUBTITULO='CIERRE DE CAJAS DEL " & Format(dtpDesde.Value, "yyyy/mm/dd") _
                 & " HASTA " & Format(dtpHasta.Value, "yyyy/mm/dd") & "'"
    
    Select Case Mid(cboRepFiltro.Text, 1, 2)
      Case "03"
        .Formulas(4) = "FXUSUARIOCIERRE = '" & glogon.Usuario & "'"
        .Formulas(5) = "fxUsuarioNombre = '" & fxUsuarioNombre(glogon.Usuario) & "'"
        
        .SubreportToChange = "subDocumentos"
        .Connect = glogon.ConectRPT
        .SelectionFormula = "CDATE({ASE_DOCUMENTOS.FECHA}) " & fxFechaReportes(1) _
                    & " AND {ASE_DOCUMENTOS.USUARIO} = '" & glogon.Usuario & "'"
        
        .SubreportToChange = "sbCKCajas"
        .Connect = glogon.ConectRPT
        .SelectionFormula = "CDATE({ASE_CK_CAJA.FECHA}) " & fxFechaReportes(1) _
                    & " AND {ASE_CK_CAJA.USUARIO} = '" & glogon.Usuario & "'"
        
        .SubreportToChange = "subOperaciones"
        .Connect = glogon.ConectRPT
        .SelectionFormula = "({vEspecialCierres.ESTADOSOL} = 'N' OR {vEspecialCierres.ESTADOSOL} ='F')" _
                          & " AND ({vEspecialCierres.EMITIR} = 'CK' OR {vEspecialCierres.EMITIR} = 'EF' OR {vEspecialCierres.EMITIR} = 'TE')" _
                          & " AND {vEspecialCierres.USERFOR} = '" & glogon.Usuario _
                          & "' AND CDATE({vEspecialCierres.FECHAFORP}) " & fxFechaReportes(1)
        
        .SubreportToChange = "sbCKCajas"
        .Connect = glogon.ConectRPT
        .SelectionFormula = "CDATE({ASE_CK_CAJA.FECHA}) " & fxFechaReportes(1) _
                    & " AND {ASE_CK_CAJA.USUARIO} = '" & glogon.Usuario & "'"
       
        .SubreportToChange = "sbFondos"
        .Connect = glogon.ConectRPT
        .SelectionFormula = "CDATE({FND_LIQUIDACION.FECHA}) " & fxFechaReportes(1) _
                    & " AND {FND_LIQUIDACION.USUARIO} = '" & glogon.Usuario & "'"
       
       
       .PrintReport
      
      Case "04"
        .Formulas(4) = "FXUSUARIOCIERRE = '" & UCase(vUsuario) & "'"
        .Formulas(5) = "fxUsuarioNombre = '" & fxUsuarioNombre(vUsuario) & "'"
        
        .SubreportToChange = "subDocumentos"
        .Connect = glogon.ConectRPT
        
        .SelectionFormula = "CDATE({ASE_DOCUMENTOS.FECHA}) " & fxFechaReportes(1) _
                    & " AND {ASE_DOCUMENTOS.USUARIO} = '" & vUsuario & "'"
        
        .SubreportToChange = "subOperaciones"
        .Connect = glogon.ConectRPT
        
        .SelectionFormula = "({vEspecialCierres.ESTADOSOL} = 'N' OR {vEspecialCierres.ESTADOSOL} ='F')" _
                          & " AND ({vEspecialCierres.EMITIR} = 'CK' OR {vEspecialCierres.EMITIR} = 'EF' OR {vEspecialCierres.EMITIR} = 'TE')" _
                          & " AND {vEspecialCierres.USERFOR} = '" & vUsuario _
                          & "' AND CDATE({vEspecialCierres.FECHAFORP}) " & fxFechaReportes(1)
       
        .SubreportToChange = "sbCKCajas"
        .Connect = glogon.ConectRPT
        .SelectionFormula = "CDATE({ASE_CK_CAJA.FECHA}) " & fxFechaReportes(1) _
                    & " AND {ASE_CK_CAJA.USUARIO} = '" & vUsuario & "'"
       
      
        .SubreportToChange = "sbFondos"
        .Connect = glogon.ConectRPT
        .SelectionFormula = "CDATE({FND_LIQUIDACION.FECHA}) " & fxFechaReportes(1) _
                    & " AND {FND_LIQUIDACION.USUARIO} = '" & vUsuario & "'"
       
       .PrintReport
       
       
       Case "05"
        .ReportFileName = SIFGlobal.fxSIFPathReportes("SIFDocumentoEspecialCierreBanco.rpt")
        .SelectionFormula = "cdate({CHEQUES.FECHA_EMISION}) " & fxFechaReportes(1) _
                          & " AND {CHEQUES.ID_BANCO} = " & cboBanco.ItemData(cboBanco.ListIndex)
       
      
        .SubreportToChange = "sbEstadistica"
        .Connect = glogon.ConectRPT
        .StoredProcParam(0) = cboBanco.ItemData(cboBanco.ListIndex)
        .StoredProcParam(1) = Format(dtpDesde.Value, "yyyy/mm/dd")
        
                  
       .PrintReport
       
    End Select
    
    Me.MousePointer = vbDefault
    Exit Sub
  
  End Select

  If chkTodasLasFechas.Value = vbUnchecked And Not optReportes(0).Value Then
    
    If chkFechaAnulacion.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND CDATE({ASE_DOCUMENTOS.FECHA_ANULACION}) " & fxFechaReportes(1) _
                & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
      .Formulas(4) = "fecha_anulacion = 'Fecha Anulación entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If
    
    If chkFechaEmision.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND CDATE({ASE_DOCUMENTOS.FECHA}) " & fxFechaReportes(1)
      .Formulas(4) = "fecha_emision = 'Fecha Emisión entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If
    
    If chkFechaTraspaso.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND CDATE({ASE_DOCUMENTOS.FECHA_TRASPASO}) " & fxFechaReportes(1) _
                & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
      .Formulas(4) = "fecha_traspaso = 'Fecha Traspaso entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If
   
   End If
   
    Select Case Mid(cboRepFiltro.Text, 1, 2)
      Case "03"
        .SelectionFormula = .SelectionFormula & " AND {ASE_DOCUMENTOS.USUARIO} = '" _
                          & glogon.Usuario & "'"
      Case "04"
        .SelectionFormula = .SelectionFormula & " AND {ASE_DOCUMENTOS.USUARIO} = '" _
                          & vUsuario & "'"
    End Select
   

   .PrintReport
   
End With

Me.MousePointer = vbDefault

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 tipo,id_documento from ase_documentos" _
           & " where Tipo = '" & fxTipoASEDoc(cboTipo.Text) & "'"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and id_documento > " & txtDocumento & " order by id_documento asc"
    Else
       strSQL = strSQL & " and id_documento < " & txtDocumento & " order by id_documento desc"
    End If
    
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If Not rs.EOF And Not rs.BOF Then
      txtDocumento = rs!Id_Documento
      Call txtDocumento_KeyPress(vbKeyReturn)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation


End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

 vModulo = 10 'Cuentas Corrientes
 vGrid.AppearanceStyle = fxGridStyle

 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 ssTab.Tab = 0
 
 dtpDesde.Value = fxFechaServidor
 dtpHasta.Value = dtpDesde.Value
 
 dtpAsientoInicio.Value = dtpDesde.Value
 dtpAsientoCorte.Value = dtpDesde.Value
 
 cboTipo.AddItem "Recibo"
 cboTipo.AddItem "Nota Credito"
 cboTipo.AddItem "Nota Debito"
 cboTipo.AddItem "Depositos"
 
 cboRepTipo.AddItem "Recibo"
 cboRepTipo.AddItem "Nota Credito"
 cboRepTipo.AddItem "Nota Debito"
 cboRepTipo.AddItem "Depositos"
 
 cboDepTipo.AddItem "Nota Credito"
 cboDepTipo.AddItem "Nota Debito"
 cboDepTipo.AddItem "Depositos"
 
 
 
 cboTipo.Text = "Recibo"
 cboRepTipo.Text = "Recibo"
 cboDepTipo.Text = "Nota Credito"
 
 
 
 cboRepFiltro.AddItem "01 - Sin Filtro de Usuario"
 cboRepFiltro.AddItem "02 - Agrupado por Usuario"
 cboRepFiltro.AddItem "03 - Solo Usuario Actual"
 cboRepFiltro.AddItem "04 - Usuario Específico"
 cboRepFiltro.AddItem "05 - Banco Específico"
 
 
 cboRepFiltro.Text = "01 - Sin Filtro de Usuario"
 
 cboBanco.Clear
 strSQL = "select id_banco,descripcion from Tes_Bancos where estado = 'A'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
  cboBanco.AddItem rs!Descripcion
  cboBanco.ItemData(cboBanco.NewIndex) = rs!id_banco
  rs.MoveNext
 Loop
 If rs.RecordCount > 0 Then
    rs.MoveFirst
    cboBanco.Text = rs!Descripcion
 End If
 rs.Close
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

 chkUsCtaTodos.Enabled = lswUsCtaAsignadas.Enabled

End Sub

Private Sub sbAsientoUnoAUno(vTipo As String, vDocumento As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vCuenta As String

On Error GoTo vError
 
 strSQL = "select CS_RE_CUENTA from ase_consecutivos"
 rs.Open strSQL, glogon.Conection, adOpenStatic
     vCuenta = rs!cs_re_cuenta
 rs.Close
    
'Borra el Asiento del Auxiliar
 strSQL = "DELETE ASE_ASIENTOS where tipo = '" & vTipo & "' and id_documento = " & vDocumento
 glogon.Conection.Execute strSQL

'Crea Nuevo Asiento en el Auxiliar
 strSQL = "Insert ASE_ASIENTOS(ID_DOCUMENTO, Tipo, RECAS_CUENTA, RECAS_MONTO,RECAS_DEBEHABER)" _
        & " VALUES(" & vDocumento & ",'" & vTipo & "','" & vCuenta & "',1,'D')"
 glogon.Conection.Execute strSQL
 
 strSQL = "Insert ASE_ASIENTOS(ID_DOCUMENTO, Tipo, RECAS_CUENTA, RECAS_MONTO,RECAS_DEBEHABER)" _
        & " VALUES(" & vDocumento & ",'" & vTipo & "','" & vCuenta & "',1,'H')"
 glogon.Conection.Execute strSQL
    
Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub


Private Sub sbAsientoReversion(rs As ADODB.Recordset)
Dim rs2 As New ADODB.Recordset, DH As String
Dim strSQL As String, intLinea As Integer
Dim vNumAsiento As String, vTipoAsiento As String

On Error GoTo CapturaError

 If fxValidaPeriodoAsiento(rs!Fecha) Then 'Verificar el Periodo Abierto en contabilidad
   'Crea Maestro
   vTipoAsiento = fxTipoAsientoDoc(rs!Tipo)
   vNumAsiento = "A." & rs!Tipo & "." & Format(rs!Id_Documento, "0000000000")
   
   strSQL = "insert CntX_asientos(cod_contabilidad,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,notas,balanceado,modulo)" _
          & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & Year(fxFechaServidor) & "," & Month(fxFechaServidor) _
          & ",'" & Format(fxFechaServidor, "yyyy/mm/dd") & "','" & rs!Concepto & "','" & rs!Cliente & "','S'," & vModulo & ")"
   glogon.Conection.Execute strSQL
    
    
    
    'Crea Detalle
    intLinea = 1
    rs2.CursorLocation = adUseServer
    rs2.Open "select * from ase_asientos where id_documento = " & rs!Id_Documento _
             & " and tipo = '" & rs!Tipo & "'", glogon.Conection, adOpenStatic
    Do While rs2.EOF = False
        strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta" _
            & ",monto_debito,monto_credito,documento,detalle,cod_unidad,cod_centro_costo,cod_divisa,tipo_cambio)" _
            & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & intLinea & ",'" & Trim(rs2!Recas_Cuenta) & "'," _
            & IIf((rs2!Recas_DebeHaber = "H"), rs2!recas_Monto, 0) & "," & IIf((rs2!Recas_DebeHaber = "H"), 0, rs2!recas_Monto) _
            & ",'" & rs!Tipo & "." & Format(rs2!Id_Documento, "00000000") & "','" & rs!Concepto & "','OC','','COL',1)"
        glogon.Conection.Execute strSQL
        intLinea = intLinea + 1
        rs2.MoveNext
    Loop
    rs2.Close

   'Ajusta el detalle del documento Original para Conciliacion
    strSQL = "insert into ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
           & " (select id_documento,tipo,recas_cuenta,recas_monto,case when recas_debehaber = 'H' then 'D' else 'H' end" _
           & " From ase_asientos where tipo = '" & rs!Tipo & "' and id_documento = " & rs!Id_Documento & ")"
    glogon.Conection.Execute strSQL

 Else
  MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado..."
 End If 'Periodo

Exit Sub

CapturaError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub sbAnulaDocumento()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vTipoDoc As String, vFecha As Date, iRespuesta As Integer
Dim vOperacion As Long

On Error GoTo vError

'Supuestos :
'           a. Solo se pueden anular recibos y del mismo día
'           b. Como son del mismo día la contabilidad no los ha registrado xq es t+1
'           c. Revisar que se encuente activo y no anulada (Evitar Duplicidad)

vTipoDoc = fxTipoASEDoc(cboTipo.Text)

If vTipoDoc <> "RE" Then
  MsgBox "Este Tipo de Documento no se puede Anular, debe recurrir a otro método", vbCritical
  Exit Sub
End If

vFecha = fxFechaServidor

strSQL = "Select * from ase_documentos where id_documento = " & txtDocumento _
       & " and tipo = '" & vTipoDoc & "' and fecha between '" & Format(vFecha, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(vFecha, "yyyy/mm/dd") & " 23:59:59' and estado <> 'A'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
   rs.Close
   
   MsgBox "No se puede Anular Documento Porque:" & vbCrLf _
          & " a. No es un recibo" & vbCrLf _
          & " b. No es un documento realizado en el día de hoy" & vbCrLf _
          & " c. ó Porque ya se encuentre anulado", vbExclamation
   
   Exit Sub
End If

   
'  ************************************
'  * Este Codigo NO Aplica con Axapta *
'  ************************************
   If rs!traspaso = "G" Then
     'Ya se generó a Contabilidad, verificar si fue durante el mes
     'y hacer asiento de reversión, de lo contrario informar al usuario
     If Year(rs!Fecha) = Year(vFecha) _
           And Month(rs!Fecha) = Month(vFecha) Then
       Call sbAsientoReversion(rs)
'       strInforma = " - Asiento de Reversión Registrado en contabilidad..."
     Else
'        strInforma = "Documento se emitió en un mes anterior, se tiene que reportar a contabilidad" _
               & " para corrección contable..."
     End If

   Else 'Traspaso

     Call sbAsientoUnoAUno(vTipoDoc, txtDocumento)

   End If
 
rs.Close

'Anula Registro del Documento
   strSQL = "update ase_documentos set estado = 'A',fecha_anulacion = '" _
          & Format(vFecha, "yyyy/mm/dd") & "',us_anula = '" _
          & glogon.Usuario & "' where id_documento = " & txtDocumento _
          & " and tipo = '" & vTipoDoc & "'"
   glogon.Conection.Execute strSQL
 
 
     'Anular Movimiento en REG_CREDITOS,CREDITOS_DT Y MOROSIDAD
       
'OJO              & "estado = 'A',fecult = " & fxFechaProcesoAnterior(rs2!fecult) & "," _


vOperacion = 0
'Rastreo en Abonos Ordinarios y Extraordinarios
strSQL = "select * from creditos_dt where estado = 'A' and tcon = '" _
       & fxTipoASENumero(vTipoDoc) & "' and ncon = '" & txtDocumento & "' and convert(varchar(30),id_solicitud) <> ncon"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
 Do While Not rs.EOF
   strSQL = "update reg_creditos set estado = 'A'" _
          & ",saldo = saldo + " & rs!amortiza & ",saldo_mes = saldo_mes + " & rs!amortiza _
          & ",amortiza = amortiza - " & rs!amortiza & ",interesc = interesc - " & rs!intcp _
          & " where id_solicitud = " & rs!Id_solicitud
   glogon.Conection.Execute strSQL
   
   vOperacion = rs!Id_solicitud
   
   rs.MoveNext
  Loop
   
   strSQL = "delete Creditos_dt where tcon = " & fxTipoASENumero(vTipoDoc) & " and ncon = '" & txtDocumento _
          & "' and convert(varchar(30),id_solicitud) <> ncon"
   glogon.Conection.Execute strSQL
   
End If
rs.Close
       
       
'Busca en Morosidad
strSQL = "select id_solicitud,coalesce(sum(abintc),0) as intc, coalesce(sum(abintm),0) as intm" _
       & ",coalesce(sum(abamortiza),0) as amortiza from morosidad" _
       & " where estado = 'C' and tcon = '" & fxTipoASENumero(vTipoDoc) & "' and ncon = '" & txtDocumento _
       & "' group by id_solicitud"

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   vOperacion = rs!Id_solicitud
   
   strSQL = "update morosidad set estado = 'A',abintc=0,abintm=0" _
          & ",abamortiza = 0 where tcon = '" _
          & fxTipoASENumero(vTipoDoc) & "' and ncon = '" & txtDocumento & "'"
   glogon.Conection.Execute strSQL
   
   strSQL = "update reg_creditos set estado = 'A',saldo = saldo + " & rs!amortiza _
          & ",saldo_mes = saldo_mes + " & rs!amortiza & ",amortiza = amortiza - " & rs!amortiza _
          & ",interesc = interesc - " & rs!intc + rs!intm _
          & " where id_solicitud = " & rs!Id_solicitud
   glogon.Conection.Execute strSQL
End If
rs.Close
         
If vOperacion > 0 Then
  strSQL = "update reg_creditos set fecUlt = dbo.fxCrdFechaProcUltMov(id_solicitud,getdate())" _
        & " where id_solicitud = " & vOperacion
  glogon.Conection.Execute strSQL
End If
         
Call Bitacora("Anula", cboTipo.Text & " #" & txtDocumento)
Call sbCargaDocumento(cboTipo.Text, txtDocumento)

MsgBox "- Documento Anulado " & vbCrLf & "Revisar la Fecha de Ultimo Pago, ya que esta no fue solucionada por el sistema", vbInformation

Exit Sub

vError:
    MsgBox Err.Description, vbCritical

End Sub


Private Sub sbAsientoTipoDiario(pTipoDocumento As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim intLinea As Integer, DH As String, vTipoAsiento As String
Dim rsTmp As New ADODB.Recordset, vFecha As Date, vNumAsiento As String


lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

On Error GoTo vError

'Determina el tipo de Asiento para Contabilidad
vTipoAsiento = fxTipoAsientoDoc(pTipoDocumento)

'Inicia Transaccion
glogon.Conection.BeginTrans

'Sacar los Documentos de Inicio y Corte

strSQL = "select year(fecha) as Anio,month(fecha) as Mes,day(fecha) as Dia,coalesce(min(id_documento),0) as Inicio, coalesce(max(id_documento),0) as Corte" _
       & " from ase_documentos where estado = 'I' and traspaso = 'P'" _
       & " and tipo in('" & pTipoDocumento & "') and Fecha between '" & Format(dtpAsientoInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpAsientoCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & " group by year(fecha),month(fecha),day(fecha)"
       
       
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
lblEstatus.Caption = "Procesando Asientos..." & pTipoDocumento
lblEstatus.Refresh

Do While Not rs.EOF
 vFecha = rs!Anio & "/" & rs!Mes & "/" & rs!Dia
 
 If fxValidaPeriodoAsiento(vFecha) Then  'Verificar el Periodo Abierto en contabilidad
    
    
    lblEstatus.Caption = "Procesando Asientos..." & pTipoDocumento & "[" & rs!inicio & "-" & rs!corte & "]"
    lblEstatus.Refresh
    
    
    vNumAsiento = "SIF." & pTipoDocumento & ".A" & rs!Anio & "M" & rs!Mes & "D" & rs!Dia
    
    
    'Crea el Maestro de Asiento
    strSQL = "insert CntX_asientos(cod_contabilidad,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado,modulo)" _
           & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & rs!Anio & "," & rs!Mes _
           & ",'" & Format(vFecha, "yyyy/mm/dd") & "','ASIENTO DE DIARIO','S'," & vModulo & ")"
    glogon.Conection.Execute strSQL
        
    intLinea = 1
    
    'Detalle del Asiento de diario
    
    strSQL = "select A.*,D.concepto,D.cliente,D.linea1,D.linea2" _
           & " from ase_documentos D inner join ase_asientos A on D.id_documento =  A.id_documento and D.tipo = A.tipo" _
           & " where A.Tipo = '" & pTipoDocumento & "' and A.id_documento between " & rs!inicio & " and " & rs!corte
    rsTmp.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rsTmp.EOF
        If UCase(rsTmp!Recas_DebeHaber) = "H" Then  'dc - dh
          DH = "C"
          strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito" _
                 & ",monto_credito,detalle,documento,cod_unidad,cod_centro_Costo,cod_divisa,tipo_cambio)" _
                 & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & intLinea & ",'" & Trim(rsTmp!Recas_Cuenta) _
                 & "',0," & rsTmp!recas_Monto & ",'" & rsTmp!Tipo & "." & Format(rsTmp!Id_Documento, "0000000000") _
                 & "','" & rsTmp!Concepto & "','OC','','COL',1)"
        
        Else
          DH = rsTmp!Recas_DebeHaber
          strSQL = "insert CntX_asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito" _
                 & ",monto_credito,detalle,documento,cod_unidad,cod_centro_Costo,cod_divisa,tipo_cambio)" _
                 & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & intLinea & ",'" & Trim(rsTmp!Recas_Cuenta) _
                 & "'," & rsTmp!recas_Monto & ",0,'" & rsTmp!Tipo & "." & Format(rsTmp!Id_Documento, "0000000000") _
                 & "','" & rsTmp!Concepto & "','OC','','COL',1)"
        End If
        
        glogon.Conection.Execute strSQL
        intLinea = intLinea + 1
      
      rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    
    'Actualiza Tabla de ASE_DOCUMENTOS
    strSQL = "Update ase_documentos set traspaso = 'G', FECHA_TRASPASO = getdate()" _
            & ",us_traspaso = '" & glogon.Usuario & "' where id_documento between " & rs!inicio _
            & " and " & rs!corte & " and tipo = '" & pTipoDocumento & "'"
    
    glogon.Conection.Execute strSQL
    
    
 Else 'Verificacion del periodo
  
   MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado...", vbInformation
  
  
 End If 'Verificacion del periodo
 
 If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

'Cierra Transaccion
glogon.Conection.CommitTrans


lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1

Call Bitacora("Aplica", "Asientos del Control de Documentos :" & pTipoDocumento)

Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    prgBar.Value = 1
    Me.MousePointer = vbDefault
    glogon.Conection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
End Sub


Private Sub sbAsientoIndividual()
Dim rs As New ADODB.Recordset, strSQL As String
Dim intLinea As Integer, DH As String, strDocumentos As String
Dim rs2 As New ADODB.Recordset, vTipoAsiento As String
Dim vFecha As Date, vDetalle As String

Me.MousePointer = vbHourglass
Me.fraTraspaso.Visible = True

lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

On Error GoTo vError

strDocumentos = ""
vFecha = fxFechaServidor


If chkAS_Depositos.Value = vbChecked Then strDocumentos = "'DP'"

If chkAS_NC.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos + ",'NC'"
  Else
    strDocumentos = "'NC'"
  End If
End If

If chkAS_ND.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos + ",'ND'"
  Else
    strDocumentos = "'ND'"
  End If
End If

If chkAS_Recibos.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos + ",'RE'"
  Else
    strDocumentos = "'RE'"
  End If
End If


If Len(strDocumentos) = 0 Then
  Me.MousePointer = vbDefault
  Exit Sub
End If


strSQL = "select * from ase_documentos where estado = 'I' and traspaso = 'P'" _
       & " and tipo in(" & strDocumentos & ") and Fecha between '" & Format(dtpAsientoInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpAsientoCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
lblEstatus.Caption = "Procesando Asientos..."
lblEstatus.Refresh

Do While Not rs.EOF
 If fxValidaPeriodoAsiento(rs!Fecha) Then 'Verificar el Periodo Abierto en contabilidad
    'Crea Maestro
   vTipoAsiento = fxTipoAsientoDoc(rs!Tipo)
   strSQL = "insert CntX_Asientos(cod_contabilidad,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado,modulo)" _
          & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & Format(rs!Id_Documento, "00000000") & "'," & Year(rs!Fecha) & "," & Month(rs!Fecha) _
          & ",'" & Format(rs!Fecha, "yyyy/mm/dd") & "','" & rs!Concepto & "','S'," & vModulo & ")"
   glogon.Conection.Execute strSQL
    
    'Crea Detalle
    intLinea = 1
    rs2.CursorLocation = adUseServer
    rs2.Open "select * from ase_asientos where id_documento = " & rs!Id_Documento _
             & " and tipo = '" & rs!Tipo & "'", glogon.Conection, adOpenStatic
    Do While Not rs2.EOF
        If UCase(rs2!Recas_DebeHaber) = "H" Then  'dc - dh
          DH = "C"
        Else
          DH = rs2!Recas_DebeHaber
        End If
        'Ahora se pone en el detalle de la cuenta el numero de deposito y luego
        'Lo que alcance del concepto
        vDetalle = ""
        If IsNull(rs!dp) Then
          vDetalle = rs!Concepto
        Else
          If Trim(rs!dp) = "" Then
              vDetalle = rs!Concepto
          Else
            vDetalle = "DP." & Trim(rs!dp) & " - " & rs!Concepto
          End If
        End If
        vDetalle = Mid(vDetalle, 1, 59)
        
        If DH = "C" Then 'Acredita
            strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                & ",detalle,documento,cod_unidad,cod_divisa,Tipo_Cambio,cod_centro_costo)" _
                & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & Format(rs2!Id_Documento, "00000000") & "'," & intLinea & "," & Trim(rs2!Recas_Cuenta) _
                & ",0," & rs2!recas_Monto & ",'" & vDetalle & "','" & Format(rs2!Id_Documento, "00000000") _
                & "','OC','COL',1,'')"
        Else 'Debita
            strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                & ",detalle,documento,cod_unidad,cod_divisa,Tipo_Cambio,cod_centro_costo)" _
                & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & Format(rs2!Id_Documento, "00000000") & "'," & intLinea & "," & Trim(rs2!Recas_Cuenta) _
                & "," & rs2!recas_Monto & ",0,'" & vDetalle & "','" & Format(rs2!Id_Documento, "00000000") _
                & "','OC','COL',1,'')"
        End If
        If Len(Trim(rs2!Recas_Cuenta)) > 0 Then
          glogon.Conection.Execute strSQL
          intLinea = intLinea + 1
        End If
        rs2.MoveNext
    Loop
    rs2.Close
    
    'Actualizar el estado del recibo
    strSQL = "Update ase_documentos set traspaso = 'G', FECHA_TRASPASO = '" & Format(vFecha, "yyyy/mm/dd") _
            & "',us_traspaso = '" & glogon.Usuario & "' where id_documento = " & rs!Id_Documento _
            & " and tipo = '" & rs!Tipo & "'"
    glogon.Conection.Execute strSQL
 Else
  MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado...", vbInformation
 End If 'Periodo
 
 If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1

Call Bitacora("Aplica", "Asientos del Control de Documentos")

MsgBox "Se realizó el pase de asientos a contabilidad ", vbInformation
Me.MousePointer = vbDefault
Me.fraTraspaso.Visible = False

Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    prgBar.Value = 1
    Me.MousePointer = vbDefault
    Me.fraTraspaso.Visible = False
    MsgBox Err.Description, vbCritical
End Sub

Private Sub sbConfiguracion()
Dim rs As New ADODB.Recordset, strSQL As String

fraConfiguracion.Visible = True
On Error Resume Next

medND.Format = GLOBALES.gstrMascara
medNC.Format = GLOBALES.gstrMascara
medDP.Format = GLOBALES.gstrMascara
medRE.Format = GLOBALES.gstrMascara


rs.Open "select * from ASE_CONSECUTIVOS", glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  'INSERTAR
  strSQL = "insert ase_consecutivos(cs_nota_credito,cs_nota_debito,cs_deposito" _
         & ",cs_recibo,cs_utilizar_recibo,cs_utilizar_reciboFlat,cs_utilizar_CuentasAuto" _
         & ",cs_nd_cuenta,cs_nc_cuenta,cs_dp_cuenta" _
         & ",cs_re_cuenta,cs_nc_asiento,cs_nd_asiento,cs_dp_asiento,cs_re_asiento) " _
         & "values(0,0,0,0,'N','N','N','','','','','','','','')"
  glogon.Conection.Execute strSQL
End If
rs.Close

rs.Open "select * from ASE_CONSECUTIVOS", glogon.Conection, adOpenStatic

medND.Text = Trim(rs!cs_nd_cuenta)
medNC.Text = Trim(rs!cs_nc_cuenta)
medDP.Text = Trim(rs!cs_dp_cuenta)
medRE.Text = Trim(rs!cs_re_cuenta)

txtConRECta = fxgCntCuentaDesc(medRE.Text)
txtConNCCta = fxgCntCuentaDesc(medNC.Text)
txtConNDCta = fxgCntCuentaDesc(medND.Text)
txtConDPCta = fxgCntCuentaDesc(medDP.Text)


txtTA_ND = Trim(rs!cs_nd_asiento)
txtTA_NC = Trim(rs!cs_nc_asiento)
txtTA_DP = Trim(rs!cs_dp_asiento)
txtTA_RE = Trim(rs!cs_re_asiento)

txtConRETA = fxgCntTipoAsientoDesc(txtTA_RE)
txtConNDTA = fxgCntTipoAsientoDesc(txtTA_ND)
txtConNCTA = fxgCntTipoAsientoDesc(txtTA_NC)
txtConDPTA = fxgCntTipoAsientoDesc(txtTA_DP)

txtID_NC = rs!cs_nota_credito
txtID_ND = rs!cs_nota_debito
txtID_DP = rs!cs_deposito
txtID_RE = rs!cs_recibo


chkUtilizaRecibo.Value = IIf((UCase(rs!cs_utilizar_recibo) = "S"), vbChecked, 0)
chkReciboFlat.Value = IIf((UCase(rs!cs_utilizar_reciboflat) = "S"), vbChecked, 0)
chkUtilizaCtaAutorizadas.Value = IIf((UCase(rs!cs_utilizar_cuentasAuto) = "S"), vbChecked, 0)

rs.Close


End Sub



Private Sub imgAnular_Click()
Dim I  As Byte

I = MsgBox("Esta seguro de que desea anular el " & cboTipo.Text & " #" & txtDocumento, vbYesNo)

If I = vbYes Then
 Call sbAnulaDocumento
End If

End Sub

Private Sub imgReImpresion_Click()
Dim strSQL As String, rs As New ADODB.Recordset

'Verificar si es un usuario autorizado, de lo contrario
'Solicitar Login de un usuario Autorizado.

On Error GoTo vError

strSQL = "select coalesce(count(*),0) as Existe from ase_usr_Autoriza where usuario = '" _
       & glogon.Usuario & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs!existe = 0 Then
  'Solicitar Usuario Autorizado
  frmCC_DocAutoImprime.Show
Else
  Call sbImprimeRecibo(txtDocumento, fxTipoASEDoc(cboTipo.Text), True)
End If
rs.Close

vError:


End Sub


Private Sub lswUsCodigos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert ase_usr_codigos(codigo) values('" & Item.Text & "')"
Else
  strSQL = "delete ase_usr_codigos where codigo = '" & Item.Text & "'"
End If
glogon.Conection.Execute strSQL

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub



Private Sub lswUsCtaAsignadas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert ase_usr_CtaAuto(cod_cuenta,usuario) values('" & Item.Text _
         & "','" & lblUSCtaActual.Caption & "')"
Else
  strSQL = "delete ase_usr_CtaAuto where cod_cuenta = '" & Item.Text _
         & "' and usuario = '" & lblUSCtaActual.Caption & "'"
End If

glogon.Conection.Execute strSQL

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub lswUsCtaAuto_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX  As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lblUSCtaActual.Caption = lswUsCtaAuto.SelectedItem.Text

strSQL = "select A.cod_cuenta,C.descripcion,A.apl_creditos,A.apl_retencion" _
       & ",A.apl_fondos,A.apl_patrimonio,U.id" _
       & " from ase_cta_autorizadas A inner join CntX_cuentas C on A.cod_cuenta = C.cod_cuenta" _
       & " and C.cod_contabilidad = " & GLOBALES.gEnlace _
       & " left join ase_usr_CtaAuto U on A.cod_cuenta = U.cod_cuenta" _
       & " and U.usuario = '" & lswUsCtaAuto.SelectedItem.Text & "'" _
       & " order by U.id desc,A.cod_cuenta"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

lswUsCtaAsignadas.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lswUsCtaAsignadas.ListItems.Add(, , rs!cod_cuenta)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = IIf((rs!apl_creditos = 1), "S", "N")
     itmX.SubItems(3) = IIf((rs!apl_retencion = 1), "S", "N")
     itmX.SubItems(4) = IIf((rs!apl_fondos = 1), "S", "N")
     itmX.SubItems(5) = IIf((rs!apl_patrimonio = 1), "S", "N")
 itmX.Checked = IIf(IsNull(rs!idConAse), vbUnchecked, vbChecked)
     
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
  Resume
End Sub

Private Sub lswUsReImprimen_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert ase_usr_Autoriza(usuario) values('" & Item.Text & "')"
Else
  strSQL = "delete ase_usr_Autoriza where usuario = '" & Item.Text & "'"
End If

glogon.Conection.Execute strSQL

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub medDP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medDP, "C")
If KeyCode = vbKeyReturn Then txtTA_DP.SetFocus
End Sub

Private Sub medNC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medNC, "C")
If KeyCode = vbKeyReturn Then txtTA_NC.SetFocus
End Sub



Private Sub medND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medND, "C")
If KeyCode = vbKeyReturn Then txtTA_ND.SetFocus
End Sub

Private Sub medRE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medRE, "C")
If KeyCode = vbKeyReturn Then txtTA_RE.SetFocus
End Sub

Private Sub tlbRecibos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer
Select Case Button.Key
  Case "anular"
     iRespuesta = MsgBox("Esta seguro de que desea anular el " & cboTipo.Text & " #" & txtDocumento, vbYesNo)
     If iRespuesta = vbYes Then
      Call sbAnulaDocumento
     End If
  Case "traspaso"
      fraTraspaso.Visible = True
  Case "configuracion"
     Call sbConfiguracion
  Case "reportes"
    fraReportes.Visible = True
End Select
End Sub

Private Sub tlbRecibos_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String, x As New clsImpresoras
Dim vDriver, vTipo As String

vTipo = fxTipoASEDoc(cboTipo.Text)

Select Case ButtonMenu.Key
 Case "repReImpresion"
                           
       With frmContenedor.Crt
          .Reset
          If vTipo = "RE" Then
            x.TipoImpresora = Recibos
            x.Reset
            .PrinterDriver = x.Controlador
            .PrinterName = x.Nombre
            .PrinterPort = x.Puerto
            .Destination = crptToPrinter
            .ReportFileName = SIFGlobal.fxSIFPathReportes("Documento.rpt")
          Else
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowState = crptMaximized
            .WindowTitle = "Reportes de Control de Documentos"
            .ReportFileName = SIFGlobal.fxSIFPathReportes("DocumentoNotas.rpt")
          End If
          .SelectionFormula = "{ASE_DOCUMENTOS.ID_DOCUMENTO} = " & Trim(txtDocumento) _
                            & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .PrintReport
       End With
       
       strSQL = "Update ASE_DOCUMENTOS set Estado='I' Where Id_DOCUMENTO=" & Trim(txtDocumento) _
                & " AND TIPO = '" & vTipo & "'"
       glogon.Conection.Execute strSQL
       
 Case "repRecibosFecha"
   fraReportes.Visible = True
End Select
End Sub

Private Sub sbCargaDocumento(vTipo As String, lngDocumento As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListItem, curDebe As Currency, curHaber As Currency
Dim strTipo As String

curDebe = 0
curHaber = 0

strTipo = fxTipoASEDoc(vTipo)

On Error Resume Next

strSQL = "select * from ase_Documentos where id_Documento = " & lngDocumento _
        & " and tipo = '" & strTipo & "'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "No se encontró Documento", vbCritical
  Exit Sub
End If
 
 txtBeneficiario = IIf(IsNull(rs!Cliente), "", rs!Cliente)
 
 Select Case rs!Estado
   Case "I" 'Impreso
      txtEstado = "Impreso"
   Case "P" 'Pendiente
      txtEstado = "Pendiente"
   Case "A" 'Anulado
      txtEstado = "Anulado"
 End Select
 
 txtFechaAnula = IIf(IsNull(rs!fecha_anulacion), "", rs!fecha_anulacion)
 txtFechaGenera = rs!Fecha
 txtFechaTraspasa = IIf(IsNull(rs!fecha_traspaso), "", rs!fecha_traspaso)
 txtConcepto = rs!Concepto
 txtMonto = Format(rs!Monto, "Standard")
 Select Case rs!tipo_pago
   Case "E"
     txtPago = "Efectivo"
   Case "M"
     txtPago = "Mixto"
   Case "C"
     txtPago = "Cheque"
   Case "D"
     txtPago = "Depósito"
 End Select

 txtTipo = rs!tipo_pago
 txtDetalle = IIf(IsNull(rs!Detalle), "", rs!Detalle)
 txtUS_Anula = IIf(IsNull(rs!us_anula), "", rs!us_anula)
 txtUS_Traspasa = IIf(IsNull(rs!us_traspaso), "", rs!us_traspaso)
 txtUS_Genera = rs!Usuario
rs.Close

strSQL = "select * from ase_asientos where id_documento=" & lngDocumento _
        & " and tipo = '" & strTipo & "'  ORDER BY Id_Ase_Asiento"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

lswAsiento.ListItems.Clear
With lswAsiento
Do While Not rs.EOF
 Set itmX = .ListItems.Add(.ListItems.Count + 1, , Format(rs!Recas_Cuenta, GLOBALES.gstrMascara))
   itmX.SubItems(1) = fxgCntCuentaDesc(Trim(rs!Recas_Cuenta))
   If rs!Recas_DebeHaber = "D" Then
      itmX.SubItems(2) = Format(rs!recas_Monto, "Standard")
      curDebe = curDebe + rs!recas_Monto
   Else
      itmX.SubItems(3) = Format(rs!recas_Monto, "Standard")
      curHaber = curHaber + rs!recas_Monto
   End If
 rs.MoveNext
Loop
 Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
  itmX.SubItems(2) = "_____________"
  itmX.SubItems(3) = "_____________"
 
 Set itmX = .ListItems.Add(.ListItems.Count + 1, , "TOTALES")
  itmX.SubItems(2) = Format(curDebe, "Standard")
  itmX.SubItems(3) = Format(curHaber, "Standard")
End With

rs.Close

End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

Select Case ssTab.Tab
 Case 0 'Documentos
 Case 1 'Reportes
 Case 2 'Depositos
 Case 3 'Usuarios
   ssTabAux.Tab = 0
   Call sbAuxUsrAutoReimpresion
 Case 4 'Traslados
 Case 5 'Configuracion
   Call sbConfiguracion
 Case 6 'Cuentas Autorizadas
   strSQL = "select A.cod_cuenta,C.descripcion,A.apl_creditos,A.apl_retencion" _
          & ",A.apl_fondos,A.apl_patrimonio" _
          & " from ase_cta_autorizadas A inner join CntX_cuentas C on A.cod_cuenta = C.cod_cuenta" _
          & " and C.cod_contabilidad = " & GLOBALES.gEnlace _
          & " order by A.cod_cuenta"
   Call sbCargaGrid(vGrid, 6, strSQL)
   
End Select

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, I As Integer
Dim lng As Long, vTemp() As Variant, x As Integer

On Error GoTo vError

ReDim vTemp(vGrid.MaxCols) As Variant

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  I = fxGuardarGrid
  If I = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


'Consular Cuenta
If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   vGrid.Text = gCuenta
   vGrid.Col = 2
   vGrid.Text = fxgCntCuentaDesc(gCuenta)
   
    
   
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  
  strSQL = "delete ase_usr_CtaAuto where cod_cuenta = '" & vGrid.Text & "'"
  glogon.Conection.Execute strSQL
  
  strSQL = "delete ase_cta_autorizadas where cod_cuenta = '" & vGrid.Text & "'"
  glogon.Conection.Execute strSQL
  
  vGrid.Col = vGrid.MaxCols
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.Col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
End If


Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Function fxGuardarGrid() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarGrid = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from ase_cta_autorizadas " _
       & " where cod_cuenta = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into ase_cta_autorizadas(cod_cuenta,apl_creditos,apl_retencion,apl_fondos,apl_patrimonio) values('" _
         & UCase(vGrid.Text) & "',"
  vGrid.Col = 3
  strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ","
  vGrid.Col = 4
  strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ","
  vGrid.Col = 5
  strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ","
  vGrid.Col = 6
  strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ")"

  glogon.Conection.Execute strSQL

  vGrid.Col = 2
  ' Call sbBitacora("Registra", "Caracteristica : " & vGrid.Text & "- " & IIf(IsNull(rs!ultimo), 0, rs!ultimo))

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update ase_cta_autorizadas set apl_creditos = "
 vGrid.Col = 3
 strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ",apl_retencion = "
 vGrid.Col = 4
 strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ",apl_fondos = "
 vGrid.Col = 5
 strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ",apl_patrimonio = "
 vGrid.Col = 6
 strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & " where cod_cuenta = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 glogon.Conection.Execute strSQL

End If
rs.Close

fxGuardarGrid = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical
 
End Function



Private Sub sbAuxUsrAutoReimpresion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select U.nombre,U.descripcion,A.usuario" _
       & " from usuarios U left join ase_usr_Autoriza A on U.nombre = A.usuario and U.estado = 'A'" _
       & " order by A.usuario desc"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

lswUsReImprimen.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lswUsReImprimen.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
 itmX.Checked = IIf(IsNull(rs!Usuario), vbUnchecked, vbChecked)
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub

Private Sub sbAuxUsrLista()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select nombre,descripcion from usuarios where estado = 'A'"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

lswUsCtaAuto.ListItems.Clear
lswUsCtaAsignadas.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswUsCtaAuto.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub


Private Sub sbAuxUsrCodigosForm()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.Codigo,C.descripcion,A.codigo as CodX" _
       & " from catalogo C left join ase_usr_codigos A on C.codigo = A.codigo" _
       & " order by A.codigo desc"
       
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

lswUsCodigos.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lswUsCodigos.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!Descripcion
 itmX.Checked = IIf(IsNull(rs!CodX), vbUnchecked, vbChecked)
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub

Private Sub ssTabAux_Click(PreviousTab As Integer)

Select Case ssTabAux.Tab
  Case 0 'Usuarios Autorizados a Reimpresion
     Call sbAuxUsrAutoReimpresion
  Case 1 'Cuentas Autorizadas por Usuarios
     Call sbAuxUsrLista
  Case 2 'Codigos de Formalizacion
     Call sbAuxUsrCodigosForm
End Select

End Sub

Private Sub txtDepDoc_Change()
txtDepActual = ""
txtDepNuevo = ""
End Sub

Private Sub txtDepDoc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If KeyCode = vbKeyReturn Then
  strSQL = "select dp from ase_documentos where tipo = '" _
         & fxTipoASEDoc(cboDepTipo.Text) & "' and id_documento = " _
         & txtDepDoc
  rs.Open strSQL, glogon.Conection, adOpenStatic
  If Not rs.EOF And Not rs.BOF Then
     txtDepActual = rs!dp & ""
  End If
  rs.Close
End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 


End Sub

Private Sub txtDocumento_Change()
 Call sbLimpiaDatos
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 Call sbCargaDocumento(cboTipo.Text, txtDocumento)
End If
End Sub

Private Sub txtID_DP_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Then cmdGuardaConfiguracion.SetFocus
vError:
End Sub

Private Sub txtID_NC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then medND.SetFocus
End Sub

Private Sub txtID_ND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then medDP.SetFocus
End Sub

Private Sub txtID_RE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then medNC.SetFocus
End Sub


Private Sub txtNumDp_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 'Nada
Else
 Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select id_documento,tipo,concepto,cliente,fecha,dp from ase_documentos where dp like '" & txtNumDp & "%'"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
lswDP.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswDP.ListItems.Add(, , rs!Tipo)
   itmX.SubItems(1) = rs!Id_Documento
   itmX.SubItems(2) = rs!Concepto & ""
   itmX.SubItems(3) = rs!Cliente & ""
   itmX.SubItems(4) = Format(rs!Fecha, "dd/mm/yyyy")
   itmX.SubItems(5) = rs!dp & ""
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

End Sub

Private Sub txtTA_DP_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_DP, "T")
If KeyCode = vbKeyReturn Then cmdGuardaConfiguracion.SetFocus
vError:
End Sub

Private Sub txtTA_NC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_NC, "T")
If KeyCode = vbKeyReturn Then txtID_NC.SetFocus
End Sub

Private Sub txtTA_ND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_ND, "T")
If KeyCode = vbKeyReturn Then txtID_ND.SetFocus
End Sub

Private Sub txtTA_RE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_RE, "T")
If KeyCode = vbKeyReturn Then txtID_RE.SetFocus
End Sub

