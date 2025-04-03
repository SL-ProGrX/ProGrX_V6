VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAF_Activos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activos Fijos"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   6525
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario que Registro Activo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3422
            MinWidth        =   3422
            Object.ToolTipText     =   "Fecha de Registro Real"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1482
            MinWidth        =   1482
            Object.ToolTipText     =   "Ultimo Periodo Depreciado"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2541
            MinWidth        =   2541
            Object.ToolTipText     =   "Depreciación Acumulada"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2187
            MinWidth        =   2187
            Object.ToolTipText     =   "Depreciación del Mes"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7380
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivosPrincipal.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivosPrincipal.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivosPrincipal.frx":0654
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivosPrincipal.frx":0BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivosPrincipal.frx":0CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6015
      Left            =   60
      TabIndex        =   10
      Top             =   480
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmActivosPrincipal.frx":0E12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label20"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label14(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label16"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label21"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label19"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label17"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label8(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label8(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblEstado"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dtpInstalacion"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "dtpAdquisicion"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtNotas"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cboVU"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtVU"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cbo"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtValorRescate"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtValorHistorico"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtCodigo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtDescripcion"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtDocCompra"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtProveedor"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtDepartamento"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtSeccion"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cboTipo"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtUDProducidas"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtUDAnio"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmActivosPrincipal.frx":0E2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "Label12(0)"
      Tab(1).Control(2)=   "Label12(1)"
      Tab(1).Control(3)=   "Label12(2)"
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(5)=   "lsw"
      Tab(1).Control(6)=   "txtModelo"
      Tab(1).Control(7)=   "txtSerie"
      Tab(1).Control(8)=   "txtMarca"
      Tab(1).Control(9)=   "txtOtrasSenas"
      Tab(1).Control(10)=   "chkResponsables"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Modificaciones"
      TabPicture(2)   =   "frmActivosPrincipal.frx":0E4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lswMod"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Historico"
      TabPicture(3)   =   "frmActivosPrincipal.frx":0E66
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblCap1(0)"
      Tab(3).Control(1)=   "lswHistorico"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Composición"
      TabPicture(4)   =   "frmActivosPrincipal.frx":0E82
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblCap1(1)"
      Tab(4).Control(1)=   "lswCompo"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Polizas"
      TabPicture(5)   =   "frmActivosPrincipal.frx":0E9E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblCap1(2)"
      Tab(5).Control(1)=   "lswPolizas"
      Tab(5).ControlCount=   2
      Begin VB.TextBox txtUDAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1380
         TabIndex        =   47
         Top             =   3120
         Width           =   1755
      End
      Begin VB.TextBox txtUDProducidas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         TabIndex        =   45
         Top             =   3120
         Width           =   1755
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmActivosPrincipal.frx":0EBA
         Left            =   1380
         List            =   "frmActivosPrincipal.frx":0EC4
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1680
         Width           =   5175
      End
      Begin VB.CheckBox chkResponsables 
         Caption         =   "Listado General"
         Height          =   255
         Left            =   -73320
         TabIndex        =   42
         Top             =   2800
         Width           =   1575
      End
      Begin VB.TextBox txtSeccion 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   41
         ToolTipText     =   "Presione F4 para consultar / Cualquier Tecla"
         Top             =   4800
         Width           =   5145
      End
      Begin VB.TextBox txtDepartamento 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   40
         ToolTipText     =   "Presione F4 para consultar / Cualquier Tecla"
         Top             =   4440
         Width           =   5145
      End
      Begin VB.TextBox txtOtrasSenas 
         Height          =   1215
         Left            =   -73320
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txtMarca 
         Height          =   315
         Left            =   -73320
         TabIndex        =   26
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtSerie 
         Height          =   315
         Left            =   -73320
         TabIndex        =   25
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox txtModelo 
         Height          =   315
         Left            =   -73320
         TabIndex        =   24
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox txtProveedor 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Presione F4 para consultar / Cualquier Tecla"
         Top             =   5160
         Width           =   5145
      End
      Begin VB.TextBox txtDocCompra 
         Height          =   315
         Left            =   1380
         TabIndex        =   17
         Top             =   5520
         Width           =   5145
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   1320
         Width           =   5145
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   2220
         TabIndex        =   0
         Top             =   600
         Width           =   1425
      End
      Begin VB.TextBox txtValorHistorico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   2040
         Width           =   1755
      End
      Begin VB.TextBox txtValorRescate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Top             =   2040
         Width           =   1755
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         ItemData        =   "frmActivosPrincipal.frx":0ED5
         Left            =   4800
         List            =   "frmActivosPrincipal.frx":0EDF
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2760
         Width           =   1755
      End
      Begin VB.TextBox txtVU 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Top             =   2760
         Width           =   705
      End
      Begin VB.ComboBox cboVU 
         Height          =   315
         ItemData        =   "frmActivosPrincipal.frx":0F01
         Left            =   2130
         List            =   "frmActivosPrincipal.frx":0F0B
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2760
         Width           =   990
      End
      Begin VB.TextBox txtNotas 
         Height          =   795
         Left            =   1380
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3600
         Width           =   5145
      End
      Begin MSComCtl2.DTPicker dtpAdquisicion 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   2400
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   54132739
         CurrentDate     =   36338
      End
      Begin MSComCtl2.DTPicker dtpInstalacion 
         Height          =   315
         Left            =   4800
         TabIndex        =   5
         Top             =   2400
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   54132739
         CurrentDate     =   36338
      End
      Begin MSComctlLib.ListView lswMod 
         Height          =   5475
         Left            =   -74940
         TabIndex        =   16
         Top             =   420
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   9657
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   3775
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Justificacion"
            Object.Width           =   7832
         EndProperty
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2715
         Left            =   -73320
         TabIndex        =   29
         Top             =   3120
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   4789
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ListView lswHistorico 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   49
         Top             =   840
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   8811
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Año"
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Mes"
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Dep.Acumulada"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Dep.Mensual"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswCompo 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   51
         Top             =   840
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   8811
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Placa"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Periodo"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Dep.Acumulada"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Dep.Mensual"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Adquisicion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Descripción"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fecha Registro"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswPolizas 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   54
         Top             =   840
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   8811
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   3599
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Número"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Inicia"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Vence"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Descripción"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cod.Poliza"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblCap1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Listado de Polizas (Seguros) Asignados al Activo"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   -74880
         TabIndex        =   55
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label lblEstado 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  xxx"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   3600
         TabIndex        =   53
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblCap1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Indica las depreciaciones del Activo y sus componentes o mejoras realizadas"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label lblCap1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Historico de Cierres de Depreciaciones Registradas al Activo"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   50
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label8 
         Caption         =   "Ud. x Año"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   48
         Top             =   3200
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "Ud. a Producir"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   44
         Top             =   3200
         Width           =   1185
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   6600
         X2              =   240
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   6600
         X2              =   240
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   3630
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Vida útil"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2790
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Adquisición"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Valor Histórico"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label19 
         Caption         =   "Tipo Activo"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label21 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   270
         TabIndex        =   34
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Sección"
         Height          =   165
         Left            =   240
         TabIndex        =   32
         Top             =   4800
         Width           =   825
      End
      Begin VB.Label Label13 
         Caption         =   "Doc.Compra"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Otras Señas"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Responsables"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   23
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Serie"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Marca"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   21
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Modelo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmActivosPrincipal.frx":0F1C
         Top             =   480
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   330
         X2              =   6480
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   300
         X2              =   6480
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label20 
         Caption         =   "Número Placa"
         Height          =   255
         Left            =   1020
         TabIndex        =   14
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Depreciación"
         Height          =   255
         Index           =   0
         Left            =   3450
         TabIndex        =   13
         Top             =   2820
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Valor de rescate"
         Height          =   255
         Left            =   3450
         TabIndex        =   12
         Top             =   2040
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Instalación"
         Height          =   255
         Left            =   3450
         TabIndex        =   11
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.PictureBox LDV 
      Height          =   435
      Left            =   7500
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LisBodegas"
                  Text            =   "Listado de Bodegas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "InvBodegas"
                  Text            =   "Inventario x Bodega"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAF_Activos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vPaso As Boolean

Private Sub cbo_Click()
If Mid(cbo.Text, 1, 1) = "U" Then
  txtUDProducidas.Locked = False
Else
  txtUDProducidas.Locked = True
End If

txtUDProducidas.ForeColor = IIf(txtUDProducidas.Locked, vbBlack, vbBlue)

txtUDAnio.Locked = txtUDProducidas.Locked
txtUDAnio.ForeColor = txtUDProducidas.ForeColor

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub cboTipo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If Not vPaso Then Exit Sub

'Llenar con valores por defecto, del tipo de activo
strSQL = "select * from af_tipo_activo where tipo_activo = '" & fxCodigoCbo(cboTipo) & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  cbo.Text = fxGMetodosDesc(rs!met_depreciacion)
  txtVU = rs!vida_util
  If rs!tipo_vida_util = "A" Then
    cboVU.Text = "Años"
  Else
    cboVU.Text = "Meses"
  End If
End If
rs.Close

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValorHistorico.SetFocus
End Sub

Private Sub cboVU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub

Private Sub chkResponsables_Click()
Call ssTab_Click(0)
End Sub

Private Sub dtpAdquisicion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInstalacion.SetFocus
End Sub

Private Sub dtpInstalacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVU.SetFocus
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 Me.Icon = MDIMenu.Icon
  
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox Err.Description, vbExclamation

  
End Sub

Private Sub sbActivaDesactiva()
Dim strSQL As String, rs As New ADODB.Recordset

'Si no hay periodos depreciados, permite la modificacion
'Total, de lo contrario, realiza bloqueo en datos de calculos
'y ubicacion

If CLng(StatusBarX.Panels(3).Text) = 0 Then
    cbo.BackColor = vbWhite
    cbo.Locked = False
Else
    cbo.BackColor = &HE0E0E0
    cbo.Locked = True
End If

txtVU.BackColor = cbo.BackColor
txtVU.Locked = cbo.Locked
cboVU.BackColor = cbo.BackColor

txtUDProducidas.BackColor = cbo.BackColor
txtUDProducidas.Locked = cbo.Locked

txtUDAnio.BackColor = cbo.BackColor
txtUDAnio.Locked = cbo.Locked

cboTipo.BackColor = cbo.BackColor
cboTipo.Locked = cbo.Locked

txtValorHistorico.BackColor = cbo.BackColor
txtValorHistorico.Locked = cbo.Locked
txtValorRescate.BackColor = cbo.BackColor
txtValorRescate.Locked = cbo.Locked

dtpAdquisicion.Enabled = IIf(cbo.Locked, False, True)
dtpInstalacion.Enabled = IIf(cbo.Locked, False, True)

txtDepartamento.BackColor = cbo.BackColor
txtSeccion.BackColor = cbo.BackColor

lsw.BackColor = cbo.BackColor
lsw.Enabled = IIf(cbo.Locked, False, True)


strSQL = "select forzar_tipoActivo from af_parametros"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  'Nada
Else
 If rs!forzar_TipoActivo = 1 Then
   cboVU.Locked = True
   txtVU.Locked = True
   cbo.Locked = True
 End If
End If
rs.Close

End Sub

Private Sub sbLimpiaPantalla()

ssTab.Tab = 0
txtDescripcion = ""

vCodigo = ""
txtCodigo = ""

lblEstado.Caption = ""

Call sbGLlenaCboMetodos(cbo)
txtVU = ""
cboVU.Text = "Años"

txtUDProducidas = 0
txtUDProducidas.Locked = True

txtUDAnio = 0
txtUDAnio.Locked = True

vPaso = False
  Call sbGLlenaCboTiposAct(cboTipo)
vPaso = True
Call cboTipo_Click

txtValorHistorico = ""
txtValorRescate = ""

dtpAdquisicion.Value = fxFechaServidor
dtpInstalacion.Value = dtpAdquisicion.Value


txtNotas = ""
txtDepartamento = ""
txtDepartamento.Tag = ""

txtSeccion = ""
txtSeccion.Tag = ""

txtDocCompra = ""
txtProveedor = ""
txtProveedor.Tag = ""

txtModelo = ""
txtSerie = ""
txtMarca = ""
txtOtrasSenas = ""
chkResponsables.Value = vbUnchecked
lsw.ListItems.Clear

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(3).Text = 0
StatusBarX.Panels(4).Text = 0
StatusBarX.Panels(5).Text = 0

Call sbActivaDesactiva

End Sub


Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

If Not vEdita Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If Item.Checked Then
  strSQL = "insert af_activo_res(num_placa,cedula,fecha,estado) values('" & vCodigo _
         & "','" & Item.Text & "',getdate(),'A')"
Else
  strSQL = "update af_activo_res set estado = 'I'  where num_placa = '" & vCodigo _
         & "' and cedula = '" & Item.Text & "' and estado = 'A'"
End If
glogon.Conection.Execute strSQL

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub



Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curDepAcum As Currency, curDepMes As Currency, curLibros As Currency
Dim itmX As ListItem, vPasoX As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case ssTab.Tab
 Case 1 'Detalle
   lsw.ListItems.Clear
   
   If chkResponsables.Value = vbChecked Then
      strSQL = "select R.cedula,R.nombre,A.num_placa" _
             & " from af_responsables R left join af_activo_Res A" _
             & " on R.cedula = A.cedula and A.num_placa = '" & vCodigo _
             & "' and A.estado = 'A' order by A.num_placa desc,R.nombre"
   Else
      strSQL = "select R.cedula,R.nombre,A.num_placa" _
             & " from af_responsables R left join af_activo_Res A" _
             & " on R.cedula = A.cedula and A.num_placa = '" & vCodigo _
             & "' and A.estado = 'A' where R.cod_departamento = '" & txtDepartamento.Tag _
             & "' and R.cod_seccion = '" & txtSeccion.Tag _
             & "' order by A.num_placa desc,R.nombre"
   End If
   rs.Open strSQL, glogon.Conection, adOpenForwardOnly
   Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!cedula)
         itmX.SubItems(1) = rs!Nombre
     If Not IsNull(rs!num_placa) Then itmX.Checked = True
     rs.MoveNext
   Loop
   rs.Close
 
   If lsw.ListItems.Count = 0 And chkResponsables.Value = vbUnchecked Then
       chkResponsables.Value = vbChecked
       Call ssTab_Click(0)
   End If
 
 Case 2 'Modificaciones
     lswMod.ListItems.Clear
     
     vPasoX = False
     
     'RETIROS
     strSQL = "select R.*,J.descripcion as Justificacion" _
            & " from af_retiro_adicion R inner join af_justificaciones J" _
            & " on R.cod_justificacion = J.cod_justificacion" _
            & " where R.tipo = 'R' and num_placa ='" & vCodigo & "' order by id"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     Do While Not rs.EOF
      If Not vPasoX Then
        Set itmX = lswMod.ListItems.Add(, , "RETIRO")
            itmX.ForeColor = vbBlue
        vPasoX = True
      End If
      Set itmX = lswMod.ListItems.Add(, , "")
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
          itmX.SubItems(3) = Format(rs!monto, "Standard")
          itmX.SubItems(4) = rs!Justificacion
      rs.MoveNext
     Loop
     rs.Close
     
     
     'Adiciones y mejoras
     vPasoX = False
     strSQL = "select R.*,J.descripcion as Justificacion" _
            & " from af_retiro_adicion R inner join af_justificaciones J" _
            & " on R.cod_justificacion = J.cod_justificacion" _
            & " where R.tipo = 'A' and num_placa ='" & vCodigo & "' order by id"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     Do While Not rs.EOF
      If Not vPasoX Then
        Set itmX = lswMod.ListItems.Add(, , "")
        Set itmX = lswMod.ListItems.Add(, , "ADICIONES/MEJORAS")
            itmX.ForeColor = vbBlue
        vPasoX = True
      End If
      Set itmX = lswMod.ListItems.Add(, , "")
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
          itmX.SubItems(3) = Format(rs!monto, "Standard")
          itmX.SubItems(4) = rs!Justificacion
      rs.MoveNext
     Loop
     rs.Close
     
     vPasoX = False
     'Revaluaciones
     vPasoX = False
     strSQL = "select R.*,J.descripcion as Justificacion" _
            & " from af_retiro_adicion R inner join af_justificaciones J" _
            & " on R.cod_justificacion = J.cod_justificacion" _
            & " where R.tipo = 'V' and num_placa ='" & vCodigo & "' order by id"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     Do While Not rs.EOF
      If Not vPasoX Then
        Set itmX = lswMod.ListItems.Add(, , "")
        Set itmX = lswMod.ListItems.Add(, , "REVALUACIONES")
            itmX.ForeColor = vbBlue
        vPasoX = True
      End If
      Set itmX = lswMod.ListItems.Add(, , "")
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
          itmX.SubItems(3) = Format(rs!monto, "Standard")
          itmX.SubItems(4) = rs!Justificacion
      rs.MoveNext
     Loop
     rs.Close
 
 
  Case 3 'Historico
     lswHistorico.ListItems.Clear
     strSQL = "select * from af_cierres where num_placa = '" & txtCodigo _
            & "' order by anio desc,mes desc"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!Anio)
           itmX.SubItems(1) = rs!Mes
           itmX.SubItems(2) = Format(rs!depreciacion_ac, "Standard")
           itmX.SubItems(3) = Format(rs!depreciacion_mes, "Standard")
       rs.MoveNext
     Loop
     rs.Close
     
  Case 4 'Composicion
     lswCompo.ListItems.Clear
     curDepAcum = 0
     curDepMes = 0
     strSQL = "select num_placa,'X' as Tipo,nombre as descripcion,depreciacion_periodo,depreciacion_acum" _
            & ",depreciacion_mes,fecha_adquisicion as Fecha, Valor_historico as Libros" _
            & " From af_activos where num_placa = '" & txtCodigo & "'" _
            & " Union " _
            & " select num_placa + '-' + CONVERT(char(3), ID) as num_Placa,'A' as Tipo,descripcion,depreciacion_periodo,depreciacion_acum" _
            & ",depreciacion_mes,fecha as Fecha, Monto as Libros" _
            & " From af_retiro_Adicion where tipo = 'A' and num_placa = '" _
            & txtCodigo & "' order by fecha asc"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     Do While Not rs.EOF
       Select Case rs!Tipo
         Case "X" 'Activo
           Set itmX = lswCompo.ListItems.Add(, , "ACTIVO")
         Case "A" 'Adicion
           Set itmX = lswCompo.ListItems.Add(, , "ADICION/MEJORA")
         Case "R" 'Revaluacion
           Set itmX = lswCompo.ListItems.Add(, , "REVALUACION")
       End Select
       
           itmX.SubItems(1) = Trim(rs!num_placa)
           itmX.SubItems(2) = rs!depreciacion_periodo
           itmX.SubItems(3) = Format(rs!depreciacion_acum, "Standard")
           itmX.SubItems(4) = Format(rs!depreciacion_mes, "Standard")
           itmX.SubItems(5) = Format(rs!Libros, "Standard")
           itmX.SubItems(6) = rs!Descripcion
           itmX.SubItems(7) = rs!fecha
        
        curDepAcum = curDepAcum + rs!depreciacion_acum
        curDepMes = curDepMes + rs!depreciacion_mes
        curLibros = curLibros + rs!Libros
        
       rs.MoveNext
     
     Loop
     rs.Close
     
     Set itmX = lswCompo.ListItems.Add(, , "")
         itmX.SubItems(3) = "_________"
         itmX.SubItems(4) = "_________"
         itmX.SubItems(5) = "_________"

     Set itmX = lswCompo.ListItems.Add(, , "TOTAL ACTIVO")
         itmX.SubItems(3) = Format(curDepAcum, "Standard")
         itmX.SubItems(4) = Format(curDepMes, "Standard")
         itmX.SubItems(5) = Format(curLibros, "Standard")
         itmX.ForeColor = vbBlue

     Set itmX = lswCompo.ListItems.Add(, , "")
     Set itmX = lswCompo.ListItems.Add(, , "")
     Set itmX = lswCompo.ListItems.Add(, , "T.ADQUIRIDO")
         itmX.ForeColor = vbBlue
         itmX.SubItems(1) = Format(curLibros, "Standard")
     Set itmX = lswCompo.ListItems.Add(, , "T.DEPRECIADO")
         itmX.ForeColor = vbBlue
         itmX.SubItems(1) = Format(curDepAcum, "Standard")
     Set itmX = lswCompo.ListItems.Add(, , "")
         itmX.SubItems(1) = "__________"
     Set itmX = lswCompo.ListItems.Add(, , "VALOR LIBROS")
         itmX.ForeColor = vbBlue
         itmX.SubItems(1) = Format(curLibros - curDepAcum, "Standard")




  Case 5 'Polizas
     lswPolizas.ListItems.Clear
     strSQL = "select P.*,T.descripcion as DescTipo" _
            & " from af_polizas_tipos T inner join af_polizas P  on T.tipo_poliza = P.tipo_poliza" _
            & " inner join af_polizas_asigna A on P.cod_poliza = A.cod_poliza " _
            & " and A.num_placa = '" & txtCodigo _
            & "' order by P.fecha_vence desc"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     Do While Not rs.EOF
       Set itmX = lswPolizas.ListItems.Add(, , rs!DescTipo)
           itmX.SubItems(1) = rs!num_poliza
           itmX.SubItems(2) = rs!Documento
           itmX.SubItems(3) = Format(rs!fecha_inicio, "dd/mm/yyyy")
           itmX.SubItems(4) = Format(rs!fecha_vence, "dd/mm/yyyy")
           itmX.SubItems(5) = rs!Descripcion
           itmX.SubItems(6) = rs!cod_poliza
       rs.MoveNext
     Loop
     rs.Close

End Select

vError:
Me.MousePointer = vbDefault

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
       gBusquedas.Consulta = "select num_placa,descripcion from af_activos"
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

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select A.*,D.descripcion as Departamento,S.descripcion as Seccion" _
       & ",P.descripcion as Proveedor,rtrim(T.tipo_activo) + ' - ' + T.descripcion as TA" _
       & " from af_activos A inner join AF_Secciones S on A.cod_departamento = S.cod_departamento" _
       & " and A.cod_seccion = S.cod_seccion inner join af_departamentos D on A.cod_departamento = D.cod_departamento" _
       & " inner join af_proveedores P on A.cod_proveedor = P.cod_proveedor" _
       & " inner join af_tipo_activo T on A.tipo_activo = T.tipo_activo" _
       & " where A.num_placa = '" & xCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vPaso = False
    
  vCodigo = rs!num_placa
  txtCodigo = rs!num_placa
 
  txtDescripcion = rs!Nombre
  cboTipo.Text = rs!ta
  cbo.Text = fxGMetodosDesc(rs!met_depreciacion)
  
  txtVU = rs!vida_util
  txtUDProducidas = Format(rs!ud_produccion, "Standard")
  txtUDAnio = Format(rs!ud_anio, "Standard")
  
  If rs!estado = "A" Then
     lblEstado.Caption = "   ACTIVO VIGENTE"
  Else
     lblEstado.Caption = "   ACTIVO RETIRADO"
  End If
  
  If rs!vida_util_en = "A" Then
    cboVU.Text = "Años"
  Else
    cboVU.Text = "Meses"
  End If
      
  txtValorHistorico = Format(rs!valor_historico, "Standard")
  txtValorRescate = Format(rs!valor_desecho, "Standard")
      
  dtpAdquisicion.Value = rs!fecha_adquisicion
  dtpInstalacion.Value = rs!fecha_instalacion
  
  txtNotas = rs!Descripcion
  txtDepartamento = rs!departamento
  txtDepartamento.Tag = rs!cod_departamento
  txtSeccion = rs!seccion
  txtSeccion.Tag = rs!cod_seccion
  txtDocCompra = rs!compra_documento
  txtProveedor = rs!Proveedor
  txtProveedor.Tag = rs!cod_proveedor
  
  txtSerie = rs!NUM_SERIE
  txtModelo = rs!modelo
  txtMarca = rs!marca
  txtOtrasSenas = rs!otras_senas
  
  ssTab.Tab = 0
  
  StatusBarX.Panels(1).Text = rs!creacion_user & ""
  StatusBarX.Panels(2).Text = rs!creacion_fecha & ""
  StatusBarX.Panels(3).Text = rs!depreciacion_periodo
  StatusBarX.Panels(4).Text = Format(rs!depreciacion_acum, "Standard")
  StatusBarX.Panels(5).Text = Format(rs!depreciacion_mes, "Standard")
  
  Call sbActivaDesactiva
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String, i As Integer, x As Boolean

x = False
vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del Activo no es válido ..."
If dtpAdquisicion.Value > dtpInstalacion.Value Then vMensaje = vMensaje & vbCrLf & " - La fecha de Adquisición no puede ser menor a la de instalacion ..."

If Not IsNumeric(txtVU) Then vMensaje = vMensaje & vbCrLf & " - Vida Util no es válida ..."
If Not IsNumeric(txtValorHistorico) Then vMensaje = vMensaje & vbCrLf & " - Valor Historico no es válido ..."
If Not IsNumeric(txtValorRescate) Then vMensaje = vMensaje & vbCrLf & " - Vida Rescate no es válido ..."
If Not IsNumeric(txtUDAnio) Or Not IsNumeric(txtUDProducidas) Then vMensaje = vMensaje & vbCrLf & " - Las unidades de producción no son validas ..."


If IsNumeric(txtUDAnio) And IsNumeric(txtUDProducidas) Then
   If CCur(txtUDAnio) > CCur(txtUDProducidas) Then
     vMensaje = vMensaje & vbCrLf & " - Las unidades de producción Anual no pueden ser mayores a las totales..."
   End If
End If

If IsNumeric(txtValorHistorico) And IsNumeric(txtValorRescate) Then
 If CCur(txtValorHistorico) < CCur(txtValorRescate) Then vMensaje = vMensaje & vbCrLf & " - Valor Historico no puede ser menor al valor de rescate (desecho) ..."
End If

If txtDepartamento.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Departamento no es válido ..."
If txtSeccion.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Sección no es válida ..."
If txtProveedor.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Proveedor no es válido ..."

'Responsables (debe haber almenos uno)
If Not vEdita Then
    For i = 1 To lsw.ListItems.Count
     If lsw.ListItems.Item(i).Checked = True Then
       x = True
     End If
    Next i
    If Not x Then vMensaje = vMensaje & vbCrLf & " - No se ha asignado ningun responsable para el activo ..."
End If

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vEdita Then


  'Si el Activo no ha depreciado ningun periodo aplicar modificacion total,
  'de lo contrario, solo actualizar los datos descriptivos.
  
  strSQL = "update af_activos set nombre = '" & UCase(txtDescripcion) _
         & "',descripcion = '" & txtNotas & "',compra_documento = '" & txtDocCompra & "',cod_proveedor = '" & txtProveedor.Tag _
         & "',num_serie = '" & txtSerie & "',marca = '" & txtMarca & "',modelo ='" & txtModelo _
         & "',otras_senas = '" & txtOtrasSenas & "'"
  
 If CLng(StatusBarX.Panels(3).Text) = 0 Then
    strSQL = strSQL & ",tipo_activo = '" & fxCodigoCbo(cboTipo) & "',met_depreciacion = '" _
           & fxGMetodosDesc(cbo.Text) & "',Vida_Util_en = '" _
           & Mid(cboVU.Text, 1, 1) & "',Vida_Util = " & txtVU _
           & ",UD_ANIO = " & CCur(txtUDAnio) & ",UD_PRODUCCION = " & CCur(txtUDProducidas) _
           & ",valor_historico = " & CCur(txtValorHistorico) & ",valor_desecho = " & CCur(txtValorRescate) _
           & ",fecha_adquisicion = '" & Format(dtpAdquisicion.Value, "yyyy/mm/dd") _
           & "',fecha_instalacion = '" & Format(dtpInstalacion.Value, "yyyy/mm/dd") _
           & "',cod_departamento = '" & txtDepartamento.Tag & "',cod_seccion = '" & txtSeccion.Tag & "'"
  End If
  
  strSQL = strSQL & " where num_placa = '" & vCodigo & "'"
  glogon.Conection.Execute strSQL
  
'  Call sbBitacora("Modifica", "Tipo Activo : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into af_activos(num_placa,nombre,tipo_activo,descripcion,met_depreciacion" _
          & ",vida_util_en,vida_util,valor_historico,valor_desecho,fecha_adquisicion,fecha_instalacion" _
          & ",cod_departamento,cod_seccion,cod_proveedor,compra_documento,num_serie,marca,modelo" _
          & ",otras_senas,estado,depreciacion_acum,depreciacion_mes,depreciacion_periodo,ud_produccion" _
          & ",ud_anio,creacion_fecha,creacion_user) " _
          & " values('" & vCodigo & "','" & UCase(txtDescripcion) & "','" & fxCodigoCbo(cboTipo) & "','" & txtNotas _
          & "','" & fxGMetodosDesc(cbo.Text) & "','" & Mid(cboVU.Text, 1, 1) & "'," & txtVU & "," & CCur(txtValorHistorico) _
          & "," & CCur(txtValorRescate) & ",'" & Format(dtpAdquisicion.Value, "yyyy/mm/dd") _
          & "','" & Format(dtpInstalacion.Value, "yyyy/mm/dd") & "','" & txtDepartamento.Tag _
          & "','" & txtSeccion.Tag & "','" & txtProveedor.Tag & "','" & txtDocCompra _
          & "','" & txtSerie & "','" & txtMarca & "','" & txtModelo & "','" _
          & txtOtrasSenas & "','A',0,0,0," & CCur(txtUDProducidas) & "," & CCur(txtUDAnio) & ",getdate(),'" & glogon.Usuario & "')"
   glogon.Conection.Execute strSQL
    
   For i = 1 To lsw.ListItems.Count
     If lsw.ListItems.Item(i).Checked = True Then
         strSQL = "insert af_activo_res(num_placa,cedula,fecha,estado) values('" & vCodigo _
                & "','" & lsw.ListItems.Item(i).Text & "',getdate(),'A')"
         glogon.Conection.Execute strSQL
     End If
   Next i
   
   
  StatusBarX.Panels(1).Text = glogon.Usuario
  StatusBarX.Panels(2).Text = fxFechaServidor
  StatusBarX.Panels(3).Text = 0
  StatusBarX.Panels(4).Text = 0
  StatusBarX.Panels(5).Text = 0
   
   ' Call sbBitacora("Registra", "Bodega: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete af_activos where num_placa = '" & vCodigo & "'"
  If CCur(StatusBarX.Panels(3).Text) = 0 Then glogon.Conection.Execute strSQL
  
'  Call sbBitacora("Elimina", "Tipo Activo : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "num_placa"
  gBusquedas.Orden = "num_placa"
  gBusquedas.Consulta = "select num_placa,nombre from af_Activos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtDepartamento_Change()
txtSeccion.Tag = ""
txtSeccion = ""
End Sub

Private Sub txtDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSeccion.SetFocus
  
If KeyCode = vbKeyF4 And CLng(StatusBarX.Panels(3).Text) = 0 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from af_departamentos"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtDepartamento.Tag) Then
       txtDepartamento.Tag = gBusquedas.Resultado
       txtDepartamento = gBusquedas.Resultado2
       txtSeccion.SetFocus
    End If
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select num_placa,nombre from af_activos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub

Private Sub txtDocCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  ssTab.Tab = 1
  txtModelo.SetFocus
End If
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOtrasSenas.SetFocus
End Sub

Private Sub txtModelo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSerie.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDepartamento.SetFocus
End Sub

Private Sub txtOtrasSenas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then lsw.SetFocus
End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocCompra.SetFocus
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_proveedor,descripcion from af_proveedores"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtProveedor.Tag) Then
       txtProveedor.Tag = gBusquedas.Resultado
       txtProveedor = gBusquedas.Resultado2
       txtDocCompra.SetFocus
    End If
End If

End Sub

Private Sub txtSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedor.SetFocus
If KeyCode = vbKeyF4 And CLng(StatusBarX.Panels(3).Text) = 0 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_seccion,descripcion from af_secciones"
    gBusquedas.Filtro = " and cod_departamento = '" & txtDepartamento.Tag & "'"
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtSeccion.Tag) Then
       txtSeccion.Tag = gBusquedas.Resultado
       txtSeccion = gBusquedas.Resultado2
       txtProveedor.SetFocus
    End If
End If
End Sub

Private Sub txtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMarca.SetFocus
End Sub


Private Sub txtUDAnio_GotFocus()
On Error GoTo vError
txtUDAnio = CCur(txtUDAnio)
vError:
End Sub

Private Sub txtUDAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUDProducidas.SetFocus
End Sub

Private Sub txtUDAnio_LostFocus()
On Error GoTo vError
txtUDAnio = Format(CCur(txtUDAnio), "Standard")
vError:
End Sub

Private Sub txtUDProducidas_GotFocus()
On Error GoTo vError
txtUDProducidas = CCur(txtUDProducidas)
vError:
End Sub

Private Sub txtUDProducidas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtUDProducidas_LostFocus()
On Error GoTo vError
txtUDProducidas = Format(CCur(txtUDProducidas), "Standard")
vError:
End Sub

Private Sub txtValorHistorico_GotFocus()
On Error GoTo vError
txtValorHistorico = CCur(txtValorHistorico)
vError:
End Sub

Private Sub txtValorHistorico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValorRescate.SetFocus
End Sub

Private Sub txtValorHistorico_LostFocus()
On Error GoTo vError
txtValorHistorico = Format(CCur(txtValorHistorico), "Standard")
vError:
End Sub

Private Sub txtValorRescate_GotFocus()
On Error GoTo vError
txtValorRescate = CCur(txtValorRescate)
vError:
End Sub

Private Sub txtValorRescate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpAdquisicion.SetFocus
Exit Sub

vError:
  txtNotas.SetFocus
End Sub

Private Sub txtValorRescate_LostFocus()
On Error GoTo vError
txtValorRescate = Format(CCur(txtValorRescate), "Standard")
vError:
End Sub

Private Sub txtVU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 If cboVU.Locked Then
    cbo.SetFocus
 Else
    cboVU.SetFocus
 End If
End If
End Sub


