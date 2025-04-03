VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmVivGastosLegales 
   Caption         =   "Gastos Legales"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5490
      Begin VB.TextBox txtValor 
         Height          =   315
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboAplicar 
         Height          =   315
         ItemData        =   "frmVIVGastosLegales.frx":0000
         Left            =   1200
         List            =   "frmVIVGastosLegales.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtdescripcion 
         Height          =   315
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.OptionButton OptInactivo 
         Caption         =   "Inactivo"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Tag             =   "I"
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optActivo 
         Caption         =   "Activo"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Tag             =   "A"
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvgastosLegales 
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageListExped"
         SmallIcons      =   "ImageListExped"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblaplicar 
         Caption         =   "Aplicar"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   855
      End
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   5670
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   4260
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   165
         TabIndex        =   10
         Top             =   30
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.ToolTipText     =   "Imprime el listado seleccionado"
               Object.Tag             =   "1"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   6
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RepActas"
                     Text            =   "Actas"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RepPreAnalisis"
                     Text            =   "Pre Analisis"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RepGarantia"
                     Text            =   "Garantía"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoleta"
                     Text            =   "Boleta"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Cheques"
                     Text            =   "Boleta de Cheques"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra esta ventana"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmVivGastosLegales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboAplicar_Click()
lblaplicar.Caption = cboAplicar.Text
End Sub

Private Sub Form_Load()
gToolBar = "00"
Call sbToolBarIconos(tlbPrincipal, False)
End Sub

