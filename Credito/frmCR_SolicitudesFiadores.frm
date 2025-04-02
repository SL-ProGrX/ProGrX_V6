VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.ShortcutBar.v19.1.0.ocx"
Begin VB.Form frmCR_SolicitudesFiadores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "xx"
   ClientHeight    =   7260
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   9948
   DrawStyle       =   1  'Dash
   HelpContextID   =   3018
   Icon            =   "frmCR_SolicitudesFiadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9948
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2412
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   9732
      _Version        =   1245185
      _ExtentX        =   17166
      _ExtentY        =   4254
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
      FlatScrollBar   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox fra 
      Height          =   3612
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9612
      _Version        =   1245185
      _ExtentX        =   16954
      _ExtentY        =   6371
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnDatos 
         Height          =   612
         Left            =   7680
         TabIndex        =   23
         Top             =   2880
         Width           =   1812
         _Version        =   1245185
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Datos Personales"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCR_SolicitudesFiadores.frx":030A
      End
      Begin XtremeSuiteControls.CheckBox chkInterno 
         Height          =   492
         Left            =   7200
         TabIndex        =   11
         Top             =   1560
         Width           =   2652
         _Version        =   1245185
         _ExtentX        =   4678
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Fiador/Co-Deudor Interno "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboCalidad 
         Height          =   312
         Left            =   4800
         TabIndex        =   3
         Top             =   2040
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
         _ExtentY        =   550
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
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   312
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2292
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   312
         Left            =   2520
         TabIndex        =   5
         Top             =   1320
         Width           =   2292
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   4800
         TabIndex        =   9
         Top             =   1320
         Width           =   2292
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   312
         Left            =   2520
         TabIndex        =   12
         Top             =   1680
         Width           =   4572
         _Version        =   1245185
         _ExtentX        =   8065
         _ExtentY        =   550
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
      End
      Begin XtremeSuiteControls.FlatEdit txtSalario 
         Height          =   312
         Left            =   4800
         TabIndex        =   15
         Top             =   2400
         Width           =   2292
         _Version        =   1245185
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
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDevengado 
         Height          =   312
         Left            =   4800
         TabIndex        =   16
         Top             =   2760
         Width           =   2292
         _Version        =   1245185
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
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLiquidez 
         Height          =   312
         Left            =   4800
         TabIndex        =   17
         Top             =   3120
         Width           =   2292
         _Version        =   1245185
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
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   4800
         TabIndex        =   25
         Top             =   600
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
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
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   6
         Left            =   2160
         TabIndex        =   26
         Top             =   600
         Width           =   2172
      End
      Begin XtremeShortcutBar.ShortcutCaption TituloOpcion 
         Height          =   360
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   9612
         _Version        =   1245185
         _ExtentX        =   16954
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Registro:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "[%] Liquidez Actual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   5
         Left            =   2520
         TabIndex        =   20
         Top             =   3120
         Width           =   2172
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "[% ]Salario Devengado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   4
         Left            =   2520
         TabIndex        =   19
         Top             =   2760
         Width           =   2172
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   3
         Left            =   2520
         TabIndex        =   18
         Top             =   2400
         Width           =   2172
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   6
         Left            =   1560
         TabIndex        =   14
         Top             =   2520
         Width           =   1212
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patrono"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   13
         Top             =   1680
         Width           =   1812
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Calidad"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   2040
         Width           =   2172
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   1
         Left            =   4800
         TabIndex        =   8
         Top             =   1080
         Width           =   2772
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   2520
         TabIndex        =   7
         Top             =   1080
         Width           =   2292
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2292
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   4800
      Top             =   0
   End
   Begin ComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   7008
      Width           =   9948
      _ExtentX        =   17547
      _ExtentY        =   445
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Registro:"
            TextSave        =   "Registro:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Modificado:"
            TextSave        =   "Modificado:"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   264
      Left            =   5400
      TabIndex        =   24
      Top             =   120
      Width           =   4296
      _ExtentX        =   7578
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "cerrar"
            Object.ToolTipText     =   "cierra esta ventana"
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption TituloOpcionesSub 
      Height          =   360
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   9732
      _Version        =   1245185
      _ExtentX        =   17166
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Fiadores / Co-Deudores Registrados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   6
   End
End
Attribute VB_Name = "frmCR_SolicitudesFiadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Integer
Dim lngFiador As Long
Dim vEspacio As Integer
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim vInstitucion As Integer

Private Sub btnDatos_Click()
If txtCedula <> "" Then
    GLOBALES.gCedulaActual = txtCedula
    Call sbFormsCall("frmCR_VerificaDatosPersonales", 1, , , False, Me)
End If
End Sub

Private Sub cboCalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtApellido1.SetFocus

End Sub

Private Sub sbInicializa()
Dim strSQL As String
 
strSQL = "select cod_institucion as 'Idx', descripcion as 'ItmX'" _
       & " from instituciones where Activa = 1"
Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

Call sbCargaFiadores

End Sub

Private Sub Form_Load()

 vModulo = 3
 Me.Caption = "Fiadores/Co-Deudores Operación :" & Operacion.Operacion
 
 
 Call sbToolBarIconos(tlbPrincipal, False)
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3600
    .Add , , "Calidad", 1200, vbCenter
    .Add , , "Salario", 1200, vbRightJustify
    .Add , , "Devengado", 1400, vbRightJustify
    .Add , , "[%] Liquidez", 1400, vbRightJustify
    .Add , , "Interno?", 1200, vbCenter
    .Add , , "Id", 800
 End With
 
 cboCalidad.AddItem "Fiador"
 cboCalidad.AddItem "Co-Deudor"
 cboCalidad.Text = "Fiador"
  
  
 With tlbPrincipal
    .Buttons(1).Enabled = True
    .Buttons(2).Enabled = False
    .Buttons(3).Enabled = False
    .Buttons(4).Enabled = False
    .Buttons(5).Enabled = False
 End With
 fra.Enabled = False
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

Private Sub sbLimpiaDatos()

txtCedula.Text = ""
txtNombre.Text = ""
txtApellido1.Text = ""
txtApellido2.Text = ""
txtSalario.Text = 0
txtDevengado.Text = 0
txtLiquidez.Text = 0
fra.Enabled = False

lngFiador = 0
chkInterno.Value = vbChecked

End Sub


Private Sub sbCargaFiadores()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select * from fiadores where id_solicitud =" & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)

With lsw
 .ListItems.Clear
 Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!cedulaf)
       itmX.SubItems(1) = rs!Nombre & ""
       If rs!Calidad = "F" Then
           itmX.SubItems(2) = "Fiador"
       Else
           itmX.SubItems(2) = "Co-Deudor"
       End If
       itmX.SubItems(3) = IIf(IsNull(rs!salario), 0, Format(rs!salario, "Standard"))
       itmX.SubItems(4) = IIf(IsNull(rs!Devengado), 0, Format(rs!Devengado, "Standard"))
       itmX.SubItems(5) = IIf(IsNull(rs!Liquidez), 0, Format(rs!Liquidez, "Standard"))
       itmX.SubItems(6) = rs!interno & ""
       itmX.SubItems(7) = rs!fia_consec
  rs.MoveNext
 Loop
End With
rs.Close

'Call RefrescaTags(Me)
Me.MousePointer = vbDefault

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Me.MousePointer = vbDefault

End Sub


Private Function fxVerificaExistencia() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from socios where cedula ='" & Trim(txtCedula) & "'"
Call OpenRecordSet(rs, strSQL)
    fxVerificaExistencia = IIf((rs!Existe = 1), True, False)
rs.Close

End Function

Private Function fxVerificaDatosFiadores() As Boolean
Dim vMensaje As String

On Error Resume Next

vMensaje = ""

If Len(txtCedula) = 0 Then vMensaje = vMensaje & " - Cédula Incorrecta" & vbCrLf
If Len(txtNombre) = 0 Then vMensaje = vMensaje & " - Nombre Incorrecto" & vbCrLf
If Trim(txtCedula) = Trim(Operacion.Cedula) Then vMensaje = vMensaje & " - El Deudor No puede Ser Fiador de su misma Operación" & vbCrLf

If Not IsNumeric(txtSalario) Then vMensaje = vMensaje & " - Salario Incorrecto" & vbCrLf
If Not IsNumeric(txtLiquidez) Then vMensaje = vMensaje & " - Liquidez Incorrecta" & vbCrLf
If Not IsNumeric(txtDevengado) Then vMensaje = vMensaje & " - Devengado Incorrecto" & vbCrLf

If CCur(txtLiquidez) > 100 Then vMensaje = vMensaje & " - Liquidez Incorrecta" & vbCrLf
   
   
If Len(vMensaje) > 0 Then
   fxVerificaDatosFiadores = False
   MsgBox vMensaje, vbCritical
Else
   fxVerificaDatosFiadores = True
End If
End Function

Private Sub sbGuardaFiadores()
Dim strSQL As String, strNombre As String

Me.MousePointer = vbHourglass

strNombre = UCase(txtApellido1.Text & " " & txtApellido2.Text & " " & txtNombre.Text)
strSQL = ""

strSQL = "exec spCrdOperacionFiadorRegistro " & Operacion.Operacion & ",'" & Operacion.Codigo _
       & "','" & Mid(cboCalidad, 1, 1) & "'," & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & ",'" & glogon.Usuario & "','" & txtCedula.Text & "','" & strNombre & "', " & chkInterno.Value _
       & ", " & CCur(txtSalario) & "," & CCur(txtDevengado) & "," & CCur(txtLiquidez) _
       & "," & vModulo & ",'" & glogon.Maquina & "','" & glogon.AppVersion & "'"

Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault

GLOBALES.gCedulaActual = txtCedula
Call sbFormsCall("frmCR_VerificaDatosPersonales", 1, , , False, Me)

MsgBox "Información Actualizada Satisfactoriamente...", vbInformation

End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

lngFiador = Item.SubItems(7)

txtCedula.Text = Item.Text

Call sbDescomponeNombre(Trim(Item.SubItems(1)))

cboCalidad.Text = Item.SubItems(2)

txtSalario.Text = Format(Item.SubItems(3), "Standard")
txtDevengado.Text = Format(Item.SubItems(4), "Standard")
txtLiquidez.Text = Format(Item.SubItems(5), "Standard")

chkInterno.Value = IIf((Item.SubItems(6) = 0), vbUnchecked, vbChecked)

Call sbDatosFiadores(lngFiador)
 

With tlbPrincipal
   .Buttons(1).Enabled = True
   .Buttons(2).Enabled = True
   .Buttons(3).Enabled = True
   .Buttons(4).Enabled = False
   .Buttons(5).Enabled = False
End With

btnDatos.Enabled = True

vError:

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer, strSQL As String

Select Case Button.Key
  Case "insertar", "nuevo"
   
   vEdita = 0
   Call sbLimpiaDatos
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
    fra.Enabled = True
    txtCedula.SetFocus
    
  Case "editar", "modificar"
   vEdita = 1
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
    fra.Enabled = True
    txtCedula.SetFocus
    btnDatos.Enabled = True
  
  Case "borrar" 'FIA_CONSEC
  
   strSQL = "delete fiadores where FIA_CONSEC=" & lngFiador
   If lngFiador > 0 Then
    iRespuesta = MsgBox("Esta seguro que desea eliminar este fiador", vbYesNo)
    If iRespuesta = vbYes Then
      Call ConectionExecute(strSQL)
      Call sbCargaFiadores
      Call sbLimpiaDatos
    Else
      Call sbLimpiaDatos
    End If
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
     End With
    
   End If
  
  Case "salvar", "guardar"
    
    If fxVerificaDatosFiadores Then
      Call sbGuardaFiadores
      Call sbCargaFiadores
      With tlbPrincipal
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
      End With
      Call sbLimpiaDatos
    Else
      MsgBox "Información Ingresada es Incorrecta por favor verifique...", vbInformation
    End If
  
  Case "deshacer"
  
    Call sbLimpiaDatos
    Call sbCargaFiadores
    
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
    End With
  
  Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
  Case "salir", "cerrar"
    Unload Me
End Select

End Sub


Private Sub txtApellido1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtApellido2.SetFocus
End If

End Sub

Private Sub txtApellido2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtNombre.SetFocus
End If
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 
 Call sbDescomponeNombre(fxNombre(txtCedula))
 vInstitucion = fxInstitucion
 
 If vInstitucion > 0 Then
  Call sbCboAsignaDato(cboInstitucion, fxXInstitucion(vInstitucion), True, CStr(vInstitucion))
  cboInstitucion.Enabled = False
 Else
   cboInstitucion.Enabled = True
 End If
 cboCalidad.SetFocus
End If
End Sub

Private Sub txtDevengado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLiquidez.SetFocus
End Sub

Private Sub txtLiquidez_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCedula.SetFocus
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cboCalidad.SetFocus
End Sub

Private Sub txtSalario_GotFocus()
On Error Resume Next
txtSalario.Text = CCur(txtSalario.Text)
End Sub

Private Sub txtSalario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDevengado.SetFocus
End Sub

Private Sub txtSalario_LostFocus()
On Error Resume Next
txtSalario.Text = Format(txtSalario.Text, "Standard")
End Sub

Private Sub txtDevengado_GotFocus()
On Error Resume Next
txtDevengado.Text = CCur(txtDevengado.Text)
End Sub

Private Sub txtDevengado_LostFocus()
On Error Resume Next
txtDevengado.Text = Format(txtDevengado.Text, "Standard")
End Sub

Private Sub txtLiquidez_GotFocus()
On Error Resume Next
txtLiquidez.Text = CCur(txtLiquidez.Text)
End Sub

Private Sub txtLiquidez_LostFocus()
On Error Resume Next
txtLiquidez.Text = Format(txtLiquidez.Text, "Standard")
End Sub

Private Sub sbDescomponeNombre(vNombre As String)
Dim i As Integer
    
vEspacio = 1
vApellido1 = ""
vApellido2 = ""
vNombre1 = ""
vNombre2 = ""
    
 For i = 1 To Len(vNombre)
     If Mid(vNombre, i, 1) <> " " Then
         Select Case vEspacio
           Case 1
             vApellido1 = vApellido1 & Mid(vNombre, i, 1)
           Case 2
             vApellido2 = vApellido2 & Mid(vNombre, i, 1)
           Case 3
             vNombre1 = vNombre1 & Mid(vNombre, i, 1)
           Case Is >= 4
             vNombre2 = vNombre2 & Mid(vNombre, i, 1)
         End Select
     Else
         vEspacio = vEspacio + 1
     End If
 Next i
 txtApellido1.Text = vApellido1
 txtApellido2.Text = vApellido2
 txtNombre.Text = vNombre1 & " " & vNombre2
 
End Sub

Private Sub sbDatosFiadores(vConsec As Long)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select F.REGISTRO_FECHA , F.ACTUALIZA_FECHA,F.CALIDAD,I.DESCRIPCION, I.cod_Institucion" _
       & " from fiadores F   inner  join INSTITUCIONES I on F.COD_INSTITUCION =  I.COD_INSTITUCION" _
       & " Where F.fia_consec =  " & vConsec & " "

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  StatusBarX.Panels(1).Text = "Registrado:... " & IIf(IsNull(rs!registro_fecha), "", Format(rs!registro_fecha, "dd/mm/yyyy"))
  StatusBarX.Panels(2).Text = "Modificado:... " & IIf(IsNull(rs!Actualiza_fecha), "", Format(rs!Actualiza_fecha, "dd/mm/yyyy"))
  Select Case rs!Calidad
    Case "F"
        cboCalidad.Text = "Fiador"
    Case "C"
        cboCalidad.Text = "Co-Deudor"
  End Select
  
  Call sbCboAsignaDato(cboInstitucion, rs!Descripcion, True, rs!cod_institucion)
  
  
Else
  StatusBarX.Panels(1).Text = "Registrado:... "
  StatusBarX.Panels(2).Text = "Modificado:... "
End If
rs.Close

End Sub

Private Function fxInstitucion() As Integer
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select COD_INSTITUCION from SOCIOS where cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
  fxInstitucion = rs!cod_institucion
Else
  fxInstitucion = 0
End If
rs.Close

End Function


