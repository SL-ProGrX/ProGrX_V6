VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSIF_Bitacora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitácora "
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   HelpContextID   =   9003
   Icon            =   "CC_Bitacora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "CC_Bitacora.frx":000C
   ScaleHeight     =   3810
   ScaleWidth      =   6030
   Begin VB.CheckBox chkHoras 
      Appearance      =   0  'Flat
      Caption         =   "&Todos los Horarios"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4080
      TabIndex        =   16
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Presione F4 para Consultar"
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.ComboBox cboModulo 
      Height          =   315
      ItemData        =   "CC_Bitacora.frx":685E
      Left            =   1320
      List            =   "CC_Bitacora.frx":688F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   4575
   End
   Begin VB.ComboBox cboMovimientos 
      Height          =   315
      ItemData        =   "CC_Bitacora.frx":695C
      Left            =   1320
      List            =   "CC_Bitacora.frx":695E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   4575
   End
   Begin VB.CheckBox chkFechas 
      Appearance      =   0  'Flat
      Caption         =   "&Todas las Fechas"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox chkUsuarios 
      Caption         =   "&Todos"
      Height          =   315
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   570
      Left            =   4560
      TabIndex        =   1
      Top             =   3120
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1005
      ButtonWidth     =   2223
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "reporte"
            Object.ToolTipText     =   "Informe de Bitacoras"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CC_Bitacora.frx":6960
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CC_Bitacora.frx":6C7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   89522179
      CurrentDate     =   36185
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   89522179
      CurrentDate     =   36185
   End
   Begin MSComCtl2.DTPicker dtpInicioTime 
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   89522178
      CurrentDate     =   36185
   End
   Begin MSComCtl2.DTPicker dtpCorteTime 
      Height          =   315
      Left            =   2640
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   89522178
      CurrentDate     =   36185
   End
   Begin VB.Label Label1 
      Caption         =   "Bitácora del Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   13
      Top             =   240
      Width           =   4215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   5880
      X2              =   0
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Módulo"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Movimiento"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Corte"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   5880
      X2              =   0
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmSIF_Bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboModulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboMovimientos.SetFocus
End Sub

Private Sub cboMovimientos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUsuario.SetFocus
End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
  dtpCorte.Enabled = False
  dtpInicio.Enabled = False
Else
  dtpCorte.Enabled = True
  dtpInicio.Enabled = True
End If
End Sub

Private Sub chkHoras_Click()
If chkHoras.Value = vbChecked Then
  dtpCorteTime.Enabled = False
  dtpInicioTime.Enabled = False
Else
  dtpCorteTime.Enabled = True
  dtpInicioTime.Enabled = True
End If
End Sub

Private Sub chkUsuarios_Click()
If chkUsuarios.Value = vbChecked Then
   txtUsuario.BackColor = vbWhite
Else
   txtUsuario.BackColor = vbGrayText
End If
End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboModulo.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 10

Me.Icon = Me.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

dtpInicioTime.Value = fxFechaServidor
dtpCorteTime.Value = dtpInicioTime.Value

With cboModulo
  .Clear
  
  strSQL = "select modulo,nombre from modulos where modulo in(1,2,3,4,5,10,18) order by modulo"
  rs.Open strSQL, glogon.Conection, adOpenForwardOnly
  Do While Not rs.EOF
   .AddItem Trim(rs!Nombre)
   .ItemData(.NewIndex) = rs!modulo
   rs.MoveNext
  Loop
  rs.Close
  
  .AddItem "[TODOS]"
  .Text = "[TODOS]"
End With

With cboMovimientos
  .Clear
  .AddItem "Registra"
  .AddItem "Modifica"
  .AddItem "Borra"
  .AddItem "Reversa"
  .AddItem "Aplica"
  .AddItem "Genera"
  .AddItem "Carga"
  .AddItem "Anula"
  .AddItem "Imprime"
  .AddItem "Consulta"
  .AddItem "[TODOS]"
  
  .Text = "[TODOS]"
End With

Call Formularios(Me)
Call RefrescaTags(Me)

chkFechas_Click
chkHoras_Click
chkUsuarios_Click

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, vSubTitulo As String

If dtpInicio.Value > dtpCorte.Value Then
   MsgBox "Verifique El Rango De Fechas", vbExclamation
   Exit Sub
End If

If dtpInicioTime.Value > dtpCorteTime.Value Then
   MsgBox "Verifique El Rango De Horas", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = ""
vSubTitulo = ""

If chkUsuarios.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = "{BITACORA.NOMBRE} = '" & txtUsuario & "'"
  vSubTitulo = "Us:" & txtUsuario
Else
  vSubTitulo = "Us:Todos"
End If

If chkFechas.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{BITACORA.FECHAHORA} in Date (" & Format(dtpInicio.Value, "yyyy,mm,dd") _
         & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
  
  vSubTitulo = vSubTitulo & " Fecha:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " a " & Format(dtpCorte.Value, "dd/mm/yyyy")
Else
  vSubTitulo = vSubTitulo & " Fecha:Todas"
End If

If chkHoras.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "ctime({BITACORA.FECHAHORA}) in Time (" & Format(dtpInicioTime.Value, "h,m,s") _
         & ") to Time (" & Format(dtpCorteTime.Value, "h,m,s") & ")"
  vSubTitulo = vSubTitulo & " Hora:" & Format(dtpInicioTime.Value, "hh:mm:ss") & " a " & Format(dtpCorteTime.Value, "hh:mm:ss")
Else
  vSubTitulo = vSubTitulo & " Hora:Todas"
End If


If cboModulo.Text <> "[TODOS]" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{BITACORA.MODULO} = " & cboModulo.ItemData(cboModulo.ListIndex)
  vSubTitulo = vSubTitulo & " Módulo:" & cboModulo.Text
Else
  vSubTitulo = vSubTitulo & " Módulo:Todos"
End If

If cboMovimientos.Text <> "[TODOS]" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{BITACORA.MOVIMIENTO} ='" & Trim(cboMovimientos.Text) & "'"
  vSubTitulo = vSubTitulo & " Mov.:" & cboMovimientos.Text
Else
  vSubTitulo = vSubTitulo & " Mov.:Todos"
End If


With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Seguridad"
     
     .ReportFileName = App.Path & "\Reportes\SIFBitacora.rpt"
     .Connect = glogon.ConectRPT
     
     .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(1) = "fxRango='" & vSubTitulo & "'"
     .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(3) = "fxServidor='SERVER:" & UCase(glogon.Servidor) & "'"
     .Formulas(4) = "fxBaseDatos='DATABASE:" & UCase(glogon.BaseDatos) & "'"
     .Formulas(5) = "fxUsuario='USER:" & UCase(glogon.Usuario) & "'"
     
     .SelectionFormula = strSQL

     .PrintReport


End With

Me.MousePointer = vbDefault

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Nombre,Descripcion From Usuarios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If

vError:

End Sub


