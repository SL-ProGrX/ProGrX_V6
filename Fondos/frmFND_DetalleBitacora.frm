VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form frmFND_DetalleBitacora 
   Caption         =   "Bitacora especial de Fondos"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmFND_DetalleBitacora.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   12360
   Begin VB.ComboBox cboOperadora 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1320
      Width           =   7215
   End
   Begin VB.CheckBox chkTodosPlan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "> Todos <"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkTodosCont 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "> Todos <"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtPlan 
      Height          =   315
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "(Presione F4)"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtContrato 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6660
      TabIndex        =   14
      Top             =   600
      Width           =   2415
   End
   Begin VB.CheckBox chkTodosUsu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "> Todos <"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6660
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmFND_DetalleBitacora.frx":6852
      Top             =   240
      Width           =   2415
   End
   Begin VB.CheckBox chkTodosMov 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "> Todos <"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox chkTodasFec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "> Todas <"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cboMovimiento 
      Height          =   315
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   2
      Top             =   1680
      Width           =   1445
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   1
      Top             =   1680
      Width           =   1445
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1980
      TabIndex        =   8
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   65077251
      CurrentDate     =   37159
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   3300
      TabIndex        =   9
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   65077251
      CurrentDate     =   37159
   End
   Begin FPSpread.vaSpread vGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   12135
      _Version        =   393216
      _ExtentX        =   21405
      _ExtentY        =   6800
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      SpreadDesigner  =   "frmFND_DetalleBitacora.frx":6862
      VisibleCols     =   6
      VisibleRows     =   500
      VScrollSpecial  =   -1  'True
      ScrollBarTrack  =   3
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plan"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   1005
      TabIndex        =   22
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plan"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   1005
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contrato"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   5700
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   5700
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Movimiento"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   1005
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fechas"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmFND_DetalleBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vTamanoForm As Double
Dim aColumnas(6) As Double
Dim vHanchoGrid As Double, vAltoGrid As Double
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub chkTodasFec_Click()

If chkTodasFec.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub

Private Sub chkTodosCont_Click()
If chkTodosCont.Value = vbChecked Then
   txtContrato.Enabled = False
 Else
   txtContrato.Enabled = True
   txtContrato = Empty
 End If
 
End Sub

Private Sub chkTodosMov_Click()
If chkTodosMov.Value = vbChecked Then
  cboMovimiento.Enabled = False
Else
  cboMovimiento.Enabled = True
End If

End Sub


Private Sub chkTodosPlan_Click()
If chkTodosPlan.Value = vbChecked Then
   txtPlan.Enabled = False
   txtPlan = "(Presione F4)"
   txtDescripcion = Empty
 Else
   txtPlan.Enabled = True
   txtPlan = "(Presione F4)"
   txtDescripcion = Empty
 End If
End Sub

Private Sub chkTodosUsu_Click()
 If chkTodosUsu.Value = vbChecked Then
   txtUsuario.Enabled = False
 Else
   txtUsuario.Enabled = True
   txtUsuario = "(Presione F4)"
 End If
 
End Sub

Private Sub cmdBuscar_Click()
Dim rs As New ADODB.Recordset

vGrid.MaxRows = 0

rs.Open fxSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 1
  vGrid.Text = rs!Usuario
  vGrid.Col = 2
  vGrid.Text = Format(rs!Fecha, "dd/mm/yyyy")
  vGrid.Col = 3
  vGrid.Text = fxMovimientos(rs!Movimiento)
  vGrid.Col = 4
  vGrid.Text = rs!Cod_operadora
  vGrid.Col = 5
  vGrid.Text = rs!cod_plan
  vGrid.Col = 6
  vGrid.Text = rs!cod_Contrato
  vGrid.Col = 7
  vGrid.Text = rs!Detalle
  
  rs.MoveNext
  
Loop
rs.Close
End Sub

Private Sub cmdExcel_Click()
Dim x As String, vRespuesta As Integer
Dim vAbrir As Long, strRuta As String

strRuta = "C:\SIF\" & Me.Caption & ".xls"
x = vGrid.ExportToExcel("C:\SIF\" & Me.Caption & ".xls", "Sheet 1", "C:\LOGFILE.TXT")

'pregunta si se desea abrir el archivo
vRespuesta = MsgBox("El archivo generado es: " & strRuta & vbCrLf & "Desa abrir este archivo...", vbYesNo)
If vRespuesta = vbYes Then vAbrir = ShellExecute(Me.hwnd, "Open", strRuta, "", "", 1)

End Sub

Private Sub cmdReporte_Click()
Me.MousePointer = vbHourglass

vGrid.PrintFooter = "fecha desde  " & Format(dtpInicio.Value, "dd/mm/yyyy") _
                  & " hasta " & Format(dtpCorte.Value, "dd/mm/yyyy") & "Usuario : " & glogon.Usuario

vGrid.PrintHeader = Me.Caption

If vGrid.MaxCols > 5 Then
    vGrid.PrintOrientation = PrintOrientationLandscape
Else
    vGrid.PrintOrientation = PrintOrientationPortrait
End If
vGrid.PrintSheet
  
Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_Load()

Me.Icon = Me.Picture
dtpInicio = fxFechaServidor
dtpCorte = dtpInicio

Call sbgFNDCargaCombos(cboOperadora, "Operadoras")

With cboMovimiento
   .AddItem "01-Cambio Cuota"
   .AddItem "02-Cambio De Plazo"
   .AddItem "03-Cambio de Inversión"
   .AddItem "04-Cambio de Subcuenntas"
   .Text = "01-Cambio Cuota"
   
End With

chkTodosMov.Value = vbChecked
cboMovimiento.Enabled = False

chkTodosUsu.Value = vbChecked
txtUsuario.Enabled = False

chkTodosCont.Value = vbChecked
txtContrato.Enabled = False

chkTodosPlan.Value = vbChecked
txtPlan.Enabled = False

End Sub




Private Function fxMovimientos(strMovimiento) As String

Select Case strMovimiento
  Case "01"
    fxMovimientos = "Cambio de cuota"
  Case "02"
    fxMovimientos = "Cambio de Plazo"
  Case "03"
    fxMovimientos = "Cambio de Inversión"
  Case "04"
    fxMovimientos = "Cambio de Subcuenntase"
  End Select

End Function

Private Function fxSQL() As String
Dim vPaso As Boolean, strSQL As String

vPaso = False

strSQL = "select * from fnd_contratos_cambios"
       
If chkTodasFec.Value = vbUnchecked Then
  If vPaso Then
    strSQL = strSQL & " and fecha between '" & Format(dtpInicio.Value, "yyyymmdd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyymmdd") & " 23:59:00'"
  Else
    vPaso = True
    strSQL = strSQL & " where fecha between '" & Format(dtpInicio.Value, "yyyymmdd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyymmdd") & " 23:59:00'"
  End If
End If 'Clausulas de Fechas

If chkTodosMov.Value = vbUnchecked Then
  If vPaso Then
    strSQL = strSQL & " and movimiento = '" & Mid(cboMovimiento.Text, 1, 2) & "'"
  Else
    vPaso = True
    strSQL = strSQL & " where movimiento = '" & Mid(cboMovimiento.Text, 1, 2) & "'"
  End If
End If

'
If chkTodosUsu.Value = vbUnchecked Then
  If txtUsuario <> "" And txtUsuario <> "(Presione F4)" Then
   If vPaso Then
     strSQL = strSQL & " and Usuario = '" & txtUsuario & "'"
   Else
     vPaso = True
     strSQL = strSQL & " where usuario = '" & txtUsuario & "'"
   End If
  End If
End If

'En caso de que sea un plan
If chkTodosPlan.Value = vbUnchecked Then
  If txtPlan <> "" And txtPlan <> "(Presione F4)" Then
   If vPaso Then
     strSQL = strSQL & " and cod_plan = '" & txtPlan & "'"
   Else
     vPaso = True
     strSQL = strSQL & " where cod_plan = '" & txtPlan & "'"
   End If
  End If
End If

'En caso de que sea contrato
If chkTodosCont.Value = vbUnchecked Then
  If txtContrato <> "" Then
   If vPaso Then
     strSQL = strSQL & " and cod_contrato = '" & txtContrato & "'"
   Else
     vPaso = True
     strSQL = strSQL & " where cod_contrato = '" & txtContrato & "'"
   End If
  End If
End If

 If vPaso Then
   strSQL = strSQL & "And Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
 Else
   strSQL = strSQL & " where Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
 End If




fxSQL = strSQL

End Function



Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 360
vGrid.Height = Me.Height - 1550

End Sub

Private Sub txtPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = " And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtPlan = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Nombre,Descripcion From Usuarios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If
End Sub




