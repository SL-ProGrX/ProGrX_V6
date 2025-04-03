VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmFNDGestiones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestiones..: Retiros, Liquidaciones y Renovaciones"
   ClientHeight    =   5928
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   9804
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5928
   ScaleWidth      =   9804
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2532
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   9612
      _Version        =   1245185
      _ExtentX        =   16954
      _ExtentY        =   4466
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
      ShowBorder      =   0   'False
   End
   Begin VB.TextBox txtConOperadora 
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
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtConPlan 
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
      Left            =   1440
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtConContrato 
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
      Left            =   2640
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtConCedula 
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
      Left            =   3840
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtConNombre 
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
      Left            =   5280
      TabIndex        =   11
      Top             =   240
      Width           =   4455
   End
   Begin VB.TextBox txtLineas 
      Alignment       =   2  'Center
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
      Left            =   8880
      TabIndex        =   10
      Text            =   "50"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtDetalle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4440
      Width           =   6480
   End
   Begin VB.TextBox txtContrato 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Width           =   4935
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   1440
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin XtremeSuiteControls.PushButton cmdLiquidar 
      Height          =   612
      Left            =   8280
      TabIndex        =   23
      Top             =   5160
      Width           =   1332
      _Version        =   1245185
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Retirar o Liquidar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmFNDGestiones.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdRenovar 
      Height          =   612
      Left            =   8280
      TabIndex        =   22
      Top             =   4440
      Width           =   1332
      _Version        =   1245185
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Renovar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmFNDGestiones.frx":0A2F
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   1440
      TabIndex        =   25
      Top             =   3360
      Width           =   6492
      _Version        =   1245185
      _ExtentX        =   11451
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   20
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   19
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Lineas"
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
      Left            =   8040
      TabIndex        =   16
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   612
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFNDGestiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Function fxVencimiento(vFechaInicio As String, vFechaRenovacion As String, vPlazo As Long) As Boolean
Dim vFecha As String

fxVencimiento = True

vFecha = Format(IIf(vFechaRenovacion <> "", vFechaRenovacion, vFechaInicio), "dd/mm/yyyy")
vFecha = DateAdd("d", vPlazo * 30, vFecha)

If vFecha > fxFechaServidor Then fxVencimiento = False

End Function

Sub sbConsultaContrato(lngCodigo As Long)
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strSQL = "select F.Cedula,S.Nombre,F.Estado,F.Plazo,F.Monto,F.Aportes,F.Rendimiento," _
   & "F.Inc_Tipo,F.Inc_Anual,Operacion from Fnd_Contratos F inner join Socios S " _
   & "on F.Cedula=S.Cedula where cod_contrato = " & lngCodigo _
   & " and Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
   & " and Cod_plan='" & Trim(txtCodigo) & "' and Estado <> 'L' "

With rs
 .Open strSQL, glogon.Conection, adOpenStatic
 
  If rs.EOF = True Then
   MsgBox "No. Contrato Incorrecto"
   txtDetalle = ""
   txtContrato = ""
   txtCedula = ""
   txtNombre = ""
   txtContrato.SetFocus
  Else
   strSQL = "CEDULA:  " & UCase(Trim(!Cedula)) & vbCrLf _
          & "NOMBRE:  " & UCase(Trim(!Nombre)) & vbCrLf _
          & "ESTADO:  " & IIf(!Estado = "A", "ACTIVO", "") & vbCrLf _
          & "PLAZO:   " & !Plazo & vbCrLf _
          & "MONTO:   " & Format(!Monto, "standard") & vbCrLf _
          & "APORTES: " & Format(!aportes, "standard") & vbCrLf _
          & "RENDIM:  " & Format(!rendimiento, "standard")
   txtDetalle = strSQL
   txtCedula = !Cedula
   txtNombre = !Nombre
  End If
 
 .Close
End With

Me.MousePointer = vbDefault

End Sub



Private Sub cboOperadora_Click()
txtCodigo_LostFocus
If Trim(txtContrato) <> "" Then Call sbConsultaContrato(Trim(txtContrato))
End Sub


Private Sub cboOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub



Private Sub cmdLiquidar_Click()

On Error GoTo vError
   
   gFondos.Cedula = txtCedula.Text
   gFondos.Contrato = txtContrato.Text
   gFondos.Operadora = cboOperadora.ItemData(cboOperadora.ListIndex)
   gFondos.Plan = txtCodigo.Text
   GLOBALES.gTag = cboOperadora.Text
   GLOBALES.gTag2 = txtDescripcion.Text


Select Case vGestion
 Case "O"
 
   If Trim(cboOperadora) <> "" Then frmFNDRetirosyLiquidaciones.Show vbModal
 
 Case "P"
   If Trim(txtCodigo) <> "" Then frmFNDRetirosyLiquidaciones.Show vbModal
 
 Case "C"
   If Trim(txtContrato) <> "" Then frmFNDRetirosyLiquidaciones.Show vbModal

End Select

txtCedula.Text = ""
txtNombre.Text = ""
txtContrato.Text = ""
txtCodigo.Text = ""
txtDescripcion.Text = ""
txtDetalle.Text = ""

Call sbLlenaLsw


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub cmdRenovar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency

Select Case vGestion
 Case "O"
   strSQL = "Select Cod_plan,Cod_Contrato,Inc_Tipo,Inc_Anual,Monto,Fecha_Inicio," _
   & "Ult_Renovacion,Plazo from Fnd_Contratos Where Cod_operadora=" _
   & cboOperadora.ItemData(cboOperadora.ListIndex) & " and Estado <> 'L' " _
   & "and Renueva='S'"
 
 Case "P"
   If Trim(txtCodigo) = "" Then Exit Sub
   
   strSQL = "Select Cod_plan,Cod_Contrato,Inc_Tipo,Inc_Anual,Monto,Fecha_Inicio," _
   & "Ult_Renovacion,Plazo from Fnd_Contratos Where Cod_operadora=" _
   & cboOperadora.ItemData(cboOperadora.ListIndex) _
   & " and Cod_plan='" & Trim(txtCodigo) & "' and Estado <> 'L' " _
   & " and Renueva='S'"
 
 Case "C"
   If Trim(txtCodigo) = "" Or Trim(txtContrato) = "" Then Exit Sub
   
    strSQL = "Select Cod_Plan,cod_Contrato,Inc_Tipo,Inc_Anual,Monto,Fecha_Inicio" _
           & ",Ult_Renovacion,Plazo from Fnd_Contratos Where cod_contrato = " _
           & Trim(txtContrato) & " and Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and Cod_plan='" & Trim(txtCodigo) & "'"
End Select



Call OpenRecordSet(rs, strSQL)

With rs
    Do While Not .EOF
      If fxVencimiento(!FECHA_INICIO, IIf(IsNull(!Ult_Renovacion), "", !Ult_Renovacion), !Plazo) = True Then
        If !inc_tipo = "P" Then
           curMonto = !Monto
           curMonto = curMonto + (curMonto * (!Inc_Anual / 100))
        Else
           curMonto = !Monto
           curMonto = curMonto + !Inc_Anual
        End If
       
        strSQL = "Update Fnd_Contratos set Monto=" & curMonto & ",Ult_renovacion=dbo.mygetdate(),ind_deduccion=0 " _
               & " Where Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
               & " and cod_plan='" & Trim(!cod_Plan) _
               & "' and cod_contrato=" & !COD_CONTRATO
        Call ConectionExecute(strSQL)
      
      End If
      .MoveNext
    Loop
    
    If .EOF = True And vGestion = "C" Then
       .Close
       MsgBox "Renovacion No Aplicada", vbExclamation
       txtDescripcion = ""
       txtContrato = ""
       txtDetalle = ""
       txtCodigo = ""
       Exit Sub
    End If
 .Close
End With

MsgBox "Renovacion Aplicada", vbExclamation
txtDescripcion = ""
txtContrato = ""
txtDetalle = ""
txtCodigo = ""


End Sub

Private Sub sbLlenaLsw()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem

On Error GoTo vError

lsw.ListItems.Clear

Me.MousePointer = vbHourglass

strSQL = "Select TOP " & txtLineas.Text & " O.Descripcion as Operadora,F.Cod_Operadora,F.Cod_plan,P.Descripcion" _
       & ",F.Cod_Contrato,F.Cedula,S.Nombre" _
       & " from Fnd_Contratos F Inner Join Fnd_Operadoras O on F.Cod_operadora = O.Cod_operadora" _
       & " inner join Fnd_planes P on F.Cod_operadora = P.Cod_operadora and F.Cod_plan = P.Cod_plan" _
       & " inner join Socios S on F.Cedula = S.Cedula " _
       & " Where F.Estado <> 'L' AND dbo.fxFndColaboradorVisualiza(F.COD_OPERADORA, F.COD_PLAN, F.cedula, S.ESTADOACTUAL , '" & glogon.Usuario & "') = 1"
 
If Trim(txtConOperadora) <> "" Then
   strSQL = strSQL & " And F.Cod_operadora=" & Trim(txtConOperadora)
End If

If Trim(txtConPlan) <> "" Then strSQL = strSQL & " And F.Cod_Plan like '%" & Trim(txtConPlan) & "%'"

If Trim(txtConContrato) <> "" Then strSQL = strSQL & " And F.Cod_Contrato=" & Trim(txtConContrato)

If Trim(txtConCedula) <> "" Then
  strSQL = strSQL & " And F.Cedula like '%" & Trim(txtConCedula) & "%'"
Else
 'Si no aplica la cedula ver por nombre
 If Trim(txtConNombre) <> "" Then
    strSQL = strSQL & " And S.nombre like '%" & Trim(txtConNombre) & "%'"
 End If
End If

Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Operadora)
        itmX.SubItems(1) = Trim(rs!cod_Plan)
        itmX.SubItems(2) = Trim(rs!Descripcion)
        itmX.SubItems(3) = rs!COD_CONTRATO
        itmX.SubItems(4) = rs!Cedula
        itmX.SubItems(5) = rs!Nombre
        itmX.Tag = rs!Cod_Operadora
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion
 
vPaso = False

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


strSQL = "select descripcion as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

'Call Formularios(Me)
'Call RefrescaTags(Me)

With lsw.ColumnHeaders
  .Clear
  .Add , , "Operadora Id", 1200
  .Add , , "Plan", 1000, vbCenter
  .Add , , "Descripción", 2500
  .Add , , "Contrato", 1200
  .Add , , "Identificación", 1400
  .Add , , "Nombre", 3000
End With


vGestion = "C"

Select Case vGestion
  Case "C"
  
  Case "P"
    txtContrato.Enabled = False
    
  Case "O"
    txtCodigo.Enabled = False
    txtDescripcion.Enabled = False
    txtContrato.Enabled = False
  
End Select

End Sub

Private Sub sbConsulta()
       
gBusquedas.Columna = "cod_contrato"
gBusquedas.Orden = "cod_contrato"

gBusquedas.Filtro = "and cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
                    & " and cod_plan='" & Trim(txtCodigo) & "'"
gBusquedas.Consulta = "select cod_Contrato,cedula from fnd_Contratos"
frmBusquedas.Show vbModal
txtContrato.SetFocus

If Trim(gBusquedas.Resultado) <> "" Then
 txtContrato = Trim(gBusquedas.Resultado)
End If
txtDetalle.SetFocus
gBusquedas.Resultado = ""

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)


gFondos.Operadora = Item.Tag
'Inicializa en "" para no consultar

txtContrato.Text = ""
txtCodigo.Text = Item.SubItems(1)
txtDescripcion.Text = Item.SubItems(2)
txtContrato.Text = Item.SubItems(3)

Call sbConsultaContrato(txtContrato.Text)


End Sub


Private Sub txtCodigo_Change()

If Trim(txtContrato) <> "" Then Call sbConsultaContrato(Trim(txtContrato))
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case vbKeyReturn
   txtDescripcion.SetFocus
 Case 48 To 57, 8
 Case Else
   KeyAscii = 0
End Select
End Sub


Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If Trim(txtCodigo) <> "" Then
   strSQL = "Select Descripcion from fnd_planes where cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   strSQL = strSQL & " And cod_plan='" & Trim(txtCodigo) & "'"
   With rs
     .Open strSQL, glogon.Conection, adOpenStatic
        If .EOF = False Then
           txtDescripcion = Trim(!Descripcion)
        Else
           MsgBox "Codigo incorrecto", vbExclamation
           txtCodigo = ""
           txtDescripcion = ""
           txtCodigo.SetFocus
        End If
     .Close
   End With
Else
  txtDescripcion = ""
End If
End Sub


Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbLlenaLsw
End Sub



Private Sub txtConContrato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbLlenaLsw
End Sub


Private Sub txtConNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbLlenaLsw
End Sub

Private Sub txtConOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbLlenaLsw
End Sub

Private Sub txtConPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbLlenaLsw
End Sub

Private Sub txtContrato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   Call sbConsulta
End If

End Sub


Private Sub txtContrato_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case vbKeyReturn
   txtContrato_LostFocus
 Case 48 To 57, 8
 Case Else
   KeyAscii = 0
End Select
End Sub


Private Sub txtContrato_LostFocus()
If Trim(txtContrato) <> "" Then
   Call sbConsultaContrato(Trim(txtContrato))
Else
   txtDetalle = ""
End If
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContrato.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub




