VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCR_APA_OperacionRenumera 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Renumeración de la Operación"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFecha 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5520
      TabIndex        =   11
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5520
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtOperacionNuevo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtSaldo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5520
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtOperacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.TextBox txtCod_Acreedor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin XtremeSuiteControls.PushButton btnRenumerar 
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   2640
      Width           =   2415
      _Version        =   1441792
      _ExtentX        =   4260
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Procesar Renumeración"
      BackColor       =   -2147483633
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
      Picture         =   "frmCR_APA_OperacionRenumera.frx":0000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Formalizado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Operación Nuevo.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Operación Actual"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Acreedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmCR_APA_OperacionRenumera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnRenumerar_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spAPA_OperacionRenumera '" & txtCod_Acreedor.Text & "','" & txtOperacion.Text _
       & "','" & txtOperacionNuevo.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Renumeracion de Operación.:" & txtOperacion.Text & " -> " & txtOperacionNuevo.Text)

txtOperacion.Text = ""
txtOperacionNuevo.Text = ""
txtMonto.Text = ""
txtSaldo.Text = ""
txtFecha.Text = ""

MsgBox "Renumeración realizada satisfactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
vModulo = 14



End Sub


Private Sub txtCod_Acreedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_ACREEDOR"
  gBusquedas.Orden = "COD_ACREEDOR"
  gBusquedas.Consulta = "SELECT COD_ACREEDOR,DESCRIPCION FROM CRD_APA_ACREEDORES"
  gBusquedas.Resultado = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     txtCod_Acreedor.Text = gBusquedas.Resultado
     txtDescripcion.Text = gBusquedas.Resultado2
  End If

End If

End Sub

Private Sub txtCod_Acreedor_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "SELECT COD_ACREEDOR,DESCRIPCION FROM CRD_APA_ACREEDORES" _
       & " where cod_acreedor = '" & txtCod_Acreedor.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
   txtDescripcion.Text = rs!Descripcion
End If
rs.Close

txtOperacion.Text = ""
txtOperacionNuevo.Text = ""
txtMonto.Text = ""
txtSaldo.Text = ""
txtFecha.Text = ""

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOperacion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "DESCRIPCION"
  gBusquedas.Orden = "DESCRIPCION"
  gBusquedas.Consulta = "SELECT COD_ACREEDOR,DESCRIPCION FROM CRD_APA_ACREEDORES"
  gBusquedas.Resultado = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     txtCod_Acreedor.Text = gBusquedas.Resultado
     txtDescripcion.Text = gBusquedas.Resultado2
  End If

End If
End Sub


Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOperacionNuevo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "OPERACION"
  gBusquedas.Orden = "OPERACION"
  gBusquedas.Filtro = " and cod_acreedor = '" & txtCod_Acreedor.Text & "'"
  gBusquedas.Consulta = "select OPERACION,MONTO, SALDO, FECHA_FORMALIZA  From CRD_APA_OPERACIONES"
  gBusquedas.Resultado = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     txtOperacion.Text = gBusquedas.Resultado
  End If

End If
End Sub

Private Sub txtOperacion_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "SELECT  OPERACION,MONTO, SALDO, FECHA_FORMALIZA FROM CRD_APA_OPERACIONES" _
       & " where cod_acreedor = '" & txtCod_Acreedor.Text _
       & "' and Operacion = '" & txtOperacion.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
   txtSaldo.Text = Format(rs!Saldo, "Standard")
   txtMonto.Text = Format(rs!Monto, "Standard")
   txtFecha.Text = Format(rs!Fecha_Formaliza, "dd/mm/yyyy")
End If
rs.Close

txtOperacionNuevo.Text = ""

End Sub
