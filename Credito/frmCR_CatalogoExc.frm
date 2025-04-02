VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_CatalogoExc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Cálculo de Diponible Creditos Sobre Excedentes"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdModifica 
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   5640
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Actualiza Tabla"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10455
      _Version        =   524288
      _ExtentX        =   18441
      _ExtentY        =   6800
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_CatalogoExc.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Porcentajes Disponibles para Crédito con Garantía en Excedentes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   300
      Width           =   6615
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCR_CatalogoExc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    vGrid.Text = CStr(rs.Fields(i - 1).Value)
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub



Private Sub cmdModifica_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

Dim pMes As Integer, pMesAcumulado As Integer, pPorc As Currency, pCapGen As Currency

vGrid.Row = 1
vGrid.col = 1

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  If vGrid.Text <> "" Then
         
    strSQL = strSQL & Space(10) & "insert EXC_DISPONIBLE(MES,ACUMULADO_MES,ACUMULADO_PORC,CAPGEN) values("
    vGrid.col = 1
    pMes = vGrid.Text
    
    strSQL = strSQL & vGrid.Text & ","
    vGrid.col = 2
    pMesAcumulado = vGrid.Text
    
    strSQL = strSQL & vGrid.Text & ","
    vGrid.col = 3
    pPorc = CCur(vGrid.Text)
    
    strSQL = strSQL & vGrid.Text & ","
    vGrid.col = 4
    pCapGen = CCur(vGrid.Text)
    strSQL = strSQL & vGrid.Text & ")"
  End If
Next i

'Procesa Lote
Call ConectionExecute(strSQL)

MsgBox "Tabla para Disponibles de Excedentes actualizada satisfactoriamente!", vbInformation
Unload Me

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle
 
 

strSQL = "select MES,ACUMULADO_MES,ACUMULADO_PORC,CAPGEN, REGISTRO_FECHA, REGISTRO_USUARIO" _
       & " from EXC_DISPONIBLE order by mes"
Call sbCargaGrid(vGrid, 6, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

If Not cmdModifica.Enabled Then vGrid.Enabled = False

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(6) As Variant, x As Integer
Dim strSQL As String, rs As New ADODB.Recordset

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.Text <> "" Then
    strSQL = "delete EXC_DISPONIBLE where MES = " & vGrid.Text
    Call ConectionExecute(strSQL)
  End If
  
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 4
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To 4
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To 4
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
End If


If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
    vGrid.col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
 If vGrid.ActiveCol = 4 Then
    If vGrid.MaxRows = vGrid.Row Then
        vGrid.MaxRows = vGrid.MaxRows + 1
        vGrid.Row = vGrid.MaxRows
    End If
 End If
 
End If

End Sub

