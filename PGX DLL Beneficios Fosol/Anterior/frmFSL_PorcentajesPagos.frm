VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_PorcentajesPagos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Fallecimiento"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   5955
      _Version        =   524288
      _ExtentX        =   10504
      _ExtentY        =   8070
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
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
      FormulaSync     =   0   'False
      MaxCols         =   4
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_PorcentajesPagos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblDescripcion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   120
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Fallecimiento"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmFSL_PorcentajesPagos.frx":06B1
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmFSL_PorcentajesPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vDevolucion, vMeses_Inicio, vMeses_Final As Integer
Dim vFactor, vCobertura, vCoberturaAcumulada As Double

Private Sub sbGuardaPorcentaje()
On Error GoTo error
   
   strSQL = "Insert FSL_PORCENTAJE_FALLECIMIENTO (COD_RANGO, MEMBRESIA_MESES_INICIO, MEMBRESIA_MESES_CORTE,COBERTURA_ACUMULADA, REGISTRO_FECHA, REGISTRO_USUARIO) " _
          & "values (" & vDevolucion & "," & vMeses_Inicio & "," & vMeses_Final & "," & vCoberturaAcumulada & ",'" & Format(fxFechaServidor, "yyyymmdd") & "','" & glogon.Usuario & "')"
   
   glogon.Conection.Execute strSQL
  
Exit Sub
error:
  MsgBox Err.Description
End Sub

Private Sub sbModificaPorcentaje()
On Error GoTo error
   
   strSQL = " Update FSL_PORCENTAJE_FALLECIMIENTO set  MEMBRESIA_MESES_INICIO=" & vMeses_Inicio & ", MEMBRESIA_MESES_CORTE=" & vMeses_Final & ", " _
          & " COBERTURA_ACUMULADA=" & vCoberturaAcumulada & " where COD_RANGO=" & vDevolucion & ""
   
   glogon.Conection.Execute strSQL

Exit Sub
error:
  MsgBox Err.Description
  
End Sub

Private Sub sbEliminaPorcentaje()
On Error GoTo error

   strSQL = "Delete FSL_PORCENTAJE_FALLECIMIENTO where COD_RANGO=" & vDevolucion & ""
   
   glogon.Conection.Execute strSQL
   
Exit Sub
error:
  MsgBox Err.Description
  
End Sub

Private Function fxValidaPorcentaje() As Boolean
With vGrid
   .Row = .ActiveRow
   .Col = 1
   If .Text = Empty Then
     strSQL = "Select isnull(max(COD_RANGO) + 1,1) as Rango from FSL_PORCENTAJE_FALLECIMIENTO"
     rs.Open strSQL, glogon.Conection, adOpenStatic
     
     vDevolucion = rs!Rango
     fxValidaPorcentaje = False
     
     rs.Close
   Else
     vDevolucion = CInt(.Text)
     fxValidaPorcentaje = True
   End If
   
   .Col = 2
   If .Text = Empty Then
     Exit Function
   Else
     vMeses_Inicio = .Text
   End If
   
   .Col = 3
      If .Text = Empty Then
     Exit Function
   Else
     vMeses_Final = .Text
   End If
   
   .Col = 4
   vCoberturaAcumulada = CDbl(IIf(.Text = "", 0, .Text))
   
End With
  
End Function

Private Sub Form_Activate()
 vModulo = 22
End Sub

Private Sub Form_Load()
  vModulo = 22
  Call sbCargaRangos
  lblDescripcion.Caption = "Registro de porcentajes de cobertura según" & vbCrLf & ""
  lblDescripcion.Caption = lblDescripcion.Caption & "la cantidad de meses transcurridos desde la primera deducción" & vbCrLf & ""
  lblDescripcion.Caption = lblDescripcion.Caption & "de la operación."
  
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
  With vGrid
    If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
       If fxValidaPorcentaje = False Then
          Call sbGuardaPorcentaje
          .MaxRows = .MaxRows + 1
          .Col = 1
          .Text = vDevolucion
       Else
          Call sbModificaPorcentaje
       End If
    ElseIf KeyCode = vbKeyDelete Then
      If fxValidaPorcentaje = True Then
         Call sbEliminaPorcentaje
         Call sbCargaRangos
      End If
    End If
  End With
End Sub

Private Sub sbCargaRangos()
  Dim i As Integer
  
  With vGrid
    .MaxRows = 1
    .Row = .MaxRows
    For i = 1 To .MaxCols
     .Col = i
     .Text = ""
    Next i
      
    strSQL = "select COD_RANGO, MEMBRESIA_MESES_INICIO, MEMBRESIA_MESES_CORTE, " _
           & "COBERTURA_ACUMULADA From FSL_PORCENTAJE_FALLECIMIENTO"
         
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    Do While Not rs.EOF
      .Row = .MaxRows
      .Col = 1
      .Text = rs!COD_RANGO
      .Col = 2
      .Text = rs!MEMBRESIA_MESES_INICIO
      .Col = 3
      .Text = rs!MEMBRESIA_MESES_CORTE
      .Col = 4
      .Text = Format(rs!COBERTURA_ACUMULADA, "standard")
      rs.MoveNext
      
      .MaxRows = .MaxRows + 1
      .Row = .ActiveRow + 1
    Loop
    rs.Close

  End With
End Sub

