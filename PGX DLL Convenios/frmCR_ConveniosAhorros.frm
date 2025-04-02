VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCR_ConveniosAhorros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Deducciones por planes de Ahorros"
   ClientHeight    =   5568
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   12492
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5568
   ScaleWidth      =   12492
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMensualidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   315
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4092
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   10332
      _Version        =   524288
      _ExtentX        =   18225
      _ExtentY        =   7218
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   483
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_ConveniosAhorros.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Index           =   8
      Left            =   9720
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Monto Aplicar"
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
      Left            =   8640
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Monto Mensualidades"
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
      Index           =   1
      Left            =   4440
      TabIndex        =   6
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Orden"
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
      Index           =   20
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio"
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
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmCR_ConveniosAhorros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vOrden As Integer
Dim vPlan As String, vContrato As Integer
Dim vEstado As String

Private Sub Form_Activate()
    vModulo = 16
End Sub

Private Sub Form_Load()
    vModulo = 16
    
    Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
    vCodigo = GLOBALES.gTag
    vOrden = GLOBALES.gTag2
    
    vGrid.MaxRows = 0
    vGrid.MaxCols = 6
    
    txtMensualidad.Text = 0
    txtTotal.Text = 0
        
    Call sbConsultaConvenio(vCodigo, vOrden)
    Call sbConsultaAhorros
End Sub


Private Sub sbCalculaTotales()
Dim curTotal As Currency, curMensualidad As Currency
Dim i As Integer

curTotal = 0
curMensualidad = 0

With vGrid
    For i = 1 To .MaxRows
     .Row = i
     .Col = 4
     curMensualidad = curMensualidad + CCur(.Text)
     
     .Col = 6
     curTotal = curTotal + CCur(.Text)
    Next i
End With

txtTotal.Text = Format(curTotal, "Standard")
txtMensualidad.Text = Format(curMensualidad, "Standard")
End Sub

Private Sub sbConsultaConvenio(vConvenio As String, vOrden As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
  
On Error GoTo vError

   strSQL = " select O.COD_CONVENIO,C.DESCRIPCION,O.COD_ORDEN,O.ESTADO" _
          & " from CRD_CONVENIOS_ORDENES O" _
          & "  inner join CRD_CONVENIOS C on O.COD_CONVENIO = C.COD_CONVENIO" _
          & " where O.COD_CONVENIO = '" & vConvenio & "' AND O.COD_ORDEN = " & vOrden & " "
   Call OpenRecordSet(rs, strSQL)

   If Not rs.EOF Then
      txtCodigo.Text = rs!COD_CONVENIO
      txtDescripcion.Text = rs!Descripcion
      txtOrden.Text = rs!cod_orden
      vEstado = rs!estado
      
      Select Case vEstado
        Case "A"
          txtEstado.Text = "Abierta"
        Case "C"
          txtEstado.Text = "Cerrada"
        Case "T"
          txtEstado.Text = "Tramitada"
      End Select
      
   End If

   rs.Close

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsultaAhorros()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPendiente As Currency, vMensualidad As Currency, vTotal As Currency

On Error GoTo vError

vMensualidad = 0
vPendiente = 0
vTotal = 0

strSQL = "exec spConvenios_Orden_Ahorros '" & vCodigo & "'," & vOrden
Call OpenRecordSet(rs, strSQL)

With vGrid

 
Do While Not rs.EOF
   
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  
  .Col = 1
  .Text = rs!cod_Contrato
  
  .Col = 2
  .Text = rs!cod_Plan
  
  .Col = 3
  .Text = rs!Descripcion
  
  .Col = 4
  .Text = Format(rs!Mensualidad, "Standard")
  

  If rs!Aporte > rs!Mensualidad Then
     vPendiente = 0
  Else
     vPendiente = rs!Mensualidad - rs!Aporte
  End If

  
  .Col = 5
  .Text = Format(vPendiente, "Standard")
  
  .Col = 6
  .Text = Format(rs!Monto, "Standard")
   
   vTotal = vTotal + rs!Monto
   vMensualidad = vMensualidad + rs!Mensualidad
  rs.MoveNext
Loop

rs.Close
End With

txtMensualidad.Text = Format(vMensualidad, "Standard")
txtTotal.Text = Format(vTotal, "Standard")

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency

On Error GoTo vError
 curMonto = 0
 vContrato = 0

 If vEstado = "C" Then
    MsgBox "No se puede modificar una Orden con estado Cerrada!!"
    Exit Sub
 End If
  
 With vGrid
   If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
     
     .Row = .ActiveRow
     
     .Col = 1
     vContrato = .Text
     
     .Col = 2
     vPlan = .Text
     
     .Col = 6
     curMonto = .Text
     
     If vContrato > 0 Then
         strSQL = "exec spConvenios_Orden_FondosPool '" & vCodigo & "'," & vOrden & ",'" & vPlan & "'," & vContrato & "," & curMonto
         Call ConectionExecute(strSQL)
     End If
  
     Call sbCalculaTotales
  
  End If
   
 End With
  
  
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Function fxValidaAhorro(vCodigo As String, vPlan As String, vContrato As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

fxValidaAhorro = True

strSQL = " Select COD_CONVENIO" _
       & " from CRD_CONVENIOS_DT_AHORROS" _
       & " where COD_CONVENIO = '" & vCodigo & "' And cod_Orden = " & vOrden _
       & " and COD_PLAN = '" & vPlan & "' and cod_contrato = " & vContrato
       
Call OpenRecordSet(rs, strSQL)

If rs.EOF Then
  fxValidaAhorro = False
End If

rs.Close

Exit Function
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

