VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmFNDRetiros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Retiros"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4155
   Icon            =   "frmFNDRetiros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread vGrid 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _Version        =   393216
      _ExtentX        =   7011
      _ExtentY        =   7011
      _StockProps     =   64
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   1
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDRetiros.frx":030A
      VisibleCols     =   500
      VisibleRows     =   500
   End
End
Attribute VB_Name = "frmFNDRetiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

vModulo = 18 'Fondo de Inversion


strSQL = "Select * From FND_Tabla_Retiros where Cod_Operadora="
strSQL = strSQL & frmFNDPlanes.cboOperadora.ItemData(frmFNDPlanes.cboOperadora.ListIndex)
strSQL = strSQL & " And Cod_Plan='" & Trim(frmFNDPlanes.txtCodigo) & "'"

With rs
 .Open strSQL, glogon.Conection, adOpenStatic
   vGrid.MaxRows = IIf(.RecordCount = 0, 1, .RecordCount)
   For i = 1 To .RecordCount
      vGrid.Row = i
      vGrid.Col = 1
      vGrid.Text = !Desde
      vGrid.Col = 2
      vGrid.Text = !Hasta
      vGrid.Col = 3
      vGrid.Text = Format(!Porcentaje, "###.00")
      
      .MoveNext
   Next i
 .Close
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strSQL As String, i As Integer
Dim vOperadora As Long
Dim vPlan As String

vOperadora = frmFNDPlanes.cboOperadora.ItemData(frmFNDPlanes.cboOperadora.ListIndex)
vPlan = Trim(frmFNDPlanes.txtCodigo)

On Error GoTo vError
glogon.Conection.BeginTrans

strSQL = "Delete From Fnd_Tabla_Retiros where Cod_Operadora = "
strSQL = strSQL & vOperadora
strSQL = strSQL & " And Cod_Plan='" & vPlan & "'"
glogon.Conection.Execute strSQL
Call Bitacora("Borra", "Lineas en Tabla Retiros Operadora:" & vOperadora & " Plan:" & vPlan)

For i = 1 To vGrid.MaxRows
 If fxVerificaLinea(i) = True Then
  vGrid.Row = i
  vGrid.Col = 1
  strSQL = "Insert into FND_Tabla_Retiros(Cod_operadora,Cod_Plan,Desde,Hasta,Porcentaje)"
  strSQL = strSQL & "Values(" & vOperadora & ",'" & vPlan & "'," & vGrid.Text & ","
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & ")"
  glogon.Conection.Execute strSQL
 End If
 Call Bitacora("Registra", "Lineas en Tabla Retiros Operadora:" & vOperadora & " Plan:" & vPlan)
Next i

glogon.Conection.CommitTrans

Exit Sub
vError:
  glogon.Conection.RollbackTrans
  MsgBox Err.Description
End Sub


Private Sub vGrid_Advance(ByVal AdvanceNext As Boolean)
Dim intI As Integer

If vGrid.ActiveRow = 1 And vGrid.MaxRows > 1 Then
   Exit Sub
End If

Select Case vGrid.ActiveCol
  Case 3
    For intI = 1 To 3
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = intI

        If Trim(vGrid.Text) = "" Then
           Exit Sub
        End If
    Next
    
    vGrid.MaxRows = vGrid.MaxRows + 1

End Select

End Sub

Function fxVerificaLinea(i As Integer)
vGrid.Row = i

vGrid.Col = 1
If vGrid.Text = "" Then
   fxVerificaLinea = False
Else
  vGrid.Col = 2
  If vGrid.Text = "" Then
     fxVerificaLinea = False
  Else
    vGrid.Col = 3
    If vGrid.Text = "" Then
       fxVerificaLinea = False
    Else
       fxVerificaLinea = True
    End If
  End If
End If


End Function
