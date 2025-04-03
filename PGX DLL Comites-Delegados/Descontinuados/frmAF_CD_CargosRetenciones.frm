VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_CD_Cargos 
   Caption         =   "Cargos y Retenciones"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _Version        =   524288
      _ExtentX        =   16748
      _ExtentY        =   5953
      _StockProps     =   64
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
      MaxCols         =   486
      SpreadDesigner  =   "frmAF_CD_CargosRetenciones.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmAF_CD_Cargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vCodigo As Integer, vEstado As Integer
Dim vDescripcion As String, vTipo As String, vCuenta As String
Dim strsql As String
Dim rs As New ADODB.Recordset


Private Sub Form_Activate()
 vModulo = 1
End Sub

Private Sub Form_Load()
  vGrid.MaxCols = 5
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)

 With vGrid
     If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
        If fxValida = False Then
           Call sbGuardaCargo
        Else
           Call sbModificaCargo
        End If
     ElseIf .ActiveCol = .MaxCols And (KeyCode = vbKeyDelete) Then
        If fxValida = True Then
           Call sbEliminaCargo
        End If
     End If
     
     If KeyCode = vbKeyF4 And (.ActiveCol = 4) Then
        Call sbgCntCuentaConsulta("D")
        .Col = .ActiveCol
        .Row = .ActiveRow
        .Text = gBusquedas.Resultado
     End If
     
  End With
 
End Sub

Private Sub sbGuardaCargo()

   strsql = "Insert AFI_CD_CARGOS_RETENCIONES (CODIGO, DESCRIPCION, CUENTA, ESTADO) " _
          & "values (" & vCodigo & ",'" & vDescripcion & "','" & vCuenta & "'," & vEstado & ")"
   
   glogon.Conection.Execute strsql
   
End Sub

Private Sub sbModificaCargo()
   
   strsql = "Update AFI_CD_CARGOS_RETENCIONES set DESCRIPCION='" & vDescripcion & "',CUENTA='" & vCuenta & "', " _
          & "ESTADO=" & vEstado & " where CODIGO = " & vCodigo & ""
   glogon.Conection.Execute strsql
   
End Sub

Private Sub sbEliminaCargo()
  
  strsql = "delete AFI_CD_CARGOS_RETENCIONES where CODIGO = " & vCodigo & ""
  glogon.Conection.Execute strsql
 
End Sub

Private Function fxValida() As Boolean
On Error GoTo Error

  With vGrid
     .Row = .ActiveRow
     .Col = 1
     
     If .Text = Empty Then
       strsql = "Select coalesce(max(CODIGO),0) + 1 as Codigo from AFI_CD_CARGOS_RETENCIONES"
       rs.Open strsql, glogon.Conection, adOpenStatic

       vCodigo = rs!Codigo
        
       rs.Close
       
       fxValida = False
       
     Else
       vCodigo = .Text
       
       fxValida = True
     End If
        
     .Col = 2
     vDescripcion = .Text
     
     .Col = 3
     vTipo = Mid(.Text, 1, 1)
     
     .Col = 4
     vCuenta = .Text
     
     .Col = 5
     vEstado = .Value
     
  End With
   
Exit Function

Error:
  MsgBox Err.Description, vbCritical

End Function
