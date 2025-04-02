VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_CD_Cargos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargos y Retenciones"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11880
   Icon            =   "frmAF_CD_Cargos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11655
      _Version        =   524288
      _ExtentX        =   20558
      _ExtentY        =   10186
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
      MaxCols         =   486
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_CD_Cargos.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Cargos Aplicables a Cuentas de Comités"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   6795
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
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
Dim strSQL As String
Dim rs As New ADODB.Recordset


Private Sub Form_Activate()
 vModulo = 40
End Sub

Private Sub Form_Load()
 
vModulo = 40
  
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.MaxCols = 4
vGrid.MaxRows = 1
Call sbCargaCargos

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)

 With vGrid
     If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
        If fxValida = False Then
           Call sbGuardaCargo
           Call sbCargaCargos
        Else
           Call sbModificaCargo
           Call sbCargaCargos
        End If
     ElseIf .ActiveCol = .MaxCols And (KeyCode = vbKeyDelete) Then
        If fxValida = True Then
           Call sbEliminaCargo
           Call sbCargaCargos
        End If
     End If
     
     If KeyCode = vbKeyF4 And (.ActiveCol = 3) Then
        Call sbgCntCuentaConsulta("D")
        .Col = .ActiveCol
        .Row = .ActiveRow
        .Text = gBusquedas.Resultado
     End If
     
  End With
 
End Sub

Private Sub sbGuardaCargo()

   strSQL = "Insert AFI_CD_CARGOS (CODIGO, DESCRIPCION, CUENTA, ESTADO) " _
          & "values (" & vCodigo & ",'" & vDescripcion & "','" & vCuenta & "'," & vEstado & ")"
   
   Call ConectionExecute(strSQL)
   
   
   
End Sub

Private Sub sbModificaCargo()
   
   strSQL = "Update AFI_CD_CARGOS set DESCRIPCION = '" & vDescripcion & "',CUENTA='" & vCuenta & "', " _
          & "ESTADO = " & vEstado & " where CODIGO = " & vCodigo & ""
   Call ConectionExecute(strSQL)
   
End Sub

Private Sub sbEliminaCargo()
  
  strSQL = "delete AFI_CD_CARGOS where CODIGO = " & vCodigo & ""
  Call ConectionExecute(strSQL)
 
End Sub

Private Function fxValida() As Boolean
On Error GoTo error

  With vGrid
     .Row = .ActiveRow
     .Col = 1
     
     If .Text = Empty Then
       strSQL = "Select coalesce(max(CODIGO),0) + 1 as Codigo from AFI_CD_CARGOS"
       Call OpenRecordSet(rs, strSQL)
       
       If Not rs.EOF Then
         vCodigo = rs!Codigo
         fxValida = False
       End If
        
       rs.Close
       
     Else
       
       vCodigo = .Text
       fxValida = True
       
     End If
        
     .Col = 2
     vDescripcion = .Text
     
     
     .Col = 3
     vCuenta = .Text
     
     .Col = 4
     vEstado = .Value
     
  End With
   
Exit Function
  
error:
  MsgBox Err.Description, vbCritical

End Function

Private Sub sbCargaCargos()
Dim i As Integer
On Error GoTo error

   With vGrid
       .MaxRows = 1
       For i = 1 To .MaxCols
         .Col = i
         .Text = ""
       Next i
            
       strSQL = "Select CODIGO, DESCRIPCION, CUENTA, ESTADO from AFI_CD_CARGOS"
       Call OpenRecordSet(rs, strSQL)
    
       Do While Not rs.EOF
          .Row = .MaxRows
          .Col = 1
          .Text = rs!Codigo
           
          .Col = 2
          .Text = rs!Descripcion
           
          .Col = 3
          .Text = rs!Cuenta
        
          .Col = 4
          .Value = rs!Estado
        
          .MaxRows = .MaxRows + 1
          rs.MoveNext
       Loop
       rs.Close
       
   End With
Exit Sub

error:
  MsgBox Err.Description, vbCritical
End Sub
