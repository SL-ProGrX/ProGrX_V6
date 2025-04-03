VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Devoluciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Devoluciones (Prioridad del proceso de aplicación)"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   7935
      _Version        =   524288
      _ExtentX        =   13996
      _ExtentY        =   5530
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
      MaxCols         =   6
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_DevolucionesTabla.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFSL_DevolucionesTabla.frx":06B0
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   7335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Devoluciones"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmFSL_DevolucionesTabla.frx":0772
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmFSL_Devoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim rs As New ADODB.Recordset

Dim vPorcentaje As Currency, vDevolucion As Integer
Dim vFechaInicio, vFechaFinal As String, vFecha As String
Dim vAplMontoFormalizado As Integer, vAplSaldo As Integer

Private Sub Form_Activate()
  vModulo = 22
End Sub

Private Sub sbGuardaFecha()
On Error GoTo error
   
   strSQL = "Insert FSL_TIPO_APLICACION(COD_DEVOLUCION, FECHA_INICIO, FECHA_FINAL" _
          & ",APL_MONTO_FORMALIZADO,APL_SALDO,PORCENTAJE,REGISTRO_FECHA, REGISTRO_USUARIO)" _
          & " Values(" & vDevolucion & ",'" & vFechaInicio & "','" & vFechaFinal & "','" & vAplMontoFormalizado & "','" & vAplSaldo & "'" _
          & ", " & vPorcentaje & ",'" & Format(fxFechaServidor, "yyyymmdd") & "','" & glogon.Usuario & "')"
          
   glogon.Conection.Execute strSQL
   
   Exit Sub
       
error:
   MsgBox Err.Description
   
End Sub

Private Sub sbModificaFecha()
On Error GoTo error
   
   strSQL = "Update FSL_TIPO_APLICACION SET FECHA_INICIO='" & vFechaInicio & "', FECHA_FINAL='" & vFechaFinal & "'" _
          & ", APL_MONTO_FORMALIZADO = '" & vAplMontoFormalizado & "',APL_SALDO = '" & vAplSaldo & "',REGISTRO_FECHA='" & Format(fxFechaServidor, "yyyymmdd") & "'" _
          & ", PORCENTAJE = " & vPorcentaje & ",REGISTRO_USUARIO='" & glogon.Usuario & "' where COD_DEVOLUCION=" & vDevolucion & ""
   
   glogon.Conection.Execute strSQL

   Exit Sub

error:
   MsgBox Err.Description

End Sub

Private Sub sbBorraFecha()
On Error GoTo error
   
   strSQL = "DELETE FROM FSL_TIPO_APLICACION" _
          & " WHERE COD_DEVOLUCION = " & vDevolucion & ""

   glogon.Conection.Execute strSQL

   Exit Sub

error:
   MsgBox Err.Description
   
End Sub

Private Function fxValidaDevolucion() As Boolean
Dim i As Integer

On Error GoTo error
 
   With vGrid
     .Row = .ActiveRow
     .Col = 1
           
     If .Text = Empty Then
       strSQL = "Select isnull(max(COD_DEVOLUCION)+ 1,1) as Rango from FSL_TIPO_APLICACION"
       rs.Open strSQL, glogon.Conection, adOpenStatic
       
       vDevolucion = rs!Rango
       fxValidaDevolucion = False
       
       rs.Close
     Else
     
       vDevolucion = .Text
       fxValidaDevolucion = True
        
     End If
    
     .Col = 2
     vFechaInicio = Format(.Text, "yyyymmdd")
        
     .Col = 3
     vFechaFinal = Format(.Text, "yyyymmdd")
     
     .Col = 4
     vAplMontoFormalizado = .Value
     
     .Col = 5
     vAplSaldo = .Value
     
     .Col = 6
     vPorcentaje = CCur(.Text)
     
   End With

   Exit Function

error:
  MsgBox Err.Description
  
End Function

Private Sub Form_Load()
   vModulo = 22
   Call sbCargaFechas
End Sub

Private Sub sbCargaFechas()
Dim i As Integer
Dim FechaIni, FechaFin As Date

On Error GoTo error
    
 With vGrid
      .MaxRows = 1
      .Row = .MaxRows
      For i = 1 To .MaxCols
       .Col = i
       .Text = ""
      Next i
     
      strSQL = "SELECT COD_DEVOLUCION, FECHA_INICIO, FECHA_FINAL, APL_MONTO_FORMALIZADO,APL_SALDO,PORCENTAJE " _
             & "From FSL_TIPO_APLICACION"
      
      rs.Open strSQL, glogon.Conection, adOpenStatic
        
      Do While Not rs.EOF
        .Row = .MaxRows
        .Col = 1
        .Text = rs!COD_DEVOLUCION
          
        .Col = 2
        .Text = rs!FECHA_INICIO
          
        .Col = 3
        .Text = rs!FECHA_FINAL
          
        .Col = 4
        .Value = rs!APL_MONTO_FORMALIZADO
        
        .Col = 5
        .Value = rs!APL_SALDO
        
        .Col = 6
        .Text = Format(rs!Porcentaje, "standard")
        
        rs.MoveNext
          
        .MaxRows = .MaxRows + 1
      Loop
  
      rs.Close
      
End With
     
Exit Sub

error:
  MsgBox Err.Description
  
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo error
  
 With vGrid
  If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
      If fxValidaDevolucion = False Then
         Call sbGuardaFecha
         .Col = 1
         .Text = vDevolucion
         .MaxRows = .MaxRows + 1
      Else
         Call sbModificaFecha
      End If
  ElseIf KeyCode = vbKeyDelete Then
    If fxValidaDevolucion = True Then
       Call sbBorraFecha
    End If
    Call sbCargaFechas
  End If
 End With
      
    
Exit Sub

error:
  MsgBox Err.Description

End Sub
