VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Gestiones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FOSOL: Tabla de Gestiones"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGridGestiones 
      Height          =   4080
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   7815
      _Version        =   524288
      _ExtentX        =   13785
      _ExtentY        =   7197
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
      MaxCols         =   3
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_Gestiones.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8160
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmFSL_Gestiones.frx":05C2
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestiones para Seguimiento de Casos"
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
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmFSL_Gestiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vCod_Gestion As Integer
Dim vDescripcion As String
Dim vEstado As String

Private Sub Form_Activate()
 vModulo = 22
End Sub

Private Sub Form_Load()
 vModulo = 22
 Call sbCargaGestiones
End Sub

Private Sub sbGuardaGestion()
On Error GoTo vError
  
strSQL = "Insert FSL_GESTIONES (COD_GESTION, DESCRIPCION, ESTADO)" _
       & " values (" & vCod_Gestion & ",'" & vDescripcion & "','" & vEstado & "')"
glogon.Conection.Execute strSQL

Exit Sub
       
vError:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub sbModificaGestion()
On Error GoTo vError
  
strSQL = "Update FSL_GESTIONES set DESCRIPCION = '" & vDescripcion & "', ESTADO = '" & vEstado & "' " _
       & " where COD_GESTION=" & vCod_Gestion & ""
glogon.Conection.Execute strSQL

Exit Sub
       
vError:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub sbEliminaGestion()
On Error GoTo vError
  
strSQL = "Delete FSL_GESTIONES where COD_GESTION =" & vCod_Gestion & ""
glogon.Conection.Execute strSQL

Exit Sub
       
vError:
   MsgBox Err.Description, vbExclamation
End Sub
'Valida si existe la Gestion
Private Function fxValidaGestion() As Boolean
On Error GoTo vError

With vGridGestiones
 .Col = 1
      
 If .Text = Empty Then
  strSQL = "Select isnull(max(COD_GESTION) + 1,1) as 'Codigo' from FSL_GESTIONES"
  rs.Open strSQL, glogon.Conection, adOpenStatic
       
  vCod_Gestion = rs!Codigo
  fxValidaGestion = False
  rs.Close
 Else
  vCod_Gestion = .Text
  fxValidaGestion = True
 End If
 .Col = 2
 vDescripcion = UCase(.Text)
        
 .Col = 3
 vEstado = .Value
End With

Exit Function
       
vError:
   MsgBox Err.Description, vbExclamation
End Function

Private Sub vGridGestiones_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
With vGridGestiones
.Row = .ActiveRow
  
 If .ActiveCol = 3 Then
 
  Select Case KeyCode
    Case vbKeyReturn 'Guardar nuevo o modificar
      If fxValidaGestion = False Then
         Call sbGuardaGestion
         .Col = 1
         .Text = vCod_Gestion
         vGridGestiones.MaxRows = vGridGestiones.MaxRows + 1
      Else
         Call sbModificaGestion
      End If
    
    Case vbKeyDelete 'Borrar
      If fxValidaGestion = True Then
        Call sbEliminaGestion
        Call sbCargaGestiones
      End If

  End Select
 End If
 
End With

Exit Sub
vError:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub sbCargaGestiones()
Dim i As Integer
On Error GoTo vError

With vGridGestiones
.MaxRows = 1
.Row = .MaxRows
For i = 1 To .MaxCols
  .Col = i
  .Text = ""
Next i

strSQL = "Select COD_GESTION, DESCRIPCION, ESTADO from FSL_GESTIONES"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF

 .Col = 1
 .Text = rs!COD_GESTION
 .Col = 2
 .Text = rs!Descripcion
 .Col = 3
 .Value = rs!Estado
 .MaxRows = .MaxRows + 1
 .Row = .MaxRows
 rs.MoveNext
Loop
rs.Close
.Row = .MaxRows
End With

Exit Sub
vError:
   MsgBox Err.Description, vbExclamation
End Sub
