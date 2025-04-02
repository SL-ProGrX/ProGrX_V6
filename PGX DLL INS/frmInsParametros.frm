VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmInsParametros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "INS: Parámetros"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInsParametros.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8415
      _Version        =   524288
      _ExtentX        =   14843
      _ExtentY        =   9128
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmInsParametros.frx":6852
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Parámetros Generales del Módulo INS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8640
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmInsParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vModulo = 17
vGrid.AppearanceStyle = fxGridStyle
 
Call Formularios(Me)


'Inicializa Parametros
strSQL = "exec spInsParametros"
glogon.Conection.Execute strSQL

strSQL = "select * from Ins_parametros" _
      & " order by cod_parametro"
Call sbCargaGridLocal(vGrid, 3, strSQL)


Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox Err.Description, vbExclamation

End Sub


Public Sub sbCargaGridLocal(pGrid As Object, MaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

With pGrid
    .MaxRows = 0
    .MaxCols = MaxCol
    rs.Open strSQL, glogon.Conection, adOpenForwardOnly
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      For i = 1 To 3
        .col = i
        Select Case i
          Case 1 'Codigo
            .CellTag = rs!Tipo & ""
            .Text = rs!Cod_Parametro
            .CellNote = "Modificado Por: " & rs!MODIFICA_USUARIO & vbCrLf & "Fecha: " & rs!MODIFICA_FECHA
          
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 2 'Descripcion
            .Text = rs!Descripcion
            .CellNote = rs!notas & ""
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 3 'Valor
            If UCase(Trim(rs!Tipo)) = "CTA" Then
                .TextTip = TextTipFixed
                .TextTipDelay = 1000
                .CellNoteIndicatorColor = vbBlue
                .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
                
                .Text = fxgCntCuentaFormato(True, Trim(rs!Valor), 0)
                .CellNote = fxgCntCuentaDesc(Trim(rs!Valor))
            Else
                .Text = rs!Valor
            End If
            
        End Select
      Next i
      rs.MoveNext
    Loop
    rs.Close

End With

End Sub




Private Sub sbGuardaParametro(pParametro As String, pValor As String _
                    , Optional pTipo As String = "DEC")
Dim strSQL As String, rs As New ADODB.Recordset
Dim Validacion As Boolean, vMensaje As String

On Error GoTo vError

Validacion = True
vMensaje = ""

Select Case Trim(pTipo)
  Case "DEC" 'Decimal
    If IsNumeric(pValor) Then
       pValor = CCur(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido...!!!"
    End If
    
  Case "NUM" 'Número Entero
    If IsNumeric(pValor) Then
       pValor = CLng(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido...!!!"
    End If
  
  Case "POR" 'Porcentaje
    If IsNumeric(pValor) Then
       pValor = CCur(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido, suministre un porcentaje ..!!!"
    End If
  
  Case "CTA" 'Cuenta Contable
    Validacion = fxgCntCuentaValida(fxgCntCuentaFormato(False, pValor))
    If Not Validacion Then
        vMensaje = "La Cuenta indicada no es válida, presiones F4 para buscar en el catálogo...!!!"
    Else
      pValor = fxgCntCuentaFormato(False, pValor)
    End If
    
  Case "CHR" 'Caracteres
    If InStr(1, pValor, "'", vbTextCompare) > 0 Then
       Validacion = False
       vMensaje = "El valor indicado contiene caracteres no válidos...!!!"
    End If
    
  Case "PSN" 'Pregunta S ó N
     If UCase(Mid(pValor, 1, 1)) = "S" Or UCase(Mid(pValor, 1, 1)) = "N" Then
       pValor = UCase(Mid(pValor, 1, 1))
     Else
       Validacion = False
       vMensaje = "El valor indicado no es válido > Indique [S] ó [N]...!!!"
     End If
     
  Case "DTS" 'Fecha
    
    If Not IsDate(pValor) Then
       Validacion = False
       vMensaje = "La Fecha indicada no es válida...!!!"
    Else
       pValor = Format(CDate(pValor), "yyyy/mm/dd")
    End If

End Select


If Not Validacion Then
  MsgBox vMensaje, vbExclamation, "Parámetros de Crédito"
  Exit Sub
End If


strSQL = "update ins_parametros set modifica_usuario = '" & glogon.Usuario & "', modifica_Fecha = dbo.MyGetdate()" _
       & ",valor = '" & Trim(pValor) & "' where cod_parametro = '" & pParametro & "'"
glogon.Conection.Execute strSQL

strSQL = "Parámetro de INS: " & pParametro & " -> " & pValor

Call Bitacora("Modifica", strSQL)

MsgBox "Parámetro actualizado satisfactoriamente...!", vbInformation

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Function fxGuardar() As Long
Dim vTemp As String

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.col = 3
vTemp = vGrid.Text


vGrid.col = 1
Call sbGuardaParametro(vGrid.Text, vTemp, vGrid.CellTag)

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

If vGrid.ActiveCol = vGrid.MaxCols And KeyCode = vbKeyF4 Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   If vGrid.CellTag = "CTA" Then
      gCuenta = ""
      frmCntX_ConsultaCuentas.Show vbModal
      If gCuenta <> "" Then
        vGrid.col = 3
        vGrid.Text = fxgCntCuentaFormato(True, gCuenta)
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicatorColor = vbBlue
        vGrid.CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
        vGrid.CellNote = fxgCntCuentaDesc(gCuenta)
        
        vGrid.col = 1
        Call sbGuardaParametro(vGrid.Text, gCuenta, "CTA")
      End If
      
   
   End If
End If

End Sub
