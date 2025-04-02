VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Parametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parámetros Generales del Módulo de Fosol"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   8535
      _Version        =   524288
      _ExtentX        =   15055
      _ExtentY        =   9975
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
      MaxCols         =   3
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_Parametros.frx":0000
      VScrollSpecialType=   2
      Appearance      =   1
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Fondos Solidarios"
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
      Height          =   492
      Index           =   2
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmFSL_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
 vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

vModulo = 7
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

 
'Inicializa Parametros
strSQL = "exec spFSL_Parametros"
Call ConectionExecute(strSQL)


strSQL = "select * from FSL_PARAMETROS" _
      & " order by cod_parametro"
Call sbCargaGridLocal(vGrid, 3, strSQL)


Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Public Sub sbCargaGridLocal(pGrid As Object, MaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

With vGrid
    .MaxRows = 0
    .MaxCols = MaxCol
    rs.Open strSQL, glogon.Conection, adOpenForwardOnly
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      For i = 1 To 3
        .Col = i
        Select Case i
          Case 1 'Codigo
            .CellTag = rs!Tipo & ""
            .Text = rs!Cod_Parametro
            .CellNote = "Modificado Por: " & rs!Registro_Usuario & vbCrLf & "Fecha: " & rs!registro_Fecha
          
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 2 'Descripcion
            .Text = rs!Detalle
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
  MsgBox vMensaje, vbExclamation, "Parámetros de FOSOL"
  Exit Sub
End If


strSQL = "update FSL_PARAMETROS set REGISTRO_USUARIO = '" & glogon.Usuario & "', REGISTRO_FECHA = getdate()" _
       & ",valor = '" & Trim(pValor) & "' where cod_parametro = '" & pParametro & "'"
Call ConectionExecute(strSQL)

strSQL = "Parámetro de FOSOL: " & pParametro & " -> " & pValor

Call Bitacora("Modifica", strSQL)

MsgBox "Parámetro actualizado satisfactoriamente...!", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxGuardar() As Long
Dim vTemp As String

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 3
vTemp = vGrid.Text


vGrid.Col = 1
Call sbGuardaParametro(vGrid.Text, vTemp, vGrid.CellTag)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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
   vGrid.Col = 1
   If vGrid.CellTag = "CTA" Then
      gCuenta = ""
      frmCntX_ConsultaCuentas.Show vbModal
      If gCuenta <> "" Then
        vGrid.Col = 3
        vGrid.Text = fxgCntCuentaFormato(True, gCuenta)
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicatorColor = vbBlue
        vGrid.CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
        vGrid.CellNote = fxgCntCuentaDesc(gCuenta)
        
        vGrid.Col = 1
        Call sbGuardaParametro(vGrid.Text, gCuenta, "CTA")
      End If
      
   
   End If
End If

End Sub

