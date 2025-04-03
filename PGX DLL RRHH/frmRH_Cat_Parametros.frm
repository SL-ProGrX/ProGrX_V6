VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Cat_Parametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RRHH: Parámetros del RRHH"
   ClientHeight    =   8328
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9888
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8328
   ScaleWidth      =   9888
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9240
      Top             =   480
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6852
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9612
      _Version        =   524288
      _ExtentX        =   16955
      _ExtentY        =   12086
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmRH_Cat_Parametros.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   6732
      _Version        =   1245187
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Parámetros Generales de RRHH"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmRH_Cat_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 23
End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


'Inicializa Parametros
strSQL = "exec spRH_Parametros"
Call ConectionExecute(strSQL)

strSQL = "select * from RH_PARAMETROS" _
      & " order by cod_parametro"
Call sbCargaGridLocal(vGrid, 3, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub Form_Load()


On Error GoTo vError

vModulo = 23

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub


Function ValidaDatos(pCta As Object, pCtaDesc As Object) As Boolean

If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, pCta.Text, 0)) Then
     ValidaDatos = False
     MsgBox "No se encontró código de cuenta especificado, Corrijalo...", vbCritical
Else
     ValidaDatos = True
     pCtaDesc.Caption = fxgCntCuentaDesc(fxgCntCuentaFormato(False, pCta.Text, 0))
End If
End Function

Function ValidaDatosGrabar(pCta As Object) As Boolean

ValidaDatosGrabar = fxgCntCuentaValida(fxgCntCuentaFormato(False, pCta.Text, 0))

End Function



Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub


Public Sub sbCargaGridLocal(pGrid As Object, MaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

With pGrid
    .MaxRows = 0
    .MaxCols = MaxCol
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
     
      For i = 1 To 3
        .Col = i
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

strSQL = "Parámetro de RRHH: " & pParametro & " -> " & pValor & "  (Ant: " & fxValor_Ant(pParametro) & ")"
Call Bitacora("Modifica", strSQL)

strSQL = "update RH_parametros set modifica_usuario = '" & glogon.Usuario & "', modifica_Fecha = dbo.MyGetdate()" _
       & ",valor = '" & Trim(pValor) & "' where cod_parametro = '" & pParametro & "'"
Call ConectionExecute(strSQL)


MsgBox "Parámetro actualizado satisfactoriamente...!", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValor_Ant(pParametro) As String
Dim pValor As String

With glogon
    .strSQL = "select Valor from RH_Parametros where Cod_Parametro = '" & pParametro & "'"
    Call OpenRecordSet(.Recordset, .strSQL)
    
    pValor = Trim(.Recordset!Valor)
End With

fxValor_Ant = pValor


End Function



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


