VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCC_Cuentas_Bancarias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Cuentas Bancarias"
   ClientHeight    =   5835
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   720
      Top             =   1200
   End
   Begin XtremeSuiteControls.RadioButton rbModulo 
      Height          =   252
      Index           =   0
      Left            =   6120
      TabIndex        =   2
      Top             =   1200
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Solo este Módulo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12615
      _Version        =   524288
      _ExtentX        =   22251
      _ExtentY        =   7223
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmCC_Cuentas_Bancarias.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.RadioButton rbModulo 
      Height          =   252
      Index           =   1
      Left            =   8400
      TabIndex        =   3
      Top             =   1200
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mostrar todas las vinculadas"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Cuentas Bancarias"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   1884
      TabIndex        =   1
      Top             =   240
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmCC_Cuentas_Bancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUltimaSeleccion As String, strUltimaSelTipo As String
Dim strSQL As String, rs As New ADODB.Recordset



Private Sub Form_Activate()
vModulo = 10

End Sub

Private Sub Form_Load()

vModulo = 10
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Me.Caption = "Cuentas Bancarias [Identificación : " & GLOBALES.gTag & "]"

strUltimaSeleccion = ""
strUltimaSelTipo = ""
   

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub sbCargaCboBancos(vCol As Integer, vRow As Long, vGrid As Object)
Dim strResultado As String, rs As New ADODB.Recordset, strSQL As String

vGrid.Col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox

strSQL = "select rtrim(Tg.COD_GRUPO) + ' - ' + rtrim(Tg.DESCRIPCION) as 'ItmX'" _
       & " from Tes_Bancos B inner join tes_banco_docs D on B.id_banco = D.id_banco and D.tipo = 'TE'" _
       & " inner join TES_BANCOS_GRUPOS Tg on B.cod_Grupo = Tg.cod_grupo" _
       & " where Tg.Activo = 1" _
       & " group by Tg.COD_GRUPO, Tg.DESCRIPCION" _
       & " order by Tg.COD_GRUPO"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And strUltimaSeleccion = "" Then
 strUltimaSeleccion = rs!itmX
End If

strResultado = ""

Do While Not rs.EOF
  If Len(strResultado) = 0 Then
    strResultado = Chr$(9) & rs!itmX
  Else
    strResultado = strResultado & Chr$(9) & rs!itmX
  End If
  rs.MoveNext
Loop
rs.Close

vGrid.TypeComboBoxList = strResultado
vGrid.TypeComboBoxEditable = False
vGrid.Text = strUltimaSeleccion

End Sub

Private Sub sbCargaCboTipos(vCol As Integer, vRow As Long, vGrid As Object)
Dim strResultado As String, rs As New ADODB.Recordset, strSQL As String

vGrid.Col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox

If strUltimaSelTipo = "" Then strUltimaSelTipo = "Corriente"

strResultado = "Corriente" & Chr$(9) & "Ahorros"

vGrid.TypeComboBoxList = strResultado
vGrid.TypeComboBoxEditable = False
vGrid.Text = strUltimaSelTipo

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strResTipo As String, vNota As String, vSQL As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vSQL = "select rtrim(Tg.COD_GRUPO) + ' - ' + rtrim(Tg.DESCRIPCION) as 'ItmX'" _
       & " from Tes_Bancos B inner join tes_banco_docs D on B.id_banco = D.id_banco and D.tipo = 'TE'" _
       & " inner join TES_BANCOS_GRUPOS Tg on B.cod_Grupo = Tg.cod_grupo" _
       & " where Tg.Activo = 1" _
       & " group by Tg.COD_GRUPO, Tg.DESCRIPCION" _
       & " order by Tg.COD_GRUPO"

Call OpenRecordSet(rs, vSQL)
  
    If Not rs.EOF And strUltimaSeleccion = "" Then
     strUltimaSeleccion = rs!itmX
    End If
    strResultado = ""
    Do While Not rs.EOF
        If Len(strResultado) = 0 Then
          strResultado = Chr$(9) & rs!itmX
        Else
          strResultado = strResultado & Chr$(9) & rs!itmX
        End If
      rs.MoveNext
    Loop
rs.Close

If strUltimaSelTipo = "" Then strUltimaSelTipo = "Corriente"

strResTipo = "Corriente" & Chr$(9) & "Ahorros"

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 1
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = strResultado
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = strUltimaSeleccion
  
  vGrid.Col = 2
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = strResTipo
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = strUltimaSelTipo
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
       vGrid.Text = rs!itmX
     
     Case 2 'Tipo
       vGrid.Text = rs!TipoDesc
        
     Case 3 'Divisa
       vGrid.Text = rs!cod_Divisa
        
     
     Case 4 'Destino
       vGrid.Text = CStr(rs!Destino & "")
     
     Case 5 'Cuenta
       vGrid.Text = CStr(rs!CUENTA_INTERNA)
     
     Case 6 'Cuenta Interbancaria
       vGrid.Text = CStr(rs!CUENTA_INTERBANCA)
     
     
     Case 7 'Cuenta Default
       vGrid.Value = rs!CUENTA_DEFAULT
     
     Case 8 'Activa
       vGrid.Value = rs!Activa
     
     Case Else
    End Select
  Next i
  
    vGrid.Col = vGrid.MaxCols
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario: " & IIf(IsNull(rs!Registro_Usuario), "...!", rs!Registro_Usuario) _
                     & vbCrLf & "Fecha: " & IIf(IsNull(rs!Registro_Fecha), "...!", rs!Registro_Fecha) _
                     & vbCrLf & "Modulo: " & rs!Modulo
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 1
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = strResultado
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = strUltimaSeleccion
  
  vGrid.Col = 2
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = strResTipo
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = strUltimaSelTipo
    
Me.MousePointer = vbDefault

End Sub

Private Function fxCodigoCboGrid(strCodigo As String) As String
Dim i As Integer, strResultado As String, blnPaso As Boolean

blnPaso = True
strResultado = ""
i = 1
Do While blnPaso
   If Mid(strCodigo, i, 1) <> "-" Then
     strResultado = strResultado & Mid(strCodigo, i, 1)
   Else
     blnPaso = False
   End If
   i = i + 1
Loop

fxCodigoCboGrid = strResultado

End Function


Private Function fxValida() As Boolean
Dim vMensaje As String
Dim pCuenta As String, pInterBanca As Integer, pDivisa As String

fxValida = True

vMensaje = ""

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 5
pCuenta = Trim(vGrid.Text)
vGrid.Col = 6
pInterBanca = vGrid.Value

vGrid.Col = 3
pDivisa = Trim(vGrid.Text)

If pDivisa = "" Then
    vMensaje = vMensaje & " - La divisa de la cuenta no es válida!" & vbCrLf
End If

vGrid.Col = 1
strSQL = "select  LCTA_Interna, LCTA_InterBancaria from tes_Bancos_Grupos where cod_grupo = '" & SIFGlobal.fxCodText(vGrid.Text) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   If pInterBanca = 0 Then
        If Len(pCuenta) > rs!LCTA_Interna Or Len(pCuenta) < 5 Then
            vMensaje = vMensaje & " - La cuenta bancaria no es válida, requiere: " & rs!LCTA_Interna & " digitos, verifique... "
        End If
   End If
   
   If pInterBanca = 1 Then
        If Len(pCuenta) <> rs!LCTA_InterBancaria Then
            vMensaje = vMensaje & " - La cuenta InterBancaria no es válida, requiere: " & rs!LCTA_InterBancaria & " digitos, verifique... "
        End If
   End If
   
   
End If
rs.Close

If Len(vMensaje) > 0 Then
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

End Function

Private Function fxGuardar() As Long
Dim pDivisa As String

On Error GoTo vError

If Not fxValida Then
   fxGuardar = 0
   Exit Function
End If


vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
strUltimaSeleccion = vGrid.Text

vGrid.Col = 2
strUltimaSelTipo = vGrid.Text

vGrid.Col = 3
pDivisa = vGrid.Text

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow

strSQL = "select count(*) as 'Existe' from SYS_CUENTAS_BANCARIAS where Identificacion = '" _
       & GLOBALES.gTag & "' and CUENTA_INTERNA = '"
vGrid.Col = 5
strSQL = strSQL & vGrid.Text & "'"
         
Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then  'Insertar
  
  vGrid.Col = 1
  strSQL = "insert SYS_CUENTAS_BANCARIAS(Identificacion,cod_banco,tipo,cod_divisa, modulo" _
         & ",DESTINO,CUENTA_INTERNA,CUENTA_INTERBANCA, CUENTA_DEFAULT, ACTIVA, REGISTRO_USUARIO,REGISTRO_FECHA) values('" _
         & GLOBALES.gTag & "','" & SIFGlobal.fxCodText(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "','" & GLOBALES.gTag2 & "','"
  vGrid.Col = 4
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.Col = 5
  strSQL = strSQL & Trim(vGrid.Text) & "',"
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"

  Call ConectionExecute(strSQL)
    
  vGrid.Col = 5
  Call Bitacora("Registra", "Cuenta Bancaria : " & vGrid.Text & " Id.: " & GLOBALES.gTag)
  
  
  fxGuardar = 1
  
Else 'Actualizar

    vGrid.Col = 1
    strSQL = "update SYS_CUENTAS_BANCARIAS set cod_banco = '" & SIFGlobal.fxCodText(vGrid.Text) & "',Tipo = '"
    vGrid.Col = 2
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',cod_Divisa = '"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & "',DESTINO  = '"
    vGrid.Col = 4
    strSQL = strSQL & Trim(vGrid.Text) & "', CUENTA_INTERNA= '"
    vGrid.Col = 5
    strSQL = strSQL & Trim(vGrid.Text) & "', CUENTA_INTERBANCA = "
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Value & ", CUENTA_DEFAULT = "
    vGrid.Col = 7
    strSQL = strSQL & vGrid.Value & ", ACTIVA = "
    vGrid.Col = 8
    strSQL = strSQL & vGrid.Value & ", Modulo = '" & GLOBALES.gTag2 & "',  registro_usuario = '" & glogon.Usuario _
            & "',registro_fecha = dbo.MyGetdate()" _
            & " where CUENTA_INTERNA = '"
    vGrid.Col = 5
    strSQL = strSQL & Trim(vGrid.Text) & "' and Identificacion = '" & GLOBALES.gTag & "'"
    
    Call ConectionExecute(strSQL)
    
    fxGuardar = 1
    
    vGrid.Col = 5
    Call Bitacora("Registra", "Cuenta Bancaria: " & vGrid.Text & " Id.: " & GLOBALES.gTag)
    
End If


vGrid.Col = 7
vGrid.TextTip = TextTipFixed
vGrid.TextTipDelay = 1000
                
vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
vGrid.CellNote = "Usuario " & glogon.Usuario _
               & vbCrLf & " Fecha " & Format(Date, "dd/mm/yyyy")

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function



Private Sub rbModulo_Click(Index As Integer)
Dim strSQL As String

Select Case Index
  Case 0 'Solo los del Modulo
        strSQL = "select rtrim(C.cod_Banco) + ' - ' + rtrim(B.Descripcion) as 'ItmX'" _
               & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
               & ",C.cod_Divisa,C.CUENTA_INTERNA,C.DESTINO, C.CUENTA_INTERBANCA, C.ACTIVA, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
               & ", C.Modulo, isnull(C.CUENTA_DEFAULT,0) as 'CUENTA_DEFAULT'" _
               & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
               & " where C.Identificacion = '" & GLOBALES.gTag & "' and C.Modulo = '" & GLOBALES.gTag2 & "'"
  
  Case 1 'Todos
        strSQL = "select rtrim(C.cod_Banco) + ' - ' + rtrim(B.Descripcion) as 'ItmX'" _
               & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
               & ",C.cod_Divisa,C.CUENTA_INTERNA,C.DESTINO, C.CUENTA_INTERBANCA, C.ACTIVA, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
               & ", C.Modulo, isnull(C.CUENTA_DEFAULT,0) as 'CUENTA_DEFAULT'" _
               & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
               & " where C.Identificacion = '" & GLOBALES.gTag & "' "
End Select


Call sbCargaGridLocal(vGrid, 8, strSQL)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call rbModulo_Click(1)
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Long, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
          Call sbCargaCboBancos(1, vGrid.MaxRows, vGrid)
          Call sbCargaCboTipos(2, vGrid.MaxRows, vGrid)
        End If
  End If 'Actualiza o Inserta
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    Call sbCargaCboBancos(1, vGrid.ActiveRow, vGrid)
    Call sbCargaCboTipos(2, vGrid.ActiveRow, vGrid)
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        strSQL = "delete SYS_CUENTAS_BANCARIAS where Identificacion = '" & GLOBALES.gTag _
                & "' and cod_Banco = '" & SIFGlobal.fxCodText(vGrid.Text) & "'"
        
        vGrid.Col = 5
        strSQL = strSQL & " and cuenta_Interna = '" & vGrid.Text & "'"
        
        Call ConectionExecute(strSQL)
        
        vGrid.Col = 5
        Call Bitacora("Elimina", "Cuenta Bancaria: " & vGrid.Text & " Id.: " & GLOBALES.gTag)
        
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub vGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim pCuenta As String

On Error GoTo vError

If vGrid.ActiveCol <> 5 Then Exit Sub

vGrid.Row = vGrid.ActiveRow
vGrid.Col = vGrid.ActiveCol

pCuenta = Trim(vGrid.Text)

If Len(pCuenta) <> 17 Then Exit Sub


Me.MousePointer = vbHourglass

If Len(pCuenta) = 17 And IsNumeric(Mid(pCuenta, 1, 2)) Then
  strSQL = "select dbo.fxSys_IBAN_Convertor('CR', '" & pCuenta & "') as 'Cuenta' "
  Call OpenRecordSet(rs, strSQL)
  pCuenta = Trim(rs!Cuenta & "")
  rs.Close
  vGrid.Text = pCuenta
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
End Sub
