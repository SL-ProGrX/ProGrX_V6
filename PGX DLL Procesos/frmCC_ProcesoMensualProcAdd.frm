VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCC_ProcesoMensualProcAdd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Planillas: Procesos Complementarios"
   ClientHeight    =   7236
   ClientLeft      =   120
   ClientTop       =   396
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7236
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   10800
      Top             =   240
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6252
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   12732
      _Version        =   524288
      _ExtentX        =   22458
      _ExtentY        =   11028
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
      MaxCols         =   8
      SpreadDesigner  =   "frmCC_ProcesoMensualProcAdd.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesos Complementarios de Planillas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   9732
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12612
   End
End
Attribute VB_Name = "frmCC_ProcesoMensualProcAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxTransaccionCod(pTransaccion As String) As String
Dim vResultado As String

Select Case Trim(pTransaccion)
 Case "Cambia Fecha de Proceso"
   vResultado = "01"
 Case "Genera deducciones"
   vResultado = "02"
 Case "Carga deducciones"
   vResultado = "03"
 Case "Desglosa deducciones"
   vResultado = "04"
 Case "Aplica Ahorros"
   vResultado = "05"
 Case "Inconsistencias de Ahorros"
   vResultado = "06"
 Case "Devoluciones de Ahorros"
   vResultado = "07"
 Case "Aplica Abonos"
   vResultado = "08"
 Case "Reporte de Inconsistencias"
   vResultado = "09"
 Case "Actualiza Intereses Moratorios"
   vResultado = "10"
 Case "Actualiza Saldo del Mes"
   vResultado = "11"
 Case Else
   vResultado = "00"
End Select
fxTransaccionCod = vResultado

End Function

Private Sub sbConsulta()
Dim strSQL As String

strSQL = "select case when transaccion = '01' then 'Cambia Fecha de Proceso'" _
       & "            when transaccion = '02' then 'Genera deducciones'" _
       & "            when transaccion = '03' then 'Carga deducciones'" _
       & "            when transaccion = '04' then 'Desglosa deducciones'" _
       & "            when transaccion = '05' then 'Aplica Ahorros'" _
       & "            when transaccion = '06' then 'Inconsistencias de Ahorros'" _
       & "            when transaccion = '07' then 'Devoluciones de Ahorros'" _
       & "            when transaccion = '08' then 'Aplica Abonos'" _
       & "            when transaccion = '09' then 'Reporte de Inconsistencias'" _
       & "            when transaccion = '10' then 'Actualiza Intereses Moratorios'" _
       & "            when transaccion = '11' then 'Actualiza Saldo del Mes' else '' end as 'Proceso'" _
       & ",PROC_NUM,EJECUCION_TIPO,EJECUCION_ORDEN,PROCEDIMIENTO,DESCRIPCION,PARAMETROS_PLANILLAS,PARAMETROS_ADD" _
       & " from prm_Procesos_add order by transaccion, EJECUCION_TIPO desc, EJECUCION_ORDEN,  PROC_NUM"

Call sbCargaGrid(vGrid, 8, strSQL)
   

End Sub

Private Sub Form_Load()
 
 vModulo = 10
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

 Call Formularios(Me)
 Call RefrescaTags(Me)

 vGrid.MaxCols = 8
 
End Sub

Private Sub Form_Resize()
On Error Resume Next


imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 550
vGrid.Height = Me.Height - (vGrid.Top + 850)

End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Call sbConsulta
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String
Dim vTransaccion As String, vNumProc As Integer, vTipo As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
        End If
  End If 'Actualiza o Inserta
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

     vGrid.Row = vGrid.ActiveRow
     vGrid.Col = 2

     If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Col = 1
        vTransaccion = fxTransaccionCod(vGrid.Text)
        vGrid.Col = 2
        vNumProc = vGrid.Text
        vGrid.Col = 3
        vTipo = Trim(vGrid.Text)


        strSQL = "delete PRM_PROCESOS_ADD" _
               & " where Transaccion = '" & vTransaccion & "' and Proc_Num = " & vNumProc & " and EJECUCION_TIPO = '" & vTipo & "'"
        Call ConectionExecute(strSQL)
        
        vGrid.Col = 2
        Call Bitacora("Elimina", "Planilla Proc.Add.: Tra: " & vTransaccion & " Tipo: " & vTipo & " Id: " & vNumProc)


        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     End If

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxNumProc(pTransaccion As String, pTipo As String) As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As Integer


strSQL = "select isnull(max(Proc_Num),0) + 1 as 'Resultado' from PRM_PROCESOS_ADD where transaccion = '" & pTransaccion _
       & "'" ' and EJECUCION_TIPO = '" & pTipo & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  vResultado = 1
Else
  vResultado = rs!Resultado
End If
rs.Close

fxNumProc = vResultado
End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTransaccion As String, vNumProc As Integer, vTipo As String
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
vTransaccion = fxTransaccionCod(vGrid.Text)
vGrid.Col = 3
vTipo = Trim(vGrid.Text)
vGrid.Col = 2


If vGrid.Text = "" Then
    
    vNumProc = fxNumProc(vTransaccion, vTipo)
    vGrid.Col = 4
    strSQL = "insert PRM_PROCESOS_ADD(TRANSACCION,PROC_NUM,EJECUCION_ORDEN,EJECUCION_TIPO,PROCEDIMIENTO,DESCRIPCION,PARAMETROS_PLANILLAS,PARAMETROS_ADD" _
           & ",REGISTRO_USUARIO,REGISTRO_FECHA) VALUES('" & vTransaccion & "'," & vNumProc & ",'" & vGrid.Text & "','" & vTipo & "','"
    vGrid.Col = 5
    strSQL = strSQL & vGrid.Text & "','"
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.Col = 7
    strSQL = strSQL & vGrid.Value & ",'"
    vGrid.Col = 8
    strSQL = strSQL & vGrid.Text & "','" & glogon.Usuario & "',dbo.MyGetdate())"
    
    Call ConectionExecute(strSQL)
  
    vGrid.Col = 2
    vGrid.Text = CStr(vNumProc)
    
    vGrid.Col = 2
    Call Bitacora("Registra", "Planilla Proc.Add.: Tra: " & vTransaccion & " Tipo: " & vTipo & " Id: " & vNumProc)
   
   Else 'Actualizar

    vGrid.Col = 2
    vNumProc = CInt(vGrid.Text)
    vGrid.Col = 4
    strSQL = "update PRM_PROCESOS_ADD set EJECUCION_ORDEN = '" & vGrid.Text & "', PROCEDIMIENTO = '"
    vGrid.Col = 5
    strSQL = strSQL & vGrid.Text & "', DESCRIPCION = '"
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Text & "', PARAMETROS_PLANILLAS = "
    vGrid.Col = 7
    strSQL = strSQL & vGrid.Value & ", PARAMETROS_ADD = '"
    vGrid.Col = 8
    strSQL = strSQL & vGrid.Text & "', ACTUALIZA_USUARIO = '" & glogon.Usuario & "', ACTUALIZA_FECHA = dbo.MyGetdate()" _
           & " where Transaccion = '" & vTransaccion & "' and Proc_Num = " & vNumProc & " and EJECUCION_TIPO = '" & vTipo & "'"
    Call ConectionExecute(strSQL)
    
    vGrid.Col = 2
    Call Bitacora("Registra", "Planilla Proc.Add.: Tra: " & vTransaccion & " Tipo: " & vTipo & " Id: " & vNumProc)
    
   
   End If


   vGrid.Col = 1
   fxGuardar = CLng(vNumProc)
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function



