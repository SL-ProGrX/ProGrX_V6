VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCO_AntiguedadTipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Antiguedad"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11685
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11292
      _Version        =   524288
      _ExtentX        =   19918
      _ExtentY        =   5948
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
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_AntiguedadTipos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin FPSpreadADO.fpSpread fpGarantias 
      Height          =   3372
      Left            =   1920
      TabIndex        =   2
      Top             =   5280
      Width           =   7452
      _Version        =   524288
      _ExtentX        =   13144
      _ExtentY        =   5948
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
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_AntiguedadTipos.frx":07C6
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scDetalle 
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   11292
      _Version        =   1310723
      _ExtentX        =   19918
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos (Franjas) de Antiguedad de Saldos"
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
      Height          =   480
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_AntiguedadTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 4

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vPaso = True

strSQL = "select COD_ANTIGUEDAD,descripcion,DIAS_DESDE,DIAS_HASTA" _
      & ", ESTIMACION_NOCUBIERTA, ESTIMACION_CUBIERTA, 0 as 'Button'" _
      & " from CBR_ANTIGUEDAD_TIPOS" _
      & " order by COD_ANTIGUEDAD"
Call sbCargaGrid(vGrid, 7, strSQL)

vPaso = False

Call sbDetalle_Limpia


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbDetalle_Limpia()

scDetalle.Caption = "..."
scDetalle.Tag = ""

fpGarantias.MaxCols = 4
fpGarantias.MaxRows = 0


End Sub

Private Sub sbDetalle_Carga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If scDetalle.Tag = "" Then Exit Sub

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spCbr_Garantia_Mitigador_Consulta '" & scDetalle.Tag & "','AS','" & glogon.Usuario & "'"
Call sbCargaGrid(fpGarantias, 3, strSQL, True)

fpGarantias.MaxRows = fpGarantias.MaxRows - 1


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from CBR_ANTIGUEDAD_TIPOS " _
       & " where COD_ANTIGUEDAD = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert CBR_ANTIGUEDAD_TIPOS(COD_ANTIGUEDAD,descripcion,DIAS_DESDE,DIAS_HASTA, ESTIMACION_NOCUBIERTA, ESTIMACION_CUBIERTA, Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.col = 4
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.col = 5
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 6
  strSQL = strSQL & CCur(vGrid.Text) & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo de Antiguedad: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update CBR_ANTIGUEDAD_TIPOS set descripcion = '" & vGrid.Text & "',DIAS_DESDE = "
 vGrid.col = 3
 strSQL = strSQL & CLng(vGrid.Text) & ", DIAS_HASTA = "
 vGrid.col = 4
 strSQL = strSQL & CLng(vGrid.Text) & ", ESTIMACION_NOCUBIERTA = "
 vGrid.col = 5
 strSQL = strSQL & CCur(vGrid.Text) & ", ESTIMACION_CUBIERTA = "
 vGrid.col = 6
 strSQL = strSQL & CCur(vGrid.Text) & " where COD_ANTIGUEDAD = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Tipo de Antiguedad: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function





Private Sub fpGarantias_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

On Error GoTo vError

If fpGarantias.ActiveCol = fpGarantias.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
   fpGarantias.Row = fpGarantias.ActiveRow
   fpGarantias.col = 1
   
   strSQL = "exec  spCbr_Garantia_Mitigador_Registra '" & scDetalle.Tag & "','AS','" & fpGarantias.Text & "'"
   fpGarantias.col = 3
   strSQL = strSQL & "," & CCur(fpGarantias.Text) & ",'" & glogon.Usuario & "'"
   
   Call ConectionExecute(strSQL)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If col <> 7 Then Exit Sub


vGrid.Row = Row
vGrid.col = 1

Call sbDetalle_Limpia

scDetalle.Tag = vGrid.Text
vGrid.col = 2
scDetalle.Caption = vGrid.Text

Call sbDetalle_Carga

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete CBR_ANTIGUEDAD_TIPOS where COD_ANTIGUEDAD = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Antiguedad: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

