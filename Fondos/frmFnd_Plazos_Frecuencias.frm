VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmFnd_Plazos_Frecuencias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plazos de Vencimientos y Frecuencia de Pago"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8175
      _Version        =   524288
      _ExtentX        =   14420
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
      MaxCols         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmFnd_Plazos_Frecuencias.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimientos y Frencuencia pago Cupón"
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
      Height          =   720
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmFnd_Plazos_Frecuencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mSheet As Integer

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call vGrid_SheetChanged(2, 1)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim vTipo As String, vMovimiento As String, pId As Long

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then  'Insertar
  vMovimiento = "Registra"
  pId = 0
Else
  vMovimiento = "Modifica"
  pId = vGrid.Text
End If
  
    Select Case mSheet
      Case 1 'Plazos
          vTipo = "Plazo de Inversión Id:"
          strSQL = "exec spFnd_CDP_Plazos_Vencimiento_Add " & pId & ", '"
      Case 2 'Frecuencias
          vTipo = "Frecuencia Pago Cupón Id:"
          strSQL = "exec spFnd_CDP_Frecuencia_Cupon_Add " & pId & ", '"
    End Select
         
         
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & ", "
  
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", '" & glogon.Usuario & "'"
  
  Call OpenRecordSet(rs, strSQL)
  vGrid.Col = 1
  vGrid.Text = CStr(rs!Id)
  
  Call Bitacora(vMovimiento, vTipo & vGrid.Text)


fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String, vTipo As String

On Error GoTo vError

vGrid.Sheet = mSheet

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
        vGrid.Col = 1
        
        Select Case mSheet
          Case 1 'Causas
                vTipo = "Causa de Morosidad: "
                strSQL = "delete CBR_CAUSAS_MOROSIDAD where cod_causa = '" & vGrid.Text & "'"
          Case 2 'Arreglos
                vTipo = "Tipo de Arreglo: "
                strSQL = "delete CBR_TIPOS_ARREGLOS where cod_arreglo = '" & vGrid.Text & "'"
        End Select
        
        Select Case mSheet
          Case 1 'Plazos
              vTipo = "Plazo de Inversión Id:"
              strSQL = "exec spFnd_CDP_Plazos_Vencimiento_Delete " & vGrid.Text & ", '" & glogon.Usuario & "'"
          Case 2 'Frecuencias
              vTipo = "Frecuencia Pago Cupón Id:"
              strSQL = "exec spFnd_CDP_Frecuencia_Cupon_Delete " & vGrid.Text & ", '" & glogon.Usuario & "'"
        End Select
        
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", vTipo & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String

vGrid.Sheet = NewSheet
lblTitulo.Caption = vGrid.SheetName

mSheet = NewSheet
 
Select Case NewSheet
   Case 1 'Plazos
        strSQL = "select ID_PLAZO, PLAZO, PLAZO_MESES, PLAZO_DIAS,  ESTADO  From FND_CDP_PLAZOS" _
              & " order by PLAZO_MESES"
   Case 2 'Frecuencias
        strSQL = "select ID_FRECUENCIACUPON, CUPON, FRECUENCIA_MESES, FRECUENCIA_DIAS, ESTADO From FND_CDP_FRECUENCIACUPONES" _
              & " order by FRECUENCIA_MESES"
End Select

vGrid.Sheet = mSheet
vGrid.ActiveSheet = mSheet

Call sbCargaGrid(vGrid, 5, strSQL)


End Sub


