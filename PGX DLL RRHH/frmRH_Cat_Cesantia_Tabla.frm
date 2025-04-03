VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Cat_Cesantia_Tabla 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla para cálculo de Cesantía"
   ClientHeight    =   6936
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8712
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6936
   ScaleWidth      =   8712
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5532
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   6492
      _Version        =   524288
      _ExtentX        =   11451
      _ExtentY        =   9758
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
      SpreadDesigner  =   "frmRH_Cat_Cesantia_Tabla.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   6732
      _Version        =   1245187
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Tabla para Cálculo de Liquidación de Cesantía"
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
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmRH_Cat_Cesantia_Tabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_ACTIVOte()
vModulo = 23
End Sub

Private Sub sbConsulta()
Dim strSQL As String

strSQL = "select ID_LINEA,TIEMPO,INICIO,CORTE,VALOR" _
      & "  from RH_CESANTIA_TABLA" _
      & " order by ID_LINEA"
Call sbCargaGrid(vGrid, 5, strSQL)

End Sub

Private Sub Form_Load()

vModulo = 23

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbConsulta

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxCodigoId() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pId As Long

strSQL = "select ISNULL(max(ID_LINEA),0) + 1 AS 'ConsecId' FROM RH_CESANTIA_TABLA"
Call OpenRecordSet(rs, strSQL)
    pId = rs!ConsecId
rs.Close

fxCodigoId = pId
End Function


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCodigoId As Long

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then  'Insertar
  pCodigoId = fxCodigoId()
  vGrid.Text = CStr(pCodigoId)
  
  strSQL = "insert into RH_CESANTIA_TABLA(ID_LINEA,TIEMPO,INICIO,CORTE,VALOR" _
            & ", REGISTRO_USUARIO, REGISTRO_FECHA) values(" _
         & vGrid.Text & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.Col = 4
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.Col = 5
  strSQL = strSQL & CCur(vGrid.Text) & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tabla Cesantía Id:" & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update RH_CESANTIA_TABLA set TIEMPO = '" & vGrid.Text & "', INICIO = "
 vGrid.Col = 3
 strSQL = strSQL & CLng(vGrid.Text) & ", CORTE = "
 vGrid.Col = 4
 strSQL = strSQL & CLng(vGrid.Text) & ", VALOR = "
 vGrid.Col = 5
 strSQL = strSQL & CCur(vGrid.Text) & " where ID_LINEA = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tabla Cesantía Id:" & vGrid.Text)

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


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

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        If IsNumeric(vGrid.Text) Then
            strSQL = "delete RH_CESANTIA_TABLA where ID_LINEA = " & vGrid.Text
            Call ConectionExecute(strSQL)
            strSQL = vGrid.Text
            vGrid.Col = 1
            Call Bitacora("Elimina", "Tabla Cesantía Id:" & vGrid.Text)
            
            Call sbConsulta
        End If
     End If
End If


End Sub
