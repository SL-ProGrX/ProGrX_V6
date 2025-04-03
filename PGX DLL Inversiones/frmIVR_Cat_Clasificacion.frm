VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmIVR_Cat_Clasificacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Clasificacion de Inversión"
   ClientHeight    =   4530
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10710
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3012
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   10092
      _Version        =   524288
      _ExtentX        =   17801
      _ExtentY        =   5313
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
      MaxCols         =   485
      ScrollBars      =   2
      SpreadDesigner  =   "frmIVR_Cat_Clasificacion.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   6732
      _Version        =   1310722
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Categoría de la Inversión"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
Attribute VB_Name = "frmIVR_Cat_Clasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub sbConsulta()
Dim strSQL As String

vPaso = True

strSQL = "select COD_CATEGORIA,descripcion, case when Afectacion = 'PAT' then 'Patrimonio' else 'Resultados' end " _
      & ", Valoriza, Activo, 0 from IVR_CATEGORIA_TIPOS" _
      & " order by COD_CATEGORIA"
Call sbCargaGrid(vGrid, 6, strSQL)

vPaso = False

End Sub

Private Sub Form_Load()

vModulo = 22

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbConsulta

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from IVR_CATEGORIA_TIPOS " _
       & " where COD_CATEGORIA = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into IVR_CATEGORIA_TIPOS(COD_CATEGORIA,DESCRIPCION, AFECTACION, VALORIZA, ACTIVO, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & Trim(UCase(vGrid.Text)) & "','"
  vGrid.Col = 2
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.Col = 3
  strSQL = strSQL & UCase(Mid(vGrid.Text, 1, 3)) & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Categoría de la Inversión:  " & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update IVR_CATEGORIA_TIPOS set descripcion = '" & vGrid.Text & "', AFECTACION = '"
  vGrid.Col = 3
  strSQL = strSQL & UCase(Mid(vGrid.Text, 1, 3)) & "', VALORIZA = "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ", ACTIVO = "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & " where COD_CATEGORIA = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Categoría de la Inversión:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If Col = 6 Then
   vGrid.Row = Row
   vGrid.Col = 1
   
   gIVR_Cuentas.Tipo = "C"
   gIVR_Cuentas.Codigo_1 = vGrid.Text
   gIVR_Cuentas.Codigo_2 = ""
   
   vGrid.Col = 2
   gIVR_Cuentas.Descripcion = vGrid.Text
   
   frmIVR_Cat_Cuentas_Contables.Show vbModal
   

End If
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = vGrid.MaxCols - 1) _
    And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  
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
        strSQL = "delete IVR_CATEGORIA_TIPOS where COD_CATEGORIA = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Categoría de la Inversión:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub



