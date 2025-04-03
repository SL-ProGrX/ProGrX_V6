VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Cat_VacacionesRegimen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Regimen de Vacaciones"
   ClientHeight    =   6852
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6852
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5532
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   9972
      _Version        =   524288
      _ExtentX        =   17590
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
      MaxCols         =   485
      ScrollBars      =   2
      SpreadDesigner  =   "frmRH_Cat_VacacionesRegimen.frx":0000
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
      _Version        =   1245187
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Régimen de Vacaciones"
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
Attribute VB_Name = "frmRH_Cat_VacacionesRegimen"
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

strSQL = "select COD_VACA_REGIMEN,descripcion, DIAS_ANIO, DEFAULT_ITM,ACTIVO" _
      & "  from RH_VACACIONES_REGIMEN" _
      & " order by COD_VACA_REGIMEN"
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


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from RH_VACACIONES_REGIMEN " _
       & " where COD_VACA_REGIMEN = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into RH_VACACIONES_REGIMEN(COD_VACA_REGIMEN,descripcion, DIAS_ANIO, DEFAULT_ITM,ACTIVO" _
            & ", REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Régimen de Vacaciones: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update RH_VACACIONES_REGIMEN set descripcion = '" & vGrid.Text & "',DIAS_ANIO = "
 vGrid.Col = 3
 strSQL = strSQL & CLng(vGrid.Text) & ", DEFAULT_ITM = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ", ACTIVO = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & " where COD_VACA_REGIMEN = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Régimen de Vacaciones: " & vGrid.Text)

End If
rs.Close

'Revisa el Default
vGrid.Col = 3
If vGrid.Value = vbChecked Then
    vGrid.Col = 1
    strSQL = "Update RH_VACACIONES_REGIMEN set DEFAULT_ITM = 0" _
           & " wHERE  COD_VACA_REGIMEN <> '" & vGrid.Text & "'"
    Call ConectionExecute(strSQL)

    vGrid.Col = 1
    Call Bitacora("Registra", "Régimen de Vacaciones: " & vGrid.Text & " como DEFAULT")
           
    Call sbConsulta
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
        strSQL = "delete RH_VACACIONES_REGIMEN where COD_VACA_REGIMEN = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Régimen de Vacaciones: " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub




