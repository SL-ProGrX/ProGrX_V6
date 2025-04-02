VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmIVR_Cat_Periocidad 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI - Tabla de Periodicidad de Inversiones"
   ClientHeight    =   6312
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9012
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6312
   ScaleWidth      =   9012
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8532
      _Version        =   524288
      _ExtentX        =   15050
      _ExtentY        =   8700
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmIVR_Cat_Periocidad.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   384
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   6732
      _Version        =   1310720
      _ExtentX        =   11874
      _ExtentY        =   677
      _StockProps     =   79
      Caption         =   "Periodo de Pago para Cupones"
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
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmIVR_Cat_Periocidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub sbConsulta()
Dim strSQL As String

strSQL = "select COD_PERIODICIDAD,descripcion,DIAS,ACTIVA from IVR_PERIODICIDAD" _
      & " order by COD_PERIODICIDAD"
Call sbCargaGrid(vGrid, 4, strSQL)

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

strSQL = "select isnull(count(*),0) as Existe from IVR_PERIODICIDAD " _
       & " where COD_PERIODICIDAD = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into IVR_PERIODICIDAD(COD_PERIODICIDAD, DESCRIPCION, DIAS, ACTIVA, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Periocidad del Cupón:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update IVR_PERIODICIDAD set descripcion = '" & vGrid.Text & "', Dias = "
 vGrid.Col = 3
 strSQL = strSQL & CLng(vGrid.Text) & ", ACTIVA = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & " where COD_PERIODICIDAD = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Periocidad del Cupón:  " & vGrid.Text)

End If
rs.Close

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
        strSQL = "delete IVR_PERIODICIDAD where COD_PERIODICIDAD = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Periocidad del Cupón:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub



