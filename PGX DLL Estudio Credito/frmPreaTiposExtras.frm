VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPreaTiposExtras 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Extras"
   ClientHeight    =   5604
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7932
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5604
   ScaleWidth      =   7932
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4092
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   6732
      _Version        =   524288
      _ExtentX        =   11875
      _ExtentY        =   7218
      _StockProps     =   64
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
      SpreadDesigner  =   "frmPreaTiposExtras.frx":0000
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Extras"
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
      Height          =   612
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaTiposExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3 'Modulo de Credito

Call Formularios(Me)
Call RefrescaTags(Me)

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select cod_extras,descripcion,prioridad from Crd_Prea_Tipos_extras" _
      & " order by cod_extras"
Call sbCargaGrid(vGrid, 3, strSQL)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Crd_Prea_Tipos_extras " _
       & " where cod_extras = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Prea_Tipos_extras(cod_extras,descripcion,prioridad) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "PreAnalisis Tipo de Extra Cod: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Crd_Prea_Tipos_extras set descripcion = '" & vGrid.Text & "',prioridad = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "' where cod_extras = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

  Call Bitacora("Modifica", "PreAnalisis Tipo de Extra Cod : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
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
'If KeyCode = vbKeyDelete Then
'  Call sbBorrar
'End If
End Sub
Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   strSQL = "delete Crd_Prea_Tipos_extras where cod_extras = '" & vGrid.Text & "'"
   Call ConectionExecute(strSQL)
   vGrid.Col = 1
   Call Bitacora("Elimina", "PreAnalisis Tipo de Extra Cod : " & vGrid.Text)
   
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
   If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1

End If



  
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




