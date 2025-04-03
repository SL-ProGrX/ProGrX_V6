VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCO_AdvertenciasX 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobros: Tipos de Advertencias"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCO_AdvertenciasTiposAnt.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   7215
      _Version        =   524288
      _ExtentX        =   12726
      _ExtentY        =   8705
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_AdvertenciasTiposAnt.frx":6852
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   10080
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Tipos de Advertencias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmCO_AdvertenciasX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 4
vGrid.AppearanceStyle = fxGridStyle

strSQL = "select COD_ADVERTENCIA,descripcion,Activa from CBR_ADVERTENCIAS_TIPO" _
      & " order by COD_ADVERTENCIA"
Call sbCargaGrid(vGrid, 3, strSQL)

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
vGrid.col = 1

strSQL = "select coalesce(count(*),0) as Existe from CBR_ADVERTENCIAS_TIPO " _
       & " where COD_ADVERTENCIA = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert CBR_ADVERTENCIAS_TIPO(COD_ADVERTENCIA,descripcion,Activa,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',getdate())"
  
  glogon.Conection.Execute strSQL

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo de Advertencia: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update CBR_ADVERTENCIAS_TIPO set descripcion = '" & vGrid.Text & "',Activa = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & " where COD_ADVERTENCIA = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 glogon.Conection.Execute strSQL

 vGrid.col = 1
 Call Bitacora("Modifica", "Tipo de Advertencia: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function



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
        strSQL = "delete CBR_ADVERTENCIAS_TIPO where COD_ADVERTENCIA = '" & vGrid.Text & "'"
        glogon.Conection.Execute strSQL
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Advertencia: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub







