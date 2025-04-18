VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmInsTiposPolizas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de P�lizas"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInsTiposPolizas.frx":0000
   ScaleHeight     =   6135
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10815
      _Version        =   524288
      _ExtentX        =   19076
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
      MaxCols         =   497
      ScrollBars      =   2
      SpreadDesigner  =   "frmInsTiposPolizas.frx":6852
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   11040
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Tipos de Seguros (P�lizas)"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmInsTiposPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 17
vGrid.AppearanceStyle = fxGridStyle

strSQL = "select Tipo_Seguro,descripcion,Tipo_Seguro_Ext,comision_porc,COMISION_PORC_PRICTA,Activo from Ins_Tipos_Seguros" _
      & " order by Tipo_Seguro"
Call sbCargaGrid(vGrid, 6, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la informaci�n de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select coalesce(count(*),0) as Existe from Ins_Tipos_Seguros " _
       & " where Tipo_Seguro = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert Ins_Tipos_Seguros(Tipo_Seguro,descripcion,Tipo_Seguro_Ext,Comision_Porc,COMISION_PORC_PRICTA,Activo,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 4
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 5
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 6
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  glogon.Conection.Execute strSQL

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo de Seguro:" & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update Ins_Tipos_Seguros set descripcion = '" & vGrid.Text & "',Tipo_Seguro_Ext = '"
 vGrid.col = 3
 strSQL = strSQL & vGrid.Text & "', Comision_Porc = "
 vGrid.col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ", COMISION_PORC_PRICTA = "
 vGrid.col = 5
 strSQL = strSQL & CCur(vGrid.Text) & ", Activo = "
 vGrid.col = 6
 strSQL = strSQL & vGrid.Value & " where Tipo_Seguro = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 glogon.Conection.Execute strSQL

 vGrid.col = 1
 Call Bitacora("Modifica", "Tipo de Seguro:" & vGrid.Text)

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
        strSQL = "delete Ins_Tipos_Seguros where Tipo_Seguro = '" & vGrid.Text & "'"
        glogon.Conection.Execute strSQL
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Seguro:" & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub







