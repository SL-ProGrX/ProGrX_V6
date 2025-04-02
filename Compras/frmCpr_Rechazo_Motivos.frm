VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCpr_Rechazo_Motivos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compras: Motivos para Rechazos"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9495
      _Version        =   524288
      _ExtentX        =   16748
      _ExtentY        =   9763
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
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frmCpr_Rechazo_Motivos.frx":0000
      VScrollSpecialType=   2
      Appearance      =   1
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivos para Rechazos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCpr_Rechazo_Motivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub Form_Activate()
vModulo = 35
End Sub

Private Sub Form_Load()

vModulo = 35

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select COD_RECHAZO, descripcion, Activo from CPR_RECHAZO_TIPOS" _
      & " order by COD_RECHAZO"
Call sbCargaGrid(vGrid, 3, strSQL)

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
If Trim(vGrid.Text) = "" Then
  MsgBox "Indique un C�digo V�lido!", vbExclamation
  Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from CPR_RECHAZO_TIPOS " _
       & " where COD_RECHAZO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into CPR_RECHAZO_TIPOS(COD_RECHAZO, descripcion, Activo, Registro_Fecha, Registro_Usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Cat�logo de Tipos de Rechazos de Compras Id: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update CPR_RECHAZO_TIPOS set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where COD_RECHAZO = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Cat�logo de Tipos de Rechazos de Compras Id: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Elimina
If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
     i = MsgBox("Est� Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
        strSQL = "delete CPR_RECHAZO_TIPOS where COD_RECHAZO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Cat�logo de Tipos de Rechazos de Compras Id: " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "select COD_RECHAZO,descripcion,Activo from CPR_RECHAZO_TIPOS" _
              & " order by COD_RECHAZO"
        Call sbCargaGrid(vGrid, 3, strSQL)
     End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub


