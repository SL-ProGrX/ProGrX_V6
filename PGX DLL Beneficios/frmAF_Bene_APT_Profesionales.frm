VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_Bene_APT_Profesionales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Profesionales para Situaciones Apremiantes"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10335
      _Version        =   524288
      _ExtentX        =   18230
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
      MaxCols         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Bene_APT_Profesionales.frx":0000
      VScrollSpecialType=   2
      Appearance      =   1
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Profesionales para Atención de Situaciones Apremiantes"
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
Attribute VB_Name = "frmAF_Bene_APT_Profesionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()

vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select ID_PROFESIONAL,IDENTIFICACION, NOMBRE, USUARIO,  ACTIVO from AFI_BENE_APT_PROFESIONALES" _
      & " order by ID_PROFESIONAL"
Call sbCargaGrid(vGrid, 5, strSQL)


End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then 'Insertar
  
  vGrid.Col = 2
  strSQL = "insert into AFI_BENE_APT_PROFESIONALES( IDENTIFICACION, NOMBRE, USUARIO,  ACTIVO, Registro_Fecha, Registro_Usuario) values('" _
         & vGrid.Text & "', '"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "', '"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)
  
  strSQL = "select isnull(max(ID_PROFESIONAL),0) as IdSeQ from AFI_BENE_APT_PROFESIONALES"
  Call OpenRecordSet(rs, strSQL)
  
  vGrid.Col = 1
  vGrid.Text = CStr(rs!IdSeQ)
        
    rs.Close
  

  Call Bitacora("Registra", "Profesional para Apremiantes Id: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update AFI_BENE_APT_PROFESIONALES set IDENTIFICACION = '" & vGrid.Text & "', NOMBRE = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "', USUARIO = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & "', ACTIVO = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where ID_PROFESIONAL = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Profesional para Apremiantes Id: " & vGrid.Text)

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
   vGrid.Col = 1
     i = MsgBox("Está Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
        strSQL = "delete AFI_BENE_APT_PROFESIONALES where ID_PROFESIONAL = " & vGrid.Text
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Profesional para Apremiantes Id: " & vGrid.Text)
        
        vGrid.Col = 1
        strSQL = "select ID_PROFESIONAL,IDENTIFICACION, NOMBRE, USUARIO,  ACTIVO from AFI_BENE_APT_PROFESIONALES" _
              & " order by ID_PROFESIONAL"
        Call sbCargaGrid(vGrid, 5, strSQL)
     End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub




