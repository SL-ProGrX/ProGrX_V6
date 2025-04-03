VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPGX_Vendedores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Vendedores"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14025
   Icon            =   "frmPGX_Vendedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   14025
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   240
      Top             =   0
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   13815
      _Version        =   524288
      _ExtentX        =   24368
      _ExtentY        =   12091
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmPGX_Vendedores.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedores"
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
      Height          =   480
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "frmPGX_Vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_activate()
vModulo = 13
End Sub

Private Sub sbInicial()
Dim strSQL As String
      
strSQL = "select cod_vendedor,identificacion,nombre,activo,comision_tipo,Comision_Cliente" _
       & ",Cuenta_Cliente,Registro_Fecha,Registro_Usuario" _
       & " from PGX_Vendedores" _
       & " order by cod_vendedor"
vPaso = True
    Call sbCargaGrid(vGrid, 9, strSQL)
vPaso = False
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 13
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


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

strSQL = "select isnull(count(*),0) as Existe from PGX_Vendedores " _
       & " where cod_Vendedor = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into PGX_Vendedores(cod_Vendedor,identificacion,nombre,activo" _
         & ",comision_Tipo,Comision_Cliente,Cuenta_Cliente,registro_fecha,registro_usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 6
  strSQL = strSQL & CCur(vGrid.Text) & ",'"
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Text & "',Getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 8
  vGrid.Text = fxFechaServidor
  vGrid.Col = 9
  vGrid.Text = glogon.Usuario
  
'  Call Bitacora("Registra", "Vendedor: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update PGX_Vendedores set Identificacion = '" & vGrid.Text & "',Nombre= '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "',Activo = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ",Comision_Tipo = '"
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Text & "',Comision_Cliente = "
 vGrid.Col = 6
 strSQL = strSQL & CCur(vGrid.Text) & ",Cuenta_Cliente = '"
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Text & "' where cod_Vendedor = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
' Call Bitacora("Modifica", "Vendedor: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 7) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete PGX_Vendedores where cod_Vendedor = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Vendedor: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

