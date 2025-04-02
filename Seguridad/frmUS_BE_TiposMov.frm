VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmUS_BE_TiposMov 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitácoras Especiales: Tipos de Movimientos"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   9255
      _Version        =   524288
      _ExtentX        =   16325
      _ExtentY        =   9128
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
      MaxCols         =   489
      ScrollBars      =   2
      SpreadDesigner  =   "frmUS_BE_TiposMov.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   345
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   5775
      _Version        =   1310723
      _ExtentX        =   10186
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modulo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   732
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Top             =   0
      Width           =   10092
   End
End
Attribute VB_Name = "frmUS_BE_TiposMov"
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
Call cbo_Click
End Sub


Private Sub cbo_Click()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select MOVIMIENTO, DESCRIPCION, REGISTRO_FECHA, REGISTRO_USUARIO" _
       & " from US_MOVIMIENTOS_BE" _
       & " where modulo = " & cbo.ItemData(cbo.ListIndex) _
       & " order by MOVIMIENTO"
Call sbCargaGrid(vGrid, 4, strSQL)
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 13

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vPaso = True
    strSQL = "select Nombre as 'ItmX', Modulo as 'IdX' from us_modulos order by modulo"
    Call sbCbo_Llena_New(cbo, strSQL, False, True)
vPaso = False


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

strSQL = "select isnull(count(*),0) as Existe from US_MOVIMIENTOS_BE " _
       & " where MOVIMIENTO = '" & vGrid.Text & "' and modulo = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into US_MOVIMIENTOS_BE(MODULO,MOVIMIENTO,descripcion,registro_fecha,registro_usuario)" _
         & " values(" & cbo.ItemData(cbo.ListIndex) & ",'" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',Getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 3
  vGrid.Text = fxFechaServidor
  vGrid.Col = 4
  vGrid.Text = glogon.Usuario
  
  Call Bitacora("Registra", "Bitácora Especial - Tipo Movimiento: " & vGrid.Text & "..Modulo: " & cbo.ItemData(cbo.ListIndex))

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update US_MOVIMIENTOS_BE set Descripcion = '" & vGrid.Text & "' where MOVIMIENTO = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "' and modulo = " & cbo.ItemData(cbo.ListIndex)
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Bitácora Especial - Tipo Movimiento: " & vGrid.Text & "..Modulo: " & cbo.ItemData(cbo.ListIndex))

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

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 2) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete US_MOVIMIENTOS_BE where MOVIMIENTO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Bitácora Especial - Tipo Movimiento: " & vGrid.Text & "..Modulo: " & cbo.ItemData(cbo.ListIndex))

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


