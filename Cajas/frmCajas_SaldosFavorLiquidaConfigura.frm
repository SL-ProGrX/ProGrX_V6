VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCajas_SaldosFavorLiquidaConfigura 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Saldos a Favor y su Liquidación"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6135
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18230
      _ExtentY        =   10821
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Tipos de Documento"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Tipos de Liquidación"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "vGridTipo"
      Item(1).Control(1)=   "Label3"
      Item(1).Control(2)=   "txtUsuario"
      Item(1).Control(3)=   "btnBuscar(0)"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5535
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   7215
         _Version        =   524288
         _ExtentX        =   12726
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
         MaxCols         =   494
         ScrollBars      =   2
         SpreadDesigner  =   "frmCajas_SaldosFavorLiquidaConfigura.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridTipo 
         Height          =   4815
         Left            =   -69880
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   10215
         _Version        =   524288
         _ExtentX        =   18018
         _ExtentY        =   8493
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
         SpreadDesigner  =   "frmCajas_SaldosFavorLiquidaConfigura.frx":05C6
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   315
         Left            =   -68920
         TabIndex        =   5
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   495
         Index           =   0
         Left            =   -66760
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCajas_SaldosFavorLiquidaConfigura.frx":0D59
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   495
         Left            =   -69760
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Usuario:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldos a Favor: Métodos de Liquidación"
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
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmCajas_SaldosFavorLiquidaConfigura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub btnBuscar_Click(Index As Integer)

strSQL = "select T.DOC_TIPO,T.DESCRIPCION, isnull(L.ENVIA_FONDO,0) , isnull(L.ENVIA_TESORERIA,0)" _
       & ", isnull(L.RET_EFECTIVO,0), isnull(L.EXCLUYE_SALDO_FAVOR,0)" _
       & " from CAJAS_SALDOS_FAVOR_TIPOS T left join CAJAS_SALDOS_FAVOR_USUARIOS_LIQUIDA L" _
       & "  on T.DOC_TIPO = L.DOC_TIPO and L.USUARIO = '" & txtUsuario.Text & "'" _
       & " Where T.ACTIVO = 1"
       
vPaso = True
    Call sbCargaGrid(vGridTipo, 6, strSQL, True)
    vGridTipo.MaxRows = vGridTipo.MaxRows - 1
vPaso = False

End Sub

Private Sub Form_Activate()
vModulo = 5
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 5
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

tcMain.Item(0).Selected = True


vGridTipo.MaxRows = 0
vGridTipo.MaxCols = 6


strSQL = "select DOC_TIPO,descripcion,Activo from CAJAS_SALDOS_FAVOR_TIPOS" _
      & " order by DOC_TIPO"
Call sbCargaGrid(vGrid, 3, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from CAJAS_SALDOS_FAVOR_TIPOS " _
       & " where DOC_TIPO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert CAJAS_SALDOS_FAVOR_TIPOS(DOC_TIPO,descripcion,Activo,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo de Saldo a Favor: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update CAJAS_SALDOS_FAVOR_TIPOS set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & " where DOC_TIPO = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Tipo de Saldo a Favor: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function





Private Sub txtUsuario_Change()
vGridTipo.MaxRows = 0
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Consulta = "select Nombre,DESCRIPCION From USUARIOS"
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Filtro = " and ESTADO = 'A' and nombre in(select USUARIO " _
                    & " from CAJAS_USUARIOS group by USUARIO )"
  frmBusquedas.Show vbModal
  txtUsuario.Text = gBusquedas.Resultado
  
  If gBusquedas.Resultado <> "" Then
    Call btnBuscar_Click(0)
  End If
  
End If
End Sub

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
        strSQL = "delete CAJAS_SALDOS_FAVOR_TIPOS where DOC_TIPO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Saldo a Favor: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGridTipo_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim vFondo As Integer, vTesoreria As Integer, vEfectivo As Integer, vTipo As String, vExcluyeSaldo As Integer

On Error GoTo vError

If vPaso Then Exit Sub

With vGridTipo
  .Row = .ActiveRow
  .col = 1
  vTipo = .Text
  .col = 3
  vFondo = .Value
  .col = 4
  vTesoreria = .Value
  .col = 5
  vEfectivo = .Value
  .col = 6
  vExcluyeSaldo = .Value
  
  strSQL = "exec spCajas_SaldoFavorTipoLiqAsigna '" & vTipo & "'," & vFondo & "," & vTesoreria & "," & vEfectivo _
         & ",'" & txtUsuario.Text & "','" & glogon.Usuario & "', " & vExcluyeSaldo
  Call ConectionExecute(strSQL)
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
