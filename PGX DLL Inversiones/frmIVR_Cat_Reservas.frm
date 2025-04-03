VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmIVR_Cat_Reservas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Reservas"
   ClientHeight    =   8136
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9228
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8136
   ScaleWidth      =   9228
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3012
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   8772
      _Version        =   1310720
      _ExtentX        =   15473
      _ExtentY        =   5313
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnCuentas 
      Height          =   312
      Index           =   0
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "Agregar Cuenta a la Reserva"
      Top             =   4680
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   556
      _StockProps     =   79
      Transparent     =   -1  'True
      Appearance      =   1
      Picture         =   "frmIVR_Cat_Reservas.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2772
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8772
      _Version        =   524288
      _ExtentX        =   15473
      _ExtentY        =   4890
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
      MaxCols         =   482
      ScrollBars      =   2
      SpreadDesigner  =   "frmIVR_Cat_Reservas.frx":0720
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   312
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   4680
      Width           =   1812
      _Version        =   1310720
      _ExtentX        =   3196
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
      Height          =   312
      Left            =   2880
      TabIndex        =   6
      Top             =   4680
      Width           =   5172
      _Version        =   1310720
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnCuentas 
      Height          =   312
      Index           =   1
      Left            =   8520
      TabIndex        =   8
      ToolTipText     =   "Eliminar Cuenta de la Reserva"
      Top             =   4680
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   556
      _StockProps     =   79
      Transparent     =   -1  'True
      Appearance      =   1
      Picture         =   "frmIVR_Cat_Reservas.frx":0D11
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   492
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   8772
      _Version        =   1310720
      _ExtentX        =   15473
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Cuenta "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   492
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   8772
      _Version        =   1310720
      _ExtentX        =   15473
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "(Seleccione una Reserva)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   7212
      _Version        =   1310720
      _ExtentX        =   12721
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Reservas"
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
      Width           =   13332
   End
End
Attribute VB_Name = "frmIVR_Cat_Reservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnCuentas_Click(Index As Integer)
Dim strSQL As String, vVincula As Integer


If scTitulo.Tag = "" Then
    Exit Sub
End If

On Error GoTo vError

Select Case Index
    Case 0 'Add
      vVincula = 1
    Case 1 'Delete
      vVincula = 0
End Select

strSQL = "exec spIVR_RESERVAS_CUENTAS_REGISTRO '" & scTitulo.Tag _
    & "', '" & txtCuenta.Text & "', " & vVincula & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbLsw_Load

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub sbConsulta()
Dim strSQL As String


vPaso = True

strSQL = "select COD_RESERVA,descripcion,ACTIVA,0 from IVR_RESERVAS" _
      & " order by COD_RESERVA"
Call sbCargaGrid(vGrid, 4, strSQL)

vPaso = False

End Sub

Private Sub Form_Load()

vModulo = 22

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
  .Clear
  .Add , , "Cuenta", 2500
  .Add , , "Descripción", 6000
End With

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

strSQL = "select isnull(count(*),0) as Existe from IVR_RESERVAS " _
       & " where COD_RESERVA = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into IVR_RESERVAS(COD_RESERVA,DESCRIPCION, ACTIVA, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipos de Reservas:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update IVR_RESERVAS set descripcion = '" & vGrid.Text & "',Activa = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where COD_RESERVA = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipos de Reservas:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbLsw_Load()
Dim strSQL As String, rs As New Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lsw.ListItems.Clear
txtCuenta.Text = ""
txtCuentaDesc.Text = ""


If scTitulo.Tag = "" Then
    Exit Sub
End If

vPaso = True

strSQL = "select Rc.*, Cta.COD_CUENTA_MASK, Cta.DESCRIPCION " _
       & "  from IVR_RESERVAS_CUENTAS Rc" _
       & " INNER JOIN vCNTX_CUENTAS_LOCAL Cta on Rc.COD_CUENTA = Cta.COD_CUENTA" _
       & " WHERE Rc.COD_RESERVA = '" & scTitulo.Tag & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!COD_CUENTA_MASK)
     itmX.SubItems(1) = rs!Descripcion
 rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:


End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtCuenta.Text = Item.Text
txtCuentaDesc.Text = Item.SubItems(1)

End Sub


Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnCuentas(0).SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuenta.Text = gCuenta
   txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text, 0)
End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub


If Col = 4 Then
   vGrid.Row = Row
   vGrid.Col = 1
   scTitulo.Tag = vGrid.Text
   vGrid.Col = 2
   scTitulo.Caption = vGrid.Text

   Call sbLsw_Load
End If


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If vGrid.ActiveCol = vGrid.MaxCols - 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete IVR_RESERVAS where COD_RESERVA = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipos de Reservas:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub


