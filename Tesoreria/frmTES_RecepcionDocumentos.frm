VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmTES_RecepcionDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Documentos"
   ClientHeight    =   6384
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10584
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6384
   ScaleWidth      =   10584
   Begin VB.ComboBox cboX 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame fraPrincipal 
      Height          =   3735
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   6375
      Begin VB.Label lblNotas 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1755
         Left            =   1680
         TabIndex        =   12
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label lblUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lblDestino 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblOrigen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarX 
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4812
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   10332
      _Version        =   524288
      _ExtentX        =   18224
      _ExtentY        =   8488
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
      SpreadDesigner  =   "frmTES_RecepcionDocumentos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   672
      Left            =   8760
      TabIndex        =   17
      Top             =   5640
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   1182
      _StockProps     =   79
      Caption         =   "&Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmTES_RecepcionDocumentos.frx":06B7
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   5640
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Remesa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgInfo 
      Height          =   255
      Left            =   4560
      Picture         =   "frmTES_RecepcionDocumentos.frx":0E8F
      Stretch         =   -1  'True
      ToolTipText     =   "Ver Información de la Remesa"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgReporte 
      Height          =   255
      Left            =   4920
      Picture         =   "frmTES_RecepcionDocumentos.frx":165E
      Stretch         =   -1  'True
      ToolTipText     =   "Reporte de Recepción"
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   10680
      X2              =   0
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   10680
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmTES_RecepcionDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As Long, vScrollX As Boolean, vScrollY As Boolean



Private Sub cboX_Click()

If vCodigo = 0 Then Exit Sub

If Trim(fxCodigoCbo(cboX)) = Trim(lblDestino.Tag) Then
  vGrid.Enabled = True
Else
  vGrid.Enabled = False
End If

End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vSolicitud As Long, vNotas As String
Dim i As Integer

On Error GoTo vError

If vCodigo = 0 Or Trim(UCase(lblOrigen.Tag)) = "R" Or Not vGrid.Enabled Then
  MsgBox "No existe la Remesa / o No se encuentra pendiente / o No esta autorizado para recibirla", vbExclamation
  Exit Sub
End If
Me.MousePointer = vbHourglass

strSQL = ""

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 1
    vSolicitud = vGrid.Text
    vGrid.col = 5
    vNotas = vGrid.Text
    
    vGrid.col = 7
    strSQL = strSQL & Space(10) & "update tes_ubi_remDet set observa_rec = '" & vNotas & "',fecha_rec = dbo.MyGetdate(), usuario_rec = '" _
           & glogon.Usuario & "',estado = " & vGrid.Value & " where cod_remesa = " & vCodigo & " and Nsolicitud = " & vSolicitud
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
    
Next i

strSQL = strSQL & Space(10) & "update tes_ubi_remesa set estado = 'R' where cod_remesa = " & vCodigo

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


Call Bitacora("Aplica", "Recepcion de la Remesa Documentos : " & vCodigo)

Call sbConsulta(vCodigo)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub Form_Activate()
 vModulo = 9
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 9
 vGrid.AppearanceStyle = fxGridStyle
 
 vScrollX = False
  FlatScrollBarX.Value = 0
 vScrollX = True
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub


Private Sub FlatScrollBarX_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then txtCodigo.Text = 0

If vScrollX Then
    strSQL = "select Top 1 cod_remesa from tes_ubi_remesa"
    
    If FlatScrollBarX.Value = 1 Then
       strSQL = strSQL & " where cod_remesa > " & txtCodigo & " order by cod_remesa asc"
    Else
       strSQL = strSQL & " where cod_remesa < " & txtCodigo & " order by cod_remesa desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_remesa
      Call sbConsulta(txtCodigo)
    End If
    rs.Close
End If

vScrollX = False
FlatScrollBarX.Value = 0
vScrollX = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub


Private Sub sbLimpiaPantalla()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Byte

vCodigo = 0
txtCodigo = ""
fraPrincipal.Visible = False

vGrid.MaxRows = 0


cboX.Clear
strSQL = "select rtrim(cod_ubicacion) + ' - ' + descripcion as Itmx from tes_ubicaciones" _
        & " where usuario = '" & glogon.Usuario & "' order by cod_ubicacion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboX.AddItem rs!itmX
 rs.MoveNext
Loop

If rs.RecordCount > 0 Then
  rs.MoveFirst
  cboX.Text = rs!itmX
End If

rs.Close

End Sub

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    
    Select Case i
      Case 1
        vGrid.Text = CStr(rs!NSolicitud)
      Case 2
        vGrid.Text = CStr(rs!id_Banco)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!BancoX
        vGrid.TextTip = TextTipFixed
      Case 3
        vGrid.Text = CStr(rs!Tipo)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!TipoX
        vGrid.TextTip = TextTipFixed
      Case 4
        vGrid.Text = CStr(rs!nDocumento) & ""
        
      Case 5
        vGrid.Text = CStr(IIf(IsNull(rs!observa_rec), "", rs!observa_rec))
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = IIf(IsNull(rs!observacion), "", rs!observacion)
        vGrid.TextTip = TextTipFixed
      
      Case 6
        vGrid.Text = CStr(IIf(IsNull(rs!usuario_rec), "", rs!usuario_rec))
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Recibo Fecha :" & IIf(IsNull(rs!fecha_rec), "", rs!fecha_rec)
        vGrid.TextTip = TextTipFixed
      
      Case 7
        vGrid.Text = CStr(rs!Estado)
    
    End Select

  
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

End Sub


Private Sub sbConsulta(xCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer, vPasoX As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select R.*,rtrim(X.cod_ubicacion) + ' - ' + X.descripcion as OUbicacion" _
       & ",rtrim(Y.cod_ubicacion) + ' - ' + Y.descripcion as DUbicacion" _
       & " from tes_ubi_remesa R inner join tes_ubicaciones X on R.cod_ubicacion = X.cod_ubicacion" _
       & " inner join tes_ubicaciones Y on R.cod_ubicacion_Destino = Y.cod_ubicacion" _
       & " where R.cod_remesa = " & xCodigo

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  
  vCodigo = rs!cod_remesa
  txtCodigo = rs!cod_remesa
  
  lblOrigen.Caption = rs!oubicacion
  lblOrigen.Tag = rs!Estado
  lblDestino.Caption = rs!dUbicacion
  lblDestino.Tag = rs!cod_ubicacion_destino
  
  
  lblFecha.Caption = Format(rs!fecha, "dd/mm/yyyy")
  lblUsuario.Caption = rs!Usuario
  
  lblNotas.Caption = rs!notas & ""
  
  strSQL = "select C.nsolicitud,C.id_banco,C.tipo,C.ndocumento,D.estado,D.observacion,D.observa_rec" _
         & ",B.descripcion as BancoX,T.descripcion as TipoX,D.fecha_rec,D.usuario_rec" _
         & " from Tes_Transacciones C inner join tes_ubi_remDet D on C.nsolicitud = D.nsolicitud" _
         & " inner join Tes_Bancos B on C.id_Banco = B.id_Banco" _
         & " inner join TES_Tipos_Doc T on C.Tipo = T.tipo" _
         & " where D.cod_remesa = " & vCodigo
  Call sbCargaGridLocal(vGrid, 7, strSQL, True)
   
  Call cboX_Click
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub imgInfo_Click()

If vCodigo = 0 Then Exit Sub
If fraPrincipal.Visible Then
   fraPrincipal.Visible = False
Else
   fraPrincipal.Visible = True
End If

End Sub

Private Sub imgReporte_Click()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
     
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_RecepcionDocumentos.rpt")
    .SelectionFormula = "{TES_UBI_REMESA.COD_REMESA} = " & vCodigo
    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
On Error GoTo vError
 If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
vError:
End Sub
