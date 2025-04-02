VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTES_TrasladosDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslados de Documentación"
   ClientHeight    =   5472
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9528
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5472
   ScaleWidth      =   9528
   Begin TabDlg.SSTab ssTab 
      Height          =   4332
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   7641
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Principal"
      TabPicture(0)   =   "frmTES_TrasladosDocumentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtUsuario"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFecha"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboDestino"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboOrigen"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtNotasX"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtEstado"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Documentos"
      TabPicture(1)   =   "frmTES_TrasladosDocumentos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vGrid"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
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
         Height          =   324
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtNotasX 
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
         Height          =   2595
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1200
         Width           =   6975
      End
      Begin VB.ComboBox cboOrigen 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   4215
      End
      Begin VB.ComboBox cboDestino 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
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
         Height          =   324
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   2  'Center
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
         Height          =   324
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3612
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   8892
         _Version        =   524288
         _ExtentX        =   15685
         _ExtentY        =   6371
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
         SpreadDesigner  =   "frmTES_TrasladosDocumentos.frx":0038
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   1
         Left            =   6000
         TabIndex        =   16
         Top             =   3840
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   6000
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   6000
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   264
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9528
      _ExtentX        =   16806
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarX 
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin VB.Label Label1 
      Caption         =   "Remesa.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1332
   End
End
Attribute VB_Name = "frmTES_TrasladosDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vPaso As Boolean
Dim vScrollX As Boolean, vScrollY As Boolean


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
 
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
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


Private Sub sbCargaCboOD(cboX As ComboBox, Optional vTipo As String = "O")
Dim strSQL As String, rs As New ADODB.Recordset

cboX.Clear

If vTipo = "O" Then
  strSQL = "select rtrim(cod_ubicacion) + ' - ' + descripcion as Itmx from tes_ubicaciones" _
         & " where usuario = '" & glogon.Usuario & "'"
Else
  strSQL = "select rtrim(cod_ubicacion) + ' - ' + descripcion as Itmx from tes_ubicaciones" _
         & " where usuario <> '" & glogon.Usuario & "'"
End If
strSQL = strSQL & " order by cod_ubicacion"

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

Private Sub sbLimpiaPantalla()
Dim i As Byte

vCodigo = 0

ssTab.Tab = 0
ssTab.TabEnabled(1) = False

txtEstado = "Pendiente"
txtEstado.Tag = "P"
txtFecha = fxFechaServidor
txtUsuario = glogon.Usuario
txtNotasX = ""
txtCodigo = ""

Call sbCargaCboOD(cboOrigen, "O")
Call sbCargaCboOD(cboDestino, "D")

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
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Fecha :" & rs!fecha_rec & " / User : " & rs!usuario_rec
        vGrid.TextTip = TextTipFixed
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
        Select Case rs!Estado
          Case 0
              vGrid.Text = "Pendiente"
          Case 1
              vGrid.Text = "Recibido"
          Case 2
              vGrid.Text = "Rechazado"
        End Select
        vGrid.CellTag = rs!Estado
        
      Case 6
        vGrid.Text = CStr(rs!observacion)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!observa_rec & ""
        vGrid.TextTip = TextTipFixed
    
    End Select

  
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub


Private Sub SSTab_Click(PreviousTab As Integer)
Dim strSQL As String

Select Case ssTab.Tab
 Case 0 'Nada
 Case 1 'Grid
  strSQL = "select C.nsolicitud,C.id_banco,C.tipo,C.ndocumento,D.estado,D.observacion,D.observa_rec" _
         & ",B.descripcion as BancoX,T.descripcion as TipoX,D.fecha_rec,D.usuario_rec" _
         & " from Tes_Transacciones C inner join tes_ubi_remDet D on C.nsolicitud = D.nsolicitud" _
         & " inner join Tes_Bancos B on C.id_Banco = B.id_Banco" _
         & " inner join TES_Tipos_Doc T on C.Tipo = T.tipo" _
         & " where D.cod_remesa = " & vCodigo
  Call sbCargaGridLocal(vGrid, 6, strSQL, True)
End Select
End Sub

Private Sub sbReporteX()

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
     
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_TraspasoDocumentos.rpt")
    .SelectionFormula = "{TES_UBI_REMESA.COD_REMESA} = " & vCodigo
    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      cboOrigen.SetFocus
      Call sbToolBar(tlb, "edicion")
      
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.Enabled = False
      cboOrigen.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
      Call sbBorrar
    
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = False
        txtCodigo.Enabled = True
        txtCodigo.SetFocus
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
    
    Case "REPORTES"
      Call sbReporteX
      
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

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
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_remesa
  txtCodigo = rs!cod_remesa
  txtCodigo.Enabled = True
  
  cboOrigen.AddItem rs!oubicacion
  cboOrigen.Text = rs!oubicacion
  
  cboDestino.AddItem rs!dUbicacion
  cboDestino.Text = rs!dUbicacion
  
  txtFecha = Format(rs!fecha, "dd/mm/yyyy")
  txtUsuario = rs!Usuario
  
  Select Case rs!Estado
    Case "P" 'Estado Inicial * Modificable
       txtEstado = "Pendiente"
    Case "X" 'Transito
       txtEstado = "Transito"
    Case "R" 'Recibo en Totalidad
       txtEstado = "Recibido"
  End Select
  txtEstado.Tag = rs!Estado
  
  txtNotasX = rs!notas & ""
  
  ssTab.Tab = 0
  ssTab.TabEnabled(1) = True
    
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

Private Function fxValida() As Boolean
Dim vMensaje As String, i As Integer

On Error GoTo vError

vMensaje = ""

fxValida = True


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vEdita Then
     
  If txtEstado.Tag = "P" Then
    strSQL = "update tes_ubi_remesa set usuario = '" & txtUsuario & "',notas = '" & txtNotasX _
           & "',cod_ubicacion = '" & fxCodigoCbo(cboOrigen) & "',cod_ubicacion_destino = '" _
           & fxCodigoCbo(cboDestino) & "' where cod_remesa = " & vCodigo
  Else
    strSQL = "update tes_ubi_remesa set usuario = '" & txtUsuario & "',notas = '" & txtNotasX _
           & "' where cod_remesa = " & vCodigo
  End If
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Remesa Traspaso : " & vCodigo)

Else
   strSQL = "select isnull(max(cod_remesa),0) as IDx from TES_UBI_REMESA"
   Call OpenRecordSet(rs, strSQL)
     vCodigo = rs!idX + 1
   rs.Close

   strSQL = "insert tes_ubi_remesa(cod_remesa,cod_ubicacion,cod_ubicacion_destino,fecha,usuario,estado,notas)" _
          & " values(" & vCodigo & ",'" & fxCodigoCbo(cboOrigen) & "','" & fxCodigoCbo(cboDestino) _
          & "',dbo.MyGetdate(),'" & txtUsuario & "','P','" & txtNotasX & "')"
   Call ConectionExecute(strSQL)

   txtCodigo = vCodigo

   Call Bitacora("Registra", "Remesa Traspaso : " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

ssTab.TabEnabled(1) = True

txtCodigo.Enabled = True
txtCodigo.SetFocus

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes And txtEstado.Tag = "P" Then

  strSQL = "delete tes_ubi_remdet where cod_remesa = " & vCodigo
  Call ConectionExecute(strSQL)

  strSQL = "delete tes_ubi_remesa where cod_remesa = " & vCodigo
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Remesa Traspaso : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboOrigen.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
On Error GoTo vError
 If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
vError:
End Sub


Private Sub sbGridCodigo(vRow As Integer, vNSolicitud As Long)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select C.nsolicitud,C.id_banco,C.tipo,C.ndocumento" _
       & ",B.descripcion as BancoX,T.descripcion as TipoX" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_Banco = B.id_Banco" _
       & " inner join TES_Tipos_Doc T on C.Tipo = T.tipo" _
       & " where C.nsolicitud = " & vNSolicitud & " and C.estado <> 'P'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   vGrid.Row = vRow
   vGrid.col = 2
   vGrid.Text = CStr(rs!id_Banco)
   vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
   vGrid.CellNote = rs!BancoX
   vGrid.TextTip = TextTipFixed
   
   vGrid.col = 3
   vGrid.Text = CStr(rs!Tipo)
   vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
   vGrid.CellNote = rs!TipoX
   vGrid.TextTip = TextTipFixed
   
   vGrid.col = 4
   vGrid.Text = CStr(rs!nDocumento) & ""
   
   vGrid.col = 5
   vGrid.Text = "Pendiente"
   vGrid.CellTag = "0"
   
   
   
Else
  MsgBox "Número de Solicitud no se encontró...", vbInformation
End If
rs.Close

End Sub

Private Sub sbBuscaDocumento(vActiveRow As Integer)
Dim vBanco As Integer, vTipo As String, vDocumento As String
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vGrid.Row = vActiveRow
vGrid.col = 4

vDocumento = vGrid.Text

vGrid.col = 2
vBanco = vGrid.Text

vGrid.col = 3
vTipo = vGrid.Text

strSQL = "select nsolicitud from Tes_Transacciones where id_Banco = " & vBanco & " and tipo = '" & vTipo _
       & "' and ndocumento = '" & vDocumento & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    Call sbGridCodigo(vActiveRow, rs!NSolicitud)
End If
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim BancoCod As Integer, BancoDesc As String, TipoCod As String
Dim TipoDesc As String

  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Filtro = ""
  
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 2 Then
 'Banco
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select id_Banco,descripcion from Tes_Bancos"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     vGrid.Row = vGrid.ActiveRow
     vGrid.col = vGrid.ActiveCol
     vGrid.Text = gBusquedas.Resultado
     vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
     vGrid.CellNote = gBusquedas.Resultado2
     vGrid.TextTip = TextTipFixed
  End If
End If

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
 'Tipo
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select tipo,descripcion from tes_tipos_doc"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     vGrid.Row = vGrid.ActiveRow
     vGrid.col = vGrid.ActiveCol
     vGrid.Text = gBusquedas.Resultado
     vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
     vGrid.CellNote = gBusquedas.Resultado2
     vGrid.TextTip = TextTipFixed
  End If
End If

If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And vGrid.ActiveCol = 1 Then
  'Consulta Solicitud
  vGrid.col = 1
  vGrid.Row = vGrid.ActiveRow
  If IsNumeric(vGrid.Text) Then
      Call sbGridCodigo(vGrid.ActiveRow, vGrid.Text)
  End If
End If

If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And vGrid.ActiveCol = 4 Then
 'Busca Documento
 Call sbBuscaDocumento(vGrid.ActiveRow)
End If

If KeyCode = vbKeyReturn And vGrid.ActiveCol = 6 Then
 'Nueva Linea y Guarda
 vGrid.Row = vGrid.ActiveRow
 
 vGrid.col = 2
 BancoCod = vGrid.Text
 BancoDesc = vGrid.CellNote
 
 vGrid.col = 3
 TipoCod = vGrid.Text
 TipoDesc = vGrid.CellNote
 
 'Guardar Aqui
 Call sbGuardarLinea
 If vGrid.ActiveRow = vGrid.MaxRows Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    vGrid.col = 2
    vGrid.Text = BancoCod
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = BancoDesc
    vGrid.TextTip = TextTipFixed
    
    vGrid.col = 3
    vGrid.Text = TipoCod
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = TipoDesc
    vGrid.TextTip = TextTipFixed
 End If
End If


End Sub


Private Function fxVerificaLinea() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vSolicitud As Long, vMensaje As String


fxVerificaLinea = True

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
vSolicitud = vGrid.Text
vMensaje = ""

'Verificar que ninguna ubicacion diferente a la actual, la tenga como recibida
       
strSQL = "select isnull(max(cod_remesa),0) as Remesa" _
       & " from tes_ubi_remdet where estado = 1 and nsolicitud = " & vSolicitud
Call OpenRecordSet(rs, strSQL)
If rs!remesa > 0 Then
  strSQL = "select isnull(count(*),0) as Existe from tes_ubi_remesa where cod_ubicacion = '" _
       & fxCodigoCbo(cboOrigen) & "' and cod_remesa = " & rs!remesa
  rs.Close
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
     vMensaje = vMensaje & " - La Solicitud : " & vSolicitud & " no se puede registrar en esta remesa" _
              & ", porque no se encuentra registrada en el Origen : " & cboOrigen.Text & vbCrLf
  End If
End If
rs.Close

If txtEstado.Tag = "R" Then vMensaje = vMensaje & " - La remesa ya fue recibida, no se pueden variar sus datos" & vbCrLf

If Len(vMensaje) > 0 Then
   fxVerificaLinea = False
   MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub sbGuardarLinea()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vSolicitud As Long, vNotas As String


On Error GoTo vError

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
vSolicitud = vGrid.Text
vGrid.col = 6
vNotas = vGrid.Text



If Not fxVerificaLinea Then Exit Sub

'Verifica si existe el documento
strSQL = "select isnull(count(*),0) as Existe from tes_ubi_remDet where nsolicitud = " & vSolicitud _
       & " and cod_remesa = " & vCodigo
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   strSQL = "insert tes_ubi_remDet(cod_remesa,nsolicitud,estado,observacion,fecha_rec,usuario_rec) values(" _
          & vCodigo & "," & vSolicitud & ",0,'" & vNotas & "',null,'')"
Else
   strSQL = "update tes_ubi_remDet set observacion = '" & vNotas & "' where cod_remesa = " & vCodigo _
          & " and Nsolicitud = " & vSolicitud
End If
rs.Close
Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub
