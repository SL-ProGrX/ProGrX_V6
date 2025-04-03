VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmSIF_Emisores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entidades Emisoras"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   6960
      Top             =   360
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6612
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8052
      _Version        =   1572864
      _ExtentX        =   14203
      _ExtentY        =   11663
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
      ItemCount       =   1
      Item(0).Caption =   "Emisores"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "scTitulo"
      Item(0).Control(2)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2892
         Left            =   120
         TabIndex        =   2
         Top             =   3600
         Width           =   7812
         _Version        =   1572864
         _ExtentX        =   13779
         _ExtentY        =   5101
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   2772
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7812
         _Version        =   524288
         _ExtentX        =   13780
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
         MaxCols         =   497
         ScrollBars      =   2
         SpreadDesigner  =   "frmSIF_Emisores.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   7812
         _Version        =   1572864
         _ExtentX        =   13779
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entidades Emisoras"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3252
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11652
   End
End
Attribute VB_Name = "frmSIF_Emisores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10
 
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", lsw.Width - (1400)
End With
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
lsw.Enabled = vGrid.Enabled
 
End Sub

Private Sub sbInicializa()
Dim strSQL As String

vPaso = True

scTitulo.Caption = ""
scTitulo.Tag = ""
lsw.ListItems.Clear


strSQL = "select cod_emisor,descripcion,activo,'' as 'Btn' from sif_emisores"
Call sbCargaGrid(vGrid, 4, strSQL)

vPaso = False

End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError


If Item.Checked Then
   strSQL = "insert into sif_emisores_tarjetas(cod_emisor,cod_tarjeta,registro_usuario,registro_fecha)" _
            & " values('" & scTitulo.Tag & "','" & Item.Text & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Item.ForeColor = vbBlue
Else
   strSQL = "Delete sif_emisores_tarjetas where cod_tarjeta ='" & Item.Text _
          & "' and cod_emisor  = '" & scTitulo.Tag & "'"
   Item.ForeColor = vbRed
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description)

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If Col = 4 Then
    vGrid.Row = Row
    vGrid.Col = 1
    scTitulo.Tag = vGrid.Text
    vGrid.Col = 2
    scTitulo.Caption = vGrid.Text
    If scTitulo.Tag = "" Then
        lsw.ListItems.Clear
    Else
        Call sbLsw_Load
    End If
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
            vGrid.Row = vGrid.ActiveRow
            vGrid.Col = 1
            If Trim(vGrid.Text) <> "" Then
             strSQL = "Delete sif_emisores where cod_emisor = '" & vGrid.Text & "'"
             Call ConectionExecute(strSQL)
            End If
            vGrid.DeleteRows vGrid.ActiveRow, 1
            vGrid.MaxRows = vGrid.MaxRows - 1
    End If
End If


If KeyCode = vbKeyInsert Then
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.InsertRows vGrid.ActiveRow, 1
  vGrid.Row = vGrid.ActiveRow
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then
    fxGuardar = 0
    Exit Function
End If


strSQL = "select isnull(count(*),0) as Existe from sif_emisores  " _
       & " where cod_emisor ='" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar

    If Trim(vGrid.Text) = "" Then Exit Function
    strSQL = "insert into sif_emisores(cod_emisor,descripcion,activo,registro_usuario,registro_fecha)" _
           & " values('" & vGrid.Text & "',"
    vGrid.Col = 2
    strSQL = strSQL & "'" & vGrid.Text & "',"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"
    
    Call ConectionExecute(strSQL)
    
    vGrid.Col = 1
    Call Bitacora("Registra", "Mantenimieto Emisores: " & vGrid.Text)

Else 'Actualizar
    
    vGrid.Col = 2
    strSQL = "update sif_emisores set descripcion= '" & vGrid.Text & "',activo = "
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & " where cod_emisor =  '"
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "'"
    
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Mantenimiento Emisores: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub sbLsw_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass
  
vPaso = True

strSQL = "select E.cod_Tarjeta as 'Codigo',E.descripcion,X.cod_Tarjeta as 'Asignado'" _
        & " from sif_Tarjetas E" _
        & " left join sif_emisores_tarjetas X on E.cod_Tarjeta = X.cod_Tarjeta" _
        & " and X.cod_Emisor = '" & scTitulo.Tag _
        & "' order by X.cod_Tarjeta desc,E.cod_Tarjeta"

lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!DESCRIPCION
 
  If Not IsNull(rs!ASIGNADO) Then
     itmX.Checked = True
  End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault
  
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
