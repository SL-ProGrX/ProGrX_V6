VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCR_ConveniosParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de Convenios"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmCR_ConveniosParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   6840
   Begin TabDlg.SSTab ssTab 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Créditos"
      TabPicture(0)   =   "frmCR_ConveniosParametros.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lsw"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdActualizaCreditos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tablas"
      TabPicture(1)   =   "frmCR_ConveniosParametros.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vGrid"
      Tab(1).Control(1)=   "Label2"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdActualizaCreditos 
         Caption         =   "&Actualizar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   4320
         Width           =   1215
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6588
         View            =   3
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7832
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   6375
         _Version        =   524288
         _ExtentX        =   11245
         _ExtentY        =   6800
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
         MaxCols         =   3
         MaxRows         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_ConveniosParametros.frx":0342
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Tabla de Membresia"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -70080
         TabIndex        =   5
         Top             =   4440
         Width           =   1575
      End
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Convenio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCR_ConveniosParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cbo_Click()
If cbo.Text <> "" Then
 Call sbLlenaLsw
End If
End Sub

Private Sub cmdActualizaCreditos_Click()
Dim strSQL As String, i As Integer, vTipo As String
Dim rs As New ADODB.Recordset, vMensaje As String

On Error GoTo vError


Me.MousePointer = vbHourglass

If Mid(cbo.Text, 1, 2) = "01" Then 'Comercial
  vTipo = "C"
Else 'Especial
  vTipo = "E"
End If

strSQL = "delete convenios_codigos where tipo = '" & vTipo & "'"
glogon.Conection.Execute strSQL

vMensaje = ""

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
    strSQL = "select coalesce(count(*),0) as Existe from convenios_codigos" _
           & " where codigo = '" & Trim(lsw.ListItems.Item(i).Text) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If rs!existe = 0 Then
        strSQL = "insert convenios_codigos(tipo,codigo) values('" & vTipo & "','" _
               & Trim(lsw.ListItems.Item(i).Text) & "')"
        glogon.Conection.Execute strSQL
    Else
        vMensaje = vMensaje & " ¦ " & Trim(lsw.ListItems.Item(i).Text)
    End If
    rs.Close
           
  End If
Next i

Me.MousePointer = vbDefault

If Len(vMensaje) > 0 Then
    MsgBox vMensaje, vbExclamation, "Códigos NO Ingresados porque Pertenecen al Otro Tipo de Convenio"
End If

MsgBox "Información Guardada Satisfactoriamente...", vbInformation

Call sbLlenaLsw

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
cbo.AddItem "01 - Convenio Comercial"
cbo.AddItem "02 - Convenio Especial"

cbo.Text = "01 - Convenio Comercial"

ssTab.Tab = 0

vModulo = 3
Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.Enabled = cmdActualizaCreditos.Enabled

End Sub


Private Sub sbLlenaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, vTipo As String


Me.MousePointer = vbHourglass

strSQL = "select C.codigo,C.descripcion,X.tipo" _
       & " from Catalogo C left join Convenios_Codigos X" _
       & " On C.codigo = X.codigo"
       
If Mid(cbo.Text, 1, 2) = "01" Then 'Comercial
  vTipo = "C"
Else 'Especial
  vTipo = "E"
End If
strSQL = strSQL & " and X.tipo = '" & vTipo & "' where C.convenio = 'S' Order by X.tipo desc,C.codigo"

rs.Open strSQL, glogon.Conection, adOpenStatic
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      itmX.Checked = IIf(IsNull(rs!Tipo), False, True)
  If itmX.Checked Then itmX.ForeColor = vbBlue
  rs.MoveNext
Loop
rs.Close

strSQL = "select Desde,Hasta,Monto from Convenios_Tablas where Tipo = '" _
       & vTipo & "' order by Desde,Hasta"

Call sbCargaGrid(vGrid, 3, strSQL)

Me.MousePointer = vbDefault

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset, vTipo As String
Dim X As Long, lng As Long, vTemp(3) As Variant

On Error GoTo vError

If Mid(cbo.Text, 1, 2) = "01" Then 'Comercial
  vTipo = "C"
Else 'Especial
  vTipo = "E"
End If

If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.Text <> "" Then
    
    strSQL = "delete convenios_tablas where tipo = '" & vTipo & "' and desde = " & vGrid.Text
    glogon.Conection.Execute strSQL
    
    For lng = vGrid.ActiveRow To vGrid.MaxRows
       vGrid.Row = lng + 1
       For X = 1 To 3
          vGrid.Col = X
          vTemp(X) = vGrid.Text
       Next X
       
       vGrid.Row = lng
       For X = 1 To 3
         vGrid.Col = X
         vGrid.Text = vTemp(X)
       Next X
    Next lng
    vGrid.MaxRows = vGrid.MaxRows - 1
    If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1

    
  End If
End If

'Guarda Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  
  strSQL = "select coalesce(count(*),0) as Existe from convenios_tablas" _
         & " where tipo = '" & vTipo & "' and desde = " & vGrid.Text
  rs.Open strSQL, glogon.Conection, adOpenStatic
  If rs!existe = 0 Then
        strSQL = "insert convenios_tablas(Tipo,Desde,Hasta,Monto) values('" & vTipo & "'," _
               & vGrid.Text & ","
        vGrid.Col = 2
        strSQL = strSQL & vGrid.Text & ","
        vGrid.Col = 3
        strSQL = strSQL & CCur(vGrid.Text) & ")"
        glogon.Conection.Execute strSQL
  Else
   'Actualiza
        vGrid.Col = 2
        strSQL = "update convenios_tablas set hasta = " & vGrid.Text
        vGrid.Col = 3
        strSQL = strSQL & ",monto = " & CCur(vGrid.Text)
        vGrid.Col = 1
        strSQL = strSQL & " where Tipo = '" & vTipo & "' and desde = " & vGrid.Text
        glogon.Conection.Execute strSQL
  
  End If
  rs.Close
  
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Da formato a las cuentas
If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And vGrid.ActiveCol < vGrid.MaxCols Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
End If

Exit Sub

vError:
MsgBox Err.Description, vbCritical

End Sub

