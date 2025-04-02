VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmTES_EntregaDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega de Documentos"
   ClientHeight    =   6360
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10704
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   10704
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   5295
   End
   Begin VB.ComboBox cboDoc 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   5052
   End
   Begin VB.CommandButton cmdSeguridad 
      Caption         =   "Seg.System"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5292
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10452
      _ExtentX        =   18436
      _ExtentY        =   9335
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ingreso Manual"
      TabPicture(0)   =   "frmTES_EntregaDocumentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista de Pendientes"
      TabPicture(1)   =   "frmTES_EntregaDocumentos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsw"
      Tab(1).Control(1)=   "chkTodos"
      Tab(1).Control(2)=   "cmdBuscar"
      Tab(1).Control(3)=   "dtpInicio"
      Tab(1).Control(4)=   "dtpCorte"
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).ControlCount=   6
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4212
         Left            =   -74880
         TabIndex        =   12
         Top             =   960
         Width           =   10212
         _Version        =   1245187
         _ExtentX        =   18013
         _ExtentY        =   7429
         _StockProps     =   77
         BackColor       =   -2147483643
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
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   492
         Left            =   -69720
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4572
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   10212
         _Version        =   524288
         _ExtentX        =   18013
         _ExtentY        =   8065
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
         MaxCols         =   487
         ScrollBars      =   2
         SpreadDesigner  =   "frmTES_EntregaDocumentos.frx":0038
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   432
         Left            =   -66120
         TabIndex        =   9
         Top             =   480
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "&Buscar"
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
         Picture         =   "frmTES_EntregaDocumentos.frx":06BB
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -72840
         TabIndex        =   10
         Top             =   480
         Width           =   1452
         _Version        =   1245187
         _ExtentX        =   2561
         _ExtentY        =   550
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -71400
         TabIndex        =   11
         Top             =   480
         Width           =   1452
         _Version        =   1245187
         _ExtentX        =   2561
         _ExtentY        =   550
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas.:"
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
         Index           =   2
         Left            =   -74160
         TabIndex        =   1
         Top             =   480
         Width           =   1212
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Documento..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Bancaria..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_EntregaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()
If vPaso Then Exit Sub

vPaso = True
Call sbTesTiposDocsCargaCbo(cboDoc, cbo.ItemData(cbo.ListIndex))
vPaso = False

Call cboDoc_Click

End Sub

Private Sub cboDoc_Click()
Dim i As Integer

If vPaso Then Exit Sub

ssTab.Tab = 0

vGrid.MaxCols = 5
vGrid.MaxRows = 1

vGrid.Row = 1
For i = 1 To vGrid.MaxCols
   vGrid.col = i
   vGrid.Text = ""
Next i

End Sub

Private Sub chkTodos_Click()
If chkTodos.Value = vbChecked Then
  dtpInicio.Enabled = False
Else
  dtpInicio.Enabled = True
End If
dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass



lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Solicitud", 1400
    .Add , , "No. Documento", 1800, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Emisión", 1800, vbCenter
    .Add , , "Beneficiario", 3500
End With



strSQL = "select nsolicitud,ndocumento,beneficiario,monto,fecha_emision from Tes_Transacciones" _
       & " where id_banco = " & cbo.ItemData(cbo.ListIndex) & " and tipo = '" _
       & SIFGlobal.fxCodText(cboDoc.Text) & "' and user_entrega is null and estado <> 'P'"
       
If chkTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and fecha_emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' order by ndocumento"
 End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
     itmX.SubItems(1) = rs!nDocumento & ""
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.SubItems(3) = rs!Fecha_Emision
     itmX.SubItems(4) = rs!Beneficiario & ""
 rs.MoveNext
Loop
rs.Close

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

vModulo = 9

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
Call sbTesBancoCargaCboAccesoGeneral(cbo)
vPaso = False

Call cbo_Click

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

ssTab.Tab = 0

Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.Enabled = cmdSeguridad.Enabled
lsw.Enabled = cmdSeguridad.Enabled


End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)


Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
    strSQL = "update Tes_Transacciones set user_entrega = '" & glogon.Usuario _
           & "', fecha_entrega = dbo.MyGetdate() where nsolicitud = " & Item.Text
Else
    strSQL = "update Tes_Transacciones set user_entrega = null, fecha_entrega = null where nsolicitud = " & Item.Text
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)

If ssTab.Tab = 1 Then
  lsw.ListItems.Clear
End If

End Sub


Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If col = 1 Then
 'Consultar
 
 vGrid.Row = Row
 vGrid.col = 1
 strSQL = "select nsolicitud,ndocumento,beneficiario,monto,Fecha_Emision from Tes_Transacciones" _
        & " where id_banco = " & cbo.ItemData(cbo.ListIndex) & " and tipo = '" _
        & SIFGlobal.fxCodText(cboDoc.Text) & "' and user_entrega is null and estado <> 'P'" _
        & " and ndocumento = '" & vGrid.Text & "'"
 Call OpenRecordSet(rs, strSQL)
 If Not rs.EOF And Not rs.BOF Then
   vGrid.col = 2
   vGrid.Text = CStr(rs!Monto)
   vGrid.col = 3
   vGrid.Text = CStr(rs!Beneficiario)
   vGrid.col = 4
   vGrid.Text = CStr(rs!NSolicitud)
   vGrid.CellTag = "S"
   vGrid.col = 5
   vGrid.Text = Format(rs!Fecha_Emision, "dd/mm/yyyy")

 Else
   vGrid.col = 4
   vGrid.CellTag = "N"
 End If
 rs.Close
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And vGrid.ActiveCol = 4 Then
 vGrid.Row = vGrid.ActiveRow
 vGrid.col = 4
 If vGrid.CellTag = "S" Then
    strSQL = "update Tes_Transacciones set user_entrega = '" & glogon.Usuario _
           & "', fecha_entrega = dbo.MyGetdate() where nsolicitud = " & vGrid.Text
    Call ConectionExecute(strSQL)
 End If
 
 If vGrid.MaxRows = vGrid.Row Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
 End If
End If

If KeyCode = vbKeyInsert And vGrid.ActiveRow > 1 Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 1
End If

End Sub


