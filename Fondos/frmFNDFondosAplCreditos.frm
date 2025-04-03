VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFNDFondosAplCreditos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicación de Fondos a Créditos"
   ClientHeight    =   7785
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   12030
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5175
      Left            =   0
      TabIndex        =   11
      Top             =   1680
      Width           =   12135
      _Version        =   1572864
      _ExtentX        =   21405
      _ExtentY        =   9128
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
      Item(0).Caption =   "Casos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "scPanel"
      Item(1).Caption =   "Fondos Disponibles"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswFondos"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4455
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   12015
         _Version        =   1572864
         _ExtentX        =   21193
         _ExtentY        =   7858
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
         Appearance      =   21
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswFondos 
         Height          =   4815
         Left            =   -70000
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1572864
         _ExtentX        =   21193
         _ExtentY        =   8493
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
         Appearance      =   21
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scPanel 
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   12135
         _Version        =   1572864
         _ExtentX        =   21405
         _ExtentY        =   661
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   255
      Left            =   8760
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Aplicar Todos los fondos"
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
      Appearance      =   21
   End
   Begin XtremeSuiteControls.RadioButton rbAccion 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1260
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Solo Morosidad"
      BackColor       =   16777215
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
      Appearance      =   21
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkExtraordinario 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1260
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6376
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aplicar Sobrante como Extraordinario"
      BackColor       =   16777215
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
      Appearance      =   21
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   840
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   840
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10393
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9240
      TabIndex        =   2
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   6840
      Width           =   12015
      _Version        =   1572864
      _ExtentX        =   21193
      _ExtentY        =   1720
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   615
         Left            =   9360
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aplicar"
         BackColor       =   16777215
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
         Appearance      =   21
         Picture         =   "frmFNDFondosAplCreditos.frx":0000
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   5292
      End
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   330
      Left            =   2040
      TabIndex        =   6
      Top             =   480
      Width           =   7095
      _Version        =   1572864
      _ExtentX        =   12515
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
   Begin XtremeSuiteControls.RadioButton rbAccion 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   9
      Top             =   1260
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Incluye Cuota en Transito"
      BackColor       =   16777215
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
      Appearance      =   21
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   330
      Left            =   2040
      TabIndex        =   15
      Top             =   120
      Width           =   7095
      _Version        =   1572864
      _ExtentX        =   12515
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   18
      Top             =   840
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Plan"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   17
      Top             =   480
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Operadora"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   16
      Top             =   120
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Institución"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15735
   End
End
Attribute VB_Name = "frmFNDFondosAplCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim vScroll As Boolean


Private Sub sbLsw_Llena(pTipo As String)

Dim curSaldo As Currency, curMorosidad As Currency, curDisponible As Currency
Dim vInstitucion As Long

On Error GoTo vError

lsw.ListItems.Clear

If chkTodos.Value = xtpUnchecked And txtCodigo.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass

If cboInstitucion.Text = "TODOS" Then
   vInstitucion = 0
Else
   vInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If

curSaldo = 0
curMorosidad = 0
curDisponible = 0

If chkTodos.Value = xtpChecked Then
    strSQL = "exec spFnd_FondosVrsCreditos_Lista " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & ",'-C-','" & pTipo & "', " & vInstitucion
Else
    strSQL = "exec spFnd_FondosVrsCreditos_Lista " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & ",'" & txtCodigo.Text & "','" & pTipo & "', " & vInstitucion
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cedula)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = Format(rs!Saldos, "Standard")
     itmX.SubItems(3) = Format(rs!Morosidad, "Standard")
     itmX.SubItems(4) = Format(rs!Disponible, "Standard")
 
 curSaldo = curSaldo + rs!Saldos
 curMorosidad = curMorosidad + rs!Morosidad
 curDisponible = curDisponible + rs!Disponible
 
 rs.MoveNext
Loop
rs.Close

scPanel.Caption = "Casos: " & lsw.ListItems.Count _
    & "     Saldos: " & Format(curSaldo, "Standard") _
    & "     Morosidad: " & Format(curMorosidad, "Standard") _
    & "     Disponible: " & Format(curDisponible, "Standard") & Space(10)


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnAplicar_Click()
Dim vAplMora As Integer, vAplCtaTransito As Integer, vAplExtra As Integer
Dim vInstitucion As Long


On Error GoTo vError

If lsw.ListItems.Count = 0 Then Exit Sub

Me.MousePointer = vbHourglass

vAplMora = IIf(rbAccion.Item(0).Value, 1, 0)
vAplCtaTransito = IIf(rbAccion.Item(1).Value, 1, 0)
vAplExtra = chkExtraordinario.Value

If cboInstitucion.Text = "TODOS" Then
   vInstitucion = 0
Else
   vInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If

If chkTodos.Value = xtpChecked Then
    strSQL = "exec spFnd_Aplicacion_Creditos_General '" & glogon.Usuario & "', " & vAplMora & ", " & vAplCtaTransito _
           & ", " & vAplExtra & ", " & vInstitucion
Else
    strSQL = "exec spFnd_Aplicacion_Creditos " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & ", '" & txtCodigo.Text & "', '" & glogon.Usuario & "', " & vAplMora & ", " & vAplCtaTransito _
           & ", " & vAplExtra & ", " & vInstitucion
End If
       
Call OpenRecordSet(rs, strSQL)

Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc)

Call Bitacora("Aplica", "Fondos: " & txtCodigo & " a Creditos - Masivo.." & rs!TipoDoc & "_" & rs!NumDoc)

rs.Close

If rbAccion(0).Value Then
    Call rbAccion_Click(0)
Else
    Call rbAccion_Click(1)
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
                
End Sub

Private Sub cboInstitucion_Click()
If vPaso Then Exit Sub

Dim pOpcion As Integer

If rbAccion(0).Value Then
  pOpcion = 0
Else
  pOpcion = 1
End If

Call rbAccion_Click(pOpcion)

End Sub

Private Sub chkTodos_Click()

Dim pOpcion As Integer

If chkTodos.Value = xtpChecked Then
    txtCodigo.BackColor = RGB(175, 210, 222)
    txtDescripcion.BackColor = RGB(175, 210, 222)
Else
    txtCodigo.BackColor = vbWhite
    txtDescripcion.BackColor = vbWhite
End If

If rbAccion(0).Value Then
  pOpcion = 0
Else
  pOpcion = 1
End If

Call rbAccion_Click(pOpcion)

End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & "  and indAplicarAmora = 1"
 
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!Cod_Plan
      txtCodigo_LostFocus
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True


With lsw.ColumnHeaders
    .Clear
    .Add , , "Identificación", 2100, vbCenter
    .Add , , "Nombre", 3000
    .Add , , "Saldos", 1800, vbRightJustify
    .Add , , "Morosidad", 1800, vbRightJustify
    .Add , , "Disponible", 1800, vbRightJustify
End With

With lswFondos.ColumnHeaders
    .Clear
    .Add , , "Código", 1500
    .Add , , "Descripción", 3500
    .Add , , "Contratos", 1500, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Divisa", 1000, vbCenter
End With



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub rbAccion_Click(Index As Integer)

Select Case Index
  Case 0 'Morosidad
    Call sbLsw_Llena("M")
  Case 1 'Mora + Cuota en Transito
    Call sbLsw_Llena("E")
End Select

End Sub


Private Sub sbFondos_Disponibles()

On Error GoTo vError

Dim curTotal As Currency, iContratos As Long

Me.MousePointer = vbHourglass

curTotal = 0
iContratos = 0

strSQL = "select R.COD_OPERADORA, R.COD_PLAN, R.COD_MONEDA, R.PLAN_DESC, R.TOTAL, R.CONTRATOS" _
       & " from vFnd_Contratos_Resumen R inner join FND_PLANES P on R.cod_Operadora = P.COD_OPERADORA and R.COD_PLAN = P.COD_PLAN" _
       & " Where P.indAplicarAmora = 1"
Call OpenRecordSet(rs, strSQL)

lswFondos.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lswFondos.ListItems.Add(, , rs!Cod_Plan)
     itmX.SubItems(1) = rs!Plan_Desc
     itmX.SubItems(2) = rs!Contratos
     itmX.SubItems(3) = Format(rs!Total, "Standard")
     itmX.SubItems(4) = rs!Cod_Moneda
         
    curTotal = curTotal + rs!Total
    iContratos = iContratos + rs!Contratos
     
 rs.MoveNext
Loop
rs.Close

 Set itmX = lswFondos.ListItems.Add(, , "Total")
     itmX.SubItems(1) = ""
     itmX.SubItems(2) = Format(iContratos, "###,##0")
     itmX.SubItems(3) = Format(curTotal, "Standard")
     itmX.Bold = True
     itmX.TextBackColor = RGB(150, 224, 249)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
  Call sbFondos_Disponibles
End If
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select cod_institucion as 'IdX', DESCRIPCION as 'ItmX'" _
       & "  From INSTITUCIONES Where ACTIVA = 1  order by DESCRIPCION"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)
 
 
strSQL = "select descripcion as 'itmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)
 
vPaso = False
 
Me.MousePointer = vbDefault

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   gBusquedas.Filtro = " And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) & " and indAplicarAmora = 1"
   
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtDescripcion.SetFocus
End Sub


Private Sub txtCodigo_LostFocus()

On Error GoTo vError

strSQL = "Select Descripcion from Fnd_Planes where Cod_Operadora = " _
       & cboOperadora.ItemData(cboOperadora.ListIndex) & " And Cod_Plan = '" & Trim(txtCodigo.Text) & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
   txtDescripcion.Text = Trim(rs!DESCRIPCION)
Else
   txtCodigo.Text = ""
   txtDescripcion.Text = ""
End If
rs.Close

Dim pOpcion As Integer

If rbAccion(0).Value Then
  pOpcion = 0
Else
  pOpcion = 1
End If

Call rbAccion_Click(pOpcion)

Exit Sub

vError:


End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal

   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub





