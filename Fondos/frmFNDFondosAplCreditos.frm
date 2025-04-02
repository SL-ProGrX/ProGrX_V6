VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFNDFondosAplCreditos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicación de Fondos a Créditos"
   ClientHeight    =   7320
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11115
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4092
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   11052
      _Version        =   1572864
      _ExtentX        =   19494
      _ExtentY        =   7218
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton rbAccion 
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1260
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Solo Morosidad"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkExtraordinario 
      Height          =   252
      Left            =   5160
      TabIndex        =   10
      Top             =   1260
      Width           =   3612
      _Version        =   1572864
      _ExtentX        =   6371
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aplicar Sobrante como Extraordinario"
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
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   5892
      _Version        =   1572864
      _ExtentX        =   10393
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8880
      TabIndex        =   3
      Top             =   360
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   10572
      _Version        =   1572864
      _ExtentX        =   18648
      _ExtentY        =   1714
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   615
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   5292
      End
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   2880
      TabIndex        =   9
      Top             =   720
      Width           =   5892
      _Version        =   1572864
      _ExtentX        =   10398
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
      Height          =   252
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   1260
      Width           =   2532
      _Version        =   1572864
      _ExtentX        =   4466
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Incluye Cuota en Transito"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scPanel 
      Height          =   372
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   11172
      _Version        =   1572864
      _ExtentX        =   19706
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.44
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
      Alignment       =   2
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFNDFondosAplCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vScroll As Boolean


Private Sub sbLsw_Llena(pTipo As String)
Dim strSQL As String, rs As New ADODB.Recordset

Dim curSaldo As Currency, curMorosidad As Currency, curDisponible As Currency
Dim itmX As ListViewItem

On Error GoTo vError

If txtCodigo.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass

curSaldo = 0
curMorosidad = 0
curDisponible = 0

lsw.ListItems.Clear

strSQL = "exec spFnd_FondosVrsCreditos_Lista " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ",'" & txtCodigo.Text & "','" & pTipo & "'"
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim vAplMora As Integer, vAplCtaTransito As Integer, vAplExtra As Integer

On Error GoTo vError

If lsw.ListItems.Count = 0 Then Exit Sub

Me.MousePointer = vbHourglass

vAplMora = IIf(rbAccion.Item(0).Value, 1, 0)
vAplCtaTransito = IIf(rbAccion.Item(1).Value, 1, 0)
vAplExtra = chkExtraordinario.Value

strSQL = "exec spFnd_Aplicacion_Creditos " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ",'" & txtCodigo.Text & "','" & glogon.Usuario & "'," & vAplMora & "," & vAplCtaTransito & "," & vAplCtaTransito
Call OpenRecordSet(rs, strSQL)
 
Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc)

Call Bitacora("Aplica", "Fondos: " & txtCodigo & " a Creditos - Masivo.." & rs!TipoDoc & "_" & rs!NumDoc)

rs.Close

Call rbAccion_Click(0)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
                
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
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
    rs.Close
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

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

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

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

 
strSQL = "select descripcion as 'itmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)
 
'Call opt_Click(0)

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
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
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select Descripcion from Fnd_Planes where Cod_Operadora="
strSQL = strSQL & cboOperadora.ItemData(cboOperadora.ListIndex) & " And "
strSQL = strSQL & "Cod_Plan='" & Trim(txtCodigo) & "'"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       txtDescripcion = Trim(!Descripcion)
    Else
       txtCodigo = ""
       txtDescripcion = ""
    End If
 .Close
End With

Call rbAccion_Click(0)

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





