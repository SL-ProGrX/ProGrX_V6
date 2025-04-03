VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmRastreoMovOp 
   Caption         =   "Analisis de Saldos de Operaciones de Crédito"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   HelpContextID   =   7003
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3492
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   5652
      _Version        =   1441792
      _ExtentX        =   9970
      _ExtentY        =   6159
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
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   600
      Top             =   360
   End
   Begin XtremeSuiteControls.CheckBox chkDiferencias 
      Height          =   200
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Width           =   200
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   0
      Top             =   5424
      Width           =   10776
      _ExtentX        =   18997
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   492
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   1452
      _Version        =   1441792
      _ExtentX        =   2561
      _ExtentY        =   868
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
      Appearance      =   17
      Picture         =   "frmRastreoMovOp.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodos 
      Height          =   312
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1812
      _Version        =   1441792
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   492
      Left            =   7320
      TabIndex        =   3
      Top             =   240
      Width           =   1452
      _Version        =   1441792
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Appearance      =   17
      Picture         =   "frmRastreoMovOp.frx":0700
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   372
      Left            =   9240
      TabIndex        =   4
      Top             =   360
      Width           =   732
      _Version        =   1441792
      _ExtentX        =   1291
      _ExtentY        =   656
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
      Text            =   "1000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4680
      Width           =   5535
      _Version        =   1441792
      _ExtentX        =   9763
      _ExtentY        =   873
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencias"
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
      Height          =   312
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodos"
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
      Height          =   312
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Líneas"
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
      Height          =   312
      Index           =   3
      Left            =   9240
      TabIndex        =   5
      Top             =   120
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmRastreoMovOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

prg.Visible = True

Call Excel_Exportar_Lsw(lsw, prg)

prg.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnBuscar_Click()
 If cboPeriodos.Text = "" Then Exit Sub
 Call sbBuscaHistorico
End Sub


Private Sub sbBuscaHistorico()

Dim lngEvaluados As Long, lngDiferencias As Long
Dim lngRegistros As Long, curSaldoFinal As Currency
Dim curTotales(5) As Currency
Dim lngLineas As Long

On Error GoTo vError


Me.MousePointer = vbHourglass

lsw.ListItems.Clear


lblEstado.Caption = vbCrLf & "****- Cargando Información (Espere) -****"

strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
Call OpenRecordSet(rs, strSQL)


strSQL = "select Top " & txtLineas.Text & " H.*,S.cedula,S.nombre " _
       & "from ase_per_cerrados H inner join reg_creditos R on H.id_solicitud = R.id_solicitud " _
       & "  left join Socios S on R.cedula = S.cedula " _
       & " inner join ase_per_Catalogo C on C.codigo = H.codigo and C.mes = H.mes and C.anio = H.anio" _
       & " where H.anio = " & rs!Anio & " and H.mes = " & rs!Mes _
       & " and C.retencion = 'N' and C.poliza = 'N'"
       
If chkDiferencias.Value = vbChecked Then
 strSQL = strSQL & " and abs(H.saldo_final - (H.saldo_inicial + H.total_debitos - H.total_creditos)) > 1"
End If
       
rs.Close

Call OpenRecordSet(rs, strSQL)

lngRegistros = rs.RecordCount
prg.Visible = True

prg.Max = rs.RecordCount + 1
prg.Value = 1
lngLineas = 1

curTotales(1) = 0
curTotales(2) = 0
curTotales(3) = 0
curTotales(4) = 0
curTotales(5) = 0

Do While Not rs.EOF
 lngEvaluados = lngEvaluados + 1
 If lngLineas > CLng(txtLineas) Then
   lngLineas = 1
   lsw.ListItems.Clear
 End If
 'Si la diferencia es mayor de 1
 
 If Abs(rs!saldo_final - (rs!saldo_inicial + rs!total_debitos - rs!total_creditos)) > 1 Then lngDiferencias = lngDiferencias + 1
     
     Set itmX = lsw.ListItems.Add(, , rs!id_solicitud)
       itmX.SubItems(1) = Trim(rs!Codigo)
       itmX.SubItems(2) = Trim(rs!Cedula & "")
       itmX.SubItems(3) = Trim(rs!Nombre & "")
       itmX.SubItems(4) = Format(rs!saldo_inicial, "Standard")
       itmX.SubItems(5) = Format(rs!saldo_final, "Standard")
       itmX.SubItems(6) = Format(rs!total_debitos, "Standard")
       itmX.SubItems(7) = Format(rs!total_creditos, "Standard")
       itmX.SubItems(8) = Format(rs!saldo_final - (rs!saldo_inicial + rs!total_debitos - rs!total_creditos), "Standard")
             
    curTotales(1) = curTotales(1) + rs!saldo_inicial
    curTotales(2) = curTotales(2) + rs!saldo_final
    curTotales(3) = curTotales(3) + rs!total_debitos
    curTotales(4) = curTotales(4) + rs!total_creditos
    curTotales(5) = curTotales(5) + rs!saldo_final - (rs!saldo_inicial + rs!total_debitos - rs!total_creditos)
    

 rs.MoveNext
 lngLineas = lngLineas + 1
 lblEstado.Caption = "Procesados  " & Format(lngEvaluados, "###,###,###,##0") _
           & "  De  " & Format(lngRegistros, "###,###,###,##0") & "  -  Porcentaje: " _
           & Round((lngEvaluados / lngRegistros) * 100, 2) & "%" & vbCrLf _
           & "Evaluados " & Format(lngEvaluados, "###,###,###,##0") & "  -  Diferencias " & Format(lngDiferencias, "###,###,###,##0")

 prg.Value = prg.Value + 1
 
 If Right(CStr(prg.Value), 2) = "00" Then DoEvents
 
Loop
rs.Close

   'TOTALES
     Set itmX = lsw.ListItems.Add()
       itmX.SubItems(4) = "_____________________"
       itmX.SubItems(5) = "_____________________"
       itmX.SubItems(6) = "_____________________"
       itmX.SubItems(7) = "_____________________"
       itmX.SubItems(8) = "_____________________"
             
     Set itmX = lsw.ListItems.Add()
       itmX.SubItems(4) = Format(curTotales(1), "Standard")
       itmX.SubItems(5) = Format(curTotales(2), "Standard")
       itmX.SubItems(6) = Format(curTotales(3), "Standard")
       itmX.SubItems(7) = Format(curTotales(4), "Standard")
       itmX.SubItems(8) = Format(curTotales(5), "Standard")
             

prg.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    prg.Visible = False
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

lsw.Width = Me.Width - 150
lsw.Height = Me.Height - (lsw.Top + lblEstado.Height + prg.Height + 480)
lblEstado.Top = lsw.Top + lsw.Height + 60
lblEstado.Left = lsw.Left
lblEstado.Width = lsw.Width

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub Form_Load()

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
       .Clear
       .Add , , "Operación", 1900
       .Add , , "Código", 1200, vbCenter
       .Add , , "Identificación", 2100, vbCenter
       .Add , , "Nombre", 3900
       .Add , , "Saldo Inicial", 2100, vbRightJustify
       .Add , , "Saldo Final", 2100, vbRightJustify
       .Add , , "Débitos", 2100, vbRightJustify
       .Add , , "Créditos", 2100, vbRightJustify
       .Add , , "Diferencia", 2100, vbRightJustify
End With

Me.BackColor = RGB(214, 234, 248)


End Sub



Private Sub sbInicial()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from ase_per_historico order by anio desc,mes desc"
Call OpenRecordSet(rs, strSQL)

cboPeriodos.Clear
Do While Not rs.EOF
 cboPeriodos.AddItem rs!Anio & " - " & fxConvierteMES(rs!Mes)
 cboPeriodos.ItemData(cboPeriodos.ListCount - 1) = CStr(rs!id_per_historico)
 rs.MoveNext
Loop

If rs.RecordCount > 0 Then
 rs.MoveFirst
 Call sbCboAsignaDato(cboPeriodos, rs!Anio & " - " & fxConvierteMES(rs!Mes), True, rs!id_per_historico)
End If
rs.Close

Me.MousePointer = vbDefault


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub
