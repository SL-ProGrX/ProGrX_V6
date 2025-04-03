VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerificaSaldosASE 
   Caption         =   "Verificación de Saldos Sistema ASE"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmBusca.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCorteSaldos 
      Caption         =   "Establecer Nuevo Corte"
      Height          =   315
      Left            =   8280
      TabIndex        =   10
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtp1 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   24707075
      CurrentDate     =   37040
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   24707075
      CurrentDate     =   37040
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   65280
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#OP"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Codigo"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cedula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Saldo Inicial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Saldo Final"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Debitos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Creditos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Diferencia"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   4155
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte de Saldos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5040
      TabIndex        =   9
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblUltimoCorte 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   8
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmVerificaSaldosASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCon As New ADODB.Connection

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rsTMP As New ADODB.Recordset
Dim curDebitos As Currency, curCreditos As Currency
Dim lngEvaluados As Long, lngDiferencias As Long
Dim lngRegistros As Long, curSaldoFinal As Currency
Dim itmX As ListItem, curTotales(5) As Currency

Me.MousePointer = vbHourglass
lsw.ListItems.Clear


lblEstado.Caption = vbCrLf & "****- Actualizado operaciones nuevas -****"
lblEstado.Refresh

strSQL = "update reg_creditos set saldo_inicial = montoapr" _
       & " where fechaforp between '" & Format(dtp1.Value, "mm/dd/yyyy") _
       & "' and '" & Format(dtp2.Value, "mm/dd/yyyy") & "' and saldo_inicial = 0" _
       & " and estado in('A','C','N')"
vCon.Execute strSQL

lblEstado.Caption = vbCrLf & "****- Cargando Información, Espere -****"
lblEstado.Refresh


strSQL = "Select R.id_solicitud,R.codigo,R.saldo,R.saldo_inicial,S.cedula,S.nombre" _
       & " From Reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " Where Saldo is not null and Saldo <> Saldo_inicial and estado in('A','C','N')"
      ' & " and fechaforp < '" & Format(dtp1.Value, "mm/dd/yyyy") & "'"
       
With rs
 .Open strSQL, vCon, adOpenStatic
    lngRegistros = .RecordCount
    prg.Max = .RecordCount + 1
    prg.Value = 1
    
    curTotales(1) = 0
    curTotales(2) = 0
    curTotales(3) = 0
    curTotales(4) = 0
    curTotales(5) = 0
    
    Do While Not .EOF
     lngEvaluados = lngEvaluados + 1
     
     If !saldo <> !saldo_inicial Then
       curDebitos = 0
       curCreditos = 0
       
       strSQL = "Select * From Creditos_dt where id_solicitud=" & !id_solicitud
       strSQL = strSQL & " And Fechas between '" & Format(dtp1, "mm/dd/yyyy") & "' And '" & Format(dtp2, "mm/dd/yyyy") & "'"
       rsTMP.Open strSQL, vCon, adOpenStatic
         Do While Not rsTMP.EOF
            Select Case rsTMP!tcon
              Case 8
                curDebitos = curDebitos + rsTMP!Amortiza
              Case Else
                curCreditos = curCreditos + rsTMP!Amortiza
            End Select
            rsTMP.MoveNext
         Loop
       rsTMP.Close
              
       strSQL = "Select * From Morosidad where id_solicitud=" & !id_solicitud _
              & " And Fecult between '" & Format(dtp1, "mm/dd/yyyy") _
              & "' And '" & Format(dtp2, "mm/dd/yyyy") & "' and estado <> 'A'"
       rsTMP.Open strSQL, vCon, adOpenStatic
         Do While Not rsTMP.EOF
            Select Case rsTMP!tcon
              Case 8
                curDebitos = curDebitos + IIf(IsNull(rsTMP!abAmortiza), 0, rsTMP!abAmortiza)
              Case Else
                curCreditos = curCreditos + IIf(IsNull(rsTMP!abAmortiza), 0, rsTMP!abAmortiza)
            End Select
            rsTMP.MoveNext
         Loop
       rsTMP.Close
       
       curSaldoFinal = !saldo_inicial + curDebitos - curCreditos
       If Abs(curSaldoFinal - !saldo) > 1 Then
         
         lngDiferencias = lngDiferencias + 1
         
         Set itmX = lsw.ListItems.Add(, , !id_solicitud)
           itmX.SubItems(1) = Trim(!Codigo)
           itmX.SubItems(2) = Trim(!Cedula)
           itmX.SubItems(3) = Trim(!nombre) & ""
           itmX.SubItems(4) = Format(!saldo_inicial, "Standard")
           itmX.SubItems(5) = Format(!saldo, "Standard")
           itmX.SubItems(6) = Format(curDebitos, "Standard")
           itmX.SubItems(7) = Format(curCreditos, "Standard")
           itmX.SubItems(8) = Format(!saldo - curSaldoFinal, "Standard")
                 
        curTotales(1) = curTotales(1) + !saldo_inicial
        curTotales(2) = curTotales(2) + !saldo
        curTotales(3) = curTotales(3) + curDebitos
        curTotales(4) = curTotales(4) + curCreditos
        curTotales(5) = curTotales(5) + !saldo - curSaldoFinal
        
       
                 
       End If
     End If
     
     .MoveNext
     lblEstado = "Procesados  " & Format(lngEvaluados, "###,###,###,##0") _
               & "  De  " & Format(lngRegistros, "###,###,###,##0") & "  -  Porcentaje: " _
               & Round((lngEvaluados / lngRegistros) * 100, 2) & "%" & vbCrLf _
               & "Evaluados " & Format(lngEvaluados, "###,###,###,##0") & "  -  Diferencias " & Format(lngDiferencias, "###,###,###,##0")
    
     prg.Value = prg.Value + 1
     lblEstado.Refresh
    Loop
 .Close
    
       'TOTALES
         Set itmX = lsw.ListItems.Add()
           itmX.SubItems(4) = "---------------------"
           itmX.SubItems(5) = "---------------------"
           itmX.SubItems(6) = "---------------------"
           itmX.SubItems(7) = "---------------------"
           itmX.SubItems(8) = "---------------------"
                 
         Set itmX = lsw.ListItems.Add()
           itmX.SubItems(4) = Format(curTotales(1), "Standard")
           itmX.SubItems(5) = Format(curTotales(2), "Standard")
           itmX.SubItems(6) = Format(curTotales(3), "Standard")
           itmX.SubItems(7) = Format(curTotales(4), "Standard")
           itmX.SubItems(8) = Format(curTotales(5), "Standard")
                 
End With

Me.MousePointer = vbDefault

End Sub

Private Sub cmdCorteSaldos_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iRespuesta As Integer


iRespuesta = MsgBox("Esta seguro que desea establecer nuevo saldo inicial, se le recuerda" _
                   & " que tiene que ser el ultimo día del mes, cuando ya no se procese información", vbYesNo)
If iRespuesta = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "update reg_creditos set saldo_inicial = saldo where estado is not null"
vCon.Execute strSQL

rs.Open "select getdate() as fecha", vCon, adOpenStatic
strSQL = "update par_ahcr set cr_corte_si ='" & Format(rs!fecha, "mm/dd/yyyy") & "'"
lblUltimoCorte.Caption = Format(rs!fecha, "dd/mmm/yyyy")
dtp1.Value = rs!fecha
dtp2.Value = rs!fecha
rs.Close

vCon.Execute strSQL

Me.MousePointer = vbDefault
MsgBox "Nuevo Corte para verificación de Saldos Establecida", vbInformation

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
vCon.Open "Driver={Sql Server};Database=ASECCSS;Server=Perseus3;uid=sa"
vCon.CommandTimeout = 2000

rs.Open "select cr_corte_si as fecha,getdate() as Actual from par_ahcr", vCon, adOpenStatic
lblUltimoCorte.Caption = Format(rs!fecha, "dd/mmm/yyyy")
dtp1.Value = rs!fecha
dtp2.Value = rs!actual
rs.Close
End Sub

Private Sub Form_Resize()
On Error Resume Next

lsw.Width = Me.Width - 100
lsw.Height = Me.Height - (800 + lblEstado.Height + prg.Height)
lblEstado.Top = lsw.Top + lsw.Height + 20
lblEstado.Width = lsw.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
vCon.Close
End Sub

