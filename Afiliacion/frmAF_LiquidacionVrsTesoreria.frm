VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAF_LiquidacionVrsTesoreria 
   Caption         =   "Liquidaciones Generadas a Tesoreria para Desembolsos"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "frmAF_LiquidacionVrsTesoreria.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   8430
   WindowState     =   2  'Maximized
   Begin VB.OptionButton opt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pendientes"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resumen"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   735
      Left            =   4200
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   765
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Liq.Fec."
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Tipo.Doc"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "#Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Banco"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Id.Liq"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdArchivo 
      Caption         =   "&Archivo"
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   195493891
      CurrentDate     =   37405
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   195493891
      CurrentDate     =   37405
   End
   Begin VB.Label Label1 
      Caption         =   "Corte"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmAF_LiquidacionVrsTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArchivo_Click()
Dim fn, strArchivo As String
Dim lng As Long, i As Integer, vCadena As String

On Error GoTo vError

fn = FreeFile
 frmContenedor.CD.InitDir = "C:\"
 frmContenedor.CD.ShowSave
 
 Open frmContenedor.CD.FileName For Output As #fn
    vCadena = "CEDULA" & vbTab & "NOMBRE" & vbTab & "FEC.LIQ" & vbTab & "MONTO" & vbTab & "TIPO.DOC" _
            & vbTab & "#DOCUMENTO" & vbTab & "BANCO" & vbTab & "ID.LIQ" & vbCrLf
    Print #fn, vCadena
    
    For lng = 1 To lsw.ListItems.Count
      vCadena = lsw.ListItems.Item(lng) & vbTab
      For i = 1 To lsw.ColumnHeaders.Count - 1
        vCadena = vCadena & lsw.ListItems.Item(lng).SubItems(i) & vbTab
      Next i
      Print #fn, vCadena
    Next lng

 Close #fn
 MsgBox "Información Guardada en " & frmContenedor.CD.FileName, vbInformation

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, vCuenta As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select CTA_LIQPAS From Par_AfAH"
Call OpenRecordSet(rs, strSQL)
  vCuenta = Trim(rs!cta_liqpas)
rs.Close

lsw.ListItems.Clear


strSQL = "select L.consec,S.cedula,S.nombre,L.Fecha_Traspaso,Tmp_Monto as Monto" _
       & " from liquidacion L inner join Socios S on L.cedula = S.cedula" _
       & " inner join Asientos_tmp A on A.tmp_caso = L.consec and A.tmp_tipo = 'LIQ'" _
       & " where A.Tmp_debehaber = 'H' and A.tmp_cuenta = '" & vCuenta & "'" _
       & " and L.Fecha_Traspaso between '" & Format(dtpInicio, "yyyy/mm/dd") & "' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 
 Set itmX = lsw.ListItems.Add(, , rs!Cedula)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = Format(rs!fecha_traspaso, "yyyy/mm/dd")
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     
 strSQL = "select C.tipo,C.ndocumento,B.descripcion" _
        & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_Banco" _
        & " where C.codigo = '" & rs!Cedula & "' and Monto = " & rs!Monto
 Call OpenRecordSet(rsTmp, strSQL, 0)
 If Not rsTmp.EOF And Not rsTmp.BOF Then
     itmX.SubItems(4) = rsTmp!Tipo
     itmX.SubItems(5) = rsTmp!nDocumento & ""
     itmX.SubItems(6) = rsTmp!Descripcion
 End If
 rsTmp.Close
 
 itmX.SubItems(7) = rs!consec
 
 
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
MsgBox "Consulta Finalizada...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Personas"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

 .Connect = glogon.ConectRPT
 
 Select Case True
   Case opt.Item(0).Value
      .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTesoreriaResumen.rpt")
        strSQL = "{LIQUIDACION.FECHA_TRASPASO} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
   Case opt.Item(1).Value
      .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTesoreriaDetalle.rpt")
        strSQL = "{LIQUIDACION.FECHA_TRASPASO} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
   Case opt.Item(2).Value
      .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTesoreriaPendientes.rpt")
        strSQL = "{LIQUIDACION.FECLIQ} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ") and {LIQUIDACION.ESTADOASIENTO} = 'P'"
 End Select


  .Formulas(2) = "SubTitulo='Del  " & Format(dtpInicio, "dd/mm/yyyy") & "  Al  " & Format(dtpCorte, "dd/mm/yyyy") & "'"
  
  .SelectionFormula = strSQL
  .PrintReport

End With

Me.MousePointer = vbDefault


End Sub

Private Sub Form_Load()

dtpInicio = fxFechaServidor
dtpCorte = dtpInicio

End Sub

Private Sub Form_Resize()
On Error Resume Next

lsw.Width = Me.Width - 130
lsw.Height = Me.Height - 1050


End Sub
