VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPres_ControlPresupuestario 
   Caption         =   "Control Presupuestario"
   ClientHeight    =   3645
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7620
   HelpContextID   =   12
   Icon            =   "frmPres_ControlPresupuestario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   582
      ButtonWidth     =   2064
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar"
            Key             =   "buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reporte"
            Key             =   "reporte"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPres_ControlPresupuestario.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPres_ControlPresupuestario.frx":0626
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboBuscar 
      Height          =   315
      ItemData        =   "frmPres_ControlPresupuestario.frx":0942
      Left            =   1200
      List            =   "frmPres_ControlPresupuestario.frx":094F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4895
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
      NumItems        =   0
   End
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Criterios"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblPeriodo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmPres_ControlPresupuestario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select (coalesce(M.saldo_inicial,0) + coalesce(M.total_Debitos,0) + coalesce(M.Total_Creditos,0)) as Real" _
       & ",P.cod_cuenta,Coalesce(P.presu_Original,0) as Presu_Original" _
       & ",Coalesce(P.presu_actual,0) as Presu_actual,Coalesce(P.ajuste_positivo,0) as Ajuste_Positivo" _
       & ",Coalesce(P.Ajuste_Negativo,0) as Ajuste_Negativo" _
       & " from presupuesto P left join Movimiento_Cuentas M on P.COD_CONTABILIDAD = P.COD_CONTABILIDAD" _
       & " and P.cod_cuenta = M.cod_cuenta and P.anio = M.anio and P.mes = M.mes" _
       & " where P.anio = " & txtAnio & " and P.mes = " & txtMes _
       & " AND P.COD_CONTABILIDAD = " & vParametros.CodigoEmpresa

Select Case Mid(cboBuscar, 1, 2)
   Case "01" 'Listado General (No hay que hacer nada)
   Case "02" 'Cuentas a Favor
     strSQL = strSQL & " and P.presu_Actual > (coalesce(M.saldo_inicial,0) + coalesce(M.total_Debitos,0) + coalesce(M.Total_Creditos,0))"
   Case "03" 'Cuenta en Contra
     strSQL = strSQL & " and P.presu_Actual < (coalesce(M.saldo_inicial,0) + coalesce(M.total_Debitos,0) + coalesce(M.Total_Creditos,0))"
End Select
strSQL = strSQL & " order by P.cod_cuenta"

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , fxFormatoCuenta(True, rs!COD_Cuenta))
     itmX.SubItems(1) = fxCuenta("D", rs!COD_Cuenta)
     itmX.SubItems(2) = Format(rs!presu_actual, "Standard")
     itmX.SubItems(3) = Format(rs!Real, "Standard")
     itmX.SubItems(4) = Format(rs!presu_actual - rs!Real, "Standard")
     itmX.SubItems(5) = Format(rs!presu_original, "Standard")
     itmX.SubItems(6) = Format(rs!ajuste_positivo, "Standard")
     itmX.SubItems(7) = Format(rs!ajuste_negativo, "Standard")
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault
MsgBox "Consulta Finalizada...", vbInformation

End Sub


Private Sub sbReportePantalla()
Dim fn, strArchivo As String, bPaso As Boolean
Dim lng As Long, i As Integer, vCadena As String


On Error GoTo vError
 bPaso = True
  
xPaso:
On Error Resume Next

fn = FreeFile
 frmContenedor.CD.InitDir = "C:\"
 frmContenedor.CD.ShowSave
 
 Kill frmContenedor.CD.FileName
 
 Open frmContenedor.CD.FileName For Output As #fn
    vCadena = "CUENTA" & vbTab & "DESCRIPCION" & vbTab & "PRESUPUESTO" & vbTab _
            & "REAL" & vbTab & "DIFERENCIA" & vbTab & "ORIGINAL" & vbTab & "(+) AJUSTES" _
            & vbTab & "(-) AJUSTES"

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
 bPaso = False
 GoTo xPaso

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
 Case "buscar"
    Call sbBuscar
 Case "reporte"
    Call sbReportePantalla
End Select
End Sub

Private Sub txtMes_Change()
On Error GoTo vError
lblPeriodo.Caption = fxPeriodoRes(txtAnio, txtMes)
vError:
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
End Sub

Private Sub txtAnio_Change()
On Error GoTo vError
lblPeriodo.Caption = fxPeriodoRes(txtAnio, txtMes)
vError:
End Sub

Private Sub Form_Load()
Set Me.MouseIcon = frmContenedor.MouseIcon

cboBuscar.Clear
cboBuscar.AddItem "01 - Listado General"
cboBuscar.AddItem "02 - Cuentas a Favor"
cboBuscar.AddItem "03 - Cuentas en Contra"
cboBuscar.Text = "01 - Listado General"

txtMes = Month(fxFechaServidor)
txtAnio = Year(fxFechaServidor)

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "Cuenta", 1200
lsw.ColumnHeaders.Add , , "Descripción", 4200
lsw.ColumnHeaders.Add , , "Presupuesto", 1400, vbRightJustify
lsw.ColumnHeaders.Add , , "Real", 1400, vbRightJustify
lsw.ColumnHeaders.Add , , "Diferencia", 1400, vbRightJustify
lsw.ColumnHeaders.Add , , "Original", 1400, vbRightJustify
lsw.ColumnHeaders.Add , , "(+) Ajustes", 1400, vbRightJustify
lsw.ColumnHeaders.Add , , "(-) Ajustes", 1400, vbRightJustify


End Sub

Private Sub Form_Resize()
On Error Resume Next
lsw.Width = Me.Width - 100
lsw.Height = Me.Height - 550 - lsw.Top
End Sub
