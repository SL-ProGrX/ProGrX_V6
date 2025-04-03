VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCR_ReporteDiferidos 
   Caption         =   "Reporte de Calculos de Diferidos"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   Icon            =   "frmCR_ReporteDiferidos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   10035
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   8
      Text            =   "CEXD"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   3
      Top             =   4740
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   37652
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#Op"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Monto (Apr)"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Monto Base"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Fec.Forma"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Fec.Final"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Dias (Total)"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Total (Dif)"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Fec.Corte"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Dias (Corte)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Mes (Dif)"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Dias (Acumulados)"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Acumulado (Dif)"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Tasa"
         Object.Width           =   1658
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   37652
   End
   Begin VB.Label Label3 
      Caption         =   " Línea Crédito"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Corte Cálculo"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicio"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   7920
      X2              =   7920
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   7920
      X2              =   7920
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   9860
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   9840
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmCR_ReporteDiferidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, curMonto As Currency, iDias As Integer
Dim vFecha As Date

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select R.id_solicitud,R.codigo,R.cedula,S.nombre,R.Montoapr,R.MONTOCALCULO" _
       & ",R.fechaforp,R.fecha_calculo_int,R.int" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado in('A','C') and R.codigo = '" & txtCodigo & "' and R.fechaforp between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
       
'       & "' and '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
'       & "' between R.fechaforp and R.fecha_calculo_int"
Call OpenRecordSet(rs, strSQL)
prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

lsw.ListItems.Clear

Do While Not rs.EOF
 If IsNull(rs!MontoCalculo) Then
   curMonto = rs!montoapr
 Else
   curMonto = rs!MontoCalculo
 End If
 
 Set itmX = lsw.ListItems.Add(, , rs!Id_Solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = Format(rs!montoapr, "Standard")
     itmX.SubItems(5) = Format(curMonto, "Standard")
     itmX.SubItems(6) = Format(rs!FechaForp, "yyyy/mm/dd")
     itmX.SubItems(7) = Format(rs!fecha_calculo_int, "yyyy/mm/dd")
     
     'Diferido Total
     iDias = DateDiff("d", rs!FechaForp, rs!fecha_calculo_int) + 1
     itmX.SubItems(8) = iDias
     itmX.SubItems(9) = Format(((iDias * curMonto * rs!Int) / 36000), "Standard")
     
     itmX.SubItems(10) = Format(dtpCorte.Value, "yyyy/mm/dd")
     'Diferido al Corte
     vFecha = CDate(Year(dtpCorte.Value) & "/" & Month(dtpCorte.Value) & "/01")
     iDias = DateDiff("d", vFecha, Format(dtpCorte.Value, "yyyy/mm/dd")) + 1
     itmX.SubItems(11) = iDias
     itmX.SubItems(12) = Format(((iDias * curMonto * rs!Int) / 36000), "Standard")
     
     'Diferido Acumulado
     iDias = DateDiff("d", rs!FechaForp, Format(dtpCorte.Value, "yyyy/mm/dd")) + 1
     itmX.SubItems(13) = iDias
     itmX.SubItems(14) = Format(((iDias * curMonto * rs!Int) / 36000), "Standard")
     itmX.SubItems(15) = rs!Int
 
 
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
MsgBox "Consulta Finalizada...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, iDias As Integer
Dim vFecha As Date, fn

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select R.id_solicitud,R.codigo,R.cedula,S.nombre,R.Montoapr,R.MONTOCALCULO" _
       & ",R.fechaforp,R.fecha_calculo_int,R.int" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado in('A','C') and R.codigo = '" & txtCodigo & "' and R.fechaforp between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)
prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

fn = FreeFile

frmContenedor.CD.ShowSave

Open frmContenedor.CD.FileName For Output As #fn   ' Crea Archivo.

Print #fn, ",Sistema SIF: Intereses Diferidos, Usuario: " & glogon.Usuario _
           & ", Fecha : " & Format(fxFechaServidor, "dd/mm/yyyy") & ",Hora : " _
           & Format(Time, "hh:mm:ss AMPM")
           
Print #fn, ",Línea: " & txtCodigo & ", Inicio : " & Format(dtpInicio.Value, "dd/mm/yyyy") _
           & ",Corte : " & Format(dtpCorte.Value, "dd/mm/yyyy")
         

Print #fn, "Operacion,codigo,cedula,nombre,Monto Apr,Monto Base,Fec.Form,Fec.Final,DiasTotal,Dif.Total" _
          & ",Fec.Corte,DiasCorte,Dif.Corte,DiasAcumulados,Dif.Acumulado,Tasa"

Do While Not rs.EOF
 If IsNull(rs!MontoCalculo) Then
   curMonto = rs!montoapr
 Else
   curMonto = rs!MontoCalculo
 End If
 
 
 strSQL = rs!Id_Solicitud & "," & rs!Codigo & "," & rs!Cedula & "," & rs!Nombre _
        & "," & rs!montoapr & "," & curMonto & "," & Format(rs!FechaForp, "yyyy/mm/dd") _
        & "," & Format(rs!fecha_calculo_int, "yyyy/mm/dd")
     'Diferido Total
     iDias = DateDiff("d", rs!FechaForp, rs!fecha_calculo_int) + 1
 strSQL = strSQL & "," & iDias & "," & ((iDias * curMonto * rs!Int) / 36000) _
        & "," & Format(dtpCorte.Value, "yyyy/mm/dd")
     'Diferido al Corte
     vFecha = CDate(Year(dtpCorte.Value) & "/" & Month(dtpCorte.Value) & "/01")
     iDias = DateDiff("d", vFecha, Format(dtpCorte.Value, "yyyy/mm/dd")) + 1
     
 strSQL = strSQL & "," & iDias & "," & ((iDias * curMonto * rs!Int) / 36000)
     'Diferido Acumulado
     iDias = DateDiff("d", rs!FechaForp, Format(dtpCorte.Value, "yyyy/mm/dd")) + 1
 strSQL = strSQL & "," & iDias & "," & ((iDias * curMonto * rs!Int) / 36000) & "," & rs!Int
 
 Print #fn, strSQL
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

Close #fn   ' Cierra Archivo.

Me.MousePointer = vbDefault
MsgBox "Archivo Guardado como : " & frmContenedor.CD.FileName, vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical




End Sub

Private Sub Form_Load()

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

End Sub

Private Sub Form_Resize()
On Error Resume Next

Line1.X1 = Me.Width
Line2.X1 = Me.Width

lsw.Width = Me.Width - 120
lsw.Height = Me.Height - 1400

End Sub
