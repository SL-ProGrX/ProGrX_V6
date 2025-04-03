VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_RemesasRetenciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesamiento de REMESAS de Retenciones"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmCR_RemesasRetenciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8535
   Begin VB.Data DaoControl 
      Caption         =   "DaoControl"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   160
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3495
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Plazo"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cuota"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nombre"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Inconsistencia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Detalle Inc"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   ">> &Buscar"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtArchivo 
      Height          =   495
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton cmdRepInconsistencias 
      Caption         =   "&Reporte de Inconsistencias"
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblCasosInc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7200
      TabIndex        =   22
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblCasosApl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7200
      TabIndex        =   21
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblSaldoInc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4440
      TabIndex        =   20
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label lblSaldoApl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4440
      TabIndex        =   19
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblCuotaInc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2160
      TabIndex        =   18
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblCuotaApl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Casos >>"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   10
      Left            =   6360
      TabIndex        =   16
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Casos >>"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   9
      Left            =   6360
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldos >>"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   8
      Left            =   3600
      TabIndex        =   14
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldos >>"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   7
      Left            =   3600
      TabIndex        =   13
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuotas >>"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   1320
      TabIndex        =   12
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuotas >>"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Total Rechazar"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total a Aplicar"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "1er.Deducción"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label1 
      Caption         =   "Detalle a Procesar"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCR_RemesasRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset, lngOperacion As Long
Dim strCodigo As String, vFecha As Date, vFechaProceso As Long
Dim vPrimera As Long, vUltima As Long, vUsuario As String
Dim lng As Long, vDetalle As String

vFecha = fxFechaServidor
vFechaProceso = GLOBALES.glngFechaCR
vUltima = GLOBALES.glngFechaCR
vPrimera = txtAnio & Format(txtMes, "00")
vUsuario = glogon.Usuario

If vFechaProceso > vPrimera Then
   MsgBox "La fecha de la primer deducción es menor a la fecha de proceso actual...", vbExclamation
   Exit Sub
End If

If lsw.ListItems.Count = 0 Then Exit Sub

Me.MousePointer = vbHourglass

For lng = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(lng).SubItems(5) = 0 Then
   'Verificar si existe en la tabla de socios, de lo contrario abrirle un registro
   ' como no socio antes de aplicar la retencion
   strSQL = "select coalesce(count(*),0) as existe from socios where cedula = '" & lsw.ListItems.Item(lng).SubItems(3) & "'"
   rs.Open strSQL, glogon.Conection, adOpenStatic
   If rs!existe = 0 Then
        strSQL = "insert socios(id_promotor,cedula,nombre,estadoactual,fechaingreso) values(" _
               & "1,'" & Trim(lsw.ListItems.Item(lng).SubItems(3)) & "','" _
               & Mid(UCase(Trim(lsw.ListItems.Item(lng).SubItems(4))), 1, 30) & "','N','" _
               & Format(vFecha, "yyyy/mm/dd") & "')"
        glogon.Conection.Execute strSQL
        strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte) values('" & Trim(lsw.ListItems.Item(lng).SubItems(3)) & "',0,0)"
        glogon.Conection.Execute strSQL
   End If
   rs.Close
    
   'Insertar la operacion
   strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
          & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
          & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
          & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
          & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol) values('" & Trim(lsw.ListItems.Item(lng).Text) & "',6,'" _
          & Trim(lsw.ListItems.Item(lng).SubItems(3)) & "'," & CCur(lsw.ListItems.Item(lng).SubItems(2)) _
          & "," & CCur(lsw.ListItems.Item(lng).SubItems(2)) & ",0," & CCur(lsw.ListItems.Item(lng).SubItems(2)) & ",0,0," _
          & CCur(lsw.ListItems.Item(lng).SubItems(2)) & "," & CCur(lsw.ListItems.Item(lng).SubItems(2)) & ",0,0," _
          & lsw.ListItems.Item(lng).SubItems(1) & ",'" & vUsuario & "','" & vUsuario & "','" & vUsuario & "'," _
          & "'" & vUsuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
          & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
          & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','N'" _
          & ",'N','OT','',0,1,0,'PROCESO AUTOMATICO : REMESAS','A'," & vPrimera _
          & "," & vUltima & ",'F')"
   glogon.Conection.Execute strSQL
    
   vDetalle = "REMESA : COD." & lsw.ListItems.Item(lng) & " CED." & lsw.ListItems.Item(lng).SubItems(3) _
            & " CTA: " & CCur(lsw.ListItems.Item(lng).SubItems(2))
   Call Bitacora("Registra", vDetalle)
    
 End If
Next lng


Me.MousePointer = vbDefault
sbLimpiaDatos
MsgBox "Casos Procesados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, vArchivo As String
Dim vPasa As Boolean


On Error GoTo vError


With frmContenedor.CD
 .InitDir = "C:\"
 .ShowOpen
 
 If .FileName = "" Then
   MsgBox "Archivo no válido...", vbExclamation
   Exit Sub
 End If
 
 If UCase(Right(.FileName, 3)) <> "DBF" Then
   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
   Exit Sub
 End If

 txtArchivo = .FileName

End With

Me.MousePointer = vbHourglass

'Procesar Información en el Lsw, y Datos de Verificación
'CAMPOS: CEDULA,NOMBRE,CODIGO,PLAZO,CUOTA

DaoControl.RecordSource = Dir(txtArchivo, vbArchive)
DaoControl.DatabaseName = Mid(txtArchivo, 1, Len(txtArchivo) - (Len(DaoControl.RecordSource) + 1))
DaoControl.Refresh

vArchivo = txtArchivo

Call sbLimpiaDatos

txtArchivo = vArchivo

With DaoControl.Recordset
  Do While Not .EOF
     DoEvents
     If Not IsNull(!Codigo) Then
     
     Set itmX = lsw.ListItems.Add(, , !Codigo)
         itmX.SubItems(1) = !Plazo
         itmX.SubItems(2) = Format(!cuota, "Standard")
         itmX.SubItems(3) = !Cedula
         itmX.SubItems(4) = !Nombre            'fxNombre(!Cedula)
         itmX.SubItems(5) = 0
      
      vPasa = True
      strSQL = "select * from Catalogo where codigo = '" & !Codigo & "'"
      rs.Open strSQL, glogon.Conection, adOpenStatic
      If rs.EOF And rs.BOF Then
        vPasa = False
        itmX.SubItems(5) = 1
      Else
        If rs!retencion = "N" Then
          vPasa = False
          itmX.SubItems(5) = 2
        End If
      End If
      rs.Close
      
      If vPasa Then
        strSQL = "select coalesce(count(*),0) as Existe from reg_creditos" _
               & " where estado = 'A' and codigo = '" & !Codigo _
               & "' and cedula = '" & !Cedula & "'"
        rs.Open strSQL, glogon.Conection, adOpenStatic
        If rs!existe > 0 Then
          vPasa = False
          itmX.SubItems(5) = 3
        End If
        rs.Close
      End If
      
      Select Case itmX.SubItems(5)
        Case 0
          itmX.SubItems(6) = "NO HAY INCONSISTENCIA"
        Case 1
          itmX.SubItems(6) = "EL CODIGO NO EXISTE EN EL CATALOGO"
        Case 2
          itmX.SubItems(6) = "EL CODIGO NO ES UNA RETENCION"
        Case 3
          itmX.SubItems(6) = "YA EXISTE UNA OPERACION EN COBRO"
      End Select
      
      If itmX.SubItems(5) > 0 Then
        itmX.ForeColor = vbRed
        lblCuotaInc.Caption = lblCuotaInc.Caption + !cuota
        lblSaldoInc.Caption = lblSaldoInc.Caption + (!cuota * !Plazo)
        lblCasosInc.Caption = lblCasosInc.Caption + 1
      Else
        lblCuotaApl.Caption = lblCuotaApl.Caption + !cuota
        lblSaldoApl.Caption = lblSaldoApl.Caption + (!cuota * !Plazo)
        lblCasosApl.Caption = lblCasosApl.Caption + 1
      End If
      
      End If 'Null
     .MoveNext
  Loop
End With


lblCuotaApl.Caption = Format(lblCuotaApl.Caption, "Standard")
lblCuotaInc.Caption = Format(lblCuotaInc.Caption, "Standard")
lblSaldoApl.Caption = Format(lblSaldoApl.Caption, "Standard")
lblSaldoInc.Caption = Format(lblSaldoInc.Caption, "Standard")


Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente....", vbInformation

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
 sbLimpiaDatos
End Sub

Private Sub cmdRepInconsistencias_Click()
Dim vCadena As String, lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

vCadena = ""

For lng = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(lng).SubItems(5) > 0 Then
    vCadena = vCadena & vbCrLf & lsw.ListItems.Item(lng) & vbTab _
            & lsw.ListItems.Item(lng).SubItems(1) & vbTab _
            & lsw.ListItems.Item(lng).SubItems(2) & vbTab & vbTab _
            & lsw.ListItems.Item(lng).SubItems(3) & vbTab _
            & lsw.ListItems.Item(lng).SubItems(4) & vbTab _
            & lsw.ListItems.Item(lng).SubItems(5) & vbTab _
            & lsw.ListItems.Item(lng).SubItems(6)
 End If
Next lng

With Printer
 Printer.Print vCadena
 .NewPage
 .EndDoc
End With

Me.MousePointer = vbDefault
MsgBox "Inconsistencias imprimiendose ...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical


End Sub

Private Sub Form_Load()
vModulo = 3
Call Formularios(Me)
Call RefrescaTags(Me)
sbLimpiaDatos

End Sub


Private Function fxPrimerDeduccion() As Long
Dim strMes As String, intAnio As Integer

strMes = Mid(Trim(str(GLOBALES.glngFechaCR)), 5, 2)
intAnio = Mid(Trim(str(GLOBALES.glngFechaCR)), 1, 4)
If Val(strMes) = 12 Then
  strMes = "01"
  intAnio = intAnio + 1
Else
  strMes = Format(Val(strMes) + 1, "00")
End If

fxPrimerDeduccion = intAnio & strMes
End Function


Private Sub sbLimpiaDatos()
lsw.ListItems.Clear
txtArchivo = ""
lblCasosApl.Caption = "0"
lblCasosInc.Caption = "0"
lblCuotaApl.Caption = "0"
lblCuotaInc.Caption = "0"
lblSaldoApl.Caption = "0"
lblSaldoInc.Caption = "0"

txtAnio = Mid(fxPrimerDeduccion, 1, 4)
txtMes = Mid(fxPrimerDeduccion, 5, 2)

End Sub
