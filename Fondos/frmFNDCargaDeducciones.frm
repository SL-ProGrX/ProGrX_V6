VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFNDCargaDeducciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga y Aplica Deducciones de Planilla"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFNDCargaDeducciones.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Picture         =   "frmFNDCargaDeducciones.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo de la Planilla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   4575
      Begin VB.ComboBox cboPeriodo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   111345667
         CurrentDate     =   37418
      End
      Begin MSComCtl2.DTPicker dtpIngreso 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   111345667
         CurrentDate     =   37418
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo"
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
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio de Acreditación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Ingreso"
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
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.ComboBox cboOperadora 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   6495
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   6
      Top             =   4275
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblEstado 
      Height          =   855
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Carga de Deducciones "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblCodigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmFNDCargaDeducciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMes As Integer, vAnio As Integer, vScroll As Boolean

Function fxContrato(pTipo As String, pValor As String) As Long
Dim strSQL As String, rs As New ADODB.Recordset

'pTipo = O: Operacion Asociada, C: Cédula de identidad

fxContrato = 0

If pTipo = "O" Then
    strSQL = "Select Cod_Contrato from Fnd_Contratos Where Operacion = " & pValor _
           & " And cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " And Cod_plan = '" & Trim(txtCodigo.Text) & "' and estado <> 'L'"
Else
    strSQL = "Select Cod_Contrato from Fnd_Contratos Where Cedula='" & pValor & "'" _
           & " And cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " And Cod_plan = '" & Trim(txtCodigo.Text) & "' and estado <> 'L'"
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then fxContrato = rs!COD_CONTRATO
rs.Close

End Function


Private Sub sbCargaDeduccion()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer, curMonto As Currency
Dim vOperadora As Long, vPlan As String
Dim vProceso As String, vContrato As Long


On Error GoTo vError

Me.MousePointer = vbHourglass


lblEstado.Caption = "Cargando Informacion..."
lblEstado.Refresh

i = 0        'Casos No identificados
curMonto = 0 'Monto No identificado

vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vPlan = Trim(txtCodigo)
vProceso = Format(Trim(vAnio), "0000") & Format(Trim(vMes), "00")

strSQL = "Select C.id_solicitud,C.Amortiza,C.Fechas,R.Cedula,C.NCon" _
       & " From Creditos_dt C inner join Reg_creditos R on C.id_solicitud = R.id_solicitud" _
       & " Where C.Tcon = '1' And C.Ncon like '" & vProceso & "%' And C.Codigo='" & Trim(lblCodigo) _
       & "' and C.Estado = 'A' and C.fechas = '" & Format(dtpIngreso.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then prgBar.Max = rs.RecordCount
     
Do While Not rs.EOF
  lblEstado.Caption = "Procesando Registro : " & prgBar.Value & " de " & prgBar.Max
  lblEstado.Refresh
        
   'Busca Inicialmente el Contrato asociada a la operacion de retencion, caso contrario x Cedula
   vContrato = fxContrato("O", rs!Id_Solicitud)
        
'Codigo Eliminado porque cuando una persona tiene dos contratos estos se ligan a dos operaciones independientes
'Por lo tanto se tiene que respetar la operacion de deduccion original, en caso de no existir o estar liquidado no aplicar
' y pasar por inconsistencia.

'   If vContrato = 0 Then
'       vContrato = fxContrato("C", Trim(rs!Cedula))
'   End If
   
   If vContrato > 0 Then
       strSQL = "Update Fnd_contratos set Aportes = Aportes + " & rs!Amortiza _
              & "Where Cod_operadora=" & vOperadora & " and Cod_plan ='" _
              & vPlan & "' and Estado='A' and Cod_contrato = " & vContrato
       Call ConectionExecute(strSQL)
       
       strSQL = "Insert Fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato," _
              & "Monto,Fecha_proceso,Fecha,Tcon,Ncon,fecha_acredita) Values(" & vOperadora & ",'" _
              & vPlan & "'," & vContrato & "," & rs!Amortiza & "," & vProceso & ",'" _
              & Format(rs!fechas, "yyyy/mm/dd") & "','1','" & rs!nCon & "','" _
              & Format(dtpFecha, "yyyy/mm/dd") & "')"
       Call ConectionExecute(strSQL)
      
      
'        Call Bitacora("Registra", "Apl.Aho Proceso:" & vProceso & " Ope:" & vOperadora & " Plan:" & vPlan & " Cont:" & vContrato & " Monto:" & rs!Amortiza)
   Else
     'Indicar En Una Variable que casos no se estan procesando...
     'Porque no existe un contrato que lo respalde...
     i = i + 1
     curMonto = curMonto + rs!Amortiza
     
   End If
   rs.MoveNext
   prgBar.Value = prgBar.Value + 1
Loop
rs.Close

 Call Bitacora("Carga", "Ded.Pla Proceso:" & vProceso & " Ope:" & vOperadora & " Plan:" & vPlan)

'Asiento
'Pendiente de Ver Con Contabilidad
 
Me.MousePointer = vbDefault

If i > 0 Then
 MsgBox "Se Encontrarón " & i & " Casos No identificados en El fondo los cuales no se Aplicaron" _
      & " por no existir un contrato activo que lo respalde, por un monto de " & Format(curMonto, "Standard") _
      & " ******* VERIFICAR Y SACAR UN REPORTE DE ESTOS CASOS PARA RESPALDO *******", vbExclamation
End If
 
lblEstado.Caption = ""
prgBar.Value = 0

MsgBox "Proceso aplicado satisfactoriamente...", vbExclamation
       
Exit Sub

vError:
  lblEstado.Caption = ""
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
       
End Sub



Private Sub cboOperadora_Click()
 txtCodigo_LostFocus
End Sub


Private Sub cboOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub



Private Sub cmdCargar_Click()
Dim strSQL As String, rs As New ADODB.Recordset, curTotal As Currency


If Trim(txtCodigo) = "" Then Exit Sub


Me.MousePointer = vbHourglass

vAnio = Mid(cboPeriodo.Text, 1, 4)
vMes = Mid(cboPeriodo.Text, 6, 2)


strSQL = "Select isnull(Count(*),0) as Total from Fnd_contratos_detalle " _
       & " where Tcon = '1' and Ncon like '" & Format(Trim(vAnio), "0000") & Format(Trim(vMes), "00") _
       & "%' And cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " and cod_plan = '" & txtCodigo & "' and fecha = '" & Format(dtpIngreso, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)
  curTotal = rs!Total
rs.Close

If curTotal > 0 Then
   MsgBox "Este Proceso ya fue generado con anterioridad", vbExclamation
Else
   Call sbCargaDeduccion
End If


Me.MousePointer = vbDefault

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & "and deducir_planilla = 1 and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & "and deducir_planilla = 1 and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_Plan
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

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion
End Sub

Private Sub Form_Load()
Dim i As Integer, vFecha As Long

vModulo = 18 'Fondo de Inversion

Call sbgFNDCargaCombos(cboOperadora, "Operadoras")

dtpFecha.Value = fxFechaServidor
dtpIngreso.Value = dtpFecha.Value

cboPeriodo.Clear
vFecha = Year(dtpFecha.Value) & Format(Month(dtpFecha.Value), "00")
cboPeriodo.AddItem Format(vFecha, "####-##")
cboPeriodo.Text = Format(vFecha, "####-##")
For i = 1 To 12
  vFecha = fxFechaProcesoAnterior(vFecha)
  cboPeriodo.AddItem Format(vFecha, "####-##")
Next i


vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub vAnio_LostFocus()
If Val(vAnio) < 1000 Then
   vAnio = 1000
End If
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "and deducir_planilla = 1 And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
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


Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select Codigo_ase,Descripcion from Fnd_Planes where Cod_Operadora="
strSQL = strSQL & cboOperadora.ItemData(cboOperadora.ListIndex) & " And "
strSQL = strSQL & "Cod_Plan='" & Trim(txtCodigo) & "'"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       txtDescripcion = Trim(!Descripcion)
       lblCodigo = Trim(!codigo_ase)
    Else
       txtCodigo = ""
       txtDescripcion = ""
       lblCodigo = ""
    End If
 .Close
End With
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPeriodo.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "and deducir_planilla = 1 And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtCodigo.SetFocus
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub




