VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmActivos_TrasladoAsientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traslado de asientos a Contabilidad"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6015
      Left            =   0
      TabIndex        =   4
      Top             =   1720
      Width           =   11295
      _Version        =   1441792
      _ExtentX        =   19923
      _ExtentY        =   10610
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      UseVisualStyle  =   0   'False
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   216
      Left            =   5760
      TabIndex        =   5
      Top             =   360
      Width           =   216
      _Version        =   1441792
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "CheckBox1"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   7830
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   1332
      _Version        =   1441792
      _ExtentX        =   2350
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   4320
      TabIndex        =   3
      Top             =   360
      Width           =   1332
      _Version        =   1441792
      _ExtentX        =   2350
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   7320
      TabIndex        =   6
      Top             =   240
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483633
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
      Picture         =   "frmActivos_TrasladoAsientos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdTrasladar 
      Height          =   492
      Left            =   8880
      TabIndex        =   7
      Top             =   240
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Trasladar"
      BackColor       =   -2147483633
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
      Picture         =   "frmActivos_TrasladoAsientos.frx":0700
   End
   Begin XtremeSuiteControls.CheckBox chkLsw 
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   1400
      Width           =   210
      _Version        =   1441792
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   11295
      _Version        =   1441792
      _ExtentX        =   19918
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Traslado de Asientos a Contabilidad"
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
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Asientos:"
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
      Height          =   252
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2532
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmActivos_TrasladoAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkLsw_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkLsw.Value
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

strSQL = "select Tipo_asiento,Num_asiento,COD_CONTABILIDAD,fecha_asiento,anio,mes,descripcion" _
       & " from Activos_Asientos where fecha_traslado is null"
If chkTodos.Value = vbUnchecked Then
  strSQL = strSQL & " and fecha_asiento between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
         & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
End If

lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Num_Asiento)
     itmX.SubItems(1) = rs!Tipo_Asiento
     itmX.SubItems(2) = Format(rs!fecha_Asiento, "yyyy/mm/dd")
     itmX.SubItems(3) = rs!Descripcion
     itmX.SubItems(4) = rs!Anio
     itmX.SubItems(5) = rs!Mes
     itmX.Tag = rs!COD_CONTABILIDAD
 itmX.Checked = chkLsw.Value
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub cmdTrasladar_Click()
Dim strSQL As String, i As Integer
Dim vPaso As Boolean, vTransac As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

prgBar.Visible = True
prgBar.Max = lsw.ListItems.Count + 2
prgBar.Value = 1

vPaso = False

With lsw.ListItems
 For i = 1 To .Count
  vTransac = False
  If .Item(i).Checked Then
  
    If fxgCntPeriodoValida(CDate(.Item(i).SubItems(2))) Then
        glogon.Conection.BeginTrans
         vTransac = True
        'Inserta Maestro
        strSQL = "insert into CntX_Asientos(COD_CONTABILIDAD,tipo_asiento,num_asiento,anio,mes,fecha_asiento" _
               & ",descripcion,balanceado,notas, Referencia, modulo,user_crea) (select COD_CONTABILIDAD,tipo_asiento" _
               & ",num_asiento,anio,mes,fecha_asiento,descripcion,'S',isnull(Notas,''), isnull(Referencia,''),  " & vModulo & ",user_crea" _
               & " from Activos_Asientos where COD_CONTABILIDAD = " & .Item(i).Tag & " and num_asiento = '" _
               & .Item(i).Text & "' and tipo_asiento = '" & .Item(i).SubItems(1) & "')"

        'Inserta Detalle
        strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(num_linea,COD_CONTABILIDAD,tipo_Asiento,num_asiento,cod_cuenta,documento" _
               & ",detalle,tipo_Cambio,monto_Debito,monto_credito,cod_unidad,cod_divisa,cod_centro_costo)" _
               & "(select num_linea,COD_CONTABILIDAD,tipo_asiento,num_asiento,cod_cuenta,documento" _
               & ",detalle,1,Monto_Debito, Monto_Credito,cod_unidad,cod_divisa,cod_centro_costo" _
               & " From Activos_Asientos_detalle" _
               & " where COD_CONTABILIDAD = " & .Item(i).Tag & " and num_asiento = '" _
               & .Item(i).Text & "' and tipo_asiento = '" & .Item(i).SubItems(1) & "')"
            
        strSQL = strSQL & Space(10) & "update Activos_Asientos set fecha_traslado = getdate(),user_traslada = '" _
               & glogon.Usuario & "' where COD_CONTABILIDAD = " & .Item(i).Tag & " and num_asiento = '" _
               & .Item(i).Text & "' and tipo_asiento = '" & .Item(i).SubItems(1) & "'"
        
        Call ConectionExecute(strSQL)
      
      glogon.Conection.CommitTrans
      vTransac = False
      
    Else
      vPaso = True
    End If
  End If
 Next i
End With
Me.MousePointer = vbDefault

prgBar.Value = 1
prgBar.Visible = False

'Actualiza listView
Call cmdBuscar_Click

If vPaso Then
  MsgBox "Existen Asientos que no se trasladan ya que el periodo fue cerrado contablemente..", vbExclamation
End If

MsgBox "Asientos Traslados Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 If vTransac Then glogon.Conection.RollbackTrans
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()
vModulo = 36

 Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

With lsw.ColumnHeaders
    .Add , , "Asiento", 2500
    .Add , , "Tipo", 1000, vbCenter
    .Add , , "Fecha", 2200
    .Add , , "Descripción", 4000
    .Add , , "Año", 1000, vbCenter
    .Add , , "Mes", 1000, vbCenter
End With


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub
