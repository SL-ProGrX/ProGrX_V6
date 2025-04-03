VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIVR_Traslado_Asientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI: Traslado de Asientos"
   ClientHeight    =   7095
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5052
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10572
      _Version        =   1441793
      _ExtentX        =   18648
      _ExtentY        =   8911
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
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   6960
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox chkLsw 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Marcar / Desmarcar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   245
      Left            =   360
      TabIndex        =   1
      Top             =   1368
      Width           =   2172
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   1332
      _Version        =   1441793
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
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   1332
      _Version        =   1441793
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
      TabIndex        =   5
      Top             =   360
      Width           =   1572
      _Version        =   1441793
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
      Appearance      =   16
      Picture         =   "frmIVR_Traslado_Asientos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdTrasladar 
      Height          =   492
      Left            =   8880
      TabIndex        =   6
      Top             =   360
      Width           =   1572
      _Version        =   1441793
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
      Appearance      =   16
      Picture         =   "frmIVR_Traslado_Asientos.frx":0A1E
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   216
      Left            =   6240
      TabIndex        =   7
      Top             =   480
      Width           =   216
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "CheckBox1"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
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
      Left            =   3480
      TabIndex        =   8
      Top             =   240
      Width           =   2532
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Asientos Localizados para Traslado al Sistema de Contabilidad   "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   10572
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmIVR_Traslado_Asientos"
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

strSQL = "select COD_CONTABILIDAD,Tipo_asiento,Num_asiento,Fecha, Referencia, Notas" _
       & " from IVR_ASIENTOS where Traslado_Fecha is null"
If chkTodos.Value = vbUnchecked Then
  strSQL = strSQL & " and FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
         & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!NUM_ASIENTO)
     itmX.SubItems(1) = rs!TIPO_ASIENTO
     itmX.SubItems(2) = Format(rs!Fecha, "yyyy/mm/dd")
     itmX.SubItems(3) = rs!REFERENCIA
     itmX.SubItems(4) = rs!NOTAS
     itmX.SubItems(5) = ""
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

PrgBar.Visible = True
PrgBar.Max = lsw.ListItems.Count + 2
PrgBar.Value = 1

vPaso = False

With lsw.ListItems
 For i = 1 To .Count
  vTransac = False
  If .Item(i).Checked Then
  
  
'    .Add , , "Asiento", 2500
'    .Add , , "Tipo", 1000, vbCenter
'    .Add , , "Fecha", 2200
'    .Add , , "Descripción", 4000
'    .Add , , "Referencia", 1000, vbCenter
'    .Add , , "Notas", 2000, vbCenter
  
    If fxgCntPeriodoValida(CDate(.Item(i).SubItems(2))) Then
        glogon.Conection.BeginTrans
         vTransac = True
        'Inserta Maestro
        strSQL = "insert into CntX_Asientos(COD_CONTABILIDAD,tipo_asiento,num_asiento,anio,mes,fecha_asiento" _
               & ",descripcion,balanceado,notas,modulo,user_crea, referencia) (select COD_CONTABILIDAD,tipo_asiento" _
               & ",num_asiento,year(fecha),month(fecha) ,fecha, substring(Num_asiento + '...' + Referencia,1,100) ,'S',Notas," & vModulo & ", registro_Usuario, referencia" _
               & " from IVR_ASIENTOS where COD_CONTABILIDAD = " & .Item(i).Tag & " and num_asiento = '" _
               & .Item(i).Text & "' and tipo_asiento = '" & .Item(i).SubItems(1) & "')"

        'Inserta Detalle
        strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(num_linea,COD_CONTABILIDAD,tipo_Asiento,num_asiento,cod_cuenta,documento" _
               & ",detalle,tipo_Cambio,monto_Debito,monto_credito,cod_unidad,cod_divisa,cod_centro_costo)" _
               & "(select Linea_Id, COD_CONTABILIDAD, tipo_asiento, num_asiento, cod_cuenta, substring(documento,1,35)" _
               & ", substring(detalle,1,100), isnull(tipo_Cambio,1)" _
               & ", case When Movimiento = 'D' Then Monto else 0 end" _
               & ", case when Movimiento = 'C' Then Monto else 0 end" _
               & ", cod_unidad, cod_divisa, cod_centro_costo" _
               & " From IVR_ASIENTOS_detalle" _
               & " where COD_CONTABILIDAD = " & .Item(i).Tag & " and num_asiento = '" _
               & .Item(i).Text & "' and tipo_asiento = '" & .Item(i).SubItems(1) & "')"
            
        strSQL = strSQL & Space(10) & "update IVR_ASIENTOS set Traslado_Fecha = getdate(), Traslado_Usuario = '" _
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

PrgBar.Value = 1
PrgBar.Visible = False

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
    .Add , , "Referencia", 4000, vbCenter
    .Add , , "Notas", 6000, vbCenter
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


