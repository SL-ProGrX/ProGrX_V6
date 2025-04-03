VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRH_Monitor_Gestiones 
   Caption         =   "RRHH: Monitor de Gestiones (Permisos, Vacaciones, Incapacidades)"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12015
   LinkTopic       =   "Form4"
   ScaleHeight     =   6585
   ScaleWidth      =   12015
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4572
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   10332
      _Version        =   1572864
      _ExtentX        =   18224
      _ExtentY        =   8064
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      FlatScrollBar   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   10080
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin MSComctlLib.Toolbar tlbAutorizacion 
      Height          =   264
      Left            =   10080
      TabIndex        =   1
      Top             =   1200
      Width           =   1608
      _ExtentX        =   2831
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Autorizar"
            Object.ToolTipText     =   "Autorizar Casos Marcados"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Desautorizar"
            Object.ToolTipText     =   "Desautorizar casos marcados"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reporte"
            Object.ToolTipText     =   "Exporta a Excel"
            ImageIndex      =   5
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Monitor_Gestiones.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Monitor_Gestiones.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Monitor_Gestiones.frx":0248
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Monitor_Gestiones.frx":036E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Monitor_Gestiones.frx":0477
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboAutorizado 
      Height          =   315
      Left            =   7680
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2760
      TabIndex        =   7
      Top             =   1200
      Width           =   2292
      _Version        =   1572864
      _ExtentX        =   4048
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
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   204
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   204
      _Version        =   1572864
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtAutorizadorId 
      Height          =   330
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   582
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Autorizador Id:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización de Solicitudes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2004
      TabIndex        =   6
      Top             =   360
      Width           =   6852
   End
   Begin XtremeShortcutBar.ShortcutCaption lblX 
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   10332
      _Version        =   1572864
      _ExtentX        =   18224
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "                                Seleccione las Solicitudes  a Autorizar o Desautorizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgBanner 
      Height          =   1125
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmRH_Monitor_Gestiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub cboAutorizado_Click()
If vPaso Then Exit Sub

Call sbBuscar

End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub

Call sbBuscar

End Sub

Private Sub chkTodos_Click()
Dim i As Long

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub Form_Activate()
vModulo = 23

End Sub

Private Sub Form_Load()

vModulo = 23

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

cboTipo.Clear
cboTipo.AddItem "Permisos"
cboTipo.AddItem "Vacaciones"
cboTipo.AddItem "Incapacidades"
cboTipo.Text = "Permisos"


cboAutorizado.Clear
cboAutorizado.AddItem "Solicitadas"
cboAutorizado.AddItem "Autorizadas"
cboAutorizado.AddItem "Denegadas"
cboAutorizado.Text = "Solicitadas"

vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

End Sub


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "Boleta Id", 1200
    .Add , , "Empleado Id", 1200
    .Add , , "Identificación", 1200
    .Add , , "Nombre", 3200
    .Add , , "Concepto", 2200
    .Add , , "Notas", 4000
    .Add , , "Usuario", 1600, vbCenter
    .Add , , "Fecha", 2100
End With

Select Case Mid(cboTipo.Text, 1, 1)
  Case "V"
        strSQL = "select * from vRH_Boleta_Vacaciones"
  
        lsw.ColumnHeaders.Add , , "Días", 900, vbCenter
        lsw.ColumnHeaders.Add , , "Salida", 1400, vbCenter
        lsw.ColumnHeaders.Add , , "Entrada", 1400, vbCenter
  
  Case "P"
        strSQL = "select * from vRH_Boleta_Permisos"
        lsw.ColumnHeaders.Add , , "Horas", 900, vbCenter
        lsw.ColumnHeaders.Add , , "Salida", 1800, vbCenter
        lsw.ColumnHeaders.Add , , "Entrada", 1800, vbCenter
  
  Case "I"
        strSQL = "select * from vRH_Boleta_Incapacidades"
        lsw.ColumnHeaders.Add , , "Días", 900, vbCenter
        lsw.ColumnHeaders.Add , , "Salida", 1400, vbCenter
        lsw.ColumnHeaders.Add , , "Entrada", 1400, vbCenter
End Select

strSQL = strSQL & " Where Estado = '" & Mid(cboAutorizado.Text, 1, 1) _
       & "' and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"


'Autorizacion vía Aplicacion
If txtAutorizadorId.Text = "" Then
    txtAutorizadorId.Text = "RH"
End If

'Autorizador> Empleados Autorizados
strSQL = strSQL & "  and dbo.fxRH_Autorizador_Valida(Empleado_ID,'" & txtAutorizadorId.Text & "') = 1"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Boleta_Id)
     itmX.SubItems(1) = rs!Empleado_ID
     itmX.SubItems(2) = rs!IDENTIFICACION
     itmX.SubItems(3) = rs!NOMBRE_COMPLETO
     itmX.SubItems(4) = rs!TipoDesc
     itmX.SubItems(5) = rs!Motivo & ""
     itmX.SubItems(6) = rs!Registro_Usuario & ""
     itmX.SubItems(7) = Format(rs!Registro_Fecha & "", "dd/mm/yyyy")
     
    Select Case Mid(cboTipo.Text, 1, 1)
      Case "V"
        itmX.SubItems(8) = rs!Dias_Disfrutados & ""
        itmX.SubItems(9) = Format(rs!Fecha_Salida & "", "dd/mm/yyyy")
        itmX.SubItems(10) = Format(rs!Fecha_Entrada & "", "dd/mm/yyyy")
      
      Case "I"
        itmX.SubItems(8) = rs!Dias & ""
        itmX.SubItems(9) = Format(rs!Fecha_Salida & "", "dd/mm/yyyy")
        itmX.SubItems(10) = Format(rs!Fecha_Entrada & "", "dd/mm/yyyy")
      
      Case "P"
        itmX.SubItems(8) = rs!Hrs_Total & ""
        itmX.SubItems(9) = Format(rs!Hora_Inicio & "", "dd/mm/yyyy hh:mm:ss")
        itmX.SubItems(10) = Format(rs!Hora_Corte & "", "dd/mm/yyyy hh:mm:ss")
    End Select
     
     
     Select Case rs!Estado
         Case "S"
         Case "A"
              itmX.Bold = True
              itmX.TextBackColor = RGB(252, 243, 207)
         Case "D"
              itmX.ForeColor = vbRed
              itmX.Bold = True
              itmX.TextBackColor = RGB(250, 219, 216)
     End Select

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAutoriza(pGestion As String)
Dim strSQL As String, i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

'spRH_Autorizaciones_Registro(@AutorizadorId varchar(20), @Tipo varchar(10), @BoletaId varchar(30), @Usuario varchar(30)
'                , @Estado char(1) = 'A', @AppCod varchar(30) = 'ProGrX' )

If pGestion = Mid(cboAutorizado.Text, 1, 1) Then Exit Sub


With lsw.ListItems
  For i = 1 To .Count
      If .Item(i).Checked Then
         strSQL = "exec spRH_Autorizaciones_Registro '" & txtAutorizadorId.Text & "','" & Mid(cboTipo.Text, 1, 1) _
                & "','" & .Item(i).Text & "','" & glogon.Usuario & "','" & pGestion & "','ProGrX'"
         Call ConectionExecute(strSQL)

         Call Bitacora("Aplica", IIf((pGestion = "A"), "Autoriza", "Deniega") & " de Boleta Id:" & .Item(i).Text _
                 & "..Empleado Id: " & .Item(i).SubItems(1) & "..Persona Id: " & .Item(i).SubItems(2))

      End If
  Next i
End With


Me.MousePointer = vbDefault
MsgBox IIf((pGestion = "A"), "Autorización", "Denegación") & " realizada satisfactoriamente.!", vbInformation

Call sbBuscar

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Resize()
On Error Resume Next


imgBanner.Width = Me.Width

lsw.Width = Me.Width - 450
lblX.Width = lsw.Width
lsw.Height = Me.Height - (lsw.Top + 880)


End Sub


Private Sub tlbAutorizacion_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbBuscar
  Case "Autorizar"
    Call sbAutoriza("A")
  Case "Desautorizar"
    Call sbAutoriza("D")
  Case "Reporte"
    Call sbReporte
End Select

End Sub

Private Sub sbReporte()

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

End Sub


Private Sub txtAutorizadorId_KeyDown(KeyCode As Integer, Shift As Integer)

   
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Col1Name = "Empleado Id"
   gBusquedas.Col2Name = "Persona Id"
   gBusquedas.Col3Name = "Nombre"
   gBusquedas.Columna = "Empleado_ID"
   gBusquedas.Orden = "Empleado_ID"
   gBusquedas.Consulta = "Select Empleado_ID,Identificacion,Nombre_Completo From Rh_Personas"
   frmBusquedas.Show vbModal
   
   txtAutorizadorId.Text = gBusquedas.Resultado
   txtAutorizadorId.ToolTipText = gBusquedas.Resultado3
    
   Call sbBuscar
End If

End Sub
