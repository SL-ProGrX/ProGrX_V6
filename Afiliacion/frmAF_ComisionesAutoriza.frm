VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.controls.v19.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.shortcutbar.v19.2.0.ocx"
Begin VB.Form frmAF_ComisionesAutoriza 
   Caption         =   "Comisiones de Afiliación: Autorización de Afiliaciones"
   ClientHeight    =   8016
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10716
   Icon            =   "frmAF_ComisionesAutoriza.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8016
   ScaleWidth      =   10716
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4572
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   10332
      _Version        =   1245186
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
   Begin MSComctlLib.Toolbar tlbAutorizacion 
      Height          =   264
      Left            =   8160
      TabIndex        =   0
      Top             =   1200
      Width           =   1608
      _ExtentX        =   2836
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
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
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_ComisionesAutoriza.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_ComisionesAutoriza.frx":012A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_ComisionesAutoriza.frx":0254
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_ComisionesAutoriza.frx":037A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_ComisionesAutoriza.frx":0483
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   1332
      _Version        =   1245186
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
      Left            =   4440
      TabIndex        =   5
      Top             =   1200
      Width           =   1332
      _Version        =   1245186
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
      Height          =   312
      Left            =   5760
      TabIndex        =   6
      Top             =   1200
      Width           =   2292
      _Version        =   1245186
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtObservaciones 
      Height          =   492
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   8892
      _Version        =   1245186
      _ExtentX        =   15684
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodPromotor 
      Height          =   312
      Left            =   3120
      TabIndex        =   9
      Top             =   1560
      Width           =   732
      _Version        =   1245186
      _ExtentX        =   1291
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombrePromotor 
      Height          =   312
      Left            =   3840
      TabIndex        =   10
      Top             =   1560
      Width           =   4212
      _Version        =   1245186
      _ExtentX        =   7429
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   3120
      TabIndex        =   11
      Top             =   1920
      Width           =   1332
      _Version        =   1245186
      _ExtentX        =   2350
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkPromotor 
      Height          =   252
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   2292
      _Version        =   1245186
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos los Promotores"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   252
      Left            =   720
      TabIndex        =   13
      Top             =   1920
      Width           =   2292
      _Version        =   1245186
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos los Usuarios"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkAportes 
      Height          =   252
      Left            =   4560
      TabIndex        =   14
      Top             =   1920
      Width           =   3492
      _Version        =   1245186
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Revisa que tengan Aporte Registrado"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   200
      Left            =   240
      TabIndex        =   16
      Top             =   3080
      Width           =   200
      _Version        =   1245186
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
   Begin XtremeShortcutBar.ShortcutCaption lblX 
      Height          =   372
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   10332
      _Version        =   1245186
      _ExtentX        =   18224
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "                                Seleccione las Afiliaciones a Autorizar o Desautorizar"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización de Afiliaciones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2004
      TabIndex        =   2
      Top             =   360
      Width           =   6852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Afiliaciones Ingresadas entre "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Image imgBanner 
      Height          =   1125
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_ComisionesAutoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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

Private Sub txtCodPromotor_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   txtCodPromotor = ""
   txtNombrePromotor = ""
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "ID_PROMOTOR"
   gBusquedas.Orden = "ID_PROMOTOR"
   gBusquedas.Consulta = "select ID_PROMOTOR ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtCodPromotor = Trim(gBusquedas.Resultado)
   txtNombrePromotor = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtNombrePromotor_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   txtNombrePromotor.Text = ""
   txtCodPromotor.Text = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select ID_PROMOTOR ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   
   frmBusquedas.Show vbModal
   txtCodPromotor.Text = Trim(gBusquedas.Resultado)
   txtNombrePromotor.Text = Trim(gBusquedas.Resultado2)
End If

End Sub


Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1500
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3000
    .Add , , "Ingreso", 1600, vbCenter
    .Add , , "Promotor", 2500
    .Add , , "Usuario", 1600, vbCenter
    .Add , , "Fecha", 2100
    .Add , , "Notas", 4000
End With

cboAutorizado.Clear
cboAutorizado.AddItem "Pendientes"
cboAutorizado.AddItem "Autorizados"
cboAutorizado.AddItem "Desautorizados"
cboAutorizado.Text = "Pendientes"

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

strSQL = "select S.*,isnull(S.Comision_Autoriza,0) as AutorizacionX,P.Nombre as PromotorX" _
       & " FROM socios S inner join promotores P on S.id_promotor = P.id_promotor" _
       & "      WHERE S.estadoactual = 'S' And S.Fecha_Comision is Null" _
       & "        and S.fechaIngreso between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & "        and P.apl_comision = 1"
       
If chkAportes.Value = vbChecked Then
  strSQL = strSQL & "  and dbo.fxAFIComisionAporte(S.FechaIngreso, S.Cedula) > 0"
End If
       
If chkPromotor.Value = vbUnchecked Then
  strSQL = strSQL & "  and S.id_promotor = " & txtCodPromotor.Text
End If

If chkUsuarios.Value = vbUnchecked Then
  strSQL = strSQL & "  and S.reg_user = '" & txtUsuario.Text & "'"
End If

       
Select Case Mid(cboAutorizado.Text, 1, 1)
  Case "A"
    strSQL = strSQL & " and S.Comision_Autoriza = 1"
  Case "P"
    strSQL = strSQL & " and isnull(S.Comision_Autoriza ,0) = 0"
  Case "D"
    strSQL = strSQL & " and S.Comision_Autoriza  = 2"
End Select
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!id_Boleta_AF & "")
     itmX.SubItems(1) = rs!Cedula
     itmX.SubItems(2) = rs!Nombre
     itmX.SubItems(3) = Format(rs!FechaIngreso, "dd/mm/yyyy")
     itmX.SubItems(4) = rs!PromotorX & ""
     itmX.SubItems(5) = rs!reg_user & ""
     itmX.SubItems(6) = Format(rs!reg_fecha & "", "dd/mm/yyyy")
     itmX.SubItems(7) = IIf(IsNull(rs!autoriza_comision_notas), "", Trim(rs!autoriza_comision_notas))
     Select Case rs!AutorizacionX
         Case 0
'             itmX.SmallIcon = 3
'             itmX.TextBackColor = vbGreen
         Case 1
              itmX.Bold = True
              itmX.TextBackColor = RGB(252, 243, 207)
         Case 2
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


Private Sub sbAutoriza(pAutoriza As Integer)
Dim strSQL As String, i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass


With lsw.ListItems
  For i = 1 To .Count
      If .Item(i).Checked Then
         strSQL = "update socios set Comision_Autoriza = " & pAutoriza _
                & ",AUTORIZA_COMISION_NOTAS = '" & UCase(txtObservaciones) & "' where cedula = '" & .Item(i).SubItems(1) & "'"
         Call ConectionExecute(strSQL)
         
         Call Bitacora("Aplica", IIf((pAutoriza = 1), "Autorización", "Desautorización") & " de Afiliación..Ced.:" & .Item(i).SubItems(1))
         
      End If
  Next i
End With


Me.MousePointer = vbDefault
MsgBox IIf((pAutoriza = 1), "Autorización", "Desautorización") & " realizada satisfactoriamente.!", vbInformation
txtObservaciones = Empty
Call sbBuscar

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Resize()
On Error Resume Next


lsw.Width = Me.Width - 450
lblX.Width = lsw.Width
lsw.Height = Me.Height - (lsw.top + 880)
txtObservaciones.Width = Me.Width - txtObservaciones.Left - 450

End Sub


Private Sub tlbAutorizacion_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbBuscar
  Case "Autorizar"
    Call sbAutoriza(1)
  Case "Desautorizar"
    If Trim(txtObservaciones) <> Empty Or Len(Trim(txtObservaciones)) < 10 Then
       Call sbAutoriza(2)
    Else
       MsgBox "Necesita una observación valida para realizar la Desautoriaciones...", vbCritical
    End If
  Case "Reporte"
    Call sbReporte
End Select

End Sub

Private Sub sbReporte()

Call sbListViewExporFileTab(lsw)

End Sub



Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   txtNombrePromotor.Text = ""
   txtCodPromotor.Text = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Nombre,Descripcion from Usuarios"
   gBusquedas.Filtro = " and Estado = 'A'"
   
   frmBusquedas.Show vbModal
   txtUsuario.Text = Trim(gBusquedas.Resultado)
   
End If

End Sub
