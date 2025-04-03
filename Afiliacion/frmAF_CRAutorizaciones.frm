VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_CRAutorizaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaciones de Renuncias"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   10590
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4572
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   10332
      _Version        =   1441793
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
      Left            =   7080
      TabIndex        =   2
      Top             =   1320
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
            Key             =   "Reportes"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   0
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
            Picture         =   "frmAF_CRAutorizaciones.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CRAutorizaciones.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CRAutorizaciones.frx":0248
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CRAutorizaciones.frx":036E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CRAutorizaciones.frx":0477
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1332
      _Version        =   1441793
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
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1332
      _Version        =   1441793
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
      Left            =   4680
      TabIndex        =   5
      Top             =   1320
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtObservaciones 
      Height          =   492
      Left            =   2040
      TabIndex        =   6
      Top             =   1800
      Width           =   8292
      _Version        =   1441793
      _ExtentX        =   14626
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   204
      Left            =   240
      TabIndex        =   9
      Top             =   2484
      Width           =   204
      _Version        =   1441793
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
      TabIndex        =   10
      Top             =   2400
      Width           =   10332
      _Version        =   1441793
      _ExtentX        =   18224
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "                                Seleccione las Renuncias a Autorizar o Desautorizar"
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
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1332
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Renuncias Rescatadas "
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
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización de Renuncias"
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
      Left            =   1752
      TabIndex        =   0
      Top             =   360
      Width           =   6732
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_CRAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cboAutorizado_Click()
lsw.ListItems.Clear
End Sub

Private Sub chkTodos_Click()
Dim i As Long

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

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
    .Add , , "Vence", 1600, vbCenter
    .Add , , "Tipo", 1250, vbCenter
    .Add , , "Usuario", 1600
    .Add , , "Resuelto?", 1500, vbCenter
    .Add , , "Estado", 1500, vbCenter
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

strSQL = "select R.*,S.Nombre,isnull(R.autorizado_estado,0) as AutorizacionX" _
       & " from afi_cr_renuncias R inner join  Socios S on R.cedula = S.cedula" _
       & " where R.resuelto_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and Estado = 'R'"
       
Select Case Mid(cboAutorizado.Text, 1, 1)
  Case "A"
    strSQL = strSQL & " and R.autorizado_estado = 1"
  Case "P"
    strSQL = strSQL & " and isnull(R.autorizado_estado,0) = 0"
  Case "D"
    strSQL = strSQL & " and R.autorizado_estado = 2"
End Select
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Renuncia)
     itmX.SubItems(1) = rs!Cedula
     itmX.SubItems(2) = rs!Nombre
     itmX.SubItems(3) = Format(rs!Vencimiento, "dd/mm/yyyy")
     itmX.SubItems(4) = IIf((rs!Tipo = "A"), "Ren.Interna", "Ren.Patronal")
     itmX.SubItems(5) = rs!resuelto_user
     itmX.SubItems(6) = Format(rs!resuelto_fecha, "dd/mm/yyyy")
     
     Select Case rs!Estado
       Case "R"
             itmX.SubItems(7) = "Rescatada"
       Case "P"
             itmX.SubItems(7) = "Perdida"
       Case "T"
             itmX.SubItems(7) = "Transito"
       Case "V"
             itmX.SubItems(7) = "Vencida"
     End Select
     
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

      itmX.SubItems(8) = IIf(IsNull(rs!Autoriza_Notas), "", rs!Autoriza_Notas)
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
         strSQL = "update afi_cr_renuncias set autoriza_notas = '" & UCase(txtObservaciones) & "',Autorizado_Estado = " & pAutoriza _
                & ", Autorizado_Fecha = dbo.MyGetdate(), Autorizado_Usuario = '" & glogon.Usuario _
                & "' where cod_renuncia = " & .Item(i).Text
         Call ConectionExecute(strSQL)
      End If
  Next i
End With


Me.MousePointer = vbDefault
MsgBox IIf((pAutoriza = 1), "Autorización", "Desautorización") & " realizada satisfactoriamente.!", vbInformation
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



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
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
  Case "Reportes"
    Call sbReporte
End Select

End Sub

Private Sub sbReporte()

Call sbListViewExporFileTab(lsw)

End Sub


