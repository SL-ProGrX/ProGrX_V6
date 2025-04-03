VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmInsConsulta 
   Caption         =   "Consulta General de Pólizas"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14220
   Icon            =   "frmInsConsulta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
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
      Left            =   11880
      TabIndex        =   17
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
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
      Left            =   6960
      TabIndex        =   16
      Top             =   360
      Width           =   4935
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
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
      Left            =   5160
      TabIndex        =   15
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtPoliza 
      Appearance      =   0  'Flat
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
      Left            =   3360
      TabIndex        =   14
      Top             =   360
      Width           =   1815
   End
   Begin VB.ComboBox cboEstado 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CheckBox chkTiposSeguros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkVendedores 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2040
      TabIndex        =   0
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin MSComctlLib.ListView lswTipoSeguros 
      Height          =   3155
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5556
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5539
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   9015
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Casos Encontrados..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Registrado..:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   8280
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   200540161
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   8640
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   200540161
      CurrentDate     =   37678
   End
   Begin MSComctlLib.ListView lswVendedores 
      Height          =   3155
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5556
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5539
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   8175
      Left            =   3360
      TabIndex        =   13
      Top             =   720
      Width           =   10695
      _Version        =   524288
      _ExtentX        =   18865
      _ExtentY        =   14420
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
      SpreadDesigner  =   "frmInsConsulta.frx":6852
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   360
      Left            =   360
      TabIndex        =   22
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar"
            Key             =   "Exportar"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Excel"
                  Text            =   "Microsoft Excel"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HTML"
                  Text            =   "HTML"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   -120
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
            Picture         =   "frmInsConsulta.frx":73CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsConsulta.frx":DC31
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsConsulta.frx":14493
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsConsulta.frx":1ACF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsConsulta.frx":1AE2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   8
      Left            =   11880
      TabIndex        =   21
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   6960
      TabIndex        =   20
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   5160
      TabIndex        =   19
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Póliza"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   3360
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblInicio 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label lblCorte 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Seguros ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblVendedores 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedores"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   1215
   End
End
Attribute VB_Name = "frmInsConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub chkTiposSeguros_Click()
Dim i As Integer

For i = 1 To lswTipoSeguros.ListItems.Count
  lswTipoSeguros.ListItems.Item(i).Checked = chkTiposSeguros.Value
Next i

End Sub

Private Sub chkVendedores_Click()
Dim i As Integer

For i = 1 To lswVendedores.ListItems.Count
  lswVendedores.ListItems.Item(i).Checked = chkVendedores.Value
Next i

End Sub

Private Sub Form_Activate()
 vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

vModulo = 17


Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.AppearanceStyle = fxGridStyle

cboEstado.Clear
cboEstado.AddItem "Pendientes"
cboEstado.AddItem "Activadas"
cboEstado.AddItem "Cerradas"
cboEstado.AddItem "[TODOS]"
cboEstado.Text = "[TODOS]"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = CDate(Year(dtpCorte.Value) & "/" & Format(Month(dtpCorte.Value), "00") & "/01")


lswTipoSeguros.ListItems.Clear
strSQL = "select Tipo_Seguro as IdX,  rtrim(Tipo_Seguro) + ' - ' + rtrim(Descripcion) as ItmX from ins_Tipos_Seguros where Activo = 1 order by Tipo_Seguro"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lswTipoSeguros.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkTiposSeguros.Value
 rs.MoveNext
Loop
rs.Close

lswVendedores.ListItems.Clear
strSQL = "select cod_vendedor,Nombre from ins_Vendedores order by nombre"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lswVendedores.ListItems.Add(, , rs!Nombre)
     itmX.Tag = rs!cod_vendedor
     itmX.Checked = chkVendedores.Value
 rs.MoveNext
Loop
rs.Close

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 3765
vGrid.Height = Me.Height - 1665

lswTipoSeguros.Height = Me.Height / 3.1188589540412   '- 6685
lswVendedores.Height = lswTipoSeguros.Height

lblVendedores.Top = lswTipoSeguros.Top + lswTipoSeguros.Height + 205
chkVendedores.Top = lblVendedores.Top

lswVendedores.Top = lblVendedores.Top + 360

cboEstado.Top = lswVendedores.Top + lswVendedores.Height + 85
lblEstado.Top = cboEstado.Top

lblInicio.Top = lblEstado.Top + 360
dtpInicio.Top = lblInicio.Top

lblCorte.Top = lblInicio.Top + 360
dtpCorte.Top = lblCorte.Top


End Sub



Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0


strSQL = "select '', NUM_POLIZA,rtrim(CEDULA),rtrim(NOMBRE),Estado_Desc" _
       & " ,MONTO,REGISTRO_FECHA,ACTIVA_FECHA,CIERRA_FECHA , PAGADO_TOTAL,COBRADO_TOTAL , Balanza_Cobro" _
       & " ,comision_Vendedor_Total,Comision_Interna_Total,isnull(Operacion,0), TIPO_SEGURO, Vendedor_NOMBRE" _
       & "  from  vInsListadoGeneral " _
       & " where Num_Poliza like '%" & txtPoliza.Text & "%'"

If Len(Trim(txtCedula.Text)) > 0 Then
   strSQL = strSQL & " and Cedula like '%" & txtCedula.Text & "%'"
End If

If Len(Trim(txtNombre.Text)) > 0 Then
   strSQL = strSQL & " and Nombre like '%" & txtNombre.Text & "%'"
End If

If Len(Trim(txtUsuario.Text)) > 0 Then
   strSQL = strSQL & " and Registro_Usuario like '%" & txtUsuario.Text & "%'"
End If


'Tipos de Seguros
iCantidad = 0
For i = 1 To lswTipoSeguros.ListItems.Count
  If lswTipoSeguros.ListItems.Item(i).Checked Then
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad <> lswTipoSeguros.ListItems.Count Then
    iCantidad = 0
    vCadena = " and Tipo_Seguro in('"
    For i = 1 To lswTipoSeguros.ListItems.Count
      If lswTipoSeguros.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswTipoSeguros.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If


'Lista de Vendedores
iCantidad = 0
For i = 1 To lswVendedores.ListItems.Count
  If lswVendedores.ListItems.Item(i).Checked Then
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad <> lswVendedores.ListItems.Count Then
    iCantidad = 0
    vCadena = " and Cod_Vendedor in(0"
    For i = 1 To lswVendedores.ListItems.Count
      If lswVendedores.ListItems.Item(i).Checked Then
        vCadena = vCadena & "," & lswVendedores.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & ")"
End If


Select Case cboEstado.Text
  Case "Pendientes"
    strSQL = strSQL & " and Estado = 'P' and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Activadas"
    strSQL = strSQL & " and Estado = 'A' and Activa_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Cerradas"
    strSQL = strSQL & " and Estado = 'C' and Cierra_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case Else
    strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  
End Select

Call sbCargaGridLocal(vGrid, 17, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim curMonto As Currency

On Error GoTo vError

vPaso = True

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.col = i
 vGrid.Text = ""
Next i

curMonto = 0

rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i

    If rs.Fields(i - 1).Type = 135 Then
        If Year(rs.Fields(i - 1).Value) > 1900 Then
           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
        End If
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
    End If
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!MONTO
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbBuscar

  Case Else
End Select

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 17
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "No.Póliza"
    vHeaders.Headers(3) = "Cédula"
    vHeaders.Headers(4) = "Nombre"
    vHeaders.Headers(5) = "Estado"
    vHeaders.Headers(6) = "Monto"
    vHeaders.Headers(7) = "Fec.Registro"
    vHeaders.Headers(8) = "Fec.Activación"
    vHeaders.Headers(9) = "Fec.Cierre"
    vHeaders.Headers(10) = "Total Pagado"
    vHeaders.Headers(11) = "Total Cobrado"
    vHeaders.Headers(12) = "Balanza Cobraza"
    vHeaders.Headers(13) = "Comisión Vendedor"
    vHeaders.Headers(14) = "Comisión Interna"
    vHeaders.Headers(15) = "No. Operación"
    vHeaders.Headers(16) = "Tipo Seguro"
    vHeaders.Headers(17) = "Vendedor"
    
Select Case ButtonMenu.Key
  Case "Excel"
      Call sbSIFGridExportar(vGrid, vHeaders, "Ins_ConsultaPolizas")
  Case "HTML"
      Call sbSIFGridExportar(vGrid, vHeaders, "Ins_ConsultaPolizas", "HTML")
End Select
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frm As Form

If vPaso Then Exit Sub

Call sbSIFForms("frmInsRegistro")

For Each frm In Forms
  If UCase(frm.Name) = UCase("frmInsRegistro") Then
    vGrid.Row = Row
    vGrid.col = 2
    Call frm.sbConsultaExterna(vGrid.Text)
    Exit For
  End If
Next frm
End Sub
