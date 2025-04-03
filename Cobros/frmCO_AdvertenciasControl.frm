VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCO_AdvertenciasControl 
   Caption         =   "Cobros: Control de Advertencias"
   ClientHeight    =   10272
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10272
   ScaleWidth      =   14400
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   7320
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   2292
      Left            =   0
      TabIndex        =   11
      Top             =   8040
      Width           =   3492
      _Version        =   1245187
      _ExtentX        =   6159
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "..."
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1680
         TabIndex        =   12
         Top             =   600
         Width           =   1692
         _Version        =   1245187
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
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
         Left            =   1680
         TabIndex        =   13
         Top             =   960
         Width           =   1692
         _Version        =   1245187
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   1692
         _Version        =   1245187
         _ExtentX        =   2985
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   732
         Left            =   1200
         TabIndex        =   18
         Top             =   1440
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCO_AdvertenciasControl.frx":0000
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   732
         Left            =   2280
         TabIndex        =   19
         Top             =   1440
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Exportar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCO_AdvertenciasControl.frx":0A1E
         TextImageRelation=   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Height          =   312
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
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
         Height          =   312
         Index           =   9
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
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
         Height          =   312
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1092
      End
      Begin VB.Image Image1 
         Height          =   10116
         Left            =   0
         Picture         =   "frmCO_AdvertenciasControl.frx":1223
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3612
      End
   End
   Begin XtremeSuiteControls.ListView lswAdvertencias 
      Height          =   3732
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   3252
      _Version        =   1245187
      _ExtentX        =   5736
      _ExtentY        =   6583
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkEstadoPersona 
      Height          =   216
      Left            =   3120
      TabIndex        =   6
      Top             =   4440
      Width           =   216
      _Version        =   1245187
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   10020
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   445
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
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   8172
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   10692
      _Version        =   524288
      _ExtentX        =   18860
      _ExtentY        =   14415
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      SpreadDesigner  =   "frmCO_AdvertenciasControl.frx":218D
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.CheckBox chkAdvertencias 
      Height          =   216
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   216
      _Version        =   1245187
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.ListView lswEstadoPersona 
      Height          =   3132
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   3252
      _Version        =   1245187
      _ExtentX        =   5736
      _ExtentY        =   5524
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   3720
      TabIndex        =   20
      Top             =   480
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   10440
      TabIndex        =   21
      Top             =   480
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   5520
      TabIndex        =   22
      Top             =   480
      Width           =   4932
      _Version        =   1245187
      _ExtentX        =   8700
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Advertencias:  "
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
      Height          =   312
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label lblEstadoPersona 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Persona: "
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
      Height          =   312
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   3012
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   1812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   7
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   1812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   8
      Left            =   10440
      TabIndex        =   2
      Top             =   240
      Width           =   1812
   End
   Begin VB.Image imgBanner 
      Height          =   9996
      Left            =   0
      Picture         =   "frmCO_AdvertenciasControl.frx":2AAC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3492
   End
End
Attribute VB_Name = "frmCO_AdvertenciasControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnBuscar_Click()

Call sbBuscar

End Sub

Private Sub btnExportar_Click()
Call sbExport_Excel
End Sub

Private Sub chkAdvertencias_Click()
Dim i As Integer

For i = 1 To lswAdvertencias.ListItems.Count
  lswAdvertencias.ListItems.Item(i).Checked = chkAdvertencias.Value
Next i

End Sub

Private Sub chkEstadoPersona_Click()
Dim i As Integer

For i = 1 To lswEstadoPersona.ListItems.Count
  lswEstadoPersona.ListItems.Item(i).Checked = chkEstadoPersona.Value
Next i

End Sub
Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


cboEstado.Clear
cboEstado.AddItem "Activa"
cboEstado.AddItem "Resuelta"
cboEstado.AddItem "Descartada"
cboEstado.AddItem "[TODOS]"
cboEstado.Text = "[TODOS]"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", (Day(dtpCorte.Value) - 1) * -1, dtpCorte.Value)


lswAdvertencias.ListItems.Clear
strSQL = "select Cod_Advertencia as IdX,   rtrim(Descripcion) as ItmX from CBR_ADVERTENCIAS_TIPO where Activa = 1 order by Cod_Advertencia"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswAdvertencias.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!idX
     itmX.Checked = chkAdvertencias.Value
 rs.MoveNext
Loop
rs.Close

lswEstadoPersona.ListItems.Clear
strSQL = "select cod_estado,descripcion from AFI_ESTADOS_PERSONA order by descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswEstadoPersona.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!cod_estado
     itmX.Checked = chkEstadoPersona.Value
 rs.MoveNext
Loop
rs.Close


End Sub

Private Sub Form_Activate()
 vModulo = 4
End Sub


Private Sub Form_Load()

vModulo = 4

vGrid.AppearanceStyle = fxGridStyle


With lswAdvertencias.ColumnHeaders
    .Clear
    .Add , , "Advertencias", lswAdvertencias.Width - 200
End With

With lswEstadoPersona.ColumnHeaders
    .Clear
    .Add , , "Estados", lswEstadoPersona.Width - 200
End With

cboEstado.Clear
cboEstado.AddItem "Activa"
cboEstado.AddItem "Resuelta"
cboEstado.AddItem "Descartada"
cboEstado.AddItem "[TODOS]"
cboEstado.Text = "[TODOS]"


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next


imgBanner.Height = Me.Height

vGrid.Width = Me.Width - (vGrid.Left + 250)
vGrid.Height = Me.Height - 1665

lswAdvertencias.Height = Me.Height / 4
lswEstadoPersona.Height = lswAdvertencias.Height



lblEstadoPersona.top = lswAdvertencias.top + lswAdvertencias.Height + 205
chkEstadoPersona.top = lblEstadoPersona.top

lswEstadoPersona.top = lblEstadoPersona.top + 360

gbFiltros.top = lswEstadoPersona.top + lswEstadoPersona.Height + 85

End Sub



Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0

strSQL = "select '', Adv.Linea, Adv.COD_ADVERTENCIA,rtrim(Adv.CEDULA), rtrim(Soc.NOMBRE), case when adv.ESTADO = 'A' then 'Activa' " _
       & "  when adv.ESTADO = 'R' then 'Resuelta' when adv.ESTADO = 'D' then 'Descargada'  else 'Activa' end" _
       & " , Tip.DESCRIPCION as 'AdvTipo', Adv.FECHA_VENCE, Adv.REGISTRO_FECHA , Adv.REGISTRO_USUARIO" _
       & " , Adv.RESOLUCION_FECHA , adv.RESOLUCION_USUARIO" _
       & "   from CBR_ADVERTENCIAS_CASOS Adv inner join Socios Soc On Adv.CEDULA = Soc.Cedula" _
       & "  inner join CBR_ADVERTENCIAS_TIPO Tip on Adv.COD_ADVERTENCIA = Tip.COD_ADVERTENCIA" _
       & "  inner join AFI_ESTADOS_PERSONA Est on Soc.ESTADOACTUAL = Est.COD_ESTADO" _
       & " Where Soc.Cedula like '%" & txtCedula.Text & "%'"

If Len(Trim(txtNombre.Text)) > 0 Then
   strSQL = strSQL & " and Soc.Nombre like '%" & txtNombre.Text & "%'"
End If

If Len(Trim(txtUsuario.Text)) > 0 Then
   strSQL = strSQL & " and Adv.Registro_Usuario like '%" & txtUsuario.Text & "%'"
End If


'Tipos de Advertencias
iCantidad = 0
For i = 1 To lswAdvertencias.ListItems.Count
  If lswAdvertencias.ListItems.Item(i).Checked Then
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad <> lswAdvertencias.ListItems.Count Then
    iCantidad = 0
    vCadena = " and Adv.Cod_Advertencia in('"
    For i = 1 To lswAdvertencias.ListItems.Count
      If lswAdvertencias.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswAdvertencias.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If


'Lista de Estados de la Persona
iCantidad = 0
For i = 1 To lswEstadoPersona.ListItems.Count
  If lswEstadoPersona.ListItems.Item(i).Checked Then
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad <> lswEstadoPersona.ListItems.Count Then
    iCantidad = 0
    vCadena = " and Soc.EstadoActual in('"
    For i = 1 To lswEstadoPersona.ListItems.Count
      If lswEstadoPersona.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswEstadoPersona.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If


Select Case cboEstado.Text
  Case "Activa"
    strSQL = strSQL & " and Adv.Estado = 'A' and Adv.Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Resuelta"
    strSQL = strSQL & " and Adv.Estado = 'R' and Adv.RESOLUCION_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Descartada"
    strSQL = strSQL & " and Adv.Estado = 'D' and Adv.RESOLUCION_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case Else
    strSQL = strSQL & " and Adv.Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  
End Select

Call sbCargaGridLocal(vGrid, 12, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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

Call OpenRecordSet(rs, strSQL)
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
  curMonto = 0 ' curMonto + rs!Monto
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbExport_Excel()

Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "No.Registro"
    vHeaders.Headers(3) = "Cod.Adv."
    vHeaders.Headers(4) = "Identificación"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Estado"
    vHeaders.Headers(7) = "Advertencia"
    vHeaders.Headers(8) = "Fec.Vence"
    vHeaders.Headers(9) = "Fec.Registro"
    vHeaders.Headers(10) = "User.Registro"
    vHeaders.Headers(11) = "Fec.Resolución"
    vHeaders.Headers(12) = "User.Resolución"
    
    Call sbSIFGridExportar(vGrid, vHeaders, "Cbr_Advertencias")

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frm As Form

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.col = 4
GLOBALES.gTag = vGrid.Text
Call sbFormsCall("frmCO_AdvertenciasRegistro", 1, , , False, Me)

End Sub


