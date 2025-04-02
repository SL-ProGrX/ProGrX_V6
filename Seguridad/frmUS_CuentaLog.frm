VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmUS_CuentaLog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitácora de Transacciones a Cuentas de Usuario"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12870
   Icon            =   "frmUS_CuentaLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5175
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   9128
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
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   18
      Top             =   0
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmUS_CuentaLog.frx":6852
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkTodas 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todas"
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
   End
   Begin VB.Frame fraRevision 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   -40
      Width           =   6375
      Begin XtremeSuiteControls.ComboBox cboRevision 
         Height          =   330
         Left            =   1200
         TabIndex        =   6
         Top             =   120
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkRevision 
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   120
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Buscar Usuario/Fecha Revisión"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Revisión:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_CuentaLog.frx":6F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_CuentaLog.frx":D7B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_CuentaLog.frx":14016
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7455
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   9735
      _Version        =   524288
      _ExtentX        =   17171
      _ExtentY        =   13150
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
      MaxCols         =   13
      SpreadDesigner  =   "frmUS_CuentaLog.frx":1A878
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
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
      Height          =   330
      Left            =   2160
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   7200
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1440
      TabIndex        =   14
      Top             =   6480
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAppName 
      Height          =   330
      Left            =   1440
      TabIndex        =   15
      Top             =   6840
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstacion 
      Height          =   330
      Left            =   1440
      TabIndex        =   16
      Top             =   7560
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAppVersion 
      Height          =   330
      Left            =   2160
      TabIndex        =   17
      Top             =   7200
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   19
      Top             =   0
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Informe"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmUS_CuentaLog.frx":1B102
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmUS_CuentaLog.frx":1B809
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   7560
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Estación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   7200
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Versión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Aplicación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fechas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUS_CuentaLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnBarra_Click(Index As Integer)
Select Case Index
  Case 0 'Buscar
    Call sbBuscar
  
  Case 1 'Reporte "Reporte"
        vGrid.PrintHeader = "Seguridad: Bitácora de Cuentas de Usuario, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Fechas Rastreo...I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
  Case 2 'Exportar
        Dim vHeaders As vGridHeaders
            vHeaders.Columnas = 13
            vHeaders.Headers(1) = "Revisado?"
            vHeaders.Headers(2) = "Fecha"
            vHeaders.Headers(3) = "Transacción"
            vHeaders.Headers(4) = "Usuario"
            vHeaders.Headers(5) = "Nombre"
            vHeaders.Headers(6) = "Notas"
            vHeaders.Headers(7) = "Aplicación"
            vHeaders.Headers(8) = "Versión"
            vHeaders.Headers(9) = "Estación"
            vHeaders.Headers(10) = "MAC Address"
            vHeaders.Headers(11) = "Gestionado por"
            vHeaders.Headers(12) = "Revisado por"
            vHeaders.Headers(13) = "Revisado Fecha"
      
      Call sbGridExportar(vGrid, vHeaders, "Seguridad_BitacoraCuenta")

End Select

End Sub


Private Sub cboRevision_Click()
If cboRevision.ListCount = 0 Then Exit Sub
Call sbBuscar
End Sub

Private Sub chkRevision_Click()
If chkRevision.Value = vbChecked Then
   txtUsuario.BackColor = cboRevision.BackColor
Else
   txtUsuario.BackColor = vbWhite
End If
End Sub

Private Sub chkTodas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodas.Value
Next i

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vGrid.MaxRows = 0

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -3, dtpCorte.Value)


strSQL = "select cod_transac,descripcion from us_transacciones order by cod_transac"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!cod_Transac
      itmX.Checked = chkTodas.Value
  rs.MoveNext
Loop
rs.Close

With lsw.ColumnHeaders
    .Clear
    .Add , , "Transacción", 4440.18
End With

cbo.Clear
cbo.AddItem "="
cbo.AddItem ">"
cbo.AddItem "<"
cbo.Text = "="

vPaso = True
cboRevision.AddItem "TODOS"
cboRevision.AddItem "Pendientes"
cboRevision.AddItem "Revisados"
cboRevision.Text = "TODOS"
vPaso = False


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbBuscar()
Dim rs As New ADODB.Recordset, strSQL As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass



'Actualiza Tabla Pivote para Reportes/Consulas
With lsw.ListItems
    strSQL = "delete us_transac_pivot where Usuario = '" & glogon.Usuario & "'"
    For i = 1 To .Count
       If .Item(i).Checked Then
          strSQL = strSQL & Space(10) & "insert us_transac_pivot(Usuario,Cod_Transac) values('" & glogon.Usuario & "','" & .Item(i).Tag & "')"
       End If
    Next i
   'Procesa Lote
   Call ConectionExecute(strSQL)

End With



strSQL = "select case when L.revisado_fecha is null then 0 else 1 end as 'Revisado'" _
       & ",L.Mov_Fecha,T.descripcion,L.Usuario,U.Nombre,L.notas,L.App_Name,L.App_Version,L.Equipo,L.Mov_User" _
       & ", L.Revisado_Usuario, L.Revisado_Fecha,L.Seq_Id,L.Cod_Transac,L.Equipo_MAC" _
       & " from us_transac_log L inner join us_transacciones T on L.cod_transac = T.cod_transac" _
       & " inner join us_transac_pivot P on T.cod_transac = P.cod_transac and P.Usuario = '" & glogon.Usuario & "'" _
       & " inner join us_Usuarios U on L.Usuario = U.Usuario"
       
'Fechas del Movimiento
If chkRevision.Value = vbChecked Then
     strSQL = strSQL & " Where L.Revisado_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
            & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
Else
     strSQL = strSQL & " Where L.mov_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
            & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
End If
       


If Len(txtEstacion.Text) > 0 Then
  strSQL = strSQL & " and L.Equipo = '" & txtEstacion.Text & "'"
End If
       
If Len(txtAppName.Text) > 0 Then
  strSQL = strSQL & " and L.App_Name = '" & txtAppName.Text & "'"
End If
       
If Len(txtAppVersion.Text) > 0 Then
  strSQL = strSQL & " and L.App_Version " & cbo.Text & "'" & txtAppVersion.Text & "'"
End If
       
      
'Usuario que Realiza el Movimiento
If Len(Trim(txtUsuario)) > 0 Then
     If chkRevision.Value = vbChecked Then
             strSQL = strSQL & " and L.Revisado_Usuario = '" & txtUsuario & "'"
     Else
             strSQL = strSQL & " and L.Usuario = '" & txtUsuario & "'"
     End If
End If


Select Case Mid(cboRevision.Text, 1, 1)
   Case "P" 'Pendientes
        strSQL = strSQL & " and L.Revisado_Fecha is null"
   Case "R" 'Revisados
        strSQL = strSQL & " and L.Revisado_Fecha is not null"
   Case "T" 'Todos
End Select



'Filtros de Seguridad
If Not gAdminAccess.Rol_AdminView Then
    gBusquedas.Filtro = " AND isnull(U.key_admin,0) = 0"
End If

If Not gAdminAccess.Rol_DirGlobal Then
    gBusquedas.Filtro = gBusquedas.Filtro & " AND U.usuario in(select usuario from PGX_CLIENTES_USERS" _
        & " Where cod_Empresa = " & gPortal.Empresa_Id & ")"
End If


If chkRevision.Value = vbChecked Then
    strSQL = strSQL & " order by L.Revisado_fecha"
Else
    strSQL = strSQL & " order by L.Mov_fecha"
End If



vPaso = True
vGrid.MaxRows = 0

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  
  vGrid.Col = 1
  vGrid.Text = rs!Revisado
  vGrid.CellTag = rs!Seq_Id
  
  vGrid.Col = 2
  vGrid.Text = rs!Mov_fecha
  vGrid.Col = 3
  vGrid.Text = rs!Descripcion & ""
  vGrid.CellTag = rs!cod_Transac
  
  vGrid.Col = 4
  vGrid.Text = rs!Usuario & ""
  
  vGrid.Col = 5
  vGrid.Text = rs!Nombre & ""
  
  
  vGrid.Col = 6
  vGrid.Text = rs!NOTAS & ""
  vGrid.Col = 7
  vGrid.Text = rs!App_Name & ""
  vGrid.Col = 8
  vGrid.Text = rs!App_Version & ""
  vGrid.Col = 9
  vGrid.Text = rs!Equipo & ""
  vGrid.Col = 10
  vGrid.Text = rs!Equipo_MAC & ""
  vGrid.Col = 11
  vGrid.Text = rs!Mov_User & ""
  vGrid.Col = 12
  vGrid.Text = rs!Revisado_Usuario & ""
  vGrid.Col = 13
  vGrid.Text = rs!Revisado_Fecha & ""
  
  rs.MoveNext
Loop

rs.Close

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Height = Me.Height - (vGrid.Top + 700)
vGrid.Width = Me.Width - (vGrid.Left + 350)

End Sub





Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "Usuario"
    gBusquedas.Orden = "Usuario"
    gBusquedas.Consulta = "select Usuario,Nombre from US_usuarios"
    gBusquedas.Filtro = ""
    
    If Not gAdminAccess.Rol_AdminView Then
        gBusquedas.Filtro = " AND isnull(key_admin,0) = 0"
    End If
    
    If Not gAdminAccess.Rol_DirGlobal Then
        gBusquedas.Filtro = gBusquedas.Filtro & " AND usuario in(select usuario from PGX_CLIENTES_USERS" _
            & " Where cod_Empresa = " & gPortal.Empresa_Id & ")"
    End If
    
    frmBusquedas.Show vbModal
    txtUsuario.Text = gBusquedas.Resultado
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

If vPaso Or Col > 1 Or Not fraRevision.Enabled Then Exit Sub
 
vGrid.Row = Row
vGrid.Col = 1
If vGrid.Value = vbChecked Then
   strSQL = "update US_TRANSAC_LOG set revisado_usuario = '" & glogon.Usuario & "', revisado_fecha = Getdate()" _
          & " where [SEQ_id]= " & vGrid.CellTag
   vGrid.Col = 3
   strSQL = strSQL & " and cod_Transac = '" & vGrid.CellTag & "'"
   vGrid.Col = 4
   strSQL = strSQL & " and Usuario = '" & vGrid.Text & "'"

   
   Call ConectionExecute(strSQL)

   vGrid.Col = 11
   vGrid.Text = glogon.Usuario
   vGrid.Col = 12
   vGrid.Text = Date
   
End If
End Sub
