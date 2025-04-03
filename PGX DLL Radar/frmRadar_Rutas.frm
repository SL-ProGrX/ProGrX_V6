VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmRadar_Rutas 
   Caption         =   "Radar: Rutas"
   ClientHeight    =   8844
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   16008
   Icon            =   "frmRadar_Rutas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8844
   ScaleWidth      =   16008
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   8652
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   11292
      _Version        =   1245185
      _ExtentX        =   19918
      _ExtentY        =   15261
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
      View            =   3
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   4692
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4332
      _Version        =   1245185
      _ExtentX        =   7641
      _ExtentY        =   8276
      _StockProps     =   79
      Caption         =   "Filtros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   312
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   2652
         _Version        =   1245185
         _ExtentX        =   4678
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCanton 
         Height          =   312
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   2652
         _Version        =   1245185
         _ExtentX        =   4678
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDistrito 
         Height          =   312
         Left            =   1680
         TabIndex        =   6
         Top             =   1080
         Width           =   2652
         _Version        =   1245185
         _ExtentX        =   4678
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   0
         TabIndex        =   10
         Top             =   1800
         Width           =   4332
         _Version        =   1245185
         _ExtentX        =   7641
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCentro 
         Height          =   312
         Left            =   0
         TabIndex        =   12
         Top             =   2520
         Width           =   4332
         _Version        =   1245185
         _ExtentX        =   7641
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDept 
         Height          =   312
         Left            =   0
         TabIndex        =   14
         Top             =   3240
         Width           =   4332
         _Version        =   1245185
         _ExtentX        =   7641
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboSeccion 
         Height          =   312
         Left            =   0
         TabIndex        =   22
         Top             =   3960
         Width           =   4332
         _Version        =   1245185
         _ExtentX        =   7641
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
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   0
         TabIndex        =   23
         Top             =   3720
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   3000
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Centro"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   2280
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Centro"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantón"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   852
      End
   End
   Begin XtremeSuiteControls.GroupBox gbRuta 
      Height          =   1692
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   4332
      _Version        =   1245185
      _ExtentX        =   7641
      _ExtentY        =   2984
      _StockProps     =   79
      Caption         =   "Crear Ruta [Unidades Seleccionadas ]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit feRutaNombre 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   4212
         _Version        =   1245185
         _ExtentX        =   7429
         _ExtentY        =   556
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnRuta 
         Height          =   492
         Left            =   2760
         TabIndex        =   19
         Top             =   1080
         Width           =   1452
         _Version        =   1245185
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Crear Ruta"
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
         Appearance      =   16
         Picture         =   "frmRadar_Rutas.frx":6852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   852
      End
   End
   Begin XtremeSuiteControls.GroupBox gbBusqueda 
      Height          =   1092
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   4332
      _Version        =   1245185
      _ExtentX        =   7641
      _ExtentY        =   1926
      _StockProps     =   79
      Caption         =   "Buscar:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   492
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   1452
         _Version        =   1245185
         _ExtentX        =   2561
         _ExtentY        =   868
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
         Transparent     =   -1  'True
         Appearance      =   16
         Picture         =   "frmRadar_Rutas.frx":7041
      End
   End
   Begin XtremeSuiteControls.ComboBox cboRuta 
      Height          =   312
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   4332
      _Version        =   1245185
      _ExtentX        =   7641
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
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta"
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
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   1092
   End
End
Attribute VB_Name = "frmRadar_Rutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub btnBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vAnd As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

vAnd = False

strSQL = "select * from vRADAR_DIRECTORIO"

If cboTipo.Text <> "TODOS" Then
  strSQL = strSQL & IIf(vAnd, " AND ", " WHERE ") & " CENTRO_TIPO = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
  vAnd = True
End If

If cboCentro.Text <> "TODOS" Then
  strSQL = strSQL & IIf(vAnd, " AND ", " WHERE ") & " COD_INSTITUCION = " & cboCentro.ItemData(cboCentro.ListIndex)
  vAnd = True
End If

If cboDept.Text <> "TODOS" Then
  strSQL = strSQL & IIf(vAnd, " AND ", " WHERE ") & " COD_DEPARTAMENTO = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
  vAnd = True
End If

If cboSeccion.Text <> "TODOS" Then
  strSQL = strSQL & IIf(vAnd, " AND ", " WHERE ") & " COD_SECCION = '" & cboSeccion.ItemData(cboSeccion.ListIndex) & "'"
  vAnd = True
End If


If cboProvincia.Text <> "TODOS" Then
  strSQL = strSQL & IIf(vAnd, " AND ", " WHERE ") & " COD_PROVINCIA = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"
  vAnd = True
End If


If cboCanton.Text <> "TODOS" Then
  strSQL = strSQL & IIf(vAnd, " AND ", " WHERE ") & " COD_CANTON = '" & cboCanton.ItemData(cboCanton.ListIndex) & "'"
  vAnd = True
End If


If cboDistrito.Text <> "TODOS" Then
  strSQL = strSQL & IIf(vAnd, " AND ", " WHERE ") & " COD_DISTRITO = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "'"
  vAnd = True
End If


With lsw.ColumnHeaders
    .Clear
    .Add , , "Provincia", 1200
    .Add , , "Cantón", 1200
    .Add , , "Distrito", 1200
    .Add , , "Req. Autor.", 1100, vbCenter
    .Add , , "Tipo Centro", 2500
    .Add , , "Centro", 3500
    .Add , , "Departamento", 3500
    .Add , , "Sección", 3500
    .Add , , "Poblado", 1500
    .Add , , "Zona", 1500
    .Add , , "Contacto", 1500
    .Add , , "Telefono 1", 1500
    .Add , , "Telefono 2", 1500
    .Add , , "Email 1", 3500
    .Add , , "Email 2", 3500
    .Add , , "Sitio Web", 3500
    .Add , , "Facebook", 3500
    .Add , , "Dirección", 3500
End With

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Provincia_Desc)
      itmX.SubItems(1) = rs!Canton_Desc
      itmX.SubItems(2) = rs!Distrito_Desc
      itmX.SubItems(3) = IIf((rs!visita_autorizada = 1), "Sí", "No")
      itmX.SubItems(4) = rs!Centro_Tipo_Desc
      itmX.SubItems(5) = rs!Institucion_Desc
      itmX.SubItems(6) = rs!Departamento_Desc
      itmX.SubItems(7) = rs!Seccion_Desc
      itmX.SubItems(8) = rs!Poblado & ""
      itmX.SubItems(9) = rs!Zona & ""
      itmX.SubItems(10) = rs!Visita_Contacto & ""
      itmX.SubItems(11) = rs!Telefono_01 & ", " & rs!Telefono_01_Ext
      itmX.SubItems(12) = rs!Telefono_02 & ", " & rs!Telefono_02_Ext
      itmX.SubItems(13) = rs!Email_01 & ""
      itmX.SubItems(14) = rs!Email_02 & ""
      itmX.SubItems(15) = rs!Sitio_Web & ""
      itmX.SubItems(16) = rs!Facebook & ""
      itmX.SubItems(17) = rs!Direccion & ""
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical





End Sub

Private Sub cboCanton_Click()
Dim strSQL As String

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, True, True)

End Sub

Private Sub cboCentro_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vCodigo As Integer

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If cboCentro.Text = "TODOS" Then
   vCodigo = 0

    strSQL = "select rtrim(Ct.CENTRO_TIPO) as 'IdX', rtrim(Ct.Descripcion) as 'ItmX'" _
           & " from RADAR_CENTROS_TIPOS Ct" _
           & " where Ct.activo = 1"
    Call sbCbo_Llena_New(cboTipo, strSQL, True, True)

Else
   vCodigo = cboCentro.ItemData(cboCentro.ListIndex)

    strSQL = "select rtrim(Ct.CENTRO_TIPO) as 'IdX', rtrim(Ct.Descripcion) as 'ItmX'" _
           & " from RADAR_CENTROS_TIPOS Ct inner join RADAR_INSTITUCION_TIPOS It on Ct.CENTRO_TIPO = It.CENTRO_TIPO" _
           & " and It.COD_INSTITUCION = " & vCodigo _
           & " where Ct.activo = 1"
    Call sbCbo_Llena_New(cboTipo, strSQL, True, True)

End If


strSQL = "select S.cod_Departamento as 'IdX', rtrim(S.descripcion) + '   [' + rtrim(S.cod_Departamento) + ']' as 'ItmX'" _
       & " from afDepartamentos S inner join RADAR_DIRECTORIO D on S.cod_Institucion = D.cod_Institucion" _
       & " and S.cod_Departamento = D.cod_Departamento" _
       & " where S.cod_institucion = " & vCodigo _
       & " group by S.descripcion, S.cod_Departamento"
       
Call sbCbo_Llena_New(cboDept, strSQL, True, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboDept_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vCodigo As Integer

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If cboCentro.Text = "TODOS" Then
   vCodigo = 0
Else
   vCodigo = cboCentro.ItemData(cboCentro.ListIndex)
End If

strSQL = "select rtrim(S.cod_Seccion) as 'Idx', rtrim(S.descripcion) as 'ItmX'" _
       & " from afSecciones S inner join RADAR_DIRECTORIO D on S.cod_Institucion = D.cod_Institucion" _
       & " and S.cod_Departamento = D.cod_Departamento and S.cod_Seccion = D.cod_Seccion" _
       & " where S.cod_institucion = " & vCodigo _
       & "   and S.cod_Departamento = '" & cboDept.ItemData(cboDept.ListIndex) _
       & "' order by S.descripcion"
      
Call sbCbo_Llena_New(cboSeccion, strSQL, True, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboProvincia_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, True, True)
vPaso = False

Call cboCanton_Click
End Sub



Private Sub Form_Activate()
vModulo = 37
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

vModulo = 37

'Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture



With lsw.ColumnHeaders
    .Clear
    .Add , , "Provincia", 1200
    .Add , , "Cantón", 1200
    .Add , , "Distrito", 1200
    .Add , , "Req. Autor.", 1100, vbCenter
    .Add , , "Tipo Centro", 2500
    .Add , , "Centro", 3500
    .Add , , "Departamento", 3500
    .Add , , "Sección", 3500
    .Add , , "Poblado", 1500
    .Add , , "Zona", 1500
    .Add , , "Contacto", 1500
    .Add , , "Telefono 1", 1500
    .Add , , "Telefono 2", 1500
    .Add , , "Email 1", 3500
    .Add , , "Email 2", 3500
    .Add , , "Sitio Web", 3500
    .Add , , "Facebook", 3500
    .Add , , "Dirección", 3500
End With

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub Form_Resize()
On Error Resume Next

lsw.Width = Me.Width - (lsw.Left + 350)
lsw.Height = Me.Height - (lsw.Top + 550)


End Sub

Private Sub PushButton1_Click()

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

On Error GoTo vError

TimerX.Interval = 0
TimerX.Enabled = False


vPaso = True
    
    strSQL = "select RUTA_ID as 'IdX', rtrim(Descripcion) as 'ItmX' from RADAR_RUTAS where ACTIVO = 1 AND RUTA_LIBRE = 0"
    Call sbCbo_Llena_New(cboRuta, strSQL, True, True)
    
    
    strSQL = "select rtrim(cod_institucion) as 'IdX', rtrim(Descripcion) as 'ItmX' from INSTITUCIONES where ACTIVA = 1"
    Call sbCbo_Llena_New(cboCentro, strSQL, True, True)

    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, True, True)
vPaso = False

Call cboCentro_Click
Call cboProvincia_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub
