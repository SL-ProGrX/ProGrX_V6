VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCR_Niveles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Niveles de Resolución (Autorización)"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "frmCR_Niveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11070
   Begin XtremeSuiteControls.ListView lswMiembros 
      Height          =   5535
      Left            =   0
      TabIndex        =   11
      Top             =   2640
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   9763
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ListView lswCodigos 
      Height          =   5535
      Left            =   5520
      TabIndex        =   12
      Top             =   2640
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   9763
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtGrupo 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   6615
      _Version        =   1441793
      _ExtentX        =   11663
      _ExtentY        =   556
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
      BackColor       =   16777215
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   270
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Realiza una consulta personalizada sobre los datos actuales"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Object.ToolTipText     =   "Imprime el listado seleccionado"
            Object.Tag             =   "1"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Grupos"
                  Text            =   "Grupos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Miembros"
                  Text            =   "Grupos y Miembros"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Derechos"
                  Text            =   "Grupos y Derechos"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "GruposTotal"
                  Text            =   "Grupos (Miembros - Derechos)"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MiembrosDerechos"
                  Text            =   "Miembros Derechos"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MiembrosGrupos"
                  Text            =   "Miembros Grupos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "cerrar"
            Object.ToolTipText     =   "Sale de esta ventana"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9000
      TabIndex        =   6
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   6615
      _Version        =   1441793
      _ExtentX        =   11668
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
   Begin XtremeSuiteControls.FlatEdit txtDesde 
      Height          =   315
      Left            =   3720
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3408
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
      BackColor       =   16777215
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtHasta 
      Height          =   315
      Left            =   6960
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   550
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMiembro 
      Height          =   435
      Left            =   0
      TabIndex        =   15
      Top             =   2190
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   435
      Left            =   5520
      TabIndex        =   16
      Top             =   2190
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitle 
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   14
      Top             =   1800
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Lineas Autorizadas"
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
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitle 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Miembros Asignados"
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
      Alignment       =   1
   End
   Begin VB.Label G 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Proceso"
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
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label G 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rangos"
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
      Height          =   315
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image imgBusqueda_Rapida 
      Height          =   255
      Index           =   1
      Left            =   9120
      Picture         =   "frmCR_Niveles.frx":000C
      Stretch         =   -1  'True
      ToolTipText     =   "Busqueda Rápida"
      Top             =   360
      Width           =   255
   End
   Begin VB.Label G 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Grupo"
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
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmCR_Niveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnEdita As Boolean, mstrGrupo As String
Dim vScroll As Boolean, vNivelTipo As String, vPaso As Boolean

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, i As Integer

Private Sub sbBuscaGrupo()

On Error GoTo vError

Me.MousePointer = vbHourglass



If txtGrupo.Tag <> "" Then
   mblnEdita = True
   tlbPrincipal.Buttons(1).Enabled = True
   tlbPrincipal.Buttons(2).Enabled = True
   tlbPrincipal.Buttons(3).Enabled = True
   tlbPrincipal.Buttons(4).Enabled = False
   tlbPrincipal.Buttons(5).Enabled = False
   
   imgBusqueda_Rapida(1).Enabled = True
    
   strSQL = "select * from nivel_grupos where nv_cod_grupo = " & txtGrupo.Tag
   Call OpenRecordSet(rs, strSQL)
        txtDesde.Text = Format(rs!nv_desde, "Standard")
        txtHasta.Text = Format(rs!nv_hasta, "Standard")
   rs.Close
   
   
   
   lswMiembros.Enabled = True
   lswCodigos.Enabled = True
   
   txtGrupo.Enabled = True
   
   Call sbCodigosAsignados
   Call sbMiembrosAsignados
   
   tlbPrincipal.Enabled = True

End If

Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbCodigosAsignados()

On Error GoTo vError

txtLineas.Text = fxSysCleanTxtInject(txtLineas.Text)

lswCodigos.ListItems.Clear

vPaso = True
strSQL = "select C.CODIGO,C.DESCRIPCION,NV_COD_GRUPO" _
       & " from catalogo C left join nivel_derechos D on C.codigo = D.codigo" _
       & " and D.nv_cod_grupo = " & txtGrupo.Tag _
       & " Where C.Retencion = 'N' and C.Poliza = 'N'" _
       & "  and (C.codigo like '%" & txtLineas.Text & "%' or C.descripcion like '%" & txtLineas.Text & "%' )" _
       & " order by D.nv_cod_grupo desc,C.codigo"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswCodigos.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      itmX.Checked = IIf(IsNull(rs!NV_Cod_Grupo), vbUnchecked, vbChecked)
 rs.MoveNext
Loop
rs.Close
   
vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbMiembrosAsignados()

On Error GoTo vError

txtMiembro.Text = fxSysCleanTxtInject(txtMiembro.Text)

lswMiembros.ListItems.Clear

vPaso = True

strSQL = "select U.Nombre,U.descripcion,M.NV_COD_GRUPO" _
       & " from Usuarios U left join nivel_Miembros M on U.nombre = M.nombre" _
       & " and M.nv_cod_grupo = " & txtGrupo.Tag _
       & " Where U.estado = 'A'" _
       & "  and (U.Nombre like '%" & txtMiembro.Text & "%' or U.descripcion like '%" & txtMiembro.Text & "%' )" _
       & " order by M.nv_cod_grupo desc,U.nombre"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswMiembros.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!NV_Cod_Grupo), vbUnchecked, vbChecked)
 rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbGuardar()

If Trim(txtGrupo) = "" Then
   MsgBox "Falta El Nombre Del Grupo", vbInformation, "No Se Puede Actualizar Registro"
   Exit Sub
End If

If mblnEdita Then
   
   strSQL = "Update Nivel_Grupos set NV_Descripcion='" & UCase(Trim(txtGrupo)) _
          & "',NV_desde = " & CCur(txtDesde) & ",NV_hasta = " & CCur(txtHasta) _
          & " Where NV_Cod_Grupo=" & txtGrupo.Tag
   Call ConectionExecute(strSQL)

Else
 
    strSQL = "Insert into Nivel_Grupos(NV_Descripcion,NV_Tipo,NV_Desde,NV_Hasta) Values('" _
           & UCase(Trim(txtGrupo)) & "','" & vNivelTipo & "'," _
           & CCur(txtDesde) & "," & CCur(txtHasta) & ")"
    Call ConectionExecute(strSQL)
    
    strSQL = "select isnull(max(nv_cod_grupo),0) as Ultimo from nivel_grupos" _
           & " where nv_tipo = '" & vNivelTipo & "'"
    Call OpenRecordSet(rs, strSQL)
      txtGrupo.Tag = rs!ultimo
    rs.Close

End If

tlbPrincipal.Buttons(1).Enabled = True
tlbPrincipal.Buttons(2).Enabled = True
tlbPrincipal.Buttons(3).Enabled = True
tlbPrincipal.Buttons(4).Enabled = False
tlbPrincipal.Buttons(5).Enabled = False

imgBusqueda_Rapida(1).Enabled = True

lswMiembros.Enabled = False
lswCodigos.Enabled = False

txtGrupo.Enabled = False

Call sbCodigosAsignados
Call sbMiembrosAsignados

mblnEdita = True

Call RefrescaTags(Me)


End Sub




Private Sub cboTipo_Click()

Select Case cboTipo.Text
  Case "Proceso de Formalización"
    vNivelTipo = "F"
  Case "Proceso de Resolución"
    vNivelTipo = "R"
  Case "Proceso de Anulación"
    vNivelTipo = "N"
End Select


lswCodigos.ListItems.Clear
lswMiembros.ListItems.Clear

txtDesde.Text = "0.00"
txtHasta.Text = "0.00"

txtGrupo.Text = ""



End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtGrupo.Tag = "" Then txtGrupo.Tag = "0"

If vScroll Then
    strSQL = "Select Top 1 NV_Cod_Grupo,NV_Descripcion From Nivel_Grupos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where nv_tipo = '" & vNivelTipo & "' and NV_Cod_Grupo > " & txtGrupo.Tag & " order by NV_Cod_Grupo asc"
    Else
       strSQL = strSQL & " where nv_tipo = '" & vNivelTipo & "' and NV_Cod_Grupo < " & txtGrupo.Tag & " order by NV_Cod_Grupo desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
        txtGrupo.Text = rs!NV_Descripcion
        txtGrupo.Tag = rs!NV_Cod_Grupo
        Call sbBuscaGrupo
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
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3



With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1800
    .Add , , "Descripción", 3100
End With

With lswCodigos.ColumnHeaders
    .Clear
    .Add , , "Código", 1300
    .Add , , "Descripción", 3600
End With


vScroll = False
FlatScrollBar.Value = 0
vScroll = True
 
Call sbToolBarIconos(tlbPrincipal, False)
Call Formularios(Me)

cboTipo.Clear
cboTipo.AddItem "Proceso de Formalización"
cboTipo.AddItem "Proceso de Resolución"
cboTipo.AddItem "Proceso de Anulación"
cboTipo.Text = "Proceso de Formalización"

tlbPrincipal.Buttons(1).Enabled = True
tlbPrincipal.Buttons(2).Enabled = False
tlbPrincipal.Buttons(3).Enabled = False
tlbPrincipal.Buttons(4).Enabled = False
tlbPrincipal.Buttons(5).Enabled = False

Call RefrescaTags(Me)

End Sub

Private Sub imgBusqueda_Rapida_Click(Index As Integer)

On Error GoTo vError


Select Case Index
  Case 1
       
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Consulta = "Select NV_Cod_Grupo as Codigo,NV_Descripcion From Nivel_Grupos"
    gBusquedas.Filtro = "And Nv_Tipo='" & vNivelTipo & "'"
    gBusquedas.Columna = "NV_Descripcion"
    gBusquedas.Orden = "NV_Descripcion"
    frmBusquedas.Show vbModal
    
    GLOBALES.gblnBuscando = True
    
    txtGrupo.Text = gBusquedas.Resultado2
    txtGrupo.Tag = gBusquedas.Resultado
    
    Call sbBuscaGrupo
    
    
End Select

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Sub

Private Sub lswCodigos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCodigos.SortKey = ColumnHeader.Index - 1
  If lswCodigos.SortOrder = 0 Then lswCodigos.SortOrder = 1 Else lswCodigos.SortOrder = 0
  lswCodigos.Sorted = True
End Sub

Private Sub lswCodigos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert nivel_derechos(nv_cod_grupo,codigo) values(" & txtGrupo.Tag _
          & ",'" & Item.Text & "')"
Else
   strSQL = "Delete from Nivel_derechos where NV_Cod_Grupo=" & txtGrupo.Tag _
          & " and codigo = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbAbortRetryIgnore

End Sub

Private Sub lswMiembros_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswMiembros.SortKey = ColumnHeader.Index - 1
  If lswMiembros.SortOrder = 0 Then lswMiembros.SortOrder = 1 Else lswMiembros.SortOrder = 0
  lswMiembros.Sorted = True
End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub
 
On Error GoTo vError

If Item.Checked Then
   strSQL = "insert nivel_miembros(nv_cod_grupo,nombre) values(" & txtGrupo.Tag _
          & ",'" & Item.Text & "')"
Else
   strSQL = "Delete from Nivel_Miembros where NV_Cod_Grupo=" & txtGrupo.Tag _
          & " and Nombre='" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbAbortRetryIgnore

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

If Button.Key <> "cerrar" Then Me.MousePointer = vbHourglass

Select Case Button.Key
  Case "insertar", "nuevo"
       txtGrupo.Tag = ""
       mblnEdita = False
       
       tlbPrincipal.Buttons(1).Enabled = False
       tlbPrincipal.Buttons(2).Enabled = False
       tlbPrincipal.Buttons(3).Enabled = False
       tlbPrincipal.Buttons(4).Enabled = True
       tlbPrincipal.Buttons(5).Enabled = True
       imgBusqueda_Rapida(1).Enabled = False
       
       txtGrupo.Enabled = True
       txtGrupo = ""
       txtDesde = ""
       txtHasta = ""
       txtGrupo.SetFocus
       
       lswCodigos.ListItems.Clear
       lswMiembros.ListItems.Clear
       
  Case "modificar", "editar"
       mblnEdita = True
       tlbPrincipal.Buttons(1).Enabled = False
       tlbPrincipal.Buttons(2).Enabled = False
       tlbPrincipal.Buttons(3).Enabled = False
       tlbPrincipal.Buttons(4).Enabled = True
       tlbPrincipal.Buttons(5).Enabled = True
       imgBusqueda_Rapida(1).Enabled = False
       
       lswMiembros.Enabled = True
       lswCodigos.Enabled = True
       
       txtGrupo.Enabled = True
       
       txtGrupo.SetFocus
       mstrGrupo = txtGrupo
       
       
  Case "borrar"
  
       If MsgBox("Registro Será Eliminado", vbInformation + vbYesNo, "Confirme Opción") = vbYes Then
          If txtGrupo.Tag <> "" Then
             strSQL = "Delete from Nivel_Derechos where NV_Cod_Grupo=" & txtGrupo.Tag
             Call ConectionExecute(strSQL)
             
             strSQL = "Delete from Nivel_Miembros where NV_Cod_Grupo=" & txtGrupo.Tag
             Call ConectionExecute(strSQL)
             
             strSQL = "Delete from Nivel_Grupos where NV_Cod_Grupo=" & txtGrupo.Tag
             Call ConectionExecute(strSQL)
             
             tlbPrincipal.Buttons(1).Enabled = True
             tlbPrincipal.Buttons(2).Enabled = False
             tlbPrincipal.Buttons(3).Enabled = False
             tlbPrincipal.Buttons(4).Enabled = False
             tlbPrincipal.Buttons(5).Enabled = False
             
             imgBusqueda_Rapida(1).Enabled = True
             
             txtGrupo.Enabled = False
             txtGrupo = ""
             lswMiembros.ListItems.Clear
             lswCodigos.ListItems.Clear
          
          End If
       End If
       
  Case "guardar"
       Call sbGuardar
       
  Case "deshacer"
       tlbPrincipal.Buttons(1).Enabled = True
       tlbPrincipal.Buttons(2).Enabled = False
       tlbPrincipal.Buttons(3).Enabled = False
       tlbPrincipal.Buttons(4).Enabled = False
       tlbPrincipal.Buttons(5).Enabled = False
       
       imgBusqueda_Rapida(1).Enabled = True
       
       lswMiembros.Enabled = False
       lswCodigos.Enabled = False
       
       txtGrupo.Enabled = False
       txtGrupo = ""
       lswMiembros.ListItems.Clear
       lswCodigos.ListItems.Clear
  
  Case "ayuda"

End Select

If Button.Key <> "cerrar" Then Me.MousePointer = vbDefault

End Sub

Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strRuta As String, strSQL As String, dateFecha As Date

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .SelectionFormula = "{NIVEL_GRUPOS.NV_TIPO} = '" & vNivelTipo & "'"
   
 .Connect = glogon.ConectRPT
   
 Select Case ButtonMenu.Key
   Case "Grupos"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_NV_Grupos.rpt")
   Case "Miembros"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_NV_GruposMiembros.rpt")
   Case "Derechos"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_NV_GruposDerechos.rpt")
   Case "GruposTotal"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_NV_GruposMiembrosDerechos.rpt")
   Case "MiembrosDerechos"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_NV_MiembrosDerechos.rpt")
   Case "MiembrosGrupos"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_NV_MiembrosGrupos.rpt")
 End Select
  
  .PrintReport
  
End With

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub txtDesde_GotFocus()
On Error GoTo vError
txtDesde = CCur(txtDesde)
vError:
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtHasta.SetFocus
End Sub

Private Sub txtDesde_LostFocus()
On Error GoTo vError
txtDesde = Format(CCur(txtDesde), "Standard")
vError:
End Sub

Private Sub txtGrupo_Change()
If GLOBALES.gblnBuscando = True Then
   Call sbBuscaGrupo
   GLOBALES.gblnBuscando = False
   Call RefrescaTags(Me)
End If
End Sub

Private Sub txtGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDesde.SetFocus
End Sub

Private Sub txtHasta_GotFocus()
On Error GoTo vError
txtHasta = CCur(txtHasta)
vError:
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtGrupo.SetFocus
End Sub

Private Sub txtHasta_LostFocus()
On Error GoTo vError
txtHasta = Format(CCur(txtHasta), "Standard")
vError:
End Sub


Private Sub txtLineas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbCodigosAsignados
End If
End Sub


Private Sub txtMiembro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbMiembrosAsignados
End If
End Sub
