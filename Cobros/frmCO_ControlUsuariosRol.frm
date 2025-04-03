VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.controls.v19.2.0.ocx"
Begin VB.Form frmCO_ControlUsuariosRol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rol de Atención de Ejecutivos de Cobros"
   ClientHeight    =   8172
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   13584
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCO_ControlUsuariosRol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8172
   ScaleWidth      =   13584
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7212
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   8772
      _Version        =   1245186
      _ExtentX        =   15473
      _ExtentY        =   12721
      _StockProps     =   68
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Asignación"
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "lswAntiguedad"
      Item(0).Control(1)=   "lswGarantias"
      Item(0).Control(2)=   "Label1(3)"
      Item(0).Control(3)=   "Label1(1)"
      Item(0).Control(4)=   "lswOficinas"
      Item(0).Control(5)=   "lswInstitucion"
      Item(0).Control(6)=   "Label1(4)"
      Item(0).Control(7)=   "Label1(2)"
      Item(1).Caption =   "Copia"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "lswCopia"
      Item(1).Control(1)=   "Label1(0)"
      Item(1).Control(2)=   "btnCopia(0)"
      Item(1).Control(3)=   "btnCopia(1)"
      Begin XtremeSuiteControls.ListView lswAntiguedad 
         Height          =   3132
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4212
         _Version        =   1245186
         _ExtentX        =   7429
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
      Begin XtremeSuiteControls.ListView lswGarantias 
         Height          =   3132
         Left            =   4440
         TabIndex        =   4
         Top             =   600
         Width           =   4212
         _Version        =   1245186
         _ExtentX        =   7429
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
      Begin XtremeSuiteControls.ListView lswOficinas 
         Height          =   3132
         Left            =   120
         TabIndex        =   7
         Top             =   4200
         Width           =   4212
         _Version        =   1245186
         _ExtentX        =   7429
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
      Begin XtremeSuiteControls.ListView lswInstitucion 
         Height          =   3132
         Left            =   4440
         TabIndex        =   8
         Top             =   4200
         Width           =   4212
         _Version        =   1245186
         _ExtentX        =   7429
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
      Begin XtremeSuiteControls.ListView lswCopia 
         Height          =   5892
         Left            =   -69880
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1245186
         _ExtentX        =   15049
         _ExtentY        =   10393
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
      Begin XtremeSuiteControls.PushButton btnCopia 
         Height          =   312
         Index           =   0
         Left            =   -67360
         TabIndex        =   14
         Top             =   6720
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1245186
         _ExtentX        =   2561
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Copia"
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
      End
      Begin XtremeSuiteControls.PushButton btnCopia 
         Cancel          =   -1  'True
         Height          =   312
         Index           =   1
         Left            =   -65920
         TabIndex        =   15
         Top             =   6720
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1245186
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Limpieza de Roles"
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
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copiar Rol del usuario actual a.: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -69880
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   3252
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Oficinas (formalizadoras).: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   3960
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Instituciones.: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   6240
         TabIndex        =   9
         Top             =   3960
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Antiguedad.: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Garantías.: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   6240
         TabIndex        =   5
         Top             =   360
         Width           =   2412
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltroUsuario 
      Height          =   312
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4452
      _Version        =   1245186
      _ExtentX        =   7853
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   7212
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   4452
      _Version        =   1245186
      _ExtentX        =   7853
      _ExtentY        =   12721
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Label lblUsuario 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13812
   End
End
Attribute VB_Name = "frmCO_ControlUsuariosRol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnCopia_Click(Index As Integer)
Dim strSQL As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Index
  Case 0  '"Copia"
     For i = 1 To lswCopia.ListItems.Count
        If lswCopia.ListItems.Item(i).Checked Then
            strSQL = "exec spCBR_UsuarioRol_Copia '" & lblUsuario.Tag & "','" _
                   & lswCopia.ListItems.Item(i).Tag & "','" & glogon.Usuario & "'"
            Call ConectionExecute(strSQL)
        End If
     Next i
     
  Case 1 'Limpia"
      strSQL = "exec spCBR_UsuarioRol_Limpia"
      Call ConectionExecute(strSQL)
End Select

Me.MousePointer = vbDefault
MsgBox "Proceso concluído satisfactoriamente!", vbInformation

Call sbInicializa

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 4

End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

lsw.ColumnHeaders.Add , , "", 4500
lswCopia.ColumnHeaders.Add , , "", 6500

lswGarantias.ColumnHeaders.Add , , "", 4000
lswAntiguedad.ColumnHeaders.Add , , "", 4000

lswOficinas.ColumnHeaders.Add , , "", 950
lswOficinas.ColumnHeaders.Add , , "", 3100

lswInstitucion.ColumnHeaders.Add , , "", 950
lswInstitucion.ColumnHeaders.Add , , "", 3100

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

lsw.ListItems.Clear
lswCopia.ListItems.Clear

lswAntiguedad.ListItems.Clear
lswGarantias.ListItems.Clear
lswOficinas.ListItems.Clear
lswInstitucion.ListItems.Clear


lblUsuario.Tag = ""
lblUsuario.Caption = ""

vPaso = True


strSQL = "select USUARIO, NOMBRE " _
       & "  From CBR_USUARIOS" _
       & " Where Estado = 1 and Nombre like '%" & Trim(txtFiltroUsuario.Text) & "%'" _
       & " ORDER BY NOMBRE"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Nombre)
      itmX.Tag = rs!Usuario
  rs.MoveNext
Loop
rs.Close

vPaso = False


strSQL = "select USUARIO, NOMBRE " _
       & "  From CBR_USUARIOS" _
       & " Where Estado = 1" _
       & " ORDER BY NOMBRE"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswCopia.ListItems.Add(, , rs!Nombre)
      itmX.Tag = rs!Usuario
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

Private Sub sbCargaListas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass


tcMain.Item(0).Selected = True

lswAntiguedad.ListItems.Clear
lswGarantias.ListItems.Clear
lswOficinas.ListItems.Clear
lswInstitucion.ListItems.Clear


vPaso = True

strSQL = "select Ant.COD_ANTIGUEDAD,Ant.DESCRIPCION,isnull(Asg.Usuario,'No-ASG') as 'Asignado'" _
       & " from CBR_ANTIGUEDAD_TIPOS Ant left join CBR_USUARIOS_ANTIGUEDADES Asg on Ant.COD_ANTIGUEDAD = Asg.COD_ANTIGUEDAD" _
       & "  AND Asg.USUARIO = '" & lblUsuario.Tag & "'" _
       & " ORDER BY Ant.COD_ANTIGUEDAD"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswAntiguedad.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!cod_Antiguedad
      If Trim(rs!Asignado) <> "No-ASG" Then
         itmX.Checked = True
      End If
  rs.MoveNext
Loop
rs.Close


strSQL = "select Ant.GARANTIA,Ant.DESCRIPCION,isnull(Asg.Usuario,'No-ASG') as 'Asignado'" _
       & " from CRD_GARANTIA_TIPOS Ant left join CBR_USUARIOS_GARANTIAS Asg on Ant.GARANTIA = Asg.GARANTIA" _
       & "  AND Asg.USUARIO = '" & lblUsuario.Tag & "'" _
       & " ORDER BY Ant.GARANTIA"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswGarantias.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!Garantia
      If Trim(rs!Asignado) <> "No-ASG" Then
         itmX.Checked = True
      End If
  rs.MoveNext
Loop
rs.Close

strSQL = "select Ant.COD_OFICINA,Ant.DESCRIPCION,isnull(Asg.Usuario,'No-ASG') as 'Asignado'" _
       & " from SIF_OFICINAS Ant left join CBR_USUARIOS_oficinas Asg on Ant.COD_OFICINA = Asg.COD_OFICINA" _
       & "  AND Asg.USUARIO = '" & lblUsuario.Tag & "'" _
       & " ORDER BY Ant.COD_OFICINA"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswOficinas.ListItems.Add(, , rs!COD_OFICINA)
      itmX.SubItems(1) = rs!Descripcion
      itmX.Tag = rs!COD_OFICINA
      If Trim(rs!Asignado) <> "No-ASG" Then
         itmX.Checked = True
      End If
  rs.MoveNext
Loop
rs.Close



strSQL = "select Ant.COD_INSTITUCION,Ant.DESCRIPCION,isnull(Asg.Usuario,'No-ASG') as 'Asignado'" _
       & " from Instituciones Ant left join CBR_USUARIOS_Institucion Asg on Ant.COD_INSTITUCION = Asg.COD_INSTITUCION" _
       & "  AND Asg.USUARIO = '" & lblUsuario.Tag & "'" _
       & " ORDER BY Ant.COD_INSTITUCION"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswInstitucion.ListItems.Add(, , rs!cod_institucion)
      itmX.SubItems(1) = rs!Descripcion
      itmX.Tag = rs!cod_institucion
      If Trim(rs!Asignado) <> "No-ASG" Then
         itmX.Checked = True
      End If
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




Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

lblUsuario.Tag = Item.Tag
lblUsuario.Caption = Item.Text

Call sbCargaListas


End Sub

Private Sub lswAntiguedad_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String


If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CBR_USUARIOS_ANTIGUEDADES(usuario,cod_antiguedad,registro_fecha,registro_usuario)" _
          & " values('" & lblUsuario.Tag & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete CBR_USUARIOS_ANTIGUEDADES where usuario = '" & lblUsuario.Tag _
          & "' and cod_antiguedad = '" & Item.Tag & "'"
End If

Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lswGarantias_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CBR_USUARIOS_GARANTIAS(usuario,GARANTIA,registro_fecha,registro_usuario)" _
          & " values('" & lblUsuario.Tag & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete CBR_USUARIOS_GARANTIAS where usuario = '" & lblUsuario.Tag _
          & "' and GARANTIA = '" & Item.Tag & "'"
End If

Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lswInstitucion_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CBR_USUARIOS_INSTITUCION(usuario,COD_INSTITUCION,registro_fecha,registro_usuario)" _
          & " values('" & lblUsuario.Tag & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete CBR_USUARIOS_INSTITUCION where usuario = '" & lblUsuario.Tag _
          & "' and COD_INSTITUCION = '" & Item.Tag & "'"
End If

Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswOficinas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CBR_USUARIOS_OFICINAS(usuario,cod_oficina,registro_fecha,registro_usuario)" _
          & " values('" & lblUsuario.Tag & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete CBR_USUARIOS_OFICINAS where usuario = '" & lblUsuario.Tag _
          & "' and cod_oficina = '" & Item.Tag & "'"
End If

Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtFiltroUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call sbInicializa
End If

End Sub
