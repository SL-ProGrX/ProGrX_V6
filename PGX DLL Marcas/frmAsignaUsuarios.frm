VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmAsignaUsuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación de Horarios"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmAsignaUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6375
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
      _Version        =   1441793
      _ExtentX        =   11245
      _ExtentY        =   11245
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
      HideSelection   =   0   'False
      View            =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   17
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   120
      Top             =   480
   End
   Begin XtremeSuiteControls.ComboBox cboHorario 
      Height          =   312
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   6372
      _Version        =   1441793
      _ExtentX        =   11245
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   6375
      _Version        =   1441793
      _ExtentX        =   11245
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   810
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   1439
      _StockProps     =   79
      Caption         =   "Indique los Usuarios vinculados con este horarios?"
      ForeColor       =   0
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Horario"
      ForeColor       =   16777215
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   996
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmAsignaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub cboHorario_Click()

If vPaso Then Exit Sub
Call sbCarga_usuarios

End Sub

Private Sub Form_Activate()
vModulo = 21
End Sub

Private Sub Form_Load()

vModulo = 21

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Usuario", 1400
lsw.ColumnHeaders.Add , , "Nombre", 3600

vPaso = True
    strSQL = "select rtrim(cod_horario) as 'IdX', rtrim(Descripcion) as 'ItmX' from marcas_horarios where estado = 1"
    Call sbCbo_Llena_New(cboHorario, strSQL, False, True)
vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

vError:

End Sub


Private Sub sbCarga_usuarios()


Me.MousePointer = vbHourglass

On Error GoTo vError

vPaso = True

lsw.ListItems.Clear

txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)


strSQL = "select U.nombre,U.descripcion,isnull(A.Usuario,'-No-') as 'Existe'" _
       & " from usuarios U left join marcas_horarios_users A" _
       & " on U.nombre = A.Usuario and  A.cod_horario = '" & cboHorario.ItemData(cboHorario.ListIndex) & "'" _
       & " Where U.estado = 'A' and U.Descripcion like '%" & txtNombre.Text & "%'" _
       & " order by A.usuario desc,U.nombre asc"
          
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Nombre)
        itmX.SubItems(1) = rs!Descripcion
    
    If rs!Existe <> "-No-" Then
        itmX.Checked = True
        itmX.ForeColor = vbBlue
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

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Then Exit Sub


If Item.Checked Then
   strSQL = "insert MARCAS_HORARIOS_USERS(usuario,cod_horario, registro_Fecha, registro_usuario) values('" & Item.Text _
            & "','" & cboHorario.ItemData(cboHorario.ListIndex) & "',dbo.Mygetdate(),'" & glogon.Usuario & "')"
   Call Bitacora("Registra", "Asignación de Horario- Usuario: " & Item.Text & " -> Horario: " & cboHorario.ItemData(cboHorario.ListIndex))
Else
   strSQL = "Delete MARCAS_HORARIOS_USERS where usuario ='" & Item.Text _
          & "' and cod_horario = '" & cboHorario.ItemData(cboHorario.ListIndex) & "'"
   Call Bitacora("Elimina", "Asignación de Horario- Usuario: " & Item.Text & " -> Horario: " & cboHorario.ItemData(cboHorario.ListIndex))
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbCarga_usuarios
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
 
If KeyCode = vbKeyReturn Then
 Call sbCarga_usuarios
End If
End Sub
