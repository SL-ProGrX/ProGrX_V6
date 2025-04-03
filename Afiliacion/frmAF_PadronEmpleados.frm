VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmAF_PadronEmpleados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de Padron de Nómina del Patrono"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   13095
      _Version        =   1441792
      _ExtentX        =   23098
      _ExtentY        =   9340
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
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkFechaIngreso 
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   1920
      Width           =   2175
      _Version        =   1441792
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ingreso a la Empresa"
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpRegInicio 
      Height          =   315
      Left            =   7560
      TabIndex        =   18
      Top             =   1920
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   195
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   11280
      Top             =   360
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      _Version        =   1441792
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   11400
      TabIndex        =   4
      Top             =   1320
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_PadronEmpleados.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnElimina 
      Height          =   375
      Left            =   11880
      TabIndex        =   5
      ToolTipText     =   "Eliiminar Registro"
      Top             =   1320
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_PadronEmpleados.frx":0700
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
      _Version        =   1441792
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtIdAlterna 
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
      _Version        =   1441792
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   1320
      Width           =   6855
      _Version        =   1441792
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   12360
      TabIndex        =   13
      ToolTipText     =   "Exportar a Excel"
      Top             =   1320
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_PadronEmpleados.frx":0CA4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   11400
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2566
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpIngInicio 
      Height          =   315
      Left            =   2160
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpIngCorte 
      Height          =   315
      Left            =   3480
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpRegCorte 
      Height          =   315
      Left            =   8880
      TabIndex        =   19
      Top             =   1920
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.CheckBox chkFechaRegistro 
      Height          =   255
      Left            =   10320
      TabIndex        =   21
      Top             =   1920
      Width           =   2175
      _Version        =   1441792
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Registro en el Padron"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   1935
      _Version        =   1441792
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   7320
      TabIndex        =   24
      Top             =   360
      Width           =   3735
      _Version        =   1441792
      _ExtentX        =   6588
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Persona"
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
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   25
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario de Registro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   1815
   End
   Begin XtremeSuiteControls.Label lblItems 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   8160
      Width           =   2175
      _Version        =   1441792
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Id Empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   13095
      _Version        =   1441792
      _ExtentX        =   23098
      _ExtentY        =   661
      _StockProps     =   14
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
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
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmAF_PadronEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean



Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnElimina_Click()

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim i As Long, x As Long

x = 0
strSQL = ""

With lsw.ListItems
    For i = 1 To .Count
        If .Item(i).Checked Then
            strSQL = strSQL & Space(10) & "delete AFI_PADRON WHERE CEDULA = '" & .Item(i).Text & "'"
            x = x + 1
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
        
    Next i
End With

'Ultimo Bloque
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


Me.MousePointer = vbDefault

MsgBox "Exclusión del Padrón realizada, Casos Afectados(" & x & ")", vbInformation

Call sbBuscar

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkFechaIngreso_Click()
If chkFechaIngreso.Value = xtpChecked Then
    dtpIngInicio.Enabled = True
Else
    dtpIngInicio.Enabled = False
End If

dtpIngCorte.Enabled = dtpIngInicio.Enabled

End Sub

Private Sub chkFechaRegistro_Click()
If chkFechaRegistro.Value = xtpChecked Then
    dtpRegInicio.Enabled = True
Else
    dtpRegInicio.Enabled = False
End If

dtpRegCorte.Enabled = dtpRegInicio.Enabled
End Sub

Private Sub chkTodos_Click()

If vPaso Then Exit Sub

Dim i As Long

For i = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


scMain.Caption = "Listado de Padron de Empleados"

With lsw.ColumnHeaders
    .Add , , "Identificación", 2000
    .Add , , "Id. Alterna", 2500
    .Add , , "Nombre", 4500
    .Add , , "Ingreso", 2000, vbCenter
    .Add , , "Reg.Usuario", 2000, vbCenter
    .Add , , "Reg.Fecha", 2000, vbCenter
    .Add , , "Estado", 2000, vbCenter
    
End With

dtpIngCorte.Value = fxFechaServidor
dtpRegCorte.Value = dtpIngCorte.Value

dtpIngInicio.Value = DateAdd("yyyy", -5, dtpIngCorte.Value)
dtpRegInicio.Value = DateAdd("d", -1, dtpIngCorte.Value)

Call chkFechaIngreso_Click
Call chkFechaRegistro_Click

Call Formularios(Me)
Call RefrescaTags(Me)
        
End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub TimerX_Timer()
Timerx.Interval = 0
Timerx.Enabled = False

strSQL = "select COD_INSTITUCION as 'Idx',  '[' + COD_DIVISA + ']  ' + DESCRIPCION as 'ItmX'" _
       & "  from INSTITUCIONES where ACTIVA = 1" _
       & "  order by COD_INSTITUCION"
Call sbCbo_Llena_New(cbo, strSQL, True, True)


strSQL = "select RTRIM(COD_ESTADO) as 'Idx',  RTRIM( DESCRIPCION ) as 'ItmX'" _
       & "  from AFI_ESTADOS_PERSONA where ACTIVO = 1" _
       & "  order by COD_ESTADO"
Call sbCbo_Llena_New(cboEstado, strSQL, True, True)

End Sub





Private Sub sbBuscar()

On Error GoTo vError

Dim vWhere As Boolean

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

vWhere = False

strSQL = "select P.* , isnull(Pe.Descripcion,'No Localizado') as 'EstadoPersona'" _
       & " from AFI_PADRON P left join Socios S on S.CEDULA in(P.ID_ALTERNO, P.CEDULA)" _
       & " left join AFI_ESTADOS_PERSONA Pe on S.EstadoActual = Pe.COD_ESTADO"


If cboEstado.Text = "TODOS" Then
    strSQL = "select P.* , isnull(Pe.Descripcion,'No Localizado') as 'EstadoPersona'" _
           & " from AFI_PADRON P left join Socios S on S.CEDULAR in(P.ID_ALTERNO, P.CEDULA)" _
           & " left join AFI_ESTADOS_PERSONA Pe on S.EstadoActual = Pe.COD_ESTADO"
Else
    strSQL = "select P.* , isnull(Pe.Descripcion,'No Localizado') as 'EstadoPersona'" _
           & " from AFI_PADRON P inner join Socios S on S.CEDULAR in(P.ID_ALTERNO, P.CEDULA)" _
           & " inner join AFI_ESTADOS_PERSONA Pe on S.EstadoActual = Pe.COD_ESTADO"
        
        If vWhere Then
            strSQL = strSQL & " AND "
        Else
            strSQL = strSQL & " WHERE "
            vWhere = True
        End If
        
        strSQL = strSQL & " S.EstadoActual = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
End If
    



If cbo.Text <> "TODOS" Then
    If vWhere Then
        strSQL = strSQL & " AND "
    Else
        strSQL = strSQL & " WHERE "
        vWhere = True
    End If
    
    strSQL = strSQL & " P.COD_INSTITUCION = " & cbo.ItemData(cbo.ListIndex)
End If
    
If Len(txtCedula.Text) > 0 Then
    If vWhere Then
        strSQL = strSQL & " AND "
    Else
        strSQL = strSQL & " WHERE "
        vWhere = True
    End If
    
    strSQL = strSQL & " P.CEDULA like '%" & txtCedula.Text & "%'"
End If
    
    
If Len(txtIdAlterna.Text) > 0 Then
    If vWhere Then
        strSQL = strSQL & " AND "
    Else
        strSQL = strSQL & " WHERE "
        vWhere = True
    End If
    
    strSQL = strSQL & " P.ID_ALTERNO like '%" & txtIdAlterna.Text & "%'"
End If
    
If Len(txtNombre.Text) > 0 Then
    If vWhere Then
        strSQL = strSQL & " AND "
    Else
        strSQL = strSQL & " WHERE "
        vWhere = True
    End If
    
    strSQL = strSQL & " P.NOMBRE like '%" & txtNombre.Text & "%'"
End If
    
 If Len(txtUsuario.Text) > 0 Then
    If vWhere Then
        strSQL = strSQL & " AND "
    Else
        strSQL = strSQL & " WHERE "
        vWhere = True
    End If
    
    strSQL = strSQL & " P.REGISTRO_USUARIO like '%" & txtUsuario.Text & "%'"
End If
       
    
If chkFechaIngreso.Value = xtpChecked Then
    If vWhere Then
        strSQL = strSQL & " AND "
    Else
        strSQL = strSQL & " WHERE "
        vWhere = True
    End If
    
    strSQL = strSQL & " P.FECHA_INGRESO between '" & Format(dtpIngInicio.Value, "yyyy-mm-dd") _
            & " 00:00:00' AND '" & Format(dtpIngCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
End If
    
If chkFechaRegistro.Value = xtpChecked Then
    If vWhere Then
        strSQL = strSQL & " AND "
    Else
        strSQL = strSQL & " WHERE "
        vWhere = True
    End If
    
    strSQL = strSQL & " P.REGISTRO_FECHA between '" & Format(dtpRegInicio.Value, "yyyy-mm-dd") _
            & " 00:00:00' AND '" & Format(dtpRegCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
End If
    
    
    

    
strSQL = strSQL & " order by P.CEDULA"



Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!CEDULA)
      itmX.SubItems(1) = rs!ID_ALTERNO & ""
      itmX.SubItems(2) = rs!Nombre & ""
      itmX.SubItems(3) = Format(rs!FECHA_INGRESO & "", "yyyy-MM-dd")
      itmX.SubItems(4) = rs!REGISTRO_USUARIO & ""
      itmX.SubItems(5) = Format(rs!REGISTRO_FECHA & "", "yyyy-MM-dd")
      itmX.SubItems(6) = rs!EstadoPersona & ""
      
  rs.MoveNext
Loop
rs.Close
    
lblItems.Caption = "Total de Líneas: " & lsw.ListItems.Count
    

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    lsw.ListItems.Clear

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And Len(txtCedula.Text) > 0 Then
    Call sbBuscar
End If
End Sub


Private Sub txtIdAlterna_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And Len(txtIdAlterna.Text) > 0 Then
    Call sbBuscar
End If

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And Len(txtNombre.Text) > 0 Then
    Call sbBuscar
End If
End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And Len(txtUsuario.Text) > 0 Then
    Call sbBuscar
End If
End Sub
