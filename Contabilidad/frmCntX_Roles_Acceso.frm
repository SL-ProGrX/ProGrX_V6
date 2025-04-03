VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCntX_Roles_Acceso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Roles de Acceso a Cuentas Contables"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   360
      Top             =   120
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltros 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   8415
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   14843
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   4
      Item(0).Caption =   "Roles"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Cuentas"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tcCtas"
      Item(2).Caption =   "Unidades y Centros de Costos"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "lswUnidades"
      Item(2).Control(1)=   "ShortcutCaption1(2)"
      Item(2).Control(2)=   "lswCentros"
      Item(2).Control(3)=   "ShortcutCaption1(3)"
      Item(2).Control(4)=   "scUnidad"
      Item(3).Caption =   "Miembros"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "lswMiembros"
      Item(3).Control(1)=   "ShortcutCaption1(4)"
      Begin XtremeSuiteControls.ListView lswUnidades 
         Height          =   3375
         Left            =   -69880
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441793
         _ExtentX        =   18865
         _ExtentY        =   5953
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
      Begin XtremeSuiteControls.ListView lswCentros 
         Height          =   3375
         Left            =   -69880
         TabIndex        =   7
         Top             =   4920
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441793
         _ExtentX        =   18865
         _ExtentY        =   5953
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
      Begin XtremeSuiteControls.ListView lswMiembros 
         Height          =   7455
         Left            =   -69880
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441793
         _ExtentX        =   18865
         _ExtentY        =   13150
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   7815
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   13785
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_Roles_Acceso.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.TabControl tcCtas 
         Height          =   7935
         Left            =   -70000
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   13996
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
         Appearance      =   4
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Asignar Cuentas"
         Item(0).ControlCount=   6
         Item(0).Control(0)=   "lswCtas"
         Item(0).Control(1)=   "Label2(2)"
         Item(0).Control(2)=   "txtCtaInicial"
         Item(0).Control(3)=   "txtCtaFinal"
         Item(0).Control(4)=   "btnCtas"
         Item(0).Control(5)=   "chkCtaTodas(0)"
         Item(1).Caption =   "Cuentas Registradas"
         Item(1).ControlCount=   3
         Item(1).Control(0)=   "lswCtaAsg"
         Item(1).Control(1)=   "chkCtaTodas(1)"
         Item(1).Control(2)=   "btnCtaElimina"
         Begin XtremeSuiteControls.ListView lswCtaAsg 
            Height          =   7095
            Left            =   -69880
            TabIndex        =   12
            Top             =   840
            Visible         =   0   'False
            Width           =   10695
            _Version        =   1441793
            _ExtentX        =   18865
            _ExtentY        =   12515
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
         Begin XtremeSuiteControls.ListView lswCtas 
            Height          =   6975
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   10695
            _Version        =   1441793
            _ExtentX        =   18865
            _ExtentY        =   12303
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
         Begin XtremeSuiteControls.CheckBox chkCtaTodas 
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   210
            _Version        =   1441793
            _ExtentX        =   370
            _ExtentY        =   370
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnCtas 
            Height          =   315
            Left            =   7800
            TabIndex        =   14
            Top             =   480
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Consultar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit txtCtaInicial 
            Height          =   315
            Left            =   2880
            TabIndex        =   15
            Top             =   480
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4254
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtCtaFinal 
            Height          =   315
            Left            =   5280
            TabIndex        =   16
            Top             =   480
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4254
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.CheckBox chkCtaTodas 
            Height          =   210
            Index           =   1
            Left            =   -69760
            TabIndex        =   17
            Top             =   525
            Visible         =   0   'False
            Width           =   210
            _Version        =   1441793
            _ExtentX        =   370
            _ExtentY        =   370
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnCtaElimina 
            Height          =   315
            Left            =   -64120
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5524
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Eliminar las Seleccionadas"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
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
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   19
            Top             =   480
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   339
            _StockProps     =   79
            Caption         =   "Rango de Cuentas:"
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
            UseMnemonic     =   0   'False
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption scUnidad 
         Height          =   375
         Left            =   -67480
         TabIndex        =   23
         Top             =   4440
         Visible         =   0   'False
         Width           =   8295
         _Version        =   1441793
         _ExtentX        =   14631
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione una Unidad!"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   2
         Left            =   -69880
         TabIndex        =   22
         Top             =   4440
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Centros de Costos para:"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   3
         Left            =   -69880
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441793
         _ExtentX        =   18865
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Unidades de Negocios:"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   4
         Left            =   -69880
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441793
         _ExtentX        =   18865
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Usuarios miembros del Rol de acceso:"
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
   End
   Begin XtremeSuiteControls.ProgressBar prgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   24
      Top             =   10560
      Width           =   11175
      _Version        =   1441793
      _ExtentX        =   19711
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Roles de Acceso a Catálogo Contable"
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
      Height          =   480
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   7335
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Rol Activo:"
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
   Begin XtremeShortcutBar.ShortcutCaption scRol 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Seleccione!"
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Filtro:"
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
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCntX_Roles_Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnCtaElimina_Click()
If vPaso Then Exit Sub


Dim i As Long

On Error GoTo vError

With lswCtaAsg.ListItems

Me.MousePointer = vbHourglass

strSQL = ""
For i = 1 To .Count
    If .Item(i).Checked Then
        strSQL = strSQL & Space(10) & "exec spCntX_AC_Cuentas_Asigna " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag _
               & "', '" & .Item(i).Text & "','" & glogon.Usuario & "','E'"
    End If

    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If

Next i

End With

'Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

Me.MousePointer = vbDefault

'Actualiza la Lista
Call sbCuentas_Consulta_Asg

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCtas_Click()
Call sbCuentas_Consulta
End Sub



Private Sub chkCtaTodas_Click(Index As Integer)
Dim i As Long

Select Case Index
    Case 0 'Cta Nuevas
        With lswCtas.ListItems
                
        Me.MousePointer = vbHourglass
        
        For i = 1 To .Count
            If .Item(i).Checked <> chkCtaTodas.Item(Index).Value Then
                .Item(i).Checked = chkCtaTodas.Item(Index).Value
            End If
        Next i
                
        Me.MousePointer = vbDefault
        
        End With
    
    Case 1 'Ctas Asignadas
        With lswCtaAsg.ListItems
                
        Me.MousePointer = vbHourglass
        
        For i = 1 To .Count
                .Item(i).Checked = chkCtaTodas.Item(Index).Value
        Next i
                
        Me.MousePointer = vbDefault
        
        End With
End Select

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Public Sub sbCargaGrid_Local(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim i As Integer

On Error GoTo vErrorLoad

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
  
vGrid.MaxRows = 1
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
  
    vGrid.Col = i
    Select Case i
        Case 1
            vGrid.Text = Trim(rs!Cod_Rol)
        Case 2
            vGrid.Text = Trim(rs!Descripcion)
        Case 3
            vGrid.Value = rs!Control
        Case 4
            vGrid.Value = rs!Activo
    End Select

  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

Exit Sub

vErrorLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub



Private Sub sbConsulta()

tcMain.Item(0).Selected = True

vPaso = True

strSQL = "exec spCntX_AC_Rol_List " & gCntX_Parametros.CodigoConta & ", '" & glogon.Usuario & "'"

Call sbCargaGrid_Local(vGrid, vGrid.MaxCols, strSQL)

vPaso = False

End Sub


Private Sub Form_Load()

vModulo = 20


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.MaxCols = 5
vGrid.MaxRows = 1

With lswCtas.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Descripción", lswCtas.Width - (2700)
End With

With lswUnidades.ColumnHeaders
    .Clear
    .Add , , "Unidad", 2500
    .Add , , "Descripción", lswUnidades.Width - (2700)
End With

With lswCentros.ColumnHeaders
    .Clear
    .Add , , "Centro", 2500
    .Add , , "Descripción", lswCentros.Width - (2700)
End With

With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Nombre", lswMiembros.Width - (2700)
End With


vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

'Seguridad a Todos los componentes
lswCentros.Enabled = vGrid.Enabled
lswUnidades.Enabled = vGrid.Enabled
lswCtas.Enabled = vGrid.Enabled
lswCtaAsg.Enabled = vGrid.Enabled
lswMiembros.Enabled = vGrid.Enabled

btnCtaElimina.Enabled = vGrid.Enabled


End Sub




Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text <> "" Then
    strSQL = "exec spCntX_AC_Rol_Add " & gCntX_Parametros.CodigoConta & ", '" & Trim(vGrid.Text) & "', '"
    vGrid.Col = 2
    strSQL = strSQL & Trim(vGrid.Text) & "', "
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Value & ", "
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Value & ", '" & glogon.Usuario & "'"
    
    Call ConectionExecute(strSQL)
    
    fxGuardar = 1
End If


Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function


Private Sub sbCuentas_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCntX_AC_Cuentas_Consulta " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag & "', '" & txtCtaInicial.Text _
       & "', '" & txtCtaFinal.Text & "', '" & txtFiltros.Text & "', '" & glogon.Usuario & "'"
       
lswCtas.ListItems.Clear
With lswCtas.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Descripción", lswCtas.Width - 2800
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswCtas.ListItems.Add(, , rs!Cod_Cuenta_Mask)
     itmX.SubItems(1) = rs!Descripcion
     
     If rs!Acepta_movimientos = 0 Then
        itmX.Bold = True
        itmX.TextBackColor = RGB(176, 211, 238)
     End If
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

vPaso = False

prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbCuentas_Consulta_Asg()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spCntX_AC_Cuentas_Consulta_Asignadas " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag & "', '" & txtFiltros.Text & "', '" & glogon.Usuario & "'"
       
lswCtaAsg.ListItems.Clear
With lswCtaAsg.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Descripción", lswCtaAsg.Width - 2800
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswCtaAsg.ListItems.Add(, , rs!Cod_Cuenta_Mask)
     itmX.SubItems(1) = rs!Descripcion
     
     If rs!Acepta_movimientos = 0 Then
        itmX.Bold = True
        itmX.TextBackColor = RGB(176, 211, 238)
     End If
     
  prgBar.Value = prgBar.Value + 1
  
 rs.MoveNext
Loop
rs.Close


vPaso = False


prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub lswCentros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spCntX_AC_Centro_Costo_Asigna " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag & "', '" & scUnidad.Tag _
       & "', '" & Item.Text & "', '" & glogon.Usuario _
       & "', '" & IIf(Item.Checked, "A", "E") & "'"
Call ConectionExecute(strSQL)


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswCtas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spCntX_AC_Cuentas_Asigna " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag _
       & "', '" & Item.Text & "', '" & glogon.Usuario _
       & "', '" & IIf(Item.Checked, "A", "E") & "'"
Call ConectionExecute(strSQL)


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbUnidades_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCntX_AC_Unidades_Consulta " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag & "', '" & txtFiltros.Text & "', '" & glogon.Usuario & "'"
       
lswUnidades.ListItems.Clear
lswCentros.ListItems.Clear

scUnidad.Tag = ""
scUnidad.Caption = "Seleccione una Unidad!"

With lswUnidades.ColumnHeaders
    .Clear
    .Add , , "Unidad", 2500
    .Add , , "Descripción", lswUnidades.Width - 2800
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswUnidades.ListItems.Add(, , rs!Cod_Unidad)
     itmX.SubItems(1) = rs!Descripcion
     
     If rs!Asignado = 1 Then
        itmX.Bold = True
        itmX.Checked = True
     End If
     
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

vPaso = False

prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spCntX_AC_Miembros_Asigna " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag _
       & "', '" & Item.Text & "', '" & glogon.Usuario _
       & "', '" & IIf(Item.Checked, "A", "E") & "'"
Call ConectionExecute(strSQL)


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswUnidades_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spCntX_AC_Unidades_Asigna " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag _
       & "', '" & Item.Text & "', '" & glogon.Usuario _
       & "', '" & IIf(Item.Checked, "A", "E") & "'"
Call ConectionExecute(strSQL)


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswUnidades_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)


If Item.Checked Then
    scUnidad.Tag = Item.Text
    scUnidad.Caption = Item.SubItems(1)
    
    Call sbCentros_Consulta
Else
    scUnidad.Tag = ""
    scUnidad.Caption = "Seleccione una Unidad Vinculada!"
End If

End Sub



Private Sub sbCentros_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCntX_AC_Centro_Costo_Consulta " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag & "', '" & scUnidad.Tag _
       & "', '" & txtFiltros.Text & "', '" & glogon.Usuario & "'"
       
lswCentros.ListItems.Clear
With lswCentros.ColumnHeaders
    .Clear
    .Add , , "Centro", 2500
    .Add , , "Descripción", lswCentros.Width - 2800
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswCentros.ListItems.Add(, , rs!cod_Centro_Costo)
     itmX.SubItems(1) = rs!Descripcion
     
     If rs!Asignado = 1 Then
        itmX.Bold = True
        itmX.Checked = True
     End If
     
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

vPaso = False

prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbMiembros_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCntX_AC_Miembros_Consulta " & gCntX_Parametros.CodigoConta & ", '" & scRol.Tag & "', '" & txtFiltros.Text & "', '" & glogon.Usuario & "'"
       
lswMiembros.ListItems.Clear

With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Miembro", 3000
    .Add , , "Descripción", lswMiembros.Width - 3200
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswMiembros.ListItems.Add(, , rs!Usuario)
     itmX.SubItems(1) = rs!USUARIO_NOMBRE
     
     If rs!Asignado = 1 Then
        itmX.Bold = True
        itmX.Checked = True
     End If
     
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

vPaso = False

prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub tcCtas_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If vPaso Then Exit Sub

Select Case Item.Index
    Case 0 'Ctas
        Call sbCuentas_Consulta
    Case 1 'Ctas Asignadas
        Call sbCuentas_Consulta_Asg
End Select

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If scRol.Tag = "" And Item.Index > 0 Then
    MsgBox "Consulte un Rol primero!", vbInformation
    tcMain.Item(0).Selected = True
    Exit Sub
End If

txtFiltros.Text = ""

Select Case Item.Index
    Case 1 'Cuentas
        If tcCtas.SelectedItem = 0 Then
            Call sbCuentas_Consulta
        Else
            Call sbCuentas_Consulta_Asg
        End If
    Case 2 'Unidades
        Call sbUnidades_Consulta
    
    Case 3 'Miembros
        Call sbMiembros_Consulta
End Select

End Sub


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

On Error GoTo vError

Call sbConsulta

Exit Sub

vError:

End Sub

Private Sub txtCtaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
     frmCntX_ConsultaCuentas.Show vbModal
     txtCtaFinal = fxCntX_CuentaFormato(True, gCuenta, 0)
End If
End Sub

Private Sub txtCtaInicial_KeyDown(KeyCode As Integer, Shift As Integer)
     
If KeyCode = vbKeyF4 Then
     frmCntX_ConsultaCuentas.Show vbModal
     txtCtaInicial = fxCntX_CuentaFormato(True, gCuenta, 0)
End If

End Sub



Private Sub txtFiltros_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
   txtFiltros.Text = fxSysCleanTxtInject(txtFiltros.Text)
   
    Select Case tcMain.SelectedItem
        Case 1 'Cuentas
            
            tcCtas.Item(0).Selected = True
            Call sbCuentas_Consulta
                
        Case 2 'Unidades
            Call sbUnidades_Consulta
        
        Case 3 'Miembros
            Call sbMiembros_Consulta
    End Select

End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

vGrid.Row = Row

vGrid.Col = 1
scRol.Tag = vGrid.Text

vGrid.Col = 2
scRol.Caption = vGrid.Text

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer


If vGrid.ActiveCol = vGrid.MaxCols - 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then
    MsgBox "No se puede registrar, verifique los datos!", vbExclamation
    Exit Sub
  End If
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "exec spCntX_AC_Rol_Delete " & gCntX_Parametros.CodigoConta & ", '" & vGrid.Text _
               & "', '" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        Call sbConsulta
     
     End If
End If

End Sub





