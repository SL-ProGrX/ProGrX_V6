VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmFNDGarantias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Garantias para Crédito"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   10335
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9720
      Top             =   240
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5892
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10092
      _Version        =   1441793
      _ExtentX        =   17801
      _ExtentY        =   10393
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
      Item(0).Caption =   "Garantías"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Contenido"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "cbo"
      Item(1).Control(1)=   "Label2(0)"
      Item(1).Control(2)=   "lsw"
      Item(1).Control(3)=   "gbGarantias"
      Item(1).Control(4)=   "cboEstado"
      Item(1).Control(5)=   "Label2(1)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2652
         Left            =   -69760
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   9804
         _Version        =   1441793
         _ExtentX        =   17293
         _ExtentY        =   4678
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5292
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   8052
         _Version        =   524288
         _ExtentX        =   14203
         _ExtentY        =   9334
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
         MaxCols         =   493
         ScrollBars      =   2
         SpreadDesigner  =   "frmFNDGarantias.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.GroupBox gbGarantias 
         Height          =   1932
         Left            =   -69160
         TabIndex        =   4
         Top             =   3960
         Visible         =   0   'False
         Width           =   8412
         _Version        =   1441793
         _ExtentX        =   14838
         _ExtentY        =   3408
         _StockProps     =   79
         Caption         =   "Registro"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPatrimonio 
            Height          =   312
            Left            =   6600
            TabIndex        =   5
            Top             =   1080
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Patrimonio?   "
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
            Enabled         =   0   'False
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   312
            Index           =   0
            Left            =   2760
            TabIndex        =   6
            Top             =   1440
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFNDGarantias.frx":058B
         End
         Begin XtremeSuiteControls.ComboBox cboOperadora 
            Height          =   312
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   6612
            _Version        =   1441793
            _ExtentX        =   11668
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
         Begin XtremeSuiteControls.FlatEdit txtPlan 
            Height          =   312
            Left            =   1680
            TabIndex        =   8
            Top             =   720
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
            Height          =   312
            Left            =   1680
            TabIndex        =   9
            Top             =   1440
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlanDesc 
            Height          =   312
            Left            =   2640
            TabIndex        =   10
            Top             =   720
            Width           =   5652
            _Version        =   1441793
            _ExtentX        =   9970
            _ExtentY        =   550
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   312
            Index           =   1
            Left            =   3120
            TabIndex        =   11
            Top             =   1440
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFNDGarantias.frx":0CAB
         End
         Begin XtremeSuiteControls.FlatEdit txtLinea 
            Height          =   312
            Left            =   5520
            TabIndex        =   12
            Top             =   1440
            Visible         =   0   'False
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMemInicio 
            Height          =   312
            Left            =   1680
            TabIndex        =   13
            Top             =   1080
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMemCorte 
            Height          =   312
            Left            =   2640
            TabIndex        =   14
            Top             =   1080
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Operadora"
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
            Index           =   4
            Left            =   600
            TabIndex        =   19
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   3
            Left            =   600
            TabIndex        =   18
            Top             =   1440
            Width           =   972
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan"
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
            Left            =   600
            TabIndex        =   17
            Top             =   720
            Width           =   852
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Membresía"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   5
            Left            =   600
            TabIndex        =   16
            Top             =   1080
            Width           =   972
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Rango en Días"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   372
            Index           =   6
            Left            =   3720
            TabIndex        =   15
            Top             =   1080
            Width           =   3252
         End
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   -69160
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   4212
         _Version        =   1441793
         _ExtentX        =   7435
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   -64960
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   4212
         _Version        =   1441793
         _ExtentX        =   7435
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Garantía"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   -69160
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Index           =   1
         Left            =   -64960
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   2412
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Definición de Garantías sobre Ahorros Extraordinarios"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   7932
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFNDGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub btnMov_Click(Index As Integer)
Dim strSQL As String, strBitacora As String
Dim pOperadora As Integer, pPlan As String, pGarantia As String, pEstado As String

On Error GoTo vError


If Not IsNumeric(txtPorcentaje.Text) Then
   MsgBox "Porcentaje no es válido", vbExclamation
   Exit Sub
Else
  If CCur(txtPorcentaje.Text) < 0 Or CCur(txtPorcentaje.Text) > 999 Then
     MsgBox "Porcentaje no es válido", vbExclamation
     Exit Sub
  End If
End If

If Not IsNumeric(txtMemInicio.Text) Then
   MsgBox "Rango de Inicio la membresía no es válido", vbExclamation
   Exit Sub
End If

If Not IsNumeric(txtMemCorte.Text) Then
   MsgBox "Rango de Corte la membresía no es válido", vbExclamation
   Exit Sub
End If


pOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
pPlan = txtPlan.Text
pGarantia = cbo.ItemData(cbo.ListIndex)
pEstado = cboEstado.ItemData(cboEstado.ListIndex)

strBitacora = "Garantia de s/Ahorros Extra, Linea: " & txtLinea.Text & ", Gar: " & pGarantia & ", Est: " & pEstado _
         & " Plan : " & pPlan & " Porcentaje : " & CCur(txtPorcentaje.Text) _
         & ", Mem.I: " & txtMemInicio.Text & ", Mem.C: " & txtMemCorte.Text

strSQL = "exec spFnd_Garantia_Ahorros_Registro '" & pGarantia & "','" & pEstado & "'," & txtLinea _
       & "," & txtMemInicio.Text & "," & txtMemCorte.Text & "," & chkPatrimonio.Value _
       & "," & pOperadora & ",'" & pPlan & "'," & CCur(txtPorcentaje.Text) & ",'" & glogon.Usuario & "'"

Select Case Index
 Case 0 'Agregar / Modificar
       
    strSQL = strSQL & ",'A'"
    Call ConectionExecute(strSQL)
       
    Call Bitacora("Registra", strBitacora)
    
 
 Case 1 'Elimina

    strSQL = strSQL & ",'E'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Elimina", strBitacora)
    
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub

If cbo.ListCount = 0 Or cboEstado.ListCount = 0 Then Exit Sub
If cbo.Text = "" Or cboEstado.Text = "" Then Exit Sub

Call sbConsulta

End Sub

Private Sub cboEstado_Click()
If vPaso Then Exit Sub

If cbo.ListCount = 0 Or cboEstado.ListCount = 0 Then Exit Sub
If cbo.Text = "" Or cboEstado.Text = "" Then Exit Sub

Call sbConsulta
End Sub

Private Sub sbConsulta()
Dim strSQL As String, rs  As New ADODB.Recordset
Dim vGarantia As String, vEstado As String
Dim itmX As ListViewItem

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

If cbo.ListCount = 0 Or cboEstado.ListCount = 0 Then Exit Sub
If cbo.Text = "" Or cboEstado.Text = "" Then Exit Sub

txtPlan.Text = ""
txtPlanDesc.Text = ""
txtPorcentaje.Text = Format(0, "Standard")
txtMemInicio.Text = "0"
txtMemCorte.Text = "99999"
txtLinea.Text = "0"
chkPatrimonio.Value = xtpUnchecked


vGarantia = cbo.ItemData(cbo.ListIndex)
vEstado = cboEstado.ItemData(cboEstado.ListIndex)

strSQL = "exec spFnd_Garantia_Ahorros_Consulta '" & vGarantia & "','" & vEstado & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Linea_Id)
      itmX.SubItems(1) = rs!cod_Operadora
      itmX.SubItems(2) = rs!cod_Plan
      itmX.SubItems(3) = rs!DESCRIPCION
      itmX.SubItems(4) = rs!MEMBRESIA_INICIO
      itmX.SubItems(5) = rs!MEMBRESIA_CORTE
      itmX.SubItems(6) = Format(rs!Porcentaje, "Standard")
      itmX.SubItems(7) = IIf(rs!Patrimonio = 1, "Sí", "No")

  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub Form_Activate()
vModulo = 18

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

txtLinea.Text = Item.Text
txtPlan.Text = Item.SubItems(2)
txtPlanDesc.Text = Item.SubItems(3)
txtPorcentaje.Text = Item.SubItems(6)

txtMemInicio.Text = Item.SubItems(4)
txtMemCorte.Text = Item.SubItems(5)

chkPatrimonio.Value = IIf(Mid(Item.SubItems(7), 1, 1) = "S", xtpChecked, xtpUnchecked)
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 1 Then
    vPaso = True

    strSQL = "select Garantia_FND as 'Idx',Descripcion as 'ItmX'" _
           & " from FND_GARANTIAS"
    
    Call sbCbo_Llena_New(cbo, strSQL, False, True)
    vPaso = False
End If

Call cbo_Click

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

vPaso = True

tcMain.Item(0).Selected = True

strSQL = "select Garantia_FND,Descripcion,Activa" _
       & " from FND_GARANTIAS"
Call sbCargaGrid(vGrid, 3, strSQL)

strSQL = "select rtrim(cod_Operadora) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  fnd_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)


strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  afi_Estados_Persona"
Call sbCbo_Llena_New(cboEstado, strSQL, False, True)


vPaso = False

End Sub



Private Sub Form_Load()

vModulo = 18

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "[Op]", 600, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Inicio", 1200, vbCenter
    .Add , , "Corte", 1200, vbCenter
    .Add , , "Porcentaje", 1200, vbRightJustify
    .Add , , "Patrimonio", 1100, vbCenter
End With


Call Formularios(Me)
Call RefrescaTags(Me)

btnMov(0).Enabled = vGrid.Enabled
btnMov(1).Enabled = vGrid.Enabled

End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

fxValida = True

vMensaje = ""


If Len(vMensaje) > 0 Then
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not fxValida Then
   fxGuardar = 0
   Exit Function
End If

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from fnd_garantias where garantia_fnd = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then   'Insertar
  
  vGrid.col = 1
  strSQL = "insert fnd_garantias(garantia_fnd,descripcion,activa) values('" & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ")"
  Call ConectionExecute(strSQL)
  
  
  vGrid.col = 1
  
  Call Bitacora("Registra", "Garantia de Fondo : " & vGrid.Text)
  
  fxGuardar = 1
  
Else 'Actualizar

    vGrid.col = 2
    strSQL = "update fnd_garantias set descripcion = '" & vGrid.Text & "', activa = "
    vGrid.col = 3
    strSQL = strSQL & vGrid.Value & " where garantia_fnd = '"
    vGrid.col = 1
    strSQL = strSQL & vGrid.Text & "'"
    
    Call ConectionExecute(strSQL)
    
    fxGuardar = 1
    
    Call Bitacora("Modifica", "Garantia de Fondo : " & vGrid.Text)
 
End If

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
        End If
  End If 'Actualiza o Inserta
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1

       If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.col = 1
        strSQL = "delete fnd_garantias where garantia_fnd = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        
        Call Bitacora("Elimina", "Garantia de Fondos : " & vGrid.Text)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If
End If

Exit Sub

vError:

  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
