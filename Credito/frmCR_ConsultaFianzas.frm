VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCR_ConsultaFianzas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fianzas y Traslados de Deudas de la persona"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11880
   HelpContextID   =   3019
   Icon            =   "frmCR_ConsultaFianzas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2292
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11652
      _Version        =   1441793
      _ExtentX        =   20553
      _ExtentY        =   4043
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
      HideSelection   =   0   'False
      View            =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   7812
      _Version        =   1441793
      _ExtentX        =   13779
      _ExtentY        =   2773
      _StockProps     =   79
      Caption         =   "Estado del Deudor: "
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtECDeudor 
         Height          =   312
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6583
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtECMembresia 
         Height          =   312
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6583
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtECCategoria 
         Height          =   312
         Left            =   1200
         TabIndex        =   14
         Top             =   1080
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6583
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtECOperaciones 
         Height          =   312
         Left            =   6000
         TabIndex        =   15
         Top             =   360
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtECSaldos 
         Height          =   312
         Left            =   6000
         TabIndex        =   16
         Top             =   720
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtECCuotas 
         Height          =   312
         Left            =   6000
         TabIndex        =   17
         Top             =   1080
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1092
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Membresía"
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
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Deudor"
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
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas"
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
         Index           =   2
         Left            =   5160
         TabIndex        =   8
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   5160
         TabIndex        =   7
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Op's"
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
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   972
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   11652
      _Version        =   1441793
      _ExtentX        =   20553
      _ExtentY        =   1714
      _StockProps     =   79
      Caption         =   "Resumen de la Fianzas: "
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkConSaldo 
         Height          =   252
         Left            =   2640
         TabIndex        =   21
         Top             =   600
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Canceladas?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   432
         Left            =   4200
         TabIndex        =   22
         Top             =   480
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   762
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuotas 
         Height          =   432
         Left            =   6480
         TabIndex        =   23
         Top             =   480
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   762
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuotas"
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
         Height          =   252
         Left            =   6480
         TabIndex        =   20
         Top             =   240
         Width           =   2172
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldos"
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
         Height          =   252
         Index           =   0
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   2292
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fianzas"
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
         Height          =   252
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   2292
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   11160
      Top             =   360
   End
   Begin XtremeSuiteControls.GroupBox gbMora 
      Height          =   1572
      Left            =   8040
      TabIndex        =   3
      Top             =   4080
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6583
      _ExtentY        =   2773
      _StockProps     =   79
      Caption         =   "Cuotas en Mora: "
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      BorderStyle     =   1
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
         _Version        =   524288
         _ExtentX        =   6376
         _ExtentY        =   2355
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   497
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_ConsultaFianzas.frx":030A
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
   Begin XtremeSuiteControls.PushButton btnMain 
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   1092
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Todos"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnMain 
      Height          =   372
      Index           =   1
      Left            =   1440
      TabIndex        =   25
      Top             =   1092
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Fianzas"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnMain 
      Height          =   372
      Index           =   2
      Left            =   2760
      TabIndex        =   26
      Top             =   1092
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Traslados"
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
      Appearance      =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   384
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   11652
      _Version        =   1441793
      _ExtentX        =   20553
      _ExtentY        =   677
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.38
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
      Alignment       =   2
   End
   Begin VB.Label LblRegistroActual 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "................"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   528
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "5"
      Top             =   156
      Width           =   11652
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmCR_ConsultaFianzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcurTotalFianzas As Currency, mcurTotalCuotas As Currency


Private Sub sbCargaFianzas(Optional pTipo As String = "F")
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass
mcurTotalFianzas = 0
mcurTotalCuotas = 0

strSQL = "exec spCrd_Consulta_Fianzas_Rsm '" & GLOBALES.gCedulaActual & "'"

If chkConSaldo.Value = 1 Then
    strSQL = strSQL & ",'C'"
Else
    strSQL = strSQL & ",'A'"
End If

strSQL = strSQL & ",'" & pTipo & "'"

Call OpenRecordSet(rs, strSQL)
      
With lsw
  .ListItems.Clear

   Do While Not rs.EOF
         Set itmX = .ListItems.Add(, , rs!Tipo)
             itmX.SubItems(1) = rs!Id_Solicitud
             itmX.SubItems(2) = rs!Codigo
             itmX.SubItems(3) = rs!Nfiadores
             itmX.SubItems(4) = rs!Cedula
             itmX.SubItems(5) = rs!Nombre
             itmX.SubItems(6) = Format(rs!montoapr, "Standard")
             itmX.SubItems(7) = Format(rs!Saldo, "Standard")
             itmX.SubItems(8) = Format(rs!Cuota, "Standard")
             itmX.SubItems(9) = "[" & rs!moraCta & "] " & Format(rs!MoraMnt, "Standard")
             mcurTotalFianzas = mcurTotalFianzas + rs!Saldo
             mcurTotalCuotas = mcurTotalCuotas + rs!Cuota
             
             If rs!moraCta > 0 Or rs!Proceso <> "N" Then
                itmX.ForeColor = vbRed
                itmX.Bold = True
                itmX.TextBackColor = RGB(250, 219, 216)
             End If
             
        rs.MoveNext
    Loop
    rs.Close
End With

txtSaldo.Text = Format(mcurTotalFianzas, "Standard")
txtCuotas.Text = Format(mcurTotalCuotas, "Standard")

scTitulo.Caption = lblTitulo.Caption & " [ Casos: " & lsw.ListItems.Count & "   Saldos: " & Format(mcurTotalFianzas, "Standard") & " ]"


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnMain_Click(Index As Integer)
Dim i As Integer


For i = 0 To btnMain.Count - 1
   btnMain.Item(i).Checked = False
Next i
btnMain.Item(Index).Checked = True


Select Case btnMain.Item(Index).Caption
  
  Case "Todos"
    gbResumen.Caption = "Resumen de las Fianzas y Traslados:"
    lblTitulo.Caption = "Fianzas y Traslados"
    scTitulo.Caption = lblTitulo.Caption
     Call sbCargaFianzas("X")

  Case "Fianzas"
    gbResumen.Caption = "Resumen de las Fianzas:"
    lblTitulo.Caption = "Fianzas"
    scTitulo.Caption = lblTitulo.Caption
     Call sbCargaFianzas("F")
  
  Case "Traslados"
    gbResumen.Caption = "Resumen de los Traslados de Deudas:"
    lblTitulo.Caption = "Traslados"
    scTitulo.Caption = lblTitulo.Caption
     Call sbCargaFianzas("T")
  
End Select

End Sub

Private Sub chkConSaldo_Click()
 If GLOBALES.gCedulaActual <> "" Then
  Call btnMain_Click(0)
  txtCuotas.Text = Format(mcurTotalCuotas, "Standard")
  txtSaldo.Text = Format(mcurTotalFianzas, "Standard")
 End If
End Sub


Private Sub Form_Load()

vGrid.AppearanceStyle = fxGridStyle
imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
  .Clear
  .Add , , "Tipo", 1300
  .Add , , "Operación", 1300
  .Add , , "Linea", 1100, vbCenter
  .Add , , "No.Fiador", 1100, vbCenter
  .Add , , "Identificación", 1800
  .Add , , "Deudor", 3500
  .Add , , "Monto", 1800, vbRightJustify
  .Add , , "Saldo", 1800, vbRightJustify
  .Add , , "Cuotas", 1800, vbRightJustify
  .Add , , "Morosidad", 2100, vbCenter

End With


End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If lsw.ListItems.Count <= 0 Then Exit Sub

txtECCategoria.Text = ""
txtECCuotas.Text = ""
txtECDeudor.Text = ""
txtECMembresia.Text = ""
txtECSaldos.Text = ""
txtECOperaciones.Text = ""


strSQL = "select S.cedula,S.fechaIngreso,S.EstadoActual,R.proceso" _
       & ",isnull(count(*),0) as Operaciones, isnull(sum(R.saldo),0) as Saldos" _
       & ",isnull(sum(R.Cuota),0) as Cuotas" _
       & ",dbo.fxCRDClasificacion(S.cedula,dbo.MyGetdate()) as Clasificacion" _
       & " from Socios S left join Reg_Creditos R on S.cedula = R.cedula and R.estado = 'A' and R.saldo > 0" _
       & " where S.cedula = '" & Item.SubItems(4) & "'" _
       & " group by S.cedula,S.fechaIngreso,S.EstadoActual,R.proceso"
Call OpenRecordSet(rs, strSQL)

    txtECDeudor.Text = Item.SubItems(4)
    txtECCategoria.Text = "Categoría : " & rs!Clasificacion
    
    If rs!EstadoActual = "S" Then
        txtECMembresia.Text = fxMembresia(rs!FechaIngreso)
    Else
        txtECMembresia.Text = "Esta persona no es Asociado"
    End If
    
    txtECCuotas.Text = Format(rs!Cuotas, "Standard")
    txtECSaldos.Text = Format(rs!Saldos, "Standard")
    txtECOperaciones.Text = rs!Operaciones
 
    If rs!Clasificacion > "B" Or rs!Proceso <> "N" Then
      txtECCategoria.ForeColor = vbRed
    Else
      txtECCategoria.ForeColor = vbBlack
    End If
 
rs.Close


 

'Carga Detalle de Mora
gbMora.Caption = "Mora Operación :" & Item.SubItems(1)
strSQL = "exec spCrd_Operacion_Mora_Cuotas " & Item.SubItems(1)

Call sbCargaGrid(vGrid, 3, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Timer1_Timer()

Me.MousePointer = vbHourglass

If GLOBALES.gCedulaActual <> "" Then
 LblRegistroActual.Caption = GLOBALES.gCedulaActual & "  -  " & fxNombre(GLOBALES.gCedulaActual)
 Call btnMain_Click(0)
End If

Me.MousePointer = vbDefault

Timer1.Interval = 0
vGrid.MaxRows = 0

End Sub
