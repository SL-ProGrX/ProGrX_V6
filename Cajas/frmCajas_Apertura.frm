VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_Apertura 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Apertura de Cajas"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbApertura 
      Height          =   4815
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   8535
      _Version        =   1572864
      _ExtentX        =   15055
      _ExtentY        =   8493
      _StockProps     =   79
      Caption         =   "Apertura:"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1335
         Left            =   2520
         TabIndex        =   24
         Top             =   3360
         Width           =   5775
         _Version        =   1572864
         _ExtentX        =   10186
         _ExtentY        =   2355
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkAP_Compartida 
         Height          =   252
         Left            =   5640
         TabIndex        =   23
         Top             =   2880
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Compartida ?"
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
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   1332
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   5892
         _Version        =   524288
         _ExtentX        =   10393
         _ExtentY        =   2350
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "frmCajas_Apertura.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtAP_Numero 
         Height          =   312
         Left            =   2520
         TabIndex        =   16
         Top             =   1800
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAP_Estado 
         Height          =   312
         Left            =   5400
         TabIndex        =   17
         Top             =   1800
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAP_Fecha 
         Height          =   312
         Left            =   2520
         TabIndex        =   18
         Top             =   2160
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAP_Usuario 
         Height          =   312
         Left            =   5400
         TabIndex        =   19
         Top             =   2160
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAP_EnUso_Fecha 
         Height          =   312
         Left            =   2520
         TabIndex        =   20
         Top             =   2520
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAP_EnUso_Usuario 
         Height          =   312
         Left            =   5400
         TabIndex        =   21
         Top             =   2520
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAP_Vence 
         Height          =   312
         Left            =   2520
         TabIndex        =   22
         Top             =   2880
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aprovisionamientos:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   8
         Left            =   360
         TabIndex        =   25
         Top             =   3600
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicial en Cajas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   792
         Index           =   7
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ultima Apertura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   1332
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "En Uso"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   5
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   1332
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento ?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   6
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   1332
      End
   End
   Begin XtremeSuiteControls.PushButton btnApertura 
      Height          =   495
      Index           =   0
      Left            =   5520
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Apertura"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCajas_Apertura.frx":05D3
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnApertura 
      Height          =   495
      Index           =   1
      Left            =   6960
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCajas_Apertura.frx":0CFA
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8916
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   960
      TabIndex        =   6
      Top             =   1200
      Width           =   2532
      _Version        =   1572864
      _ExtentX        =   4466
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   5160
      TabIndex        =   7
      Top             =   1200
      Width           =   3252
      _Version        =   1572864
      _ExtentX        =   5736
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmCajas_Apertura.frx":1410
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Caja Asignada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   -120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   492
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   8772
      _Version        =   1572864
      _ExtentX        =   15473
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmCajas_Apertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub btnApertura_Click(Index As Integer)
On Error GoTo vError

Select Case Index
   Case 0 'Aplicar
     If vGrid.Enabled Then
         Call sbCreaApertura
     Else
        MsgBox "Digite una clave correcta para desbloquer?", vbExclamation
     End If
   
   Case 1 'Cancelar
     Unload Me
     Exit Sub
End Select

Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbTEF_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "exec spCajas_TE_Consulta '" & cbo.ItemData(cbo.ListIndex) & "', 'D', '', 'P', Null, Null"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!TRASLADO_ID)
     itmX.SubItems(1) = rs!cod_Divisa
     itmX.SubItems(2) = Format(rs!Importe, "Standard")
     itmX.SubItems(3) = rs!Cod_Caja
     itmX.SubItems(4) = rs!REGISTRO_USUARIO
     itmX.SubItems(5) = rs!Cod_Apertura
     itmX.SubItems(6) = rs!REGISTRO_FECHA
     itmX.SubItems(7) = rs!TIPO_CAMBIO
     itmX.SubItems(8) = Format(rs!Monto, "Standard")
     itmX.SubItems(9) = rs!Notas


 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub cbo_Click()


If vPaso Or cbo.ListCount = 0 Then Exit Sub


strSQL = "select *, Case when Estado = 'A' then 'Abierta' else 'Cerrada' end as 'Estado'" _
       & " from Cajas_Aperturas_Main" _
       & " where cod_Caja = '" & cbo.ItemData(cbo.ListIndex) _
       & "' and Cod_Apertura in(select max(cod_Apertura) from Cajas_Aperturas_Main" _
       & " where cod_Caja = '" & cbo.ItemData(cbo.ListIndex) & "')"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtAP_Numero.Text = rs!Cod_Apertura
  txtAP_Estado.Text = rs!Estado
  txtAP_Fecha.Text = rs!Apertura_Fecha
  txtAP_Usuario.Text = rs!Apertura_Usuario
  
  txtAP_EnUso_Fecha.Text = rs!En_Uso_Fecha & ""
  txtAP_EnUso_Usuario.Text = rs!En_Uso_Usuario & ""
  
  txtAP_Vence.Text = rs!Apertura_Vence & ""
  chkAP_Compartida.Value = rs!Apertura_Compartida
  
Else
  txtAP_Numero.Text = "0"
  txtAP_Estado.Text = ""
  txtAP_Fecha.Text = ""
  txtAP_Usuario.Text = ""

  txtAP_EnUso_Fecha.Text = ""
  txtAP_EnUso_Usuario.Text = ""
  
  txtAP_Vence.Text = ""
  chkAP_Compartida.Value = vbUnchecked
End If
rs.Close

Call sbTEF_Consulta


End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Sub Form_Load()

vModulo = 5
 
txtUsuario = glogon.Usuario
txtClave = ""

ModuloCajas.mApertura = 0

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Tramite Id", 1400
    .Add , , "Divisa", 1000, vbCenter
    .Add , , "Importe", 2500, vbRightJustify
    .Add , , "Caja Id", 1400, vbCenter
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Apertura Id", 1400, vbCenter
    .Add , , "Fecha", 2500
    .Add , , "T.C.", 2500, vbRightJustify
    .Add , , "Monto", 2500, vbRightJustify
    .Add , , "Notas", 3500
End With

strSQL = "select rtrim(C.cod_caja) as 'Idx',rtrim(C.Descripcion) as itmX" _
        & " from cajas_definicion C inner join cajas_usuarios U on C.cod_caja = U.cod_caja and U.usuario = '" & glogon.Usuario & "'" _
        & " where C.Activa = 1 order by C.cod_caja"
vPaso = True
    
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vPaso = False

Call cbo_Click

strSQL = "Select cod_divisa,0 as 'Efectivo',0 as 'Documentos'" _
       & " from CNTX_DIVISAS where COD_CONTABILIDAD = " & GLOBALES.gEnlace
Call sbCargaGrid(vGrid, 3, strSQL)

vGrid.Enabled = False
If vGrid.MaxRows > 0 Then vGrid.MaxRows = vGrid.MaxRows - 1
  
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyReturn Then
    If cbo.ListCount = 0 Then Exit Sub
    
    strSQL = "select count(*) as 'Aceptado' from cajas_usuarios where usuario= '" & txtUsuario.Text _
            & "' and contrasena = '" & SIFGlobal.fxStringCifrado(txtClave.Text) _
            & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "'"
    
    Call OpenRecordSet(rs, strSQL)
    
    If rs!aceptado > 0 Then
         'verifica que la caja no tenga ninguna apertura
         rs.Close
         vGrid.Enabled = True
         vGrid.SetFocus
         
    Else
       vGrid.Enabled = False
       MsgBox "No se encuentra autorizado para utilizar esta caja...", vbCritical
       rs.Close
    End If
'   rs.Close
End If

End Sub


Private Function fxCuentaDevolucion(vCaja As String) As String

strSQL = "Select cod_cuenta_dev from cajas_definicion where cod_caja = '" & vCaja & "'"
Call OpenRecordSet(rs, strSQL)
fxCuentaDevolucion = rs!Cod_Cuenta_Dev
rs.Close

End Function

Private Function fxValidaCaja(vCaja As String) As Boolean
Dim vMensaje As String

vMensaje = ""
fxValidaCaja = True

strSQL = "select count(*) as Existe from CAJAS_FORMAS_PAGO where cod_caja = '" & vCaja & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then
  vMensaje = "Aun no se definen formas de pago para esta caja..."
End If
rs.Close

strSQL = "select count(*) as Existe from CAJAS_DOCUMENTOS where cod_caja ='" & vCaja & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then
  vMensaje = "Aun no se definen documentos para esta caja..."
End If
rs.Close

strSQL = "select count(*) as Existe from cajas_servicios_asignados where cod_caja ='" & vCaja & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then
  vMensaje = "Aun no se definen servicios para esta caja..."
End If
rs.Close


If Len(vMensaje) > 0 Then
  fxValidaCaja = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbCreaApertura()
Dim i As Integer, iDiasVence As Long

On Error GoTo vError

If cbo.ListCount = 0 Then Exit Sub
If Not fxValidaCaja(cbo.ItemData(cbo.ListIndex)) Then Exit Sub


strSQL = "Select count(*) as existe from cajas_aperturas_main where cod_caja  = '" & cbo.ItemData(cbo.ListIndex) & "' and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   MsgBox "La caja se encuentra abierta", vbInformation
   Exit Sub
End If
rs.Close
        
'Indicador de Apertura  Compartida entre usuarios
strSQL = "Select Apertura_Compartida,Cierre_Periocidad" _
       & " from cajas_definicion" _
       & " where cod_caja  = '" & cbo.ItemData(cbo.ListIndex) & "' and activa = 1"
Call OpenRecordSet(rs, strSQL)
i = rs!Apertura_Compartida
Select Case Trim(rs!Cierre_Periocidad)
  Case "A" 'Abierto
    iDiasVence = 0
  Case "D" 'Diario
    iDiasVence = 1
  Case "S" 'Semanal
    iDiasVence = 7
  Case "Q" 'Quincenal
    iDiasVence = 15
  Case "M" 'Mensual
    iDiasVence = 30
  Case Else
    iDiasVence = 0
End Select

rs.Close
        
        

strSQL = "select isnull(max(cod_apertura),0) as 'Ultimo' from cajas_aperturas_main where cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
ModuloCajas.mApertura = rs!ultimo + 1
rs.Close

strSQL = "insert cajas_aperturas_main(cod_apertura,cod_caja,apertura_usuario,apertura_fecha,apertura_compartida,apertura_vence, estado)" _
       & " values(" & ModuloCajas.mApertura & ",'" & cbo.ItemData(cbo.ListIndex) & "','" & glogon.Usuario _
       & "',dbo.MyGetdate()," & i

If iDiasVence = 0 Then
   strSQL = strSQL & ",NULL,'A')"
Else
   strSQL = strSQL & ",dateadd(d," & iDiasVence & ",dbo.MyGetdate()),  'A')"
End If
       
Call ConectionExecute(strSQL)


For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 1
    
    If Trim(vGrid.Text) <> "" Then
        strSQL = "insert cajas_aperturas_cierres(cod_apertura,cod_caja,si_efectivo,si_documentos,cod_divisa)" _
               & " values(" & ModuloCajas.mApertura & ",'" & cbo.ItemData(cbo.ListIndex) & "',"
        vGrid.Col = 2
        strSQL = strSQL & " " & CCur(vGrid.Text) & ", "
        vGrid.Col = 3
        strSQL = strSQL & "" & CCur(vGrid.Text) & ","
        vGrid.Col = 1
        strSQL = strSQL & "'" & Trim(vGrid.Text) & "')"
        
        Call ConectionExecute(strSQL)
        
     End If
Next i


'strSQL = "update cajas_aperturas_main set estado = 'B' where cod_apertura = " & ModuloCajas.mApertura & " and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "'"
'Call ConectionExecute(strSQL)


'---Procesa Entredas> Aprovisionamientos

With lsw.ListItems
For i = 1 To .Count
  If .Item(i).Checked Then
        
        strSQL = "exec spCajas_TE_Resolucion " & .Item(i).Text & ", 'A', '" & cbo.ItemData(cbo.ListIndex) _
               & "', '" & txtUsuario.Text & "', " & ModuloCajas.mApertura & ", '" & glogon.Usuario & "', 1"
        Call ConectionExecute(strSQL)
 End If
Next i

End With


ModuloCajas.mCaja = cbo.ItemData(cbo.ListIndex)
ModuloCajas.mCuentaConta = fxCuentaDevolucion(cbo.ItemData(cbo.ListIndex))

MsgBox "Apertura # " & ModuloCajas.mApertura & " registrada satisfactoriamente...", vbInformation

Unload Me
frmCajas_Acceso.Show vbModal

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

