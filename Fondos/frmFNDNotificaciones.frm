VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDNotificaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Notificaciones"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   10575
      _Version        =   1572864
      _ExtentX        =   18653
      _ExtentY        =   5741
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   8160
      Width           =   10455
      _Version        =   1572864
      _ExtentX        =   18441
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   0
         Left            =   7920
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Nueva"
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
         Picture         =   "frmFNDNotificaciones.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   1
         Left            =   9120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Picture         =   "frmFNDNotificaciones.frx":0632
      End
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   255
      Left            =   8160
      TabIndex        =   14
      Top             =   5040
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Activa Notificación"
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9960
      TabIndex        =   0
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _Version        =   1572864
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   5175
      _Version        =   1572864
      _ExtentX        =   9128
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipoMov 
      Height          =   330
      Left            =   1800
      TabIndex        =   8
      Top             =   5040
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtIdNotifica 
      Height          =   435
      Left            =   1800
      TabIndex        =   9
      Top             =   4440
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   330
      Left            =   5520
      TabIndex        =   12
      Top             =   5040
      Width           =   2055
      _Version        =   1572864
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotifica_Desc 
      Height          =   330
      Left            =   1800
      TabIndex        =   15
      Top             =   5520
      Width           =   8775
      _Version        =   1572864
      _ExtentX        =   15478
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotifica_1 
      Height          =   570
      Left            =   1800
      TabIndex        =   17
      Top             =   6000
      Width           =   8775
      _Version        =   1572864
      _ExtentX        =   15478
      _ExtentY        =   1005
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotifica_2 
      Height          =   570
      Left            =   1800
      TabIndex        =   19
      Top             =   6720
      Width           =   8775
      _Version        =   1572864
      _ExtentX        =   15478
      _ExtentY        =   1005
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotifica_3 
      Height          =   570
      Left            =   1800
      TabIndex        =   21
      Top             =   7440
      Width           =   8775
      _Version        =   1572864
      _ExtentX        =   15478
      _ExtentY        =   1005
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   20
      Top             =   7440
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Notificación No 2"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Notificación No 2"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   6000
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Notificación No 1"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Descripción"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Mnt Supervisado"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo Movimiento"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Id Notificación"
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "frmFNDNotificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vScroll As Boolean


Private Sub sbLimpia()

 txtIdNotifica.Text = ""
 txtMonto.Text = "0"
 chkActivo.Value = xtpUnchecked
 txtNotifica_Desc.Text = ""
 txtNotifica_1.Text = ""
 txtNotifica_2.Text = ""
 txtNotifica_3.Text = ""

End Sub


Private Sub sbNotifica_List()

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

lsw.ListItems.Clear

strSQL = "exec spFnd_Notifica_List " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ", '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!COD_NOTIFICACION)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = rs!Tipo_Mov_Desc
      itmX.SubItems(3) = Format(rs!RANGO, "Standard")
      itmX.SubItems(4) = rs!Activo
      itmX.SubItems(5) = rs!NOTIFICACION1 & ""
      itmX.SubItems(6) = rs!NOTIFICACION2 & ""
      itmX.SubItems(7) = rs!NOTIFICACION3 & ""
      
      itmX.SubItems(8) = rs!Registro_Usuario & ""
      itmX.SubItems(9) = rs!Registro_Fecha & ""
      itmX.SubItems(10) = rs!MODIFICA_USUARIO & ""
      itmX.SubItems(11) = rs!MODIFICA_FECHA & ""

  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbNotifica_Load(pNotifica As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spFnd_Notifica_Load '" & pNotifica & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 txtIdNotifica.Text = rs!COD_NOTIFICACION
 txtMonto.Text = Format(rs!RANGO, "Standard")
 chkActivo.Value = rs!Activo
 txtNotifica_Desc.Text = rs!Descripcion
 txtNotifica_1.Text = rs!NOTIFICACION1
 txtNotifica_2.Text = rs!NOTIFICACION2
 txtNotifica_3.Text = rs!NOTIFICACION3
 Call sbCboAsignaDato(cboTipoMov, rs!Tipo_Mov_Desc, True, rs!COD_TIPO_MOVIMENTO)
End If

rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbNotifica_Update()

On Error GoTo vError


txtNotifica_1.Text = fxSysCleanTxtInject(txtNotifica_1.Text)
txtNotifica_2.Text = fxSysCleanTxtInject(txtNotifica_2.Text)
txtNotifica_3.Text = fxSysCleanTxtInject(txtNotifica_3.Text)
txtNotifica_Desc.Text = fxSysCleanTxtInject(txtNotifica_Desc.Text)

If Len(txtNotifica_Desc.Text) = 0 Then
 MsgBox "Indique una Descripción para Esta notificación!", vbExclamation
 Exit Sub
End If

If Not IsNumeric(txtMonto.Text) Then
 MsgBox "El Monto no es válido!", vbExclamation
 Exit Sub
End If



Me.MousePointer = vbHourglass

If txtIdNotifica.Text = "" Or Not IsNumeric(txtIdNotifica.Text) Then
   txtIdNotifica.Text = "0"
End If


strSQL = "exec spFnd_Notifica_Add " & txtIdNotifica.Text & ", " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ", '" & txtCodigo.Text & "', '" & txtNotifica_Desc.Text & "', " & CCur(txtMonto.Text) & ", " & chkActivo.Value _
       & ", " & cboTipoMov.ItemData(cboTipoMov.ListIndex) & ", '" & txtNotifica_1.Text & "', '" & txtNotifica_2.Text _
       & "', '" & txtNotifica_3.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Call Bitacora("Registra", "Notificación Cnf: " & rs!IdNotifica)

Call sbNotifica_List

Me.MousePointer = vbDefault


MsgBox "Notificación Registrada satisfactoriamente!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnAccion_Click(Index As Integer)

Select Case Index
    Case 0 'Nueva
        Call sbLimpia
    
    Case 1 'Guardar
        Call sbNotifica_Update
End Select

End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan,descripcion from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Cod_Plan
      txtDescripcion.Text = rs!Descripcion
      Call sbNotifica_List
    End If
End If

vScroll = False
    FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Call sbNotifica_Load(Item.Text)
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo.Text = Trim(gBusquedas.Resultado)
      txtDescripcion.Text = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub


Private Sub txtCodigo_LostFocus()

If Trim(txtCodigo) <> "" Then
   
   strSQL = "Select Descripcion" _
          & " from fnd_planes where cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & " And cod_plan='" & Trim(txtCodigo) & "'"
   Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
           txtDescripcion.Text = Trim(rs!Descripcion)
        Else
           MsgBox "Codigo incorrecto", vbExclamation
           txtCodigo.Text = ""
           txtDescripcion.Text = ""
           txtCodigo.SetFocus
        End If
     rs.Close

Else
  txtDescripcion.Text = ""
End If

Call sbNotifica_List

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
      
      Call sbNotifica_List
  
   End If
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If


End Sub




Private Sub Form_Load()

vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

strSQL = "select COD_TIPO_MOVIMENTO as 'IdX', DESCRIPCION as 'ItmX'" _
       & " From FND_SEG_TIPOSMOVIMIENTOS"
Call sbCbo_Llena_New(cboTipoMov, strSQL, False, True)

vScroll = False
     FlatScrollBar.Value = 0
vScroll = True

With lsw.ColumnHeaders
    .Add , , "Id Notifica", 1100
    .Add , , "Descripción", 3000
    .Add , , "Tipo Mov", 1800
    .Add , , "Mnt Limite", 2100, vbRightJustify
    .Add , , "Activo?", 1000, vbCenter
    .Add , , "Notifica No.1", 3000
    .Add , , "Notifica No.2", 3000
    .Add , , "Notifica No.3", 3000
    .Add , , "Reg.Fecha", 1800
    .Add , , "Reg.Usuario", 1800, vbCenter
    .Add , , "Mod.Fecha", 1800
    .Add , , "Mod.Usuario", 1800, vbCenter
End With

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtCodigo_Change()

Call sbLimpia
lsw.ListItems.Clear

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

Exit Sub
vError:

End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")

Exit Sub
vError:
End Sub
