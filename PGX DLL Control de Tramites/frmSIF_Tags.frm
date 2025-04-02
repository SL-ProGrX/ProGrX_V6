VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmSIF_Tags 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etiquetas (Tag's) de Control del Sistema"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11505
   Icon            =   "frmSIF_Tags.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbNotificacion 
      Height          =   6375
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   11535
      _Version        =   1441793
      _ExtentX        =   20346
      _ExtentY        =   11245
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.GroupBox gbCorreos 
         Height          =   1575
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   10575
         _Version        =   1441793
         _ExtentX        =   18653
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Para: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboTagPara 
            Height          =   330
            Left            =   2040
            TabIndex        =   7
            Top             =   360
            Width           =   8055
            _Version        =   1441793
            _ExtentX        =   14208
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
         Begin XtremeSuiteControls.FlatEdit txtCorreosPara 
            Height          =   615
            Left            =   2040
            TabIndex        =   8
            Top             =   840
            Width           =   8055
            _Version        =   1441793
            _ExtentX        =   14208
            _ExtentY        =   1085
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   10
            Top             =   360
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Etiqueta"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   9
            Top             =   720
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Lista de Correos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbCorreos 
         Height          =   1575
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   10575
         _Version        =   1441793
         _ExtentX        =   18653
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Con Copia: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboTagCC 
            Height          =   330
            Left            =   2040
            TabIndex        =   12
            Top             =   360
            Width           =   8055
            _Version        =   1441793
            _ExtentX        =   14208
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
         Begin XtremeSuiteControls.FlatEdit txtCorreosCC 
            Height          =   615
            Left            =   2040
            TabIndex        =   13
            Top             =   840
            Width           =   8055
            _Version        =   1441793
            _ExtentX        =   14208
            _ExtentY        =   1085
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   375
            Index           =   2
            Left            =   600
            TabIndex        =   15
            Top             =   720
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Lista de Correos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   14
            Top             =   360
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Etiqueta"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeSuiteControls.GroupBox gbCorreos 
         Height          =   855
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   5400
         Width           =   10575
         _Version        =   1441793
         _ExtentX        =   18653
         _ExtentY        =   1508
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnNotificacion 
            Height          =   375
            Index           =   0
            Left            =   8040
            TabIndex        =   17
            Top             =   360
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   661
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
            Appearance      =   17
            Picture         =   "frmSIF_Tags.frx":6852
         End
         Begin XtremeSuiteControls.PushButton btnNotificacion 
            Height          =   375
            Index           =   1
            Left            =   9240
            TabIndex        =   18
            Top             =   360
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   661
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmSIF_Tags.frx":6F79
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtMensaje 
         Height          =   1095
         Left            =   2280
         TabIndex        =   19
         Top             =   4200
         Width           =   8055
         _Version        =   1441793
         _ExtentX        =   14208
         _ExtentY        =   1931
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
      Begin XtremeShortcutBar.ShortcutCaption scNotificacion 
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   11535
         _Version        =   1441793
         _ExtentX        =   20346
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Notificaciones:                                                                               "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   20
         Top             =   4200
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mensaje"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6135
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   8655
      _Version        =   524288
      _ExtentX        =   15266
      _ExtentY        =   10821
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
      SpreadDesigner  =   "frmSIF_Tags.frx":768F
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnNotifica 
      Height          =   255
      Index           =   0
      Left            =   9240
      TabIndex        =   1
      Top             =   1095
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Agrega"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnNotifica 
      Height          =   255
      Index           =   1
      Left            =   10320
      TabIndex        =   2
      Top             =   1095
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Elimina"
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
      Appearance      =   6
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Etiquetas de Control"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   6252
   End
   Begin XtremeShortcutBar.ShortcutCaption lbl 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   11535
      _Version        =   1441793
      _ExtentX        =   20346
      _ExtentY        =   556
      _StockProps     =   14
      Caption         =   "Notificación:                                                                               "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.01
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   2
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmSIF_Tags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim EditaNotificacion As Boolean
Dim vTag As String

Private Sub btnNotifica_Click(Index As Integer)
    Select Case Index
    Case 0 '"NOTIFICACION"
        Call sbCargarNotificacion
    Case 1 '"ELIMINAR"
        Call sbEliminarNotificacion
    End Select
End Sub

Private Sub btnNotificacion_Click(Index As Integer)
Select Case Index
    Case 0 '"GUARDAR"
        Call sbGuardaNotificacion
        Call sbLimpiarNotificacion
    
    Case 1 '"CANCELAR"
        Call sbLimpiarNotificacion
        gbNotificacion.Visible = False
End Select

End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()

vModulo = 8

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

gbNotificacion.Visible = False

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "exec dbo.spSifTags"
Call ConectionExecute(strSQL)

Call sbCargaTags

vGrid.MaxRows = vGrid.MaxRows
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "NOTIFICACION"
        Call sbCargarNotificacion
    Case "ELIMINAR"
        Call sbEliminarNotificacion
    End Select
End Sub

Private Sub tlbNotificaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case UCase(Button.Key)
    Case "CANCELAR"
        Call sbLimpiarNotificacion
        gbNotificacion.Visible = False
    Case "GUARDAR"
        Call sbGuardaNotificacion
        Call sbLimpiarNotificacion
    End Select
End Sub

Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)
    
    vGrid.Row = Row
    vGrid.Col = 1
    scNotificacion.Tag = Trim(vGrid.Text)
    vGrid.Col = 2
    scNotificacion.Caption = "Notificaciones para: " & vGrid.Text

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
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
    vGrid.Col = 3


End If

'Borrar Linea
End Sub
Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


strSQL = "select isnull(count(*),0) as Existe from SIF_TAGS" _
       & " where TAG_CODIGO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)


If Mid(vGrid.Text, 1, 1) = "S" Then
  MsgBox "Este tag es de sistema, no puede modificarse...", vbInformation
  Call sbCargaTags
Exit Function
End If


If rs!Existe = 0 Then 'Insertar
    If Trim(vGrid.Text) = "" Then Exit Function
    strSQL = "insert SIF_TAGS(TAG_CODIGO,descripcion,activo) values('" _
            & UCase(vGrid.Text) & "','"
    vGrid.Col = 2
    strSQL = strSQL & UCase(vGrid.Text) & "'," & ""
    vGrid.Col = 3
    strSQL = strSQL & UCase(vGrid.Text) & ")"
    Call ConectionExecute(strSQL)
    Call Bitacora("Inserta", "SIF Tipo de Etiqueta : " & vGrid.Text)

Else

    vGrid.Col = 2
    strSQL = "update SIF_TAGS set descripcion = '" & vGrid.Text
    vGrid.Col = 3
    strSQL = strSQL & "',Activo= "
    strSQL = strSQL & vGrid.Value & " where TAG_CODIGO = '"
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
    vGrid.Col = 1
    Call Bitacora("Modifica", "SIF Tipo de Etiqueta : " & vGrid.Text)
End If

vGrid.Col = 3

fxGuardar = 1

Call sbCargaTags

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function

Private Sub sbCargarNotificacion()
    
    On Error GoTo vError
    
    gbNotificacion.Left = 0
    gbNotificacion.Top = 1440
    
    
    strSQL = "select TAG_CODIGO as  'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
         & " from SIF_TAGS"
         
    Call sbCbo_Llena_New(cboTagPara, strSQL, False, True)
    Call sbCbo_Llena_New(cboTagCC, strSQL, False, True)
    
    cboTagPara.AddItem " "
    cboTagPara.Text = " "
    
    cboTagCC.AddItem " "
    cboTagCC.Text = " "
    
    strSQL = "select CT.PARA_TAG as 'Para_Tag_Id',  rtrim(TP.DESCRIPCION) as 'PARA_TAG'" _
        & ", CT.PARA_EMAIL, CT.CC_TAG as 'CC_Tag_Id', rtrim(TC.DESCRIPCION) as 'CC_TAG', CT.CC_EMAIL, CT.MENSAJE " _
        & " from SIF_TAGS_AVISOS CT LEFT JOIN SIF_TAGS TP ON CT.PARA_TAG = TP.TAG_CODIGO " _
        & " LEFT JOIN SIF_TAGS TC ON CT.CC_TAG = TC.TAG_CODIGO " _
        & " WHERE CT.TAG_CODIGO = '" & scNotificacion.Tag & "'"

    Call OpenRecordSet(rs, strSQL)

    If Not rs.EOF Then
    
        EditaNotificacion = True
        
        If IsNull(rs!Para_Tag) = False Then
            Call sbCboAsignaDato(cboTagPara, rs!Para_Tag, True, rs!Para_Tag_Id)
        End If
        
        txtCorreosPara.Text = IIf(IsNull(rs!PARA_EMAIL), "", rs!PARA_EMAIL)
        If IsNull(rs!CC_Tag) = False Then
            Call sbCboAsignaDato(cboTagCC, rs!CC_Tag, True, rs!CC_Tag_Id)
        End If
        txtCorreosCC.Text = IIf(IsNull(rs!CC_EMAIL), "", rs!CC_EMAIL)
        txtMensaje.Text = IIf(IsNull(rs!Mensaje), "", rs!Mensaje)
        
    Else

        EditaNotificacion = False
        
    End If
    
    gbNotificacion.Visible = True
    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbEliminarNotificacion()
    
On Error GoTo vError

If scNotificacion.Tag = Empty Or scNotificacion.Tag = "" Then
    Exit Sub
End If

If MsgBox("Desea eliminar la notificación a la etiqueta: " & scNotificacion.Tag, vbOKCancel) = vbOK Then
    strSQL = "DELETE SIF_TAGS_AVISOS WHERE TAG_CODIGO = '" & scNotificacion.Tag & "'"
    Call ConectionExecute(strSQL)
    MsgBox "La notificación a sido eliminada"
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
        
End Sub

Private Sub sbLimpiarNotificacion()
    cboTagPara.Clear
    txtCorreosPara.Text = Empty
    cboTagCC.Clear
    txtCorreosCC.Text = Empty
    txtMensaje.Text = Empty
End Sub


Private Sub sbGuardaNotificacion()
    
    Dim EtiquetaPara As String, EtiquetaCC As String
    
    On Error GoTo vError
    
    If scNotificacion.Tag = Empty Or scNotificacion.Tag = "" Then
        Exit Sub
    End If
    
    If Trim(cboTagPara.Text) = Empty Or cboTagPara.Text = "" Then
        EtiquetaPara = Empty
    Else
        EtiquetaPara = cboTagPara.ItemData(cboTagPara.ListIndex)
    End If
    
    If Trim(cboTagCC.Text) = Empty Or cboTagCC.Text = "" Then
        EtiquetaCC = Empty
    Else
        EtiquetaCC = cboTagCC.ItemData(cboTagCC.ListIndex)
    End If
    
    If EditaNotificacion = False Then
        strSQL = "Insert SIF_TAGS_AVISOS (TAG_CODIGO,PARA_TAG,PARA_EMAIL,CC_TAG,CC_EMAIL,MENSAJE) VALUES " _
            & " ('" & scNotificacion.Tag & "','" _
            & Trim(EtiquetaPara) & "','" _
            & txtCorreosPara.Text & "','" _
            & Trim(EtiquetaCC) & "','" _
            & txtCorreosCC.Text & "','" _
            & txtMensaje.Text & "')"
    Else
        strSQL = "update SIF_TAGS_AVISOS set PARA_TAG = '" & Trim(EtiquetaPara) & "', PARA_EMAIL ='" _
            & Trim(txtCorreosPara.Text) & "', CC_TAG = '" _
            & Trim(EtiquetaCC) & "', CC_EMAIL = '" _
            & Trim(txtCorreosCC.Text) & "',MENSAJE ='" _
            & Trim(txtMensaje.Text) & "' WHERE TAG_CODIGO = '" & scNotificacion.Tag & "'"
    End If
    Call ConectionExecute(strSQL)
    gbNotificacion.Visible = False
    
    
    
    MsgBox "Información almacenada con éxito"
    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
        
End Sub

Private Sub sbCargaTags()

strSQL = "select tag_codigo,descripcion,activo from sif_Tags" _
       & " order by tag_codigo"
Call sbCargaGrid(vGrid, 3, strSQL)


End Sub
