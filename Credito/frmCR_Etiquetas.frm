VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_Etiquetas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Etiquetas para Créditos"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11760
   Begin XtremeSuiteControls.GroupBox gbNotificacion 
      Height          =   6375
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   11535
      _Version        =   1572864
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
         TabIndex        =   7
         Top             =   600
         Width           =   10575
         _Version        =   1572864
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
            TabIndex        =   10
            Top             =   360
            Width           =   8055
            _Version        =   1572864
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
            TabIndex        =   11
            Top             =   840
            Width           =   8055
            _Version        =   1572864
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
            Index           =   0
            Left            =   600
            TabIndex        =   9
            Top             =   720
            Width           =   1095
            _Version        =   1572864
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
            Index           =   1
            Left            =   600
            TabIndex        =   8
            Top             =   360
            Width           =   1095
            _Version        =   1572864
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
         Height          =   1575
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   10575
         _Version        =   1572864
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
            TabIndex        =   13
            Top             =   360
            Width           =   8055
            _Version        =   1572864
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
            TabIndex        =   14
            Top             =   840
            Width           =   8055
            _Version        =   1572864
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
            Index           =   3
            Left            =   600
            TabIndex        =   16
            Top             =   360
            Width           =   1095
            _Version        =   1572864
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
            Index           =   2
            Left            =   600
            TabIndex        =   15
            Top             =   720
            Width           =   1095
            _Version        =   1572864
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
         Height          =   855
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   5400
         Width           =   10575
         _Version        =   1572864
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
            TabIndex        =   20
            Top             =   360
            Width           =   1215
            _Version        =   1572864
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
            Picture         =   "frmCR_Etiquetas.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnNotificacion 
            Height          =   375
            Index           =   1
            Left            =   9240
            TabIndex        =   21
            Top             =   360
            Width           =   1215
            _Version        =   1572864
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
            Picture         =   "frmCR_Etiquetas.frx":0727
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtMensaje 
         Height          =   1095
         Left            =   2280
         TabIndex        =   18
         Top             =   4200
         Width           =   8055
         _Version        =   1572864
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
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   19
         Top             =   4200
         Width           =   1095
         _Version        =   1572864
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
      Begin XtremeShortcutBar.ShortcutCaption scNotificacion 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   11535
         _Version        =   1572864
         _ExtentX        =   20346
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Notificaciones:                                                                               "
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
      End
   End
   Begin XtremeSuiteControls.PushButton btnNotifica 
      Height          =   255
      Index           =   0
      Left            =   9240
      TabIndex        =   3
      Top             =   1095
      Width           =   1095
      _Version        =   1572864
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
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11535
      _Version        =   524288
      _ExtentX        =   20346
      _ExtentY        =   10610
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
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Etiquetas.frx":0E3D
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnNotifica 
      Height          =   255
      Index           =   1
      Left            =   10320
      TabIndex        =   4
      Top             =   1095
      Width           =   1095
      _Version        =   1572864
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
   Begin XtremeShortcutBar.ShortcutCaption lbl 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   11535
      _Version        =   1572864
      _ExtentX        =   20346
      _ExtentY        =   556
      _StockProps     =   14
      Caption         =   "Notificación:                                                                               "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Etiquetas para Operaciones de Crédito"
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
      TabIndex        =   1
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmCR_Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset

Dim mUlltimoRequisitoSel As String, strUltimaSelTipo As String, mListaRequisitos As String
Dim EditaNotificacion As Boolean

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
vModulo = 3
End Sub

Private Sub Form_Load()


vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

gbNotificacion.Visible = False

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select rtrim(cod_requisito) + ' - ' + descripcion as Requisito" _
       & " from requisitos_adicionales" _
       & " order by cod_requisito"

Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF And mUlltimoRequisitoSel = "" Then
     mUlltimoRequisitoSel = rs!Requisito
    End If
    
    mListaRequisitos = ""
    Do While Not rs.EOF
        If Len(mListaRequisitos) = 0 Then
          mListaRequisitos = Chr$(9) & rs!Requisito
        Else
          mListaRequisitos = mListaRequisitos & Chr$(9) & rs!Requisito
        End If
      rs.MoveNext
    Loop
rs.Close

strSQL = "select T.TAG_CODIGO,T.descripcion,isnull(rtrim(T.COD_REQUISITO) + ' - ' + R.Descripcion,'') as Requisito" _
      & ",T.NOTA_LARGO, T.ESPERA_ACTIVA, T.ESPERA_DESACTIVA, T.activo" _
      & "  from CRD_TAGS T left join requisitos_adicionales R on T.cod_requisito = R.cod_requisito" _
      & " order by T.TAG_CODIGO"
      
Call sbCargaGridLocal(vGrid, 7, strSQL)


End Sub

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim strResTipo As String, vNota As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 3
  vGrid.cellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mListaRequisitos
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mUlltimoRequisitoSel
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1 'Codigo de Tag
       vGrid.Text = CStr(rs!TAG_CODIGO)
     
     Case 2 'Descripcion
       vGrid.Text = CStr(rs!Descripcion)
     Case 3 'Tipo
        vGrid.Text = rs!Requisito
     
     Case 4 'Largo de la Nota
        vGrid.Text = CStr(rs!Nota_Largo)
        
     Case 5 'ESPERA_ACTIVA
'       vGrid.Text = CStr(rs!ESPERA_ACTIVA) & ""
     
     Case 6 'ESPERA_DESACTIVA
'       vGrid.Text = CStr(rs!ESPERA_DESACTIVA) & ""
     
     Case 7 'Estado
       vGrid.Text = CStr(rs!Activo)
     
     Case Else
    
    End Select
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 3
  vGrid.cellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mListaRequisitos
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mUlltimoRequisitoSel

Me.MousePointer = vbDefault

End Sub



Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from CRD_TAGS " _
       & " where TAG_CODIGO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  If Mid(Trim(vGrid.Text), 1, 1) = "S" Then
     MsgBox ("No se puede agregar etiquetas que inicien con 'S', resevadas para el sistema")
     Exit Function
  End If
  
  
  strSQL = "insert into CRD_TAGS(TAG_CODIGO, descripcion, cod_requisito, Nota_Largo, ESPERA_ACTIVA, ESPERA_DESACTIVA, ACTIVO) values('" _
         & vGrid.Text & "',' "
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 3
  If Len(Trim(vGrid.Text)) = 0 Then
      strSQL = strSQL & "null" & ", "
  Else
      strSQL = strSQL & "'" & SIFGlobal.fxCodText(vGrid.Text) & "', "
  End If
  
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & ", "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Value & ")"
  
  

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipo de Etiqueta : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CRD_TAGS set descripcion = '" & vGrid.Text
  vGrid.Col = 3
  If Len(Trim(vGrid.Text)) = 0 Then
     strSQL = strSQL & "', cod_requisito = Null, NOTA_LARGO = "
  Else
     strSQL = strSQL & "', cod_requisito = '" & SIFGlobal.fxCodText(vGrid.Text) & "', NOTA_LARGO = "
  End If
  
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & ", ESPERA_ACTIVA = "
 
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", ESPERA_DESACTIVA = "
 
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & ", ACTIVO = "
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Value & " where TAG_CODIGO = '"
 
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipo de Etiqueta : " & vGrid.Text)

End If
rs.Close

vGrid.Col = 3
mUlltimoRequisitoSel = vGrid.Text


fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub sbCargarNotificacion()
    
    On Error GoTo vError
   
    strSQL = "select TAG_CODIGO as  'IdX', rtrim(DESCRIPCION) as ItmX" _
         & " from CRD_TAGS"
         
    Call sbCbo_Llena_New(cboTagPara, strSQL, False, True)
    Call sbCbo_Copia(cboTagPara, cboTagCC)
    
    cboTagPara.AddItem " "
    cboTagCC.AddItem " "
    
    cboTagPara.Text = " "
    cboTagCC.Text = " "
    
    strSQL = "select CT.PARA_TAG, rtrim(TP.DESCRIPCION) as 'PARA_TAG_DESC'" _
        & ", CT.PARA_EMAIL, CT.CC_TAG, rtrim(TC.DESCRIPCION) as 'CC_TAG_DESC', CT.CC_EMAIL, CT.MENSAJE" _
        & " from CRD_TAGS_AVISOS CT LEFT JOIN CRD_TAGS TP ON CT.PARA_TAG = TP.TAG_CODIGO" _
        & " LEFT JOIN CRD_TAGS TC ON CT.CC_TAG = TC.TAG_CODIGO" _
        & " WHERE CT.TAG_CODIGO = '" & scNotificacion.Tag & "'"

     

    Call OpenRecordSet(rs, strSQL)

    If Not rs.EOF Then
    
        EditaNotificacion = True
        
        If IsNull(rs!Para_Tag_Desc) = False Then
            Call sbCboAsignaDato(cboTagPara, rs!Para_Tag_Desc, True, rs!Para_Tag)
        End If
        
        txtCorreosPara.Text = IIf(IsNull(rs!PARA_EMAIL), "", rs!PARA_EMAIL)
        
        If IsNull(rs!CC_TAG_DESC) = False Then
            Call sbCboAsignaDato(cboTagCC, rs!CC_TAG_DESC, True, rs!CC_TAG)
        End If
        txtCorreosCC.Text = IIf(IsNull(rs!CC_EMAIL), "", rs!CC_EMAIL)
        txtMensaje.Text = IIf(IsNull(rs!Mensaje), "", rs!Mensaje)
        
    Else

        EditaNotificacion = False
        
    End If
    
gbNotificacion.top = 1080
gbNotificacion.Left = 0
gbNotificacion.Visible = True
    
    
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
    
    If scNotificacion.Tag = Empty Then
        Exit Sub
    End If
    
    If Trim(cboTagPara.Text) = Empty Then
        EtiquetaPara = Empty
    Else
        EtiquetaPara = cboTagPara.ItemData(cboTagPara.ListIndex)
    End If
    
    If Trim(cboTagCC.Text) = Empty Then
        EtiquetaCC = Empty
    Else
        EtiquetaCC = cboTagCC.ItemData(cboTagCC.ListIndex)
    End If
    
    If EditaNotificacion = False Then
        strSQL = "Insert CRD_TAGS_AVISOS (TAG_CODIGO,PARA_TAG,PARA_EMAIL,CC_TAG,CC_EMAIL,MENSAJE) VALUES " _
            & " ('" & scNotificacion.Tag & "','" _
            & Trim(EtiquetaPara) & "','" _
            & txtCorreosPara & "','" _
            & Trim(EtiquetaCC) & "','" _
            & txtCorreosCC & "','" _
            & txtMensaje & "')"
    Else
        strSQL = "update CRD_TAGS_AVISOS set PARA_TAG = '" & Trim(EtiquetaPara) & "', PARA_EMAIL ='" _
            & Trim(txtCorreosPara) & "', CC_TAG = '" _
            & Trim(EtiquetaCC) & "', CC_EMAIL = '" _
            & Trim(txtCorreosCC) & "',MENSAJE ='" _
            & Trim(txtMensaje) & "' WHERE TAG_CODIGO = '" & scNotificacion.Tag & "'"
    End If
    Call ConectionExecute(strSQL)
    
    gbNotificacion.Visible = False
    
    MsgBox "Información almacenada con éxito!", vbInformation
    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
        
End Sub

Private Sub sbEliminarNotificacion()
    
    On Error GoTo vError
    
    If scNotificacion.Tag = Empty Then
        Exit Sub
    End If
    
    If MsgBox("Desea eliminar la notificación a la etiqueta: " & scNotificacion.Tag, vbOKCancel) = vbOK Then
        strSQL = "DELETE CRD_TAGS_AVISOS WHERE TAG_CODIGO = '" & scNotificacion.Tag & "'"
        Call ConectionExecute(strSQL)
        
        MsgBox "La notificación ha sido eliminada!", vbInformation
    End If
    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
        
End Sub

Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)
    vGrid.Col = 1
    vGrid.Row = Row
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
  
    vGrid.Col = 3
    vGrid.cellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mListaRequisitos
    vGrid.TypeComboBoxEditable = False
    vGrid.Text = mUlltimoRequisitoSel
  
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow

    vGrid.Col = 3
    vGrid.cellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mListaRequisitos
    vGrid.TypeComboBoxEditable = False
    vGrid.Text = mUlltimoRequisitoSel

End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete CRD_TAGS where TAG_CODIGO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Etiqueta : " & vGrid.Text)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        
        If vGrid.MaxRows <= 0 Then
          vGrid.MaxRows = 1
        End If

     End If
End If


End Sub

