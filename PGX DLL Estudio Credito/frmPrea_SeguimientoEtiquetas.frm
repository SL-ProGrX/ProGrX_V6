VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPrea_SeguimientoEtiquetas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Tag's del Estudio de Crédito"
   ClientHeight    =   6576
   ClientLeft      =   108
   ClientTop       =   408
   ClientWidth     =   9744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6576
   ScaleWidth      =   9744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5292
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   9612
      _Version        =   1245187
      _ExtentX        =   16954
      _ExtentY        =   9334
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
      Item(0).Caption =   "Historial"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "txtNotaTag"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "cboTag"
      Item(1).Control(1)=   "txtAsignadoIdentificacion"
      Item(1).Control(2)=   "txtAsignadoClave"
      Item(1).Control(3)=   "cmdAplicar"
      Item(1).Control(4)=   "Label1(6)"
      Item(1).Control(5)=   "Label1(5)"
      Item(1).Control(6)=   "Label1(4)"
      Item(1).Control(7)=   "Label1(3)"
      Item(1).Control(8)=   "Label1(2)"
      Item(1).Control(9)=   "txtNotas"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3372
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9492
         _Version        =   1245187
         _ExtentX        =   16743
         _ExtentY        =   5948
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
         FullRowSelect   =   -1  'True
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTag 
         Height          =   312
         Left            =   -67960
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   6972
         _Version        =   1245187
         _ExtentX        =   12298
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtNotaTag 
         Height          =   1332
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Width           =   9492
         _Version        =   1245187
         _ExtentX        =   16743
         _ExtentY        =   2350
         _StockProps     =   77
         ForeColor       =   0
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1332
         Left            =   -67960
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   6972
         _Version        =   1245187
         _ExtentX        =   12298
         _ExtentY        =   2350
         _StockProps     =   77
         ForeColor       =   0
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsignadoIdentificacion 
         Height          =   312
         Left            =   -66280
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1245187
         _ExtentX        =   9334
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsignadoClave 
         Height          =   312
         Left            =   -67960
         TabIndex        =   6
         Top             =   3120
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1245187
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         Alignment       =   2
         PasswordChar    =   "*"
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   612
         Left            =   -62560
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1245187
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "&Aplicar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmPrea_SeguimientoEtiquetas.frx":0000
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   252
         Index           =   6
         Left            =   -66400
         TabIndex        =   12
         Top             =   2880
         Visible         =   0   'False
         Width           =   5412
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Clave"
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
         Height          =   252
         Index           =   5
         Left            =   -67960
         TabIndex        =   11
         Top             =   2880
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Left            =   -69280
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Asignado a"
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
         Left            =   -69160
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Etiqueta"
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
         Left            =   -69280
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   972
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   315
      Left            =   3000
      TabIndex        =   16
      Top             =   600
      Width           =   6612
      _Version        =   1245187
      _ExtentX        =   11663
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación :"
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
      Height          =   252
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label lblTitulo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2652
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmPrea_SeguimientoEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mId_Solicitud As String



Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If cboTag.ListCount = 0 Then Exit Sub



On Error GoTo vError:
If Val(mId_Solicitud) > 0 Then

    strSQL = "select isnull(max(Linea),0) + 1 as Linea from CRD_OPERACION_TAGS" _
           & " where id_solicitud = " & txtOperacion.Text
    Call OpenRecordSet(rs, strSQL)
    
    strSQL = "insert CRD_OPERACION_TAGS(Linea,Tag_Codigo,Codigo,Id_Solicitud,Registro_Fecha,Registro_Usuario,Asignado_A,Notas)" _
           & " values(" & rs!Linea & ",'" & cboTag.ItemData(cboTag.ListIndex) & "','" & txtOperacion.Tag & "'," _
           & txtOperacion.Text & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtAsignadoClave.Text & "','" & txtNotas.Text & "')"
Else

    strSQL = "select isnull(max(Linea),0) + 1 as Linea from CRD_PREA_TAGS" _
           & " where  COD_PREANALISIS = " & txtOperacion.Text
    Call OpenRecordSet(rs, strSQL)
    
    strSQL = "insert CRD_PREA_TAGS(Linea,Tag_Codigo,Codigo,Cod_preanalisis,Registro_Fecha,Registro_Usuario,Asignado_A,Notas)" _
           & " values(" & rs!Linea & ",'" & cboTag.ItemData(cboTag.ListIndex) & "','" & txtOperacion.Tag & "'," _
           & txtOperacion.Text & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtAsignadoClave.Text & "','" & txtNotas.Text & "')"
End If
Call ConectionExecute(strSQL)

rs.Close

MsgBox "Etiqueta Registrada Satisfactoriamente...", vbInformation

tcMain.Item(0).Selected = True

Call sbLswEtiquetas

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError



Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture
 

With lsw.ColumnHeaders
    .Clear
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Etiqueta", 2600
    .Add , , "Asignado a:", 2100, vbCenter
    .Add , , "Notas", 3600
End With

strSQL = "Select rtrim(T.Tag_Codigo) as  'IdX', rtrim(T.Descripcion)  as 'ItmX'" _
        & " from CRD_TAGS T " _
        & " inner join CRD_TAGS_GRUPOS TG on TG.TAG_CODIGO = T.TAG_CODIGO " _
        & " inner join CRD_GRPUSERS GU on GU.COD_GRUPO = TG.COD_GRUPO " _
        & " where T.ACTIVO = 1 and GU.USUARIO = '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboTag, strSQL, False, True)

If Val(mId_Solicitud) > 0 Then

    lblTitulo.Caption = "Solicitud:"

    strSQL = "select S.cedula,S.nombre,R.id_solicitud,R.codigo" _
           & " from socios S inner join reg_creditos R on S.cedula = R.cedula" _
           & " where R.id_solicitud = " & mId_Solicitud

    Call OpenRecordSet(rs, strSQL)
    txtOperacion.Text = rs!ID_SOLICITUD
    txtOperacion.Tag = rs!Codigo
    txtIdentificacion.Text = "[ " & rs!cedula & " ] " & rs!nombre
    
Else

    lblTitulo.Caption = "Estudio de Crédito:"

    strSQL = "select S.cedula,S.nombre,R.cod_preanalisis,R.cod_linea" _
           & " from socios S inner join CRD_PREA_PREANALISIS R on S.cedula = R.cedula" _
           & " where R.cod_preanalisis = '" & gPreAnalisis.Expediente & "'"
           
    Call OpenRecordSet(rs, strSQL)
    txtOperacion.Text = rs!cod_preanalisis
    txtOperacion.Tag = rs!Cod_Linea
    txtIdentificacion.Text = "[ " & rs!cedula & " ] " & rs!nombre
           
End If

rs.Close

Call sbLswEtiquetas

vError:

End Sub


Private Sub sbLswEtiquetas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


txtNotaTag.Text = ""
lsw.ListItems.Clear

If Val(mId_Solicitud) > 0 Then
    strSQL = "select O.*,T.descripcion as Etiqueta" _
           & " from CRD_OPERACION_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
           & " where O.id_solicitud = " & txtOperacion.Text
Else
    strSQL = "select O.*,T.descripcion as Etiqueta" _
           & " from CRD_PREA_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
           & " where O.COD_PREANALISIS = " & txtOperacion.Text
End If
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!registro_Fecha)
     itmX.SubItems(1) = rs!registro_usuario
     itmX.SubItems(2) = rs!Etiqueta
     itmX.SubItems(3) = rs!Asignado_A
     itmX.SubItems(4) = rs!Notas
     itmX.Tag = rs!Linea
 rs.MoveNext
Loop
rs.Close


End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
txtNotaTag.Text = ""

If lsw.ListItems.Count = 0 Then Exit Sub

txtNotaTag.Text = Item.SubItems(4)

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Then
   Call sbLswEtiquetas
Else
   txtAsignadoClave.Text = ""
   txtAsignadoIdentificacion.Text = ""
   txtNotas.Text = ""
End If

End Sub
