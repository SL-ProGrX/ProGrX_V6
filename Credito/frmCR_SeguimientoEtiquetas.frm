VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_SeguimientoEtiquetas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etiquetas de la Operación"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9960
   Icon            =   "frmCR_SeguimientoEtiquetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   9735
      _Version        =   1441793
      _ExtentX        =   17171
      _ExtentY        =   9340
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
         Height          =   3375
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
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
         View            =   3
         FullRowSelect   =   -1  'True
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTag 
         Height          =   312
         Left            =   -67960
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   6972
         _Version        =   1441793
         _ExtentX        =   12303
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
      Begin XtremeSuiteControls.FlatEdit txtNotaTag 
         Height          =   1335
         Left            =   0
         TabIndex        =   12
         Top             =   3840
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1332
         Left            =   -67960
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   6972
         _Version        =   1441793
         _ExtentX        =   12298
         _ExtentY        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtAsignadoIdentificacion 
         Height          =   312
         Left            =   -66280
         TabIndex        =   15
         Top             =   3120
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsignadoClave 
         Height          =   312
         Left            =   -67960
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   612
         Left            =   -62560
         TabIndex        =   4
         Top             =   4080
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "&Aplicar"
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
         Picture         =   "frmCR_SeguimientoEtiquetas.frx":000C
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
         TabIndex        =   11
         Top             =   840
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
         TabIndex        =   10
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
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
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   972
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
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   1572
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
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   5412
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   3000
      TabIndex        =   16
      Top             =   600
      Width           =   5652
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1332
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
      TabIndex        =   0
      Top             =   600
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmCR_SeguimientoEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub CmdAplicar_Click()
Dim i As Integer

If Len(cboTag.Text) = 0 Then
    Exit Sub
End If

On Error GoTo vError

txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)

strSQL = "select isnull(Nota_Largo,0) as 'Nota_Largo' from Crd_Tags where Tag_Codigo = '" & cboTag.ItemData(cboTag.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Nota_Largo > Len(txtNotas.Text) Then
    MsgBox "Este tipo de Etiqueta Requiere que la Nota sea de al menos " & rs!Nota_Largo & " caracteres!", vbExclamation
    Exit Sub
End If

Call sbCrdOperacionTags(txtOperacion.Text, txtOperacion.Tag, cboTag.ItemData(cboTag.ListIndex), txtAsignadoIdentificacion.Text, txtNotas.Text)

i = MsgBox("Etiqueta Registrada Satisfactoriamente, desea salir de la pantalla de registro de etiquetas?", vbYesNo, "Etiqueta Registrada")
If i = vbYes Then
   Unload Me
Else
    tcMain.Item(0).Selected = True
    Call sbLswEtiquetas
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

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

strSQL = "select rtrim(T.Tag_Codigo) as  'IdX', rtrim( T.Descripcion)  as 'ItmX'" _
        & " from CRD_TAGS T " _
        & " inner join CRD_TAGS_GRUPOS TG on TG.TAG_CODIGO = T.TAG_CODIGO " _
        & " inner join CRD_GRPUSERS GU on GU.COD_GRUPO = TG.COD_GRUPO " _
        & " where T.ACTIVO = 1 and GU.USUARIO = '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboTag, strSQL, False, True)

strSQL = "select S.cedula,S.nombre,R.id_solicitud,R.codigo" _
       & " from socios S inner join reg_creditos R on S.cedula = R.cedula" _
       & " where R.id_solicitud = " & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)
  txtOperacion.Text = rs!Id_Solicitud
  txtOperacion.Tag = rs!Codigo
  lblNombre.Caption = "[ " & rs!Cedula & " ] " & rs!Nombre
rs.Close


tcMain.Item(0).Selected = True

Call sbLswEtiquetas

vError:

End Sub


Private Sub sbLswEtiquetas()

txtNotaTag.Text = ""
lsw.ListItems.Clear

strSQL = "select O.*,T.descripcion as Etiqueta" _
       & " from CRD_OPERACION_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
       & " where O.id_solicitud = " & txtOperacion.Text & " order by O.Registro_Fecha"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!REGISTRO_FECHA)
     itmX.SubItems(1) = rs!REGISTRO_USUARIO
     itmX.SubItems(2) = rs!Etiqueta
     itmX.SubItems(3) = rs!Asignado_A & ""
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
