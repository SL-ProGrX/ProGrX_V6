VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFSL_SeguimientoEtiquetas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etiquetas"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtExpediente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtIdentificacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   4095
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Etiquetas Registradas"
      TabPicture(0)   =   "frmFSL_SeguimientoEtiquetas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNota"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsw"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Crear Etiqueta"
      TabPicture(1)   =   "frmFSL_SeguimientoEtiquetas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAplicar"
      Tab(1).Control(1)=   "cboTag"
      Tab(1).Control(2)=   "txtAsignadoClave"
      Tab(1).Control(3)=   "txtAsignadoIdentificacion"
      Tab(1).Control(4)=   "txtNotas"
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(6)=   "Label1(3)"
      Tab(1).Control(7)=   "Label1(4)"
      Tab(1).Control(8)=   "Line1(1)"
      Tab(1).Control(9)=   "Label1(5)"
      Tab(1).Control(10)=   "Label1(6)"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -68280
         Picture         =   "frmFSL_SeguimientoEtiquetas.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox cboTag 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtAsignadoClave 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73320
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtAsignadoIdentificacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   6375
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Etiqueta"
            Object.Width           =   5010
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Asignado A"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Notas"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label lblNota 
         Caption         =   "Nota :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   7695
      End
      Begin VB.Label Label1 
         Caption         =   "Etiqueta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Asignado a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74520
         TabIndex        =   11
         Top             =   1920
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -75000
         X2              =   -66240
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   -71760
         TabIndex        =   9
         Top             =   1200
         Width           =   4575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Expediente :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   16
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Identificación :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   15
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   8880
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmFSL_SeguimientoEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()
Dim i As Integer

If Len(cboTag.Text) = 0 Then
    Exit Sub
End If

On Error GoTo vError

'Call sbCrdOperacionTags(txtExpediente.Text, SIFGlobal.fxSIFCodText(cboTag.Text), txtAsignadoIdentificacion.Text, txtNotas.Text)

i = MsgBox("Etiqueta Registrada Satisfactoriamente, desea salir de la pantalla de registro de etiquetas?", vbYesNo, "Etiqueta Registrada")
If i = vbYes Then
   Unload Me
Else
    ssTab.Tab = 0
    Call sbLswEtiquetas
End If


Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError:

ssTab.Tab = 0
txtExpediente.Text = GLOBALES.gTag
txtIdentificacion.Text = GLOBALES.gTag2

strSQL = "select T.Tag_Codigo + ' - ' + T.Descripcion  as ItmX from FSL_TAGS T " _
        & " inner join FSL_TAGS_GRUPOS TG on TG.TAG_CODIGO = T.TAG_CODIGO " _
        & " inner join FSL_GRPUSERS GU on GU.COD_GRUPO = TG.COD_GRUPO " _
        & " where T.ACTIVO = 1 and GU.USUARIO = '" & glogon.Usuario & "'"
Call sbLlenaCbo(cboTag, strSQL, False, False)

strSQL = "select S.cedula,S.nombre,E.Cod_Expediente,R.codigo" _
       & " from socios S inner join FSL_EXPEDIENTES E on S.cedula = E.cedula" _
       & " where E.Cod_Expediente = " & GLOBALES.gTag
rs.Open strSQL, glogon.Conection, adOpenStatic
  txtExpediente.Text = rs!COD_EXPEDIENTE
  txtIdentificacion.Text = "[ " & rs!Cedula & " ] " & rs!Nombre
rs.Close

Call sbLswEtiquetas

vError:

End Sub

Private Sub sbLswEtiquetas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem


lblNota.Caption = ""
lsw.ListItems.Clear

strSQL = "select O.*,T.descripcion as Etiqueta" _
       & " from FSL_OPERACION_TAGS O inner join FSL_Tags T on O.Tag_codigo = T.Tag_Codigo" _
       & " where O.id_solicitud = " & txtExpediente.Text
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Registro_Fecha)
     itmX.SubItems(1) = rs!Registro_Usuario
     itmX.SubItems(2) = rs!Etiqueta
     itmX.SubItems(3) = rs!Asignado_A
     itmX.SubItems(4) = rs!Notas
     itmX.Tag = rs!Linea
 rs.MoveNext
Loop
rs.Close


End Sub


Private Sub lsw_Click()

lblNota.Caption = ""

If lsw.ListItems.Count = 0 Then Exit Sub

lblNota.Caption = lsw.SelectedItem.SubItems(4)

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)

If ssTab.Tab = 0 Then
   Call sbLswEtiquetas
Else
   txtAsignadoClave.Text = ""
   txtAsignadoIdentificacion.Text = ""
   txtNotas.Text = ""
End If

End Sub

Public Sub sbFSLExpedientesTags(pOperacion As Long, pTag As String _
        , Optional pAsignado As String = "", Optional pNotas As String = "")
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spFSLExpedienteTagRegistra " & pOperacion & ",'" & pTag & "','" & glogon.Usuario _
                & "','" & Mid(pAsignado, 1, 30) & "','" & Mid(pNotas, 1, 1000) & "'"
glogon.Conection.Execute strSQL

Exit Sub

vError:
  MsgBox Err.Description, vbCritical


End Sub

