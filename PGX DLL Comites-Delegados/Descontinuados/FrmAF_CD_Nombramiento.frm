VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAF_CD_Nombramiento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nombramientos"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAF_CD_Nombramiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Crt 
      Left            =   8385
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox TxtCodC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1935
      TabIndex        =   25
      Top             =   585
      Width           =   525
   End
   Begin VB.TextBox TxtComites 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2445
      Locked          =   -1  'True
      TabIndex        =   24
      ToolTipText     =   "Para buscar el comite presione la tecla  F4 "
      Top             =   585
      Width           =   4515
   End
   Begin VB.CommandButton CmdMiembros 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   23
      ToolTipText     =   "Miembros del Comite Seleccionado"
      Top             =   585
      Width           =   960
   End
   Begin TabDlg.SSTab SstComites 
      Height          =   4380
      Left            =   150
      TabIndex        =   2
      Top             =   1485
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Asignar Nombramientos"
      TabPicture(0)   =   "FrmAF_CD_Nombramiento.frx":3482
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblNombre"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ChkDesm"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ChkMiembro"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CmdAplica"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CmdNuevo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CboPuesto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtNotas"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtCedula"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Consultas"
      TabPicture(1)   =   "FrmAF_CD_Nombramiento.frx":349E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkHistorial"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "LswMiembros"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Comites Activos"
      TabPicture(2)   =   "FrmAF_CD_Nombramiento.frx":34BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ChkTodos"
      Tab(2).Control(1)=   "OptNoact"
      Tab(2).Control(2)=   "OptActivos"
      Tab(2).Control(3)=   "CmdImp"
      Tab(2).Control(4)=   "Picture1"
      Tab(2).Control(5)=   "LswComi"
      Tab(2).ControlCount=   6
      Begin VB.CheckBox ChkTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos los Comites"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -70875
         TabIndex        =   31
         Top             =   3705
         Width           =   1890
      End
      Begin VB.OptionButton OptNoact 
         Caption         =   "Imp. Miembros no Activos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72990
         TabIndex        =   30
         Top             =   3705
         Width           =   1905
      End
      Begin VB.OptionButton OptActivos 
         Caption         =   "Imp. Miembros Activos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74805
         TabIndex        =   29
         Top             =   3705
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.CommandButton CmdImp 
         Height          =   840
         Left            =   -68250
         Picture         =   "FrmAF_CD_Nombramiento.frx":34D6
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3345
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         Height          =   1440
         Left            =   -68610
         Picture         =   "FrmAF_CD_Nombramiento.frx":3805
         ScaleHeight     =   1380
         ScaleWidth      =   1620
         TabIndex        =   27
         ToolTipText     =   "Presione Click izquierdo del mouse para actualizar comites"
         Top             =   1080
         Width           =   1680
      End
      Begin MSComctlLib.ListView LswComi 
         Height          =   3000
         Left            =   -74805
         TabIndex        =   26
         Top             =   495
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   5292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comite"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Provincia"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.TextBox TxtCedula 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   22
         ToolTipText     =   "Para buscar el asociado presione la tecla  F4 "
         Top             =   855
         Width           =   1365
      End
      Begin VB.TextBox TxtNotas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   870
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "FrmAF_CD_Nombramiento.frx":45EB
         Top             =   3195
         Width           =   5610
      End
      Begin VB.ComboBox CboPuesto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   20
         Top             =   1650
         Width           =   6960
      End
      Begin VB.CommandButton CmdNuevo 
         Height          =   795
         Left            =   6525
         Picture         =   "FrmAF_CD_Nombramiento.frx":45ED
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Nueva Asignación"
         Top             =   3285
         Width           =   930
      End
      Begin VB.CommandButton CmdAplica 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7455
         Picture         =   "FrmAF_CD_Nombramiento.frx":485E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3285
         Width           =   930
      End
      Begin VB.CheckBox ChkMiembro 
         Alignment       =   1  'Right Justify
         Caption         =   "Miembro Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   15
         Top             =   2760
         Width           =   1590
      End
      Begin VB.CheckBox ChkDesm 
         Alignment       =   1  'Right Justify
         Caption         =   "Desembolsos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4605
         TabIndex        =   14
         Top             =   2775
         Width           =   1575
      End
      Begin VB.CheckBox ChkHistorial 
         Alignment       =   1  'Right Justify
         Caption         =   "Historial"
         Height          =   195
         Left            =   -68130
         TabIndex        =   8
         Top             =   480
         Width           =   1170
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nombramientos"
         Height          =   1635
         Left            =   -74670
         TabIndex        =   3
         Top             =   2580
         Width           =   7695
         Begin VB.TextBox TxtNotasCon 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   2025
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Text            =   "FrmAF_CD_Nombramiento.frx":4B0B
            Top             =   435
            Width           =   5565
         End
         Begin VB.CheckBox ChkDesembolso 
            Alignment       =   1  'Right Justify
            Caption         =   "Desembolsos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   225
            TabIndex        =   5
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox ChkActivo 
            Alignment       =   1  'Right Justify
            Caption         =   "Miembro Activo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2025
            TabIndex        =   7
            Top             =   210
            Width           =   585
         End
      End
      Begin MSComctlLib.ListView LswMiembros 
         Height          =   1815
         Left            =   -74685
         TabIndex        =   9
         Top             =   720
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Puesto"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   1575
         X2              =   8385
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Label Label6 
         Caption         =   "Nombramientos"
         Height          =   210
         Left            =   165
         TabIndex        =   19
         Top             =   2385
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   17
         Top             =   3150
         Width           =   555
      End
      Begin VB.Label LblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   13
         Top             =   855
         Width           =   5595
      End
      Begin VB.Label Label5 
         Caption         =   "Cédula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   12
         Top             =   885
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   11
         Top             =   1665
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Miembros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74700
         TabIndex        =   10
         Top             =   480
         Width           =   795
      End
   End
   Begin MSComctlLib.StatusBar staUse 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   5895
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "26/03/2014"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Left            =   195
      Picture         =   "FrmAF_CD_Nombramiento.frx":4B0D
      Stretch         =   -1  'True
      Top             =   105
      Width           =   1530
   End
   Begin VB.Label Label3 
      Caption         =   "Comites"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1920
      TabIndex        =   0
      Top             =   345
      Width           =   720
   End
End
Attribute VB_Name = "FrmAF_CD_Nombramiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vLista As Boolean
Dim strSQL As String
Dim rs As New ADODB.Recordset


Function FxNomComite(vUnidad As String)
   Dim rs As New ADODB.Recordset
   Dim strSQL As String
  
   strSQL = "select U.descripcion from uprogramatica U right join afi_cd_agrupacomites A " _
            & "on U.codigo = A.id_pricomite" _
            & " where A.id_pricomite = '" & vUnidad & "'"
            rs.Open strSQL, glogon.Conection, adOpenStatic
   If rs.EOF Then
      FxNomComite = "No existe unidad definida en Comites y Delegados"
   Else
      FxNomComite = rs!Descripcion
   End If

End Function
Private Function fxCons(vCampo As String, vTabla As String) As Integer

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max( " & vCampo & " ),0) as Consecutivo from " & vTabla & ""

rs.Open strSQL, glogon.Conection, adOpenStatic
  fxCons = rs!consecutivo + 1
rs.Close

End Function

Sub sbComite()
If TxtCodC.Text <> "" Then
strSQL = "select id_promotor,nombre from promotores where id_promotor = '" & TxtCodC.Text & "'"
                 rs.Open strSQL, glogon.Conection, adOpenStatic
               If Not rs.EOF Then
                 TxtCodC.Text = rs!id_promotor
                 TxtComites.Text = rs!Nombre
                 TxtCedula.SetFocus
               End If
        rs.Close
End If
End Sub

Sub sbComN()
  Dim strSQL As String
  Dim rs As New ADODB.Recordset
   
        strSQL = "select id_promotor,nombre from promotores where id_promotor = '" & TxtCodC.Text & "'"
                 rs.Open strSQL, glogon.Conection, adOpenStatic
               If Not rs.EOF Then
                 TxtComites.Text = rs!Nombre
                 TxtCedula.SetFocus
               End If
 rs.Close
End Sub

Sub sbGuardaCom(Codigo As Integer, Cedula As String)

Dim strSQL As String, rs As New ADODB.Recordset
Dim TxtCodC As Integer
Dim Fecha As String
Dim Cons As Integer
Dim IDpuesto As String

IDpuesto = fxCodigoCbo(CboPuesto)

Fecha = Format(fxFechaServidor, "yyyymmdd")
Cons = fxCons("id_seq", "afi_cd_nombramientos_h")
'TxtCodC = CboComites.ItemData(CboComites.ListIndex)

strSQL = "select id_pricomite,cedula,cod_puesto from afi_cd_nombramientos where id_pricomite = '" & Codigo & "'" _
         & "and cedula = '" & Cedula & "' and cod_puesto = '" & Trim(IDpuesto) & "'"
         rs.Open strSQL, glogon.Conection, adOpenStatic


If rs.EOF Then
    
    'Ingresa datos en AFI_CD_NOMBRAMIENTOS
    strSQL = "insert into afi_cd_nombramientos (id_pricomite,CEDULA,COD_PUESTO,APL_DESEMBOLSOS,FECHA,USUARIO,NOTAS,ESTADO)" _
              & " values('" & Codigo & "','" & Cedula & "','" & Trim(IDpuesto) & "'," _
              & " '" & ChkDesm.Value & "','" & Fecha & "','" & glogon.Usuario & "','" & UCase(txtNotas.Text) & "','" & ChkMiembro.Value & "')"
              glogon.Conection.Execute strSQL
    
    'Ingresa datos en AFI_CD_NOMBRAMIENTOS_H
    strSQL = "insert into afi_cd_nombramientos_h (ID_SEQ,id_pricomite,CEDULA,COD_PUESTO,APL_DESEMBOLSOS,FECHA,USUARIO,NOTAS,ESTADO)" _
              & " values('" & Cons & " ','" & Codigo & "','" & Cedula & "','" & IDpuesto & "'," _
              & " '" & ChkDesm.Value & "','" & Fecha & "','" & glogon.Usuario & "','" & UCase(txtNotas.Text) & "','" & ChkMiembro.Value & "')"
              glogon.Conection.Execute strSQL
              
              MsgBox "Los datos fueron ingresados correctamente", vbInformation, "Información"
Else
     
    'Actualizar datos en AFI_CD_NOMBRAMIENTOS
    strSQL = "update afi_cd_nombramientos set APL_DESEMBOLSOS = '" & ChkDesm.Value & "'" _
             & ",USUARIO ='" & glogon.Usuario & "',NOTAS='" & UCase(txtNotas.Text) & "'" _
             & ",ESTADO='" & ChkMiembro.Value & "' where id_pricomite ='" & Codigo & "' and " _
             & "CEDULA='" & Cedula & "' and COD_PUESTO = '" & Trim(IDpuesto) & "'"
             glogon.Conection.Execute strSQL
     
     
     
     ' Actualizar datos en AFI_CD_NOMBRAMIENTOS_H
    strSQL = "update afi_cd_nombramientos_H set APL_DESEMBOLSOS = '" & ChkDesm.Value & "'" _
             & ",USUARIO ='" & glogon.Usuario & "',NOTAS='" & UCase(txtNotas.Text) & "'" _
             & ",ESTADO='" & ChkMiembro.Value & "' where id_pricomite ='" & Codigo & "' and CEDULA='" & Cedula & "' and COD_PUESTO = '" & Trim(IDpuesto) & "'"
             glogon.Conection.Execute strSQL
    
             MsgBox "Los datos fueron actualizados correctamente", vbInformation, "Información"
End If
rs.Close
Call sbLimpia

End Sub

Sub sbListaComite()

Dim vProv As String

strSQL = "select distinct A.id_pricomite,U.descripcion,U.provincia from afi_cd_agrupacomites A left join uprogramatica U " _
         & "on A.id_pricomite = U.codigo"
         rs.Open strSQL, glogon.Conection, adOpenStatic
         LswComi.ListItems.Clear
        
        While Not rs.EOF
          Set itmX = LswComi.ListItems.Add(, , rs!id_pricomite)
          itmX.SubItems(1) = IIf(IsNull(rs!Descripcion), "Sin Nombre", rs!Descripcion)
          vProv = IIf(Not IsNull(rs!provincia = 1), "SAN JOSE", "NO DEFINIDA")
          vProv = IIf(Not IsNull(rs!provincia = 2), "ALAJUELA", "NO DEFINIDA")
          vProv = IIf(Not IsNull(rs!provincia = 3), "CARTAGO", "NO DEFINIDA")
          vProv = IIf(Not IsNull(rs!provincia = 4), "HEREDIA", "NO DEFINIDA")
          vProv = IIf(Not IsNull(rs!provincia = 5), "GUANACASTE", "NO DEFINIDA")
          vProv = IIf(Not IsNull(rs!provincia = 6), "PUNTARENAS", "NO DEFINIDA")
          vProv = IIf(Not IsNull(rs!provincia = 7), "LIMON", "NO DEFINIDA")
          itmX.SubItems(2) = vProv
          rs.MoveNext
        Wend
        rs.Close
End Sub

Sub sbLlamaCom()
    
End Sub

Sub sbLlamaPromo()
 Dim strSQL As String, rs As New ADODB.Recordset
  
'strsql = " select id_promotor,nombre,estado from promotores " _
'                & "where comite = 1"
'                rs.Open strsql, glogon.Conection, adOpenStatic
'While Not rs.EOF
'    CboComites.AddItem rs!Nombre
'    CboComites.ItemData(CboComites.NewIndex) = rs!id_promotor
'    rs.MoveNext
'Wend
'
'If rs.RecordCount > 0 Then
'    rs.MoveFirst
'    CboComites.Text = rs!Nombre
'End If
'rs.Close


End Sub
Sub sbMiembros(Codigo As Integer)
    
Dim itmX As ListItem
Dim TxtCodC As Integer
ChkActivo.Value = 0
ChkDesembolso.Value = 0
TxtNotasCon.Text = ""
SstComites.Tab = 1

If SstComites.Tab = 1 Then
LswMiembros.ListItems.Clear

strSQL = " select N.cedula,S.nombre,N.cod_puesto,P.descripcion,N.fecha" _
       & " from socios S right join afi_cd_nombramientos_h N" _
       & " on S.cedula = N.cedula inner join afi_cd_puestos P" _
       & " on n.cod_puesto = P.cod_puesto where N.id_pricomite = '" & Codigo & "'"
 
 Select Case ChkHistorial.Value
   Case 1
     rs.Open strSQL, glogon.Conection, adOpenStatic
   Case 0
     strSQL = strSQL & " And N.Estado = 1"
     rs.Open strSQL, glogon.Conection, adOpenStatic
 End Select
    
   While Not rs.EOF
      Set itmX = LswMiembros.ListItems.Add(, , Trim(rs!Cedula))
      itmX.SubItems(1) = rs!Nombre
      itmX.SubItems(2) = rs!Descripcion
      itmX.SubItems(3) = Format(rs!Fecha, "dd/mm/yyyy")
      rs.MoveNext
   Wend
  rs.Close
 End If
End Sub

Sub sbConsOpc()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select estado,apl_desembolsos,notas from afi_cd_nombramientos" _
             & " where cedula = '" & LswMiembros.SelectedItem & "'"
             rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF Then
 If LswMiembros.SelectedItem.Selected = True Then
    If Not rs.EOF Then
      ChkActivo.Value = rs!Estado
      ChkDesembolso.Value = rs!apl_desembolsos
      TxtNotasCon.Text = rs!notas
    End If
 End If
End If
rs.Close

Exit Sub

vError:
    MsgBox "No se encuentran ningún miembro para seleccionar", vbInformation, "Información"
End Sub

Sub SbPuestos()

Dim strSQL As String, rs As New ADODB.Recordset
CboPuesto.Clear

strSQL = " select cod_puesto,descripcion from afi_cd_puestos "
            rs.Open strSQL, glogon.Conection, adOpenStatic
While Not rs.EOF
 CboPuesto.AddItem Trim(rs!cod_puesto) & " - " & Trim(rs!Descripcion)
 rs.MoveNext
Wend
rs.Close

End Sub


Public Function fxCodigoCbo(cbo As ComboBox) As String

Dim i As Integer, vPaso As Boolean
Dim x As Integer

If cbo.ListCount = 0 Then
  fxCodigoCbo = ""
  Exit Function
End If

vPaso = True
i = 1
x = Len(cbo.Text)
Do While vPaso
  If Mid(cbo.Text, i, 1) = "-" Then
    vPaso = False
    i = i - 1
  Else
    i = i + 1
  End If
  If i = x Then Exit Do
Loop
fxCodigoCbo = Trim(Mid(cbo.Text, 1, i))

End Function

Sub sbLimpia()

TxtCedula.Text = ""
LblNombre.Caption = ""
ChkMiembro.Value = 0
ChkDesm.Value = 0
txtNotas.Text = ""


End Sub


Private Sub ChkHistorial_Click()
 
 If TxtCodC.Text <> "" Then
  Call sbMiembros(TxtCodC)
  If ChkHistorial.Value = 0 Then
   ChkActivo.Value = 0
   ChkDesembolso.Value = 0
   TxtNotasCon.Text = ""
  End If
End If
End Sub

Private Sub cmdAplica_Click()
 If fxValida Then Call sbGuardaCom(TxtCodC, Trim(TxtCedula))
End Sub

Private Sub CmdImp_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
 .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_Cd_MiembrosComite.rpt")
  
 If ChkTodos.Value = 1 Then
    Select Case True
       Case OptActivos.Value = True
         .WindowTitle = "Reporte de Comites y sus Miembros Activos"
         .Formulas(3) = "fxTitulo='Comites con sus Miembros Activos'"
         .SelectionFormula = "{afi_cd_nombramientos_h.estado} = '1'"
       Case OptNoact.Value = True
        .WindowTitle = "Reporte de Comites y sus Miembros No Activos"
        .Formulas(3) = "fxTitulo='Comites con sus Miembros No Activos'"
        .SelectionFormula = "{afi_cd_nombramientos_h.estado} = '0'"
    End Select
 Else
    Select Case True
       Case OptActivos.Value = True
         .WindowTitle = "Reporte de Comites y sus Miembros Activos"
         .SelectionFormula = "{afi_cd_nombramientos.id_pricomite} = '" & LswComi.SelectedItem & "' " _
         & "and {afi_cd_nombramientos_h.estado} = '1'"
         .Formulas(3) = "fxTitulo='Comites con sus Miembros Activos'"
       Case OptNoact.Value = True
         .WindowTitle = "Reporte de Comites y sus Miembros No Activos"
         .SelectionFormula = "{afi_cd_nombramientos.id_pricomite} = '" & LswComi.SelectedItem & "' " _
         & "and {afi_cd_nombramientos_h.estado} = '0'"
         .Formulas(3) = "fxTitulo='Comites con sus Miembros No Activos'"
    End Select
 End If
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub CmdMiembros_Click()
 SstComites.Tab = 1
 If Not TxtCodC.Text = "" Then Call sbMiembros(TxtCodC)
  If TxtCodC.Text = "" Then
    LswMiembros.ListItems.Clear
  End If
End Sub

Private Sub CmdNuevo_Click()
 Call sbLimpia
End Sub
Private Sub Form_Load()
 
 staUse.Panels.Item(1) = glogon.Usuario
 Call sbLlamaPromo
 Call SbPuestos
 SstComites.Tab = 0
 vLista = True

End Sub

Private Function fxValida() As Boolean

Dim vMensaje As String

vMensaje = ""
fxValida = True

If CboPuesto.Text = "" Then vMensaje = vMensaje & vbCrLf & "No se definio los puestos"
If TxtComites.Text = "" Then vMensaje = vMensaje & vbCrLf & "No se definio el comite"
If TxtCedula.Text = "" Then vMensaje = vMensaje & vbCrLf & "No se especifico una cédula"


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function



Private Sub LblCodC_Click()

End Sub

Private Sub LswComi_DblClick()

Dim i As Integer
  
 For i = 1 To LswComi.ListItems.Count
  If LswComi.ListItems.Item(i).Selected = True Then
   TxtCodC.Text = LswComi.ListItems.Item(i)
   SstComites.Tab = 1
   LswMiembros.ListItems.Clear
   TxtCodC.SetFocus
   TxtComites.SetFocus
   TxtComites.Text = FxNomComite(TxtCodC.Text)
   Call CmdMiembros_Click
   
   End If
 Next i

End Sub


Private Sub LswMiembros_Click()
 Call sbConsOpc
End Sub

Private Sub LswMiembros_KeyDown(KeyCode As Integer, Shift As Integer)
Call sbConsOpc
End Sub


Private Sub LswMiembros_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbConsOpc
End Sub


Private Sub Picture1_Click()
 Call sbListaComite
End Sub

Private Sub SstComites_Click(PreviousTab As Integer)

 Dim strSQL As String, rs As New ADODB.Recordset
 Dim itmX As ListItem
 
 Select Case True
   Case SstComites.Tab = 0
     Call SbPuestos
   Case SstComites.Tab = 1
     Call sbLimpia
   Case SstComites.Tab = 2
     If vLista = True Then
      Call sbListaComite
      vLista = False
     End If
 End Select

End Sub

 
Private Sub TxtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "cedula"
       gBusquedas.Consulta = "select top 1000 cedula,nombre from socios"
       gBusquedas.Filtro = " and estadoactual = 'S'"
       frmBusquedas.Show vbModal
       TxtCedula.SetFocus
       vCodigo = gBusquedas.Resultado
       TxtCedula = gBusquedas.Resultado
       LblNombre.Caption = gBusquedas.Resultado2
End If
End Sub


Private Sub TxtCedula_KeyPress(KeyAscii As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

Select Case KeyAscii
  Case 48 To 57, 8
  Case 13
   If TxtCedula.Text = "" Then
        MsgBox "No se defino la cedula del asociado", vbInformation + vbOKOnly, "Información"
     Else
       strSQL = " select cedula,nombre,estadoactual from socios " _
                & "where estadoactual = 'S' and cedula = '" & Trim(TxtCedula.Text) & "'"
                rs.Open strSQL, glogon.Conection, adOpenStatic
                  
                  If Not rs.EOF Then
                     TxtCedula.Text = rs!Cedula
                     LblNombre.Caption = rs!Nombre
                  End If
        rs.Close
             
       strSQL = " select N.id_pricomite,N.cedula,S.nombre,N.estado,N.apl_desembolsos,N.estado,P.descripcion,N.cod_puesto,N.notas " _
                & "from afi_cd_puestos P inner join afi_cd_nombramientos N on P.cod_puesto = N.cod_puesto " _
                & "left join Socios S on N.cedula = S.cedula " _
                & "where N.cedula = '" & Trim(TxtCedula.Text) & "'"
                rs.Open strSQL, glogon.Conection, adOpenStatic
                  
                If Not rs.EOF Then
                     TxtCedula.Text = rs!Cedula
                     LblNombre.Caption = rs!Nombre
                     ChkMiembro.Value = rs!Estado
                     CboPuesto.Text = rs!cod_puesto & " - " & rs!Descripcion
                     ChkDesm.Value = rs!apl_desembolsos
                     txtNotas.Text = rs!notas
                     TxtCodC.Text = rs!id_pricomite
                     Call sbComN
'                   For I = 0 To CboComites.ListCount - 1
'                    CboComites.ListIndex = I
'                    If CboComites.ItemData(CboComites.ListIndex) = rs!id_promotor Then Exit For
'                   Next I
                End If
        rs.Close
      
    End If
 Case Else
   KeyAscii = 0
End Select
End Sub


Private Sub TxtCodC_Click()
 TxtComites.Text = ""
 LswMiembros.ListItems.Clear
End Sub

Private Sub TxtCodC_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        TxtComites.Text = FxNomComite(TxtCodC.Text)
      Case Else
       KeyAscii = 0
    End Select
End Sub

Private Sub TxtCodC_LostFocus()
If TxtCodC.Text = Empty Or TxtComites.Text = Empty Then
 Call sbComite
End If
End Sub

Private Sub TxtComites_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select distinct A.id_pricomite,U.descripcion from afi_cd_agrupacomites A " _
                             & "left join uprogramatica U on A.id_pricomite = U.codigo " _
                             '& "group by A.id_pricomite,U.descripcion"
       
       'gBusquedas.Filtro = " and comite = 1"
       frmBusquedas.Show vbModal
       TxtCodC.Text = gBusquedas.Resultado
       TxtComites.Text = gBusquedas.Resultado2
       TxtCedula.SetFocus
       LswMiembros.ListItems.Clear
End If
End Sub

