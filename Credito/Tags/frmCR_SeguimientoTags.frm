VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_SeguimientoTags 
   Caption         =   "Aplicación de Etiquetas"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   12720
   Begin VB.TextBox txtNota 
      Height          =   495
      Left            =   1320
      TabIndex        =   18
      Top             =   7680
      Width           =   9015
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Digitación"
      TabPicture(0)   =   "frmCR_SeguimientoTags.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tlbDigitacion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lswOperaciones"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAgregar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtOperacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "frmCR_SeguimientoTags.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "tlbBuscar"
      Tab(1).Control(4)=   "dtpFFin"
      Tab(1).Control(5)=   "dtpFInicio"
      Tab(1).Control(6)=   "vGrid"
      Tab(1).Control(7)=   "chkEspera"
      Tab(1).Control(8)=   "chkRevisados"
      Tab(1).Control(9)=   "cboEstado"
      Tab(1).Control(10)=   "cboDocumentacion"
      Tab(1).Control(11)=   "chkTodos"
      Tab(1).ControlCount=   12
      Begin VB.TextBox txtOperacion 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1440
         TabIndex        =   23
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
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
         Left            =   -74760
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cboDocumentacion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -66120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Estado de Operación"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboEstado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Estado de Operación"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkRevisados 
         Caption         =   "Solo Revisados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkEspera 
         Caption         =   "Solo en Espera"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70080
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   11
         Top             =   1560
         Width           =   12015
         _Version        =   524288
         _ExtentX        =   21193
         _ExtentY        =   7646
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   493
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_SeguimientoTags.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpFInicio 
         Height          =   330
         Left            =   -73560
         TabIndex        =   12
         ToolTipText     =   "Fecha Inicio Búsqueda"
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   358154243
         CurrentDate     =   40361
      End
      Begin MSComCtl2.DTPicker dtpFFin 
         Height          =   330
         Left            =   -71880
         TabIndex        =   13
         ToolTipText     =   "Fecha Fin Búsqueda"
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   358154243
         CurrentDate     =   40361
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   330
         Left            =   -64200
         TabIndex        =   14
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         ButtonWidth     =   1640
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswOperaciones 
         Height          =   4695
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   8281
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Operación"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Línea"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cédula"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Oficina"
            Object.Width           =   8114
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbDigitacion 
         Height          =   312
         Left            =   9720
         TabIndex        =   25
         Top             =   600
         Width           =   2268
         _ExtentX        =   3995
         _ExtentY        =   556
         ButtonWidth     =   1852
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "Eliminar"
               Object.ToolTipText     =   "Eliminar Crédito"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Limpiar"
               Key             =   "Limpiar"
               Object.ToolTipText     =   "Limpiar Lista"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Formalización:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70080
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Documentación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67440
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboEtiquetas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCR_SeguimientoTags.frx":0F6B
      Left            =   2040
      List            =   "frmCR_SeguimientoTags.frx":0F6D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   5295
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   12000
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTags.frx":0F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTags.frx":77D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTags.frx":E033
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTags.frx":14895
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAplicar 
      Height          =   570
      Left            =   10920
      TabIndex        =   19
      Top             =   7680
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1005
      ButtonWidth     =   2117
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Aplicar Etiqueta"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label lblObservacion 
      Caption         =   "Observación"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblUsuario 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmCR_SeguimientoTags.frx":1B0F7
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmCR_SeguimientoTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem
Private mCargaGrid As Boolean

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub chkTodos_Click()
    If chkTodos.Value = vbChecked Then
        Call sbGridMarcarTodo(True)
    Else
        Call sbGridMarcarTodo(False)
    End If
End Sub

Private Sub cmdAgregar_Click()
    Call sbCargaOperacion
End Sub

Private Sub Form_Activate()

vModulo = 8
End Sub

Private Sub Form_Resize()
'' Procedimiento para posicionar los controles al max y minimizar la pantalla
On Error GoTo vError
        ssTab.Width = Me.Width - 600
        vGrid.Width = ssTab.Width - 450
        lswOperaciones.Width = vGrid.Width

        ssTab.Height = Me.Height - 3000
        vGrid.Height = ssTab.Height - 1800
        lswOperaciones.Height = vGrid.Height + 300
        

        tlbAplicar.top = ssTab.top + ssTab.Height + 300
        txtNota.top = tlbAplicar.top
        lblObservacion.top = tlbAplicar.top
    Exit Sub

vError:

End Sub

Private Sub Form_Load()


vModulo = 8

    mCargaGrid = True

    lblUsuario.Caption = glogon.Usuario
    lblNombre.Caption = fxNombreUsuario
    
    vGrid.AppearanceStyle = fxGridStyle
    ssTab.Tab = 0
    
    vGrid.MaxRows = 0
    vGrid.MaxCols = 11
    
    Call sbCargarCombos
    
    Me.Width = 12960
    Me.Height = 8970
       
End Sub

Private Function fxNombreUsuario() As String
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

    strSQL = "select DESCRIPCION from USUARIOS where Nombre = '" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxNombreUsuario = rs.Fields(0)
    Else
        fxNombreUsuario = Empty
    End If

Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function


Private Sub sbCargarCombos()
Dim strSQL As String, rs As New ADODB.Recordset, strPrimero As String

On Error GoTo vError

    strSQL = "SELECT CT.TAG_CODIGO as llave,CT.DESCRIPCION as describe FROM CRD_TAGS CT INNER JOIN CRD_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
           & " INNER JOIN CRD_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
           & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
           & "' order by CT.TAG_CODIGO"
    Call OpenRecordSet(rs, strSQL)

    If Not rs.EOF And Not rs.BOF Then strPrimero = Trim(rs!llave) & " - " & Trim(rs!describe)

    Do While Not rs.EOF
      cboEtiquetas.AddItem Trim(rs!llave) & " - " & Trim(rs!describe)
      rs.MoveNext
    Loop
    If strPrimero <> "" Then cboEtiquetas.Text = strPrimero
    rs.Close
    
    cboEstado.Clear
    cboEstado.AddItem ("Todos")
    cboEstado.AddItem ("Recibida")
    cboEstado.AddItem ("Pendiente")
    cboEstado.AddItem ("Formalizada")
    cboEstado.Text = "Todos"
    
    cboDocumentacion.Clear
    cboDocumentacion.AddItem ("Todos")
    cboDocumentacion.AddItem ("Recepción")
    cboDocumentacion.AddItem ("Devolución")
    cboDocumentacion.Text = "Todos"
    
    dtpFInicio.Value = fxFechaServidor
    dtpFFin.Value = dtpFInicio.Value

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarListaSolicitudes()
' Carga Lista de operaciones
    Dim strSQL As String
    
On Error GoTo error
    'Consulta la lista de las Operaciones
    
    Me.MousePointer = vbHourglass
    
    mCargaGrid = True
    
    strSQL = "SELECT RC.ID_SOLICITUD,RC.FECHAFORP,S.NOMBRE,RC.CODIGO,RC.MONTOSOL,RC.CUOTA,RC.PLAZO,RC.INT," _
            & "case RC.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else RC.ESTADOSOL end, RC.FECHASOL " _
            & "FROM REG_CREDITOS RC INNER JOIN SOCIOS S ON RC.CEDULA = S.CEDULA " _
            & "WHERE RC.FECHAFORP BETWEEN '" & Format(dtpFInicio, "yyyy/mm/dd") & "' AND '" _
            & Format(dtpFFin, "yyyy/mm/dd") & "'"
            
    Select Case cboEstado.Text
    Case "Todos"
        strSQL = strSQL & " and RC.ESTADOSOL in ('P','R','F')"
    Case "Recibida"
        strSQL = strSQL & " and RC.ESTADOSOL = 'R'"
    Case "Pendiente"
        strSQL = strSQL & " and RC.ESTADOSOL = 'P'"
    Case "Formalizada"
        strSQL = strSQL & " and RC.ESTADOSOL = 'F'"
    End Select
    
    If Not cboDocumentacion.Text = "Todos" Then
        Select Case cboDocumentacion.Text
        Case "Recepción"
            strSQL = strSQL & " and isnull(RC.ANALISTAS_RECEPCION,'P') = 'R'"
        Case "Devolución"
            strSQL = strSQL & " and isnull(RC.ANALISTAS_RECEPCION,'P') = 'D'"
        End Select
    End If
    
    If chkRevisados.Value = vbChecked Then
        strSQL = strSQL & " and isnull(ANALISTAS_REVISION,0) = 1"
    End If
        
    If chkEspera.Value = vbChecked Then
        strSQL = strSQL & " and not EN_ESPERA_FECHA is null"
    End If
    
    strSQL = strSQL & " order by RC.ID_SOLICITUD"
        
    Call sbCargaGridCheckIni(vGrid, 10, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
    
    mCargaGrid = False
    chkTodos.Value = Unchecked
    Me.MousePointer = vbDefault
    Exit Sub
error:
    mCargaGrid = False
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub

Public Sub sbCargaGridCheckIni(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
'Procedimiento para cargar grids con el check en la primera columna
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

    vGrid.MaxCols = vGridMaxCol + 1
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.col = i
     vGrid.Text = ""
    Next i
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      vGrid.Row = vGrid.MaxRows
      For i = 4 To vGrid.MaxCols
        vGrid.col = i
        vGrid.Text = CStr(rs.Fields(i - 4).Value & "")
      Next i
      vGrid.MaxRows = vGrid.MaxRows + 1
      rs.MoveNext
    Loop
    rs.Close
    Exit Sub

vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    Select Case ssTab.Tab
    Case 0
        txtOperacion.Text = Empty
        lswOperaciones.ListItems.Clear
    Case 1
        vGrid.MaxRows = 0
        vGrid.MaxCols = 11
    End Select
End Sub




Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)

    If MsgBox("Está seguro que sea aplicar la etiqueta en las operaciones seleccionadas", vbExclamation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If ssTab.Tab = 0 Then
        Call sbIncluirEtiquetasLista
    Else
        If vGrid.MaxRows = 0 Then Exit Sub
        Call sbIncluirEtiquetasGrid
        Call sbCargarListaSolicitudes
    End If
    
    txtNota.Text = Empty
    
End Sub

Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
    chkEspera.SetFocus
    Call sbCargarListaSolicitudes
End Sub

Private Sub sbIncluirEtiquetasLista()
Dim i As Integer, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

With lswOperaciones.ListItems
    For i = 1 To .Count
        Call sbCrdOperacionTags(.Item(i).Text, .Item(i).SubItems(1), SIFGlobal.fxCodText(cboEtiquetas), "", txtNota.Text)
        Call Sleep(100)  'Esperar 100 milisegundos para evitar que se asigne la misma fecha de revisión a dos líneas diferentes
    Next i
    .Clear
End With
Me.MousePointer = vbDefault

MsgBox "Proceso concluido con éxito...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbIncluirEtiquetasGrid()
'' Procedimiento para Ingresar las etiquetas marcadas en las operaciones seleccionadas

Dim IdSolicitud As Long, Linea As String, i As Integer

On Error GoTo vError

    vGrid.Row = 1
    vGrid.col = 1
    For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        
        If vGrid.Value = vbChecked Then
        
                vGrid.col = 4
                IdSolicitud = vGrid.Value
                
                If IdSolicitud <> Empty Then
                
                    vGrid.col = 7
                    Linea = vGrid.Value
                    
                    Call sbCrdOperacionTags(IdSolicitud, Linea, SIFGlobal.fxCodText(cboEtiquetas), "", txtNota.Text)
                    
                End If
                
                vGrid.col = 1
        End If
    Next i
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub tlbDigitacion_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "LIMPIAR"
        lswOperaciones.ListItems.Clear
    Case "ELIMINAR"
        If lswOperaciones.ListItems.Count > 0 Then
            If lswOperaciones.SelectedItem.Index <> 0 Then
                lswOperaciones.ListItems.Remove (lswOperaciones.SelectedItem.Index)
            End If
        End If
    End Select
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargaOperacion
    End If
End Sub

Private Sub sbCargaOperacion()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

    If Not IsNumeric(txtOperacion) Then
        Exit Sub
    End If
    
    If fxValidaNoDuplicados = True Then
        MsgBox "La operación se ya fue digitada"
        txtOperacion.Text = Empty
        txtOperacion.SetFocus
        Exit Sub
    End If
    
    strSQL = "SELECT R.ID_SOLICITUD,R.CODIGO,R.CEDULA,R.FECHAFORF,isnull(O.DESCRIPCION,'') as DESCRIPCION FROM REG_CREDITOS R LEFT JOIN SIF_OFICINAS O ON R.COD_OFICINA_R = O.COD_OFICINA WHERE R.ID_SOLICITUD = " & Trim(txtOperacion)
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
         Set itmX = lswOperaciones.ListItems.Add(, , rs!Id_Solicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!Cedula
        itmX.SubItems(3) = rs!Descripcion
    End If
    rs.Close

    txtOperacion.Text = Empty
    txtOperacion.SetFocus

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim frm As Form
    If Not mCargaGrid Then
    
        Dim IdSolicitud As Long
        
        vGrid.Row = Row
        vGrid.col = 4
        
        If vGrid.Value = Empty Then Exit Sub
        
        IdSolicitud = vGrid.Value
        
        If IdSolicitud = Empty Then Exit Sub
        
        Select Case col
        Case 2
            Call sbFormsCall("frmCR_SeguimientoTramites")
            For Each frm In Forms
                If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbConsultaExterna(IdSolicitud)
                    Exit For
                End If
            Next frm
        Case 3
            Operacion.Operacion = IdSolicitud
            If Operacion.Operacion > 0 Then
                Call sbFormsCall("frmCR_SeguimientoEtiquetas", 1, , , False, Me)
            End If
        End Select
    
    End If
End Sub

Public Sub sbGridMarcarTodo(Habilita As Boolean)
' Procedimiento para marcar todos los checks en un grid
Dim i As Long

On Error GoTo vError

    vGrid.Row = 1
    vGrid.col = 1
    For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        vGrid.col = 1
        If Habilita Then
            vGrid.Value = vbChecked
        Else
            vGrid.Value = Unchecked
        End If
    Next i
    Exit Sub
vError:
        MsgBox fxSys_Error_Handler(Err.Description)
    
End Sub


Private Function fxValidaNoDuplicados() As Boolean
Dim i As Integer

    fxValidaNoDuplicados = False

    For i = 1 To lswOperaciones.ListItems.Count

        If lswOperaciones.ListItems(i).Text = Trim(txtOperacion.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function



