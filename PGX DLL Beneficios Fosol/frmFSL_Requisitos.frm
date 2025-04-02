VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Requisitos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Catálogo de Requisitos"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Definición"
      TabPicture(0)   =   "frmFSL_Requisitos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asignación"
      TabPicture(1)   =   "frmFSL_Requisitos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboCausa"
      Tab(1).Control(1)=   "cboTipo"
      Tab(1).Control(2)=   "vGridAsg"
      Tab(1).Control(3)=   "Label1(1)"
      Tab(1).Control(4)=   "Label1(0)"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox cboCausa 
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
         Left            =   -72120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   6015
      End
      Begin VB.ComboBox cboTipo 
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
         Left            =   -72120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   6015
      End
      Begin FPSpreadADO.fpSpread vGridAsg 
         Height          =   4095
         Left            =   -74280
         TabIndex        =   2
         Top             =   1440
         Width           =   8535
         _Version        =   524288
         _ExtentX        =   15055
         _ExtentY        =   7223
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Requisitos.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5055
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   7515
         _Version        =   524288
         _ExtentX        =   13256
         _ExtentY        =   8916
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   3
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Requisitos.frx":064D
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Causa"
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
         Left            =   -74160
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Aplicación"
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
         Left            =   -74160
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -74160
         TabIndex        =   4
         Top             =   2760
         Width           =   8055
      End
      Begin VB.Label Label3 
         Caption         =   "Nivel de Aplicación del Requisito"
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
         Left            =   -74160
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmFSL_Requisitos.frx":0C5F
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Requisitos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmFSL_Requisitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean



Private Sub cboCausa_Click()
Dim strSQL As String

If vPaso Then Exit Sub
If cboCausa.ListCount = 0 Or cboTipo.ListCount = 0 Then Exit Sub

vPaso = True
strSQL = "select rtrim(cod_Causa) + ' - ' + descripcion as 'ItmX'" _
       & " from FSL_Planes_Causas where activa = 1 and cod_plan = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) & "'"

strSQL = "select Rq.COD_REQUISITO, Rq.DESCRIPCION, isnull(Rc.OPCIONAL,0) as 'Opcional' , isnull(Rc.ASIGNADO,0) as 'Asignado'" _
       & " from FSL_REQUISITOS Rq left join FSL_REQUISITOS_CAUSAS Rc on Rq.COD_REQUISITO = Rc.COD_REQUISITO" _
       & " and Rc.COD_PLAN = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) _
       & "' and Rc.COD_CAUSA = '" & SIFGlobal.fxSIFCodText(cboCausa.Text) _
       & "' Where Rq.ACTIVO = 1"

Call sbCargaGrid(vGridAsg, 4, strSQL, True)

vPaso = False

End Sub

Private Sub cboTipo_Click()
Dim strSQL As String

If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

vPaso = True
strSQL = "select rtrim(cod_Causa) + ' - ' + descripcion as 'ItmX'" _
       & " from FSL_Planes_Causas where activa = 1 and cod_plan = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) & "'"

Call sbLlenaCbo(cboCausa, strSQL, False, False)

vPaso = False

Call cboCausa_Click

End Sub

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 22
vGrid.AppearanceStyle = fxGridStyle

ssTab.Tab = 0

strSQL = "select COD_REQUISITO,descripcion,Activo from FSL_REQUISITOS order by COD_REQUISITO"
Call sbCargaGrid(vGrid, 3, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pExiste As Long, pCodigo As String, pTabla As String

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

pCodigo = "COD_REQUISITO"
pTabla = "FSL_REQUISITOS"

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then Exit Function

strSQL = "select isnull(count(*),0) as Existe from " & pTabla _
       & " where " & pCodigo & " = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic


If rs!existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & " ,Descripcion, Activo,registro_fecha,registro_usuario) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Requisitos (Lista)  Id.:" & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update " & pTabla & " set Descripcion = '" & vGrid.Text & "', Activo = "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & " where " & pCodigo & " = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Modifica", "Requisitos (Lista)  Id.:" & vGrid.Text)

End If

rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function





Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

If ssTab.Tab = 0 Then Exit Sub

vPaso = True
strSQL = "select rtrim(cod_Plan) + ' - ' + descripcion as 'ItmX' from FSL_Planes where activo = 1 "

Call sbLlenaCbo(cboTipo, strSQL, False, False)

vPaso = False

Call cboTipo_Click

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

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
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete FSL_REQUISITOS where COD_REQUISITO = '" & vGrid.Text & "'"
        glogon.Conection.Execute strSQL

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Requisitos (Lista)  Id.:" & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


Private Sub vGridAsg_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, vMovimiento As String
Dim vTempo As Integer, pPlan As String, pCausa As String

If vPaso Then Exit Sub


pPlan = SIFGlobal.fxSIFCodText(cboTipo.Text)
pCausa = SIFGlobal.fxSIFCodText(cboCausa.Text)


With vGridAsg

   .Row = Row
   .Col = Col
   
   If Col = 4 Then 'Ultima Columna
      If .Value = 1 Then
         .Col = 3
         vTempo = .Value
         .Col = 1
         vMovimiento = "Registra"
         strSQL = "insert FSL_REQUISITOS_CAUSAS(COD_PLAN,COD_CAUSA,cod_requisito,opcional,asignado,registro_fecha,registro_usuario)" _
                & " values('" & pPlan & "','" & pCausa & "','" & .Text & "'," & vTempo & ",1,getdate(),'" & glogon.Usuario & "')"
      Else
         .Col = 1
         vMovimiento = "Borrar"
         strSQL = "delete FSL_REQUISITOS_CAUSAS where COD_PLAN = '" & pPlan & "' and cod_Causa = '" & pCausa _
                & "' and cod_requisito = '" & .Text & "'"
         
       End If
       
       glogon.Conection.Execute strSQL
       Call Bitacora(vMovimiento, "Requisito : " & .Text & " Plan: " & pPlan & " Causa: " & pCausa)
   End If

   If Col = 3 Then 'Columna de Opcional
      .Col = 3
      vTempo = .Value
      .Col = 4
      If .Value = 1 Then
          .Col = 1
          vMovimiento = "Modifica"
          strSQL = "update FSL_REQUISITOS_CAUSAS set Opcional = " & vTempo _
                 & " where COD_PLAN = '" & pPlan & "' and cod_Causa = '" & pCausa _
                 & "' and cod_requisito = '" & .Text & "'"
          
          glogon.Conection.Execute strSQL
          Call Bitacora(vMovimiento, "Requisito : " & .Text & " Plan: " & pPlan & " Causa: " & pCausa)
      End If
   End If


End With


End Sub

