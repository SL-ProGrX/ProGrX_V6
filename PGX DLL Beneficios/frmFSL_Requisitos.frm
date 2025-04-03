VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmFSL_Requisitos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Catálogo de Requisitos"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   8775
      _Version        =   1441793
      _ExtentX        =   15478
      _ExtentY        =   12515
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
      Item(0).Caption =   "Requisitos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "cboCausa"
      Item(1).Control(1)=   "cboTipo"
      Item(1).Control(2)=   "vGridAsg"
      Item(1).Control(3)=   "Label1(1)"
      Item(1).Control(4)=   "Label1(0)"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6375
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   7515
         _Version        =   524288
         _ExtentX        =   13256
         _ExtentY        =   11245
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   3
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Requisitos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridAsg 
         Height          =   5295
         Left            =   -69880
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   8535
         _Version        =   524288
         _ExtentX        =   15055
         _ExtentY        =   9340
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Requisitos.frx":065C
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   -67720
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
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
      Begin XtremeSuiteControls.ComboBox cboCausa 
         Height          =   330
         Left            =   -67720
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
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
      Begin VB.Label Label1 
         Caption         =   "Tipo de Aplicación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -69760
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Causa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -69760
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requisitos"
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
      Height          =   492
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
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

strSQL = "select Rq.COD_REQUISITO, Rq.DESCRIPCION, isnull(Rc.OPCIONAL,0) as 'Opcional' , isnull(Rc.ASIGNADO,0) as 'Asignado'" _
       & " from FSL_REQUISITOS Rq left join FSL_REQUISITOS_CAUSAS Rc on Rq.COD_REQUISITO = Rc.COD_REQUISITO" _
       & " and Rc.COD_PLAN = '" & cboTipo.ItemData(cboTipo.ListIndex) _
       & "' and Rc.COD_CAUSA = '" & cboCausa.ItemData(cboCausa.ListIndex) _
       & "' Where Rq.ACTIVO = 1"

Call sbCargaGrid(vGridAsg, 4, strSQL, True)

vPaso = False

End Sub

Private Sub cboTipo_Click()
Dim strSQL As String

If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

vPaso = True
strSQL = "select rtrim(cod_Causa) as 'idx', rtrim(descripcion) as 'ItmX'" _
       & " from FSL_Planes_Causas where activa = 1 and cod_plan = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"

Call sbCbo_Llena_New(cboCausa, strSQL, False, True)

vPaso = False

Call cboCausa_Click

End Sub

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 7
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

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
Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & " ,Descripcion, Activo,registro_fecha,registro_usuario) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Requisitos (Lista)  Id.:" & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update " & pTabla & " set Descripcion = '" & vGrid.Text & "', Activo = "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & " where " & pCodigo & " = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Requisitos (Lista)  Id.:" & vGrid.Text)

End If

rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 0 Then Exit Sub

vPaso = True
strSQL = "select rtrim(cod_Plan) as 'IdX', rtrim(descripcion) as 'ItmX' from FSL_Planes where activo = 1 "

Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

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
        Call ConectionExecute(strSQL)

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
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGridAsg_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, vMovimiento As String
Dim vTempo As Integer, pPlan As String, pCausa As String

If vPaso Then Exit Sub


pPlan = cboTipo.ItemData(cboTipo.ListIndex)
pCausa = cboCausa.ItemData(cboCausa.ListIndex)


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
       
       Call ConectionExecute(strSQL)
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
          
          Call ConectionExecute(strSQL)
          Call Bitacora(vMovimiento, "Requisito : " & .Text & " Plan: " & pPlan & " Causa: " & pCausa)
      End If
   End If


End With


End Sub

