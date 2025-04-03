VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmSYS_Educacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Mantenimientos de Centros Educativos"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   7200
      Top             =   360
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7095
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   9015
      _Version        =   1572864
      _ExtentX        =   15901
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
      ItemCount       =   1
      Item(0).Caption =   "Educación"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "scTitulo"
      Item(0).Control(2)=   "lsw"
      Item(0).Control(3)=   "btnAsigna(0)"
      Item(0).Control(4)=   "btnAsigna(1)"
      Item(0).Control(5)=   "btnAsigna(2)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2895
         Left            =   0
         TabIndex        =   2
         Top             =   4200
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   5106
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAsigna 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   3720
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Niveles"
         BackColor       =   14737632
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   2775
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   9015
         _Version        =   524288
         _ExtentX        =   15901
         _ExtentY        =   4895
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
         MaxCols         =   497
         ScrollBars      =   2
         SpreadDesigner  =   "frmSYS_Educacion.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnAsigna 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   3720
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Carreras"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.PushButton btnAsigna 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   3720
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Especialidades"
         BackColor       =   14737632
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   3240
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "..."
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
         Alignment       =   1
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento de Centros Educativos"
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
      Height          =   372
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   3252
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSYS_Educacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim mSheet As Integer
Dim vPaso As Boolean

Private Sub btnAsigna_Click(Index As Integer)
Dim i As Integer
Dim vTipoId As String

For i = 0 To 2
    btnAsigna(i).Checked = False
Next i

btnAsigna(Index).Checked = True

Select Case Index
  Case 0 'Niveles
    vTipoId = "N"
  Case 1 'Carreras
    vTipoId = "C"
  Case 2 'Especialidades
    vTipoId = "E"
End Select

lsw.ListItems.Clear
lsw.Checkboxes = True

vPaso = True

strSQL = "exec spSys_Educacion_Asigna_Consulta '" & scTitulo.Tag & "', '" & vTipoId & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_Educ)
      itmX.SubItems(1) = rs!DESCRIPCION
      
      itmX.Checked = IIf(rs!ASIGNADO = 1, True, False)
      
  rs.MoveNext
Loop
rs.Close

vPaso = False


End Sub

Private Sub Form_Load()
vModulo = 10
 
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", lsw.Width - (1400)
End With
 

Call vGrid_SheetChanged(2, 1)
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
lsw.Enabled = vGrid.Enabled
End Sub


Private Function fxGuardar() As Long

Dim vTipo As String, vTipoId As String

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


Select Case mSheet
  Case 1 'Universidades
        vTipoId = "U"
        vTipo = "Universidades: "
  Case 2 'Nivel
        vTipoId = "N"
        vTipo = "Nivel Educativo: "
  Case 3 'Carreras
        vTipoId = "C"
        vTipo = "Carreras Educativas: "
  Case 4 'Especialidades
        vTipoId = "E"
        vTipo = "Especialidades: "
End Select

If Trim(vGrid.Text) = "" Then Exit Function

strSQL = "select isnull(count(*),0) as Existe from SYS_EDUCACION_CFG " _
       & " where cod_Educ = '" & vGrid.Text & "' and Tipo <> '" & vTipoId & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 1 Then
   MsgBox "El código no puede ser utilizado! Ya existe un item diferente con su uso!", vbExclamation
   Exit Function
End If


strSQL = "select isnull(count(*),0) as Existe from SYS_EDUCACION_CFG " _
       & " where cod_Educ = '" & vGrid.Text & "' and Tipo = '" & vTipoId & "'"

Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  
  strSQL = "insert SYS_EDUCACION_CFG(cod_Educ, Tipo, descripcion, Activa, Registro_Usuario, Registro_Fecha) values('"

         
  strSQL = strSQL & vGrid.Text & "', '" & vTipoId & "', '"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ", '" & glogon.Usuario & "' ,dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", vTipo & vGrid.Text)

Else 'Actualizar
    
    vGrid.Col = 2
    strSQL = "update SYS_EDUCACION_CFG set descripcion = '" & vGrid.Text & "', Activa = "
    
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Value & " where cod_Educ = '"
          
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "'"
    
    Call ConectionExecute(strSQL)
    
    vGrid.Col = 1
    Call Bitacora("Modifica", vTipo & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spSys_Educacion_Asigna '" & scTitulo.Tag & "', '" & Item.Text & "', '" & glogon.Usuario & "', '" & IIf(Item.Checked, "A", "E") & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If vGrid.Sheet = 1 Or vGrid.Sheet = 3 Then
    vGrid.Row = Row
    vGrid.Col = 1
    scTitulo.Tag = vGrid.Text
    vGrid.Col = 2
    scTitulo.Caption = vGrid.Text
    
    If vGrid.Sheet = 1 Then
        btnAsigna(0).Enabled = True
        btnAsigna(1).Enabled = True
        btnAsigna(2).Enabled = False
        Call btnAsigna_Click(0)
    End If

    If vGrid.Sheet = 3 Then
        btnAsigna(0).Enabled = False
        btnAsigna(1).Enabled = False
        btnAsigna(2).Enabled = True
        Call btnAsigna_Click(2)
    End If

End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, vTipo As String, vTipoId As String

On Error GoTo vError

vGrid.Sheet = mSheet

If vGrid.ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        
        Select Case mSheet
          Case 1 'Universidades
                vTipoId = "U"
                vTipo = "Universidades: "
          Case 2 'Nivel
                vTipoId = "N"
                vTipo = "Nivel Educativo: "
          Case 3 'Carreras
                vTipoId = "C"
                vTipo = "Carreras Educativas: "
          Case 4 'Especialidades
                vTipoId = "E"
                vTipo = "Especialidades: "
        End Select
        
        strSQL = "delete SYS_EDUCACION_CFG where cod_Educ = '" & vGrid.Text & "' and Tipo = '" & vTipoId & "'"
        
        
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", vTipo & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)

Dim vTipoId As String

vGrid.Sheet = NewSheet
mSheet = NewSheet
tcMain.Item(0).Caption = vGrid.SheetName
 
Select Case mSheet
  Case 1 'Universidades
        vTipoId = "U"
  Case 2 'Nivel
        vTipoId = "N"
  Case 3 'Carreras
        vTipoId = "C"
  Case 4 'Especialidades
        vTipoId = "E"
End Select
 

vGrid.Sheet = mSheet
vGrid.ActiveSheet = mSheet


vPaso = True

If vTipoId = "U" Or vTipoId = "C" Then
    strSQL = "select cod_Educ, descripcion, Activa, 0 as btn from SYS_EDUCACION_CFG" _
          & " Where Tipo = '" & vTipoId & "' order by cod_Educ"
    Call sbCargaGrid(vGrid, 4, strSQL)
Else
    strSQL = "select cod_Educ, descripcion, Activa from SYS_EDUCACION_CFG" _
          & " Where Tipo = '" & vTipoId & "' order by cod_Educ"
    Call sbCargaGrid(vGrid, 3, strSQL)
End If

vPaso = False


scTitulo.Caption = ""
scTitulo.Tag = ""
btnAsigna(0).Enabled = False
btnAsigna(1).Enabled = False
btnAsigna(2).Enabled = False
lsw.ListItems.Clear

End Sub


