VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_CD_PeriodoActividades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descripción de Actividades"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   13185
   Icon            =   "FrmAF_CD_PeriodoActividades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SStabact 
      Height          =   5475
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   9657
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Asignacion de Cuentas a las Actividades"
      TabPicture(0)   =   "FrmAF_CD_PeriodoActividades.frx":3482
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGridact"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "OptAct"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "OptDes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdImprimir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAplicar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAno"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Asignación de Montos a las actividades"
      TabPicture(1)   =   "FrmAF_CD_PeriodoActividades.frx":349E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblActividad"
      Tab(1).Control(1)=   "vGridMontos"
      Tab(1).Control(2)=   "lswAct"
      Tab(1).Control(3)=   "Cmdguarda"
      Tab(1).Control(4)=   "cmdActualiza"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdAno 
         Caption         =   "Actualizar Año en Curso"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8580
         TabIndex        =   9
         Top             =   510
         Width           =   2115
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar Actividades"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73170
         TabIndex        =   8
         Top             =   540
         Width           =   1905
      End
      Begin VB.CommandButton Cmdguarda 
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
         Height          =   840
         Left            =   -63600
         Picture         =   "FrmAF_CD_PeriodoActividades.frx":34BA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Aplicar"
         Top             =   2400
         Width           =   930
      End
      Begin MSComctlLib.ListView lswAct 
         Height          =   4470
         Left            =   -74820
         TabIndex        =   5
         Top             =   810
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   7885
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Actividad"
            Object.Width           =   8290
         EndProperty
      End
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
         Height          =   825
         Left            =   12000
         Picture         =   "FrmAF_CD_PeriodoActividades.frx":35F2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aplicar"
         Top             =   1830
         Width           =   840
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Reporte"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   12000
         Picture         =   "FrmAF_CD_PeriodoActividades.frx":372A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2655
         Width           =   840
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Desembolsos Trimestrales"
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   405
         Value           =   -1  'True
         Width           =   2160
      End
      Begin VB.OptionButton OptAct 
         Caption         =   "Actividades Especiales"
         Height          =   270
         Left            =   2655
         TabIndex        =   1
         Top             =   465
         Width           =   2040
      End
      Begin FPSpreadADO.fpSpread vGridact 
         Height          =   4440
         Left            =   285
         TabIndex        =   10
         Top             =   840
         Width           =   11610
         _Version        =   524288
         _ExtentX        =   20479
         _ExtentY        =   7832
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   7
         ScrollBars      =   2
         SpreadDesigner  =   "FrmAF_CD_PeriodoActividades.frx":389D
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridMontos 
         Height          =   4185
         Left            =   -69285
         TabIndex        =   11
         Top             =   1095
         Width           =   5235
         _Version        =   524288
         _ExtentX        =   9234
         _ExtentY        =   7382
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         SpreadDesigner  =   "FrmAF_CD_PeriodoActividades.frx":3F34
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label LblActividad 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -69315
         TabIndex        =   6
         Top             =   795
         Width           =   5265
      End
   End
End
Attribute VB_Name = "frmAF_CD_PeriodoActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Nuevo As Boolean
Dim Filas As Integer, Columnas As Integer, Inc As Integer, Can As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem
Dim vCod As Integer
Dim vActivo As Boolean


 
Sub sbApliMonto()

'Dim InfoMonto As String
ReDim InfoMonto(vGridMontos.MaxCols)

Dim Consec As Integer
Dim A As Integer, B As Integer

For A = 1 To vGridMontos.MaxRows
    vGridMontos.Row = A
      
      For B = 1 To vGridMontos.MaxCols
          vGridMontos.Col = B
          InfoMonto(B) = vGridMontos.Text
      Next B
  
  
  strSQL = "select cod_actividad from afi_cd_actividades_rangos " _
           & "where cod_actividad ='" & vCod & "'"
           rs.Open strSQL, glogon.Conection, adOpenStatic
   
                 
         If rs.EOF Then
            If InfoMonto(2) <> "" And InfoMonto(3) <> "" Then
              'Consec = fxConsMonto
              strSQL = "insert into afi_cd_actividades_rangos (cod_actividad,monto,minimo,maximo,cod_monto) " _
                       & "values(" & vCod & "," & CCur(InfoMonto(1)) & "," & InfoMonto(2) & "," & InfoMonto(3) & "," & InfoMonto(4) & ")"
                       glogon.Conection.Execute strSQL
            Else
              MsgBox "Hay Campos vacios", vbExclamation, "Información"
              rs.Close
              Exit Sub
            End If
         
         Else
          
              
              strSQL = "update afi_cd_actividades_rangos " _
                       & "set monto = " & CCur(InfoMonto(1)) & ", " _
                       & "minimo = " & InfoMonto(2) & "," _
                       & "maximo = " & InfoMonto(3) & " " _
                       & "where cod_actividad = " & vCod & " and cod_monto = " & InfoMonto(4) & ""
                       glogon.Conection.Execute strSQL
                   
         End If
rs.Close
Next A
MsgBox "Se actualizaron los montos correctamente", vbInformation, "Información"
End Sub

Sub sbCallAct()

If vActivo = False Then
 If SStabact.Tab = 1 Then
 lswAct.SetFocus
 strSQL = "select cod_actividad,descripcion from afi_cd_actividades"
              rs.Open strSQL, glogon.Conection, adOpenStatic
        lswAct.ListItems.Clear
        While Not rs.EOF
          Set itmX = lswAct.ListItems.Add(, , rs!Cod_actividad)
           itmX.SubItems(1) = IIf(IsNull(rs!Descripcion), "Sin Nombre", rs!Descripcion)
          rs.MoveNext
        Wend
        rs.Close
  Call sbSelec
 End If
 vActivo = True
End If

End Sub

Sub sbCambiaAno()

Dim fecPeriodo As String
Dim fecLiq As String
Dim i As Integer, S As Integer

S = MsgBox("Desea cambiar el año en curso", vbYesNo + vbInformation, "Información")

If S = vbYes Then
 For i = 1 To vGridact.MaxRows
   vGridact.Row = i
   vGridact.Col = 5
   fecPeriodo = Format(vGridact.Text, "dd/mm/")
   vGridact.Text = fecPeriodo & Year(fxFechaServidor)
   vGridact.Col = 6
   fecLiq = Format(vGridact.Text, "dd/mm/")
   vGridact.Text = fecLiq & Year(fxFechaServidor)
   fecPeriodo = ""
   fecLiq = ""
 Next i
End If

End Sub

Sub sbCargaCuenta()

Dim i As Integer

For i = 1 To vGridact.MaxRows

    vGridact.Row = i
    vGridact.Col = 3
    If vGridact.Col = 3 And vGridact.Text = Empty Then
        frmCC_ConsultaCuentas.Show vbModal
        
        vGridact.Col = 3
        vGridact.Text = gCuenta
        
        vGridact.Col = 3
        vGridact.Text = fxgCntCuentaFormato(True, vGridact.Text, 0)
        
        vGridact.Col = 4
        vGridact.Text = fxgCntCuentaDesc(gCuenta)
     
     Exit Sub
    End If
 Next i

End Sub


Sub sbCargaGrid()

Dim vTipo As String

Select Case True
  Case OptDes.Value = True
    vTipo = "D"
  Case OptAct.Value = True
    vTipo = "A"
End Select


strSQL = " select * from afi_cd_actividades where tipo = '" & vTipo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
vGridact.MaxRows = rs.RecordCount

While Not rs.EOF
  For Filas = 1 To vGridact.MaxRows
     
     vGridact.Row = Filas
          vGridact.Col = 1
          vGridact.Text = rs!Cod_actividad
          vGridact.Col = 2
          vGridact.Text = rs!Descripcion
          vGridact.Col = 3
          vGridact.Text = fxgCntCuentaFormato(True, rs!cod_cuenta, 0)
          vGridact.Col = 4
          vGridact.Text = fxgCntCuentaDesc(rs!cod_cuenta)
          vGridact.Col = 5
          vGridact.Text = Format(rs!fechaperiocidad, "dd/mm/yyyy")
          vGridact.Col = 6
          vGridact.Text = Format(rs!fechaliq, "dd/mm/yyyy")
          vGridact.Col = 7
          vGridact.Text = rs!activa
     
     rs.MoveNext
  Next Filas
Wend
rs.Close

End Sub


Sub sbSelec()

Dim i As Integer


For i = 1 To lswAct.ListItems.Count
  If lswAct.ListItems.Item(i).Selected = True Then
     vCod = lswAct.ListItems.Item(i)
     LblActividad.Caption = Trim(lswAct.SelectedItem.SubItems(1))
  End If
Next i

strSQL = " select * from afi_cd_actividades_rangos where cod_actividad = " & vCod & ""
rs.Open strSQL, glogon.Conection, adOpenStatic

vGridMontos.MaxRows = rs.RecordCount

While Not rs.EOF
  
  For i = 1 To vGridMontos.MaxRows
     
          vGridMontos.Row = i
          vGridMontos.Col = 1
          vGridMontos.Text = Format(rs!Monto, "standard")
          vGridMontos.Col = 2
          vGridMontos.Text = rs!minimo
          vGridMontos.Col = 3
          vGridMontos.Text = rs!maximo
          vGridMontos.Col = 4
          vGridMontos.Text = rs!cod_monto
          
     rs.MoveNext
  Next i
Wend
rs.Close


End Sub

Private Sub cmdActualiza_Click()
 vActivo = False
 Call sbCallAct
End Sub


Private Sub CmdAno_Click()
 Call sbCambiaAno
End Sub

Private Sub cmdAplicar_Click()

Dim strSQL As String, rs As New ADODB.Recordset
Dim InsInfo() As String
ReDim InsInfo(vGridact.MaxCols)
Dim A As Integer, B As Integer
Dim vTipo As String
Dim vCuenta As Currency


Select Case True
 Case OptDes.Value = True
    vTipo = "D"
 Case OptAct.Value = True
    vTipo = "A"
End Select



For A = 1 To vGridact.MaxRows
    vGridact.Row = A
      
      For B = 1 To vGridact.MaxCols
          vGridact.Col = B
          InsInfo(B) = vGridact.Text
      Next B
  
  strSQL = "select cod_actividad from afi_cd_actividades " _
           & "where cod_actividad ='" & InsInfo(1) & "'"
           rs.Open strSQL, glogon.Conection, adOpenStatic
   
         vCuenta = fxgCntCuentaFormato(False, InsInfo(3))
         
         If rs.EOF Then
            If InsInfo(1) <> "" And InsInfo(2) <> "" Then
              
              strSQL = "insert into afi_cd_actividades (cod_actividad,descripcion,cod_cuenta,activa,fechaperiocidad,tipo,fechaliq) " _
                       & "values(" & InsInfo(1) & ",'" & UCase(Trim(InsInfo(2))) & "','" & vCuenta & "','" & (InsInfo(7)) & "','" & Format(InsInfo(5), "yyyymmdd") & "','" & vTipo & "','" & Format(InsInfo(6), "yyyymmdd") & "')"
                       glogon.Conection.Execute strSQL
            Else
              MsgBox "Hay Campos vacios", vbExclamation, "Información"
              Exit Sub
            End If
         
         Else
              
              strSQL = "update afi_cd_actividades " _
                       & "set descripcion = '" & UCase(Trim(InsInfo(2))) & "'," _
                       & " cod_cuenta = " & vCuenta & ",activa ='" & InsInfo(7) & "'" _
                       & ",fechaperiocidad = '" & Format(InsInfo(5), "yyyymmdd") & "',fechaliq = '" & Format(InsInfo(6), "yyyymmdd") & "'" _
                       & "where cod_actividad = '" & InsInfo(1) & "'"
                       glogon.Conection.Execute strSQL
         End If
rs.Close
Next A
 Call sbCargaGrid
 MsgBox "Información Aplicada", vbInformation, "Información"
End Sub

Private Sub Cmdguarda_Click()
Call sbApliMonto
End Sub

Private Sub Cmdimprimir_Click()
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
 .Connect = "pwd=" & glogon.RootKey
 .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_Cd_Actividades.rpt")
 .WindowTitle = "Reporte Actividades y sus montos"
 
' .SelectionFormula = "{afi_cd_nombramientos.id_pricomite} = '" & LswComi.SelectedItem & "' " _
' & "and {afi_cd_nombramientos_h.estado} = '1'"
  
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='ACTIVIDADES Y SUS MONTOS'"
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
 Call sbCargaGrid
 vActivo = False
 SStabact.Tab = 0
End Sub

Private Function fxConsMonto() As Long

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max(cod_monto),0) as Consecutivo from afi_cd_actividades_rangos "
rs.Open strSQL, glogon.Conection, adOpenStatic
fxConsMonto = rs!consecutivo + 1
rs.Close

End Function
Private Sub TabStrip1_Click()

End Sub

Private Sub LswAct_Click()
 Call sbSelec
End Sub

Private Sub LswAct_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbSelec
End Sub


Private Sub OptAct_Click()
 Call sbCargaGrid
End Sub

Private Sub OptDes_Click()
 Call sbCargaGrid
End Sub

Private Sub SSTab1_DblClick()






End Sub

Private Sub SStabact_Click(PreviousTab As Integer)
 Call sbCallAct
End Sub

Private Sub vGridactact_DblClick(ByVal Col As Long, ByVal Row As Long)
 Call sbCargaCuenta
End Sub


Private Sub vGridactact_KeyDown(KeyCode As Integer, Shift As Integer)

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim Conse As Integer, Inc As Integer

strSQL = "select coalesce(max(codtipo),0) + 1 as Ultimo from afi_cd_periocidadactividades"
          rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
 Conse = rs!ultimo
End If
rs.Close

If KeyCode = vbKeyInsert Then
    vGridact.MaxRows = vGridact.MaxRows + 1
    Inc = vGridact.MaxRows
    vGridact.InsertRows vGridact.ActiveRow + 1, 1
    vGridact.MaxRows = vGridact.MaxRows
    vGridact.SetActiveCell 0, vGridact.MaxRows
    vGridact.Row = vGridact.ActiveRow
    vGridact.Col = 1
    vGridact.Text = Conse
    Nuevo = True
    Call sbCargaCuenta
       
End If

If KeyCode = vbKeyDelete Then
  If MsgBox("¿Desea eliminar esta linea?", vbYesNo Or vbQuestion, "") = vbYes Then
    vGridact.Row = vGridact.ActiveRow
    vGridact.Col = 2
     If vGridact.MaxRows > 0 Then
       vGridact.MaxRows = vGridact.MaxRows - 1
       vGridact.DeleteRows vGridact.ActiveRow + 1, 1
       vGridact.Row = vGridact.ActiveRow
     End If
End If




End If

End Sub

Private Sub vGridactact_KeyPress(KeyAscii As Integer)
 Call sbCargaCuenta
End Sub


Private Sub vGridact_KeyDown(KeyCode As Integer, Shift As Integer)

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim Conse As Integer, Inc As Integer



strSQL = "select coalesce(max(cod_actividad),0) + 1 as Ultimo from afi_cd_actividades"
          rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
 Conse = rs!ultimo
End If
rs.Close

If KeyCode = vbKeyInsert Then
    vGridact.MaxRows = vGridact.MaxRows + 1
    Inc = vGridact.MaxRows
    vGridact.InsertRows vGridact.ActiveRow + 1, 1
    vGridact.MaxRows = vGridact.MaxRows
    vGridact.SetActiveCell 0, vGridact.MaxRows
    vGridact.Row = vGridact.ActiveRow
    vGridact.Col = 1
    vGridact.Text = Conse
    Nuevo = True
    Call sbCargaCuenta
       
End If


If KeyCode = vbKeyDelete Then
  If MsgBox("¿Desea eliminar esta linea?", vbYesNo Or vbQuestion, "") = vbYes Then
    vGridact.Row = vGridact.ActiveRow
    vGridact.Col = 2
    If vGridact.MaxRows > 0 Then
       vGridact.MaxRows = vGridact.MaxRows - 1
       vGridact.DeleteRows vGridact.ActiveRow + 1, 1
       vGridact.Row = vGridact.ActiveRow
    End If
End If




End If


End Sub


Private Sub vGridact_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 Call sbCargaCuenta
End Sub


Private Sub vGridMontos_DblClick(ByVal Col As Long, ByVal Row As Long)
 Call sbCargaCuenta
End Sub

Private Sub vGridMontos_KeyDown(KeyCode As Integer, Shift As Integer)
    

If KeyCode = vbKeyInsert Then
    
    vGridMontos.MaxRows = vGridMontos.MaxRows + 1
    Inc = vGridMontos.MaxRows
    vGridMontos.InsertRows vGridMontos.ActiveRow + 1, 1
    vGridMontos.MaxRows = vGridMontos.MaxRows
    vGridMontos.SetActiveCell 0, vGridMontos.MaxRows
    vGridMontos.Row = vGridMontos.ActiveRow
    vGridMontos.Col = 4
    vGridMontos.Text = Inc
   
   Nuevo = True
         
End If

If KeyCode = vbKeyDelete Then
  If MsgBox("¿Desea eliminar esta linea?", vbYesNo Or vbQuestion, "") = vbYes Then
    vGridMontos.Row = vGridact.ActiveRow
    vGridMontos.Col = 2
    If vGridMontos.MaxRows > 0 Then
       vGridMontos.MaxRows = vGridMontos.MaxRows - 1
       vGridMontos.DeleteRows vGridMontos.ActiveRow + 1, 1
       vGridMontos.Row = vGridMontos.ActiveRow
    End If
 End If
End If

End Sub


