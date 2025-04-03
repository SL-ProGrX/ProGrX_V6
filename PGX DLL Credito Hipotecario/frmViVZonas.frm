VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmVivZonas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de zonas"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9060
   Icon            =   "frmViVZonas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   8535
      _Version        =   1310723
      _ExtentX        =   15055
      _ExtentY        =   11668
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
      Item(0).Caption =   "Zonas"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Cobertura"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "cboProvincia"
      Item(1).Control(1)=   "Label2(0)"
      Item(1).Control(2)=   "cboZonas"
      Item(1).Control(3)=   "Label2(1)"
      Item(1).Control(4)=   "chkSoloAsignadas"
      Item(1).Control(5)=   "lsw"
      Item(1).Control(6)=   "btnCantones"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4575
         Left            =   -69760
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   8055
         _Version        =   1310723
         _ExtentX        =   14208
         _ExtentY        =   8070
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkSoloAsignadas 
         Height          =   375
         Left            =   -68680
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   3255
         _Version        =   1310723
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ver solo los cantones asignados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5895
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   8295
         _Version        =   524288
         _ExtentX        =   14631
         _ExtentY        =   10398
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
         MaxCols         =   483
         ScrollBars      =   2
         SpreadDesigner  =   "frmViVZonas.frx":1982
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   315
         Left            =   -68680
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   4455
         _Version        =   1310723
         _ExtentX        =   7858
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboZonas 
         Height          =   315
         Left            =   -68680
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   4455
         _Version        =   1310723
         _ExtentX        =   7858
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCantones 
         Height          =   615
         Left            =   -63640
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Asignar todos los Cantones!"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   315
         Index           =   1
         Left            =   -69640
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   855
         _Version        =   1310723
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Zona"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   315
         Index           =   0
         Left            =   -69640
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   855
         _Version        =   1310723
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Provincia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zonas y Coberturas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmVivZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub btnCantones_Click()

On Error GoTo vError

strSQL = "insert into ViviendaZonaAsigna(idZona, Provincia,Canton, Distrito, RegistroFecha, RegistroUsuario)" _
       & "( select " & cboZonas.ItemData(cboZonas.ListIndex) _
       & ", C.Provincia, C.Canton, '',getdate(), '" & glogon.Usuario & "'" _
       & " from Cantones C " _
       & " left join ViviendaZonaAsigna A on A.idZona = " & cboZonas.ItemData(cboZonas.ListIndex) _
       & "   and C.provincia = A.provincia and C.canton = A.canton" _
       & " where C.provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' and isnull(A.idZona,0) = 0 )"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "CrdHip Zonas / Coberturas Todos los Cantones, P." & cboProvincia.ItemData(cboProvincia.ListIndex) & ", Z." & cboZonas.ItemData(cboZonas.ListIndex))

MsgBox "Asignación de todos los cantones a esta zona realizado satisfactoriamente!", vbInformation

Call cboProvincia_Click

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboProvincia_Click()


If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If chkSoloAsignadas.Value = vbChecked Then
    strSQL = "select C.Canton, rtrim(C.Descripcion) as Descripcion, case when isnull(A.idZona,0) = 0 then 0 else 1 end as 'Check'" _
           & " from Cantones C " _
           & " inner join ViviendaZonaAsigna A on A.idZona = " & cboZonas.ItemData(cboZonas.ListIndex) _
           & "   and C.provincia = A.provincia and C.canton = A.canton" _
           & " where C.provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by C.Canton"

Else
    strSQL = "select C.Canton, rtrim(C.Descripcion) as Descripcion, case when isnull(A.idZona,0) = 0 then 0 else 1 end as 'Check'" _
           & " from Cantones C " _
           & " left join ViviendaZonaAsigna A on A.idZona = " & cboZonas.ItemData(cboZonas.ListIndex) _
           & "   and C.provincia = A.provincia and C.canton = A.canton" _
           & " where C.provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by C.Canton"
End If
Call OpenRecordSet(rs, strSQL)

vPaso = True
lsw.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Canton)
      itmX.SubItems(1) = rs!Descripcion
      itmX.Checked = rs!Check
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cboZonas_Click()
If vPaso Or cboZonas.ListCount <= 0 Then Exit Sub

Call cboProvincia_Click

End Sub

Private Sub chkSoloAsignadas_Click()
Call cboProvincia_Click
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

With lsw.ColumnHeaders
    .Clear
    .Add , , "Cantón", 2000
    .Add , , "Descripción", 5000

End With

strSQL = "select IdZona,descripcion,Activa from ViviendaZonas" _
      & " order by IdZona"
Call sbCargaGrid(vGrid, 3, strSQL)

vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False

tcMain.Item(0).Selected = True


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then  'Insertar
  
  strSQL = "select isnull(max(IdZona),0) + 1 as 'ZonaId' from ViviendaZonas"
  Call OpenRecordSet(rs, strSQL)
      vGrid.Text = CStr(rs!ZonaId)
  rs.Close
  
  strSQL = "insert into ViviendaZonas(IdZona,descripcion,Activa, RegistroFecha, RegistroUsuario) values(" _
         & vGrid.Text & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Credito Hipotecario Zona Id:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update ViviendaZonas set descripcion = '" & vGrid.Text & "', Activa = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where IdZona = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Credito Hipotecario Zona Id:  " & vGrid.Text)

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Or lsw.ListItems.Count <= 0 Then Exit Sub

If Item.Checked Then
    strSQL = "insert ViviendaZonaAsigna(idZona, Provincia, Canton, Distrito, RegistroFecha, RegistroUsuario)" _
          & " Values(" & cboZonas.ItemData(cboZonas.ListIndex) & ",'" _
          & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & Item.Text _
          & "','',getdate(),'" & glogon.Usuario & "')"
Else
    strSQL = "delete ViviendaZonaAsigna where idZona = " & cboZonas.ItemData(cboZonas.ListIndex) _
           & " and Provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
           & "' and Canton = '" & Item.Text & "' and Distrito = ''"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)


Select Case Item.Index
  Case 0 'Zonas
  
    strSQL = "select IdZona,descripcion,Activa from ViviendaZonas" _
           & " order by IdZona"
    Call sbCargaGrid(vGrid, 3, strSQL)

  Case 1 'Coberturas
     vPaso = True
        strSQL = "select IdZona as 'IdX', rtrim(descripcion) as 'ItmX' from ViviendaZonas" _
               & " order by IdZona"
        Call sbCbo_Llena_New(cboZonas, strSQL, False, True)
     vPaso = False
     
     Call cboProvincia_Click
     
End Select

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer


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

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete ViviendaZonas where IdZona = " & vGrid.Text
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Credito Hipotecario Zona Id:  " & vGrid.Text)
        
        strSQL = "select IdZona,descripcion,Activa from ViviendaZonas" _
              & " order by IdZona"
        Call sbCargaGrid(vGrid, 3, strSQL)
     
     End If
End If


End Sub

