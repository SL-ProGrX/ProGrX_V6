VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCR_PolizasRegiones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación de Regiones para Pólizas"
   ClientHeight    =   6060
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   7455
      _ExtentX        =   13145
      _ExtentY        =   8911
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Regiones"
      TabPicture(0)   =   "frmCR_PolizasRegiones.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asignación Cantones"
      TabPicture(1)   =   "frmCR_PolizasRegiones.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "lblRegionMontos"
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(3)=   "vGridCantones"
      Tab(1).Control(4)=   "FlatScrollBar"
      Tab(1).Control(5)=   "txtRegion"
      Tab(1).Control(6)=   "optTodos"
      Tab(1).Control(7)=   "optSoloAsignados"
      Tab(1).Control(8)=   "OptNoAsignados"
      Tab(1).Control(9)=   "cboProvincia"
      Tab(1).ControlCount=   10
      Begin VB.ComboBox cboProvincia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         ItemData        =   "frmCR_PolizasRegiones.frx":0038
         Left            =   -70560
         List            =   "frmCR_PolizasRegiones.frx":0051
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton OptNoAsignados 
         Caption         =   "No Asignados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69240
         TabIndex        =   11
         Top             =   4560
         Width           =   1455
      End
      Begin VB.OptionButton optSoloAsignados 
         Caption         =   "Solo Asignados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70920
         TabIndex        =   10
         Top             =   4560
         Width           =   1455
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         TabIndex        =   9
         Top             =   4560
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtRegion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73800
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4092
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   6972
         _Version        =   524288
         _ExtentX        =   12298
         _ExtentY        =   7218
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_PolizasRegiones.frx":009A
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   315
         Left            =   -68400
         TabIndex        =   5
         Top             =   600
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   550
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin FPSpreadADO.fpSpread vGridCantones 
         Height          =   3012
         Left            =   -74760
         TabIndex        =   7
         Top             =   1440
         Width           =   6972
         _Version        =   524288
         _ExtentX        =   12298
         _ExtentY        =   5313
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_PolizasRegiones.frx":06DF
         VScrollSpecialType=   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblRegionMontos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72240
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Región:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Label lblPoliza 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   240
      Picture         =   "frmCR_PolizasRegiones.frx":0CB4
      Top             =   120
      Width           =   384
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Poliza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCR_PolizasRegiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vScroll As Boolean
Dim vCarga As Boolean, vPaso As Boolean

Private Sub sbLlenarComboProvincias()
    Dim strSQL As String
    
    strSQL = "select rtrim(PROVINCIA) + ' - ' + rtrim(descripcion) as Itmx" _
       & " from PROVINCIAS order by DESCRIPCION"
    vPaso = True
    Call sbLlenaCbo(cboProvincia, strSQL, True, False)
    vPaso = False
    
    Call cboProvincia_Click
    
End Sub

Private Sub cboProvincia_Click()
If vPaso Or cboProvincia.ListCount = 0 Then Exit Sub
    
Call sbCargaRegionesDetalle

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 COD_REGION from CRD_POLIZAS_REGION"
    
    If Len(txtRegion.Text) > 0 Then
    
        If FlatScrollBar.Value = 1 Then
           strSQL = strSQL & " where COD_POLIZA = '" & lblPoliza.Tag & "' " _
                & " and COD_REGION > '" & txtRegion.Text & "' order by COD_REGION asc"
        Else
           strSQL = strSQL & " where COD_POLIZA = '" & lblPoliza.Tag & "' " _
                & " and COD_REGION < '" & txtRegion.Text & "' order by COD_REGION desc"
        End If
        
    Else
        strSQL = strSQL & " where COD_POLIZA = '" & lblPoliza.Tag & "' order by COD_REGION asc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtRegion.Text = rs!COD_REGION
      txtRegion_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
    
    vGrid.AppearanceStyle = fxGridStyle

    lblPoliza.Tag = GLOBALES.gTag
    lblPoliza.Caption = GLOBALES.gTag2
    
    SSTab1.Tab = 0
    
    Call sbCargaRegiones
    

    
End Sub


Private Sub sbCargaRegiones()
    Dim strSQL As String

    strSQL = "select COD_REGION,MONTO_COMERCIAL,MONTO_PERSONAL,MODIFICA_FECHA,MODIFICA_USUARIO,REGISTRO_USUARIO,REGISTRO_FECHA from CRD_POLIZAS_REGION where COD_POLIZA = '" & lblPoliza.Tag & "'" _
        & " order by COD_REGION"
    
    Call sbCargaGridLocal(3, strSQL)
    
    
End Sub

Public Sub sbCargaGridLocal(vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer


vGrid.Sheet = 3
vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
  
    vGrid.Col = i
    Select Case i
       Case 1
            vGrid.Text = rs!COD_REGION
            
       Case 2
            vGrid.Value = CDbl(rs!MONTO_COMERCIAL)
       Case 3
            vGrid.Value = CDbl(rs!MONTO_PERSONAL)
            
            vGrid.TextTip = TextTipFixed
            vGrid.TextTipDelay = 1000
            vGrid.CellNote = "Registro : " & rs!registro_usuario & "[" & rs!registro_fecha & "]" & vbCrLf _
                           & "Actualizado: " & rs!MODIFICA_USUARIO & "[" & rs!MODIFICA_FECHA & "]"
    End Select
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim Codigo_Region As Integer

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

'strSQL = "select isnull(count(*),0) as Existe from CRD_POLIZAS_REGION" _
'       & " where COD_REGION = '" & vGrid.Text & "'"
'Call OpenRecordSet(rs, strSQL)

If Len(Trim(vGrid.Text)) = 0 Then  'Insertar

    strSQL = "select isnull(max(COD_REGION),0)+1 as Codigo from CRD_POLIZAS_REGION where COD_POLIZA = '" & Trim(lblPoliza.Tag) & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
        Codigo_Region = rs!Codigo
    End If

    
    strSQL = "insert into CRD_POLIZAS_REGION (COD_POLIZA,COD_REGION,MONTO_COMERCIAL,MONTO_PERSONAL,REGISTRO_USUARIO,REGISTRO_FECHA) values('" _
            & lblPoliza.Tag & "'," & Codigo_Region & ","
            
    vGrid.Col = 2
    
    If Len(vGrid.Text) = 0 Then
        MsgBox "Debe completar el monto comercial"
        Exit Function
    End If
    
    strSQL = strSQL & CDbl(vGrid.Text) & ","

    vGrid.Col = 3
    
    If Len(vGrid.Text) = 0 Then
        MsgBox "Debe completar el monto personal"
        Exit Function
    End If
    
    strSQL = strSQL & CDbl(vGrid.Text) & ",'" & glogon.Usuario & "',dbo.MyGetdate())"

    Call ConectionExecute(strSQL)

    vGrid.Col = 1
    Call Bitacora("Registra", "Región: " & Codigo_Region & ", Póliza: " & lblPoliza.Tag)
    
    rs.Close
Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CRD_POLIZAS_REGION set MONTO_COMERCIAL = " & CDbl(vGrid.Text) & ","
 
 vGrid.Col = 3
 strSQL = strSQL & "MONTO_PERSONAL = " & CDbl(vGrid.Text) & "," _
            & "MODIFICA_USUARIO = '" & glogon.Usuario & "'," _
            & "MODIFICA_FECHA = dbo.MyGetdate() " _
            & "WHERE COD_POLIZA = '" & lblPoliza.Tag & "' AND COD_REGION = "
 
 vGrid.Col = 1
 strSQL = strSQL & Trim(vGrid.Text)
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Región: " & Trim(vGrid.Text) & ", Póliza: " & lblPoliza.Tag)


End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub OptNoAsignados_Click()
    Call sbCargaRegionesDetalle
End Sub

Private Sub optSoloAsignados_Click()
    Call sbCargaRegionesDetalle
End Sub

Private Sub optTodos_Click()
    Call sbCargaRegionesDetalle
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
    Case 1
        vGrid.Col = 1
        vGrid.Row = 1
        txtRegion = vGrid.Text
        If Len(Trim(txtRegion)) = 0 Then
            lblRegionMontos = Empty
            Exit Sub
        End If
        vGrid.Col = 2
        lblRegionMontos = "MC: " & vGrid.Text
        vGrid.Col = 3
        lblRegionMontos = lblRegionMontos.Caption & " MP: " & vGrid.Text
        
        Call sbLlenarComboProvincias
        
'        Call sbCargaRegionesDetalle
        
    End Select

End Sub



Private Sub txtRegion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGridCantones.SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "COD_REGION"
        gBusquedas.Orden = "COD_REGION"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select COD_REGION,MONTO_COMERCIAL,MONTO_PERSONAL from CRD_POLIZAS_REGION"
        gBusquedas.Filtro = " AND COD_POLIZA = '" & lblPoliza.Tag & "'"
        frmBusquedas.Show vbModal
        txtRegion = gBusquedas.Resultado
        vGridCantones.SetFocus
    End If
End Sub

Private Sub sbCargarDatosRegion()

    If Len(txtRegion) > 0 Then
        
        Dim strSQL As String, rs As New ADODB.Recordset

       strSQL = "select MONTO_COMERCIAL,MONTO_PERSONAL from CRD_POLIZAS_REGION where COD_POLIZA = '" & lblPoliza.Tag & "'" _
                & " and COD_REGION = " & Trim(txtRegion)
                
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            
            lblRegionMontos = "MC: " & rs!MONTO_COMERCIAL
            vGrid.Col = 3
            lblRegionMontos = lblRegionMontos.Caption & " MP: " & rs!MONTO_PERSONAL
            
        Else
            lblRegionMontos.Caption = Empty
        End If
        
        Call sbCargaRegionesDetalle
    Else
        lblRegionMontos.Caption = Empty
        vGridCantones.MaxRows = 0
    End If


End Sub

Private Sub txtRegion_LostFocus()
    Call sbCargarDatosRegion
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, rs As New ADODB.Recordset, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
  Call sbCargaRegiones
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1

        If vGrid.Value = "" Then Exit Sub

        i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
        If i = vbYes Then
        
            strSQL = "select count(*) as Cantidad from CRD_POLIZAS_REGION_DETALLE where COD_POLIZA = '" & Trim(lblPoliza.Tag) & "' and COD_REGION = " & vGrid.Value
            Call OpenRecordSet(rs, strSQL)
            If Not rs.EOF And Not rs.BOF Then
                If rs!cantidad > 0 Then
                    MsgBox "No se puede borrar la región, tiene cantones asignados"
                    Exit Sub
                End If
            End If

       
            strSQL = "delete CRD_POLIZAS_REGION where COD_POLIZA = '" & lblPoliza.Tag & "' and COD_REGION = " & vGrid.Text
            Call ConectionExecute(strSQL)
            
            Call Bitacora("Elimina", "Pólizas región: " & vGrid.Text & " Póliza: " & lblPoliza)

        
            vGrid.DeleteRows vGrid.ActiveRow, 1
            vGrid.MaxRows = vGrid.MaxRows - 1
            If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
        End If
End If

End Sub

Private Sub sbCargaRegionesDetalle()
    Dim strSQL As String
    
    If Len(txtRegion) = 0 Then
        Exit Sub
    End If
    
    If OptNoAsignados.Value = vbUnchecked Then
    
        If optSoloAsignados.Value = vbUnchecked Then
        
            strSQL = "select C.CANTON,C.DESCRIPCION as NCANTON,P.PROVINCIA,P.DESCRIPCION as NPROVINCIA,RD.CANTON as ASIGNADO,RD.REGISTRO_FECHA,RD.REGISTRO_USUARIO " _
                & "from CANTONES C inner join PROVINCIAS P on C.PROVINCIA = P.PROVINCIA " _
                & "left join CRD_POLIZAS_REGION_DETALLE RD on C.CANTON = RD.CANTON and C.PROVINCIA = RD.PROVINCIA and RD.COD_POLIZA = '" & lblPoliza.Tag & "' " _
                & "and  RD.COD_REGION = " & txtRegion
                
            If cboProvincia.Text = "TODOS" Then
                strSQL = strSQL & " order by P.DESCRIPCION, C.DESCRIPCION"
            Else
                strSQL = strSQL & " where P.PROVINCIA = " & SIFGlobal.fxCodText(cboProvincia.Text) _
                    & " order by P.DESCRIPCION, C.DESCRIPCION "
            End If
                
        Else
        
            strSQL = "select C.CANTON,C.DESCRIPCION as NCANTON,P.PROVINCIA,P.DESCRIPCION as NPROVINCIA,RD.CANTON as ASIGNADO,RD.REGISTRO_FECHA,RD.REGISTRO_USUARIO " _
                & "from CANTONES C inner join PROVINCIAS P on C.PROVINCIA = P.PROVINCIA " _
                & "inner join CRD_POLIZAS_REGION_DETALLE RD on C.CANTON = RD.CANTON and C.PROVINCIA = RD.PROVINCIA and RD.COD_POLIZA = '" & lblPoliza.Tag & "' " _
                & "and  RD.COD_REGION = " & txtRegion
                
            If SIFGlobal.fxCodText(cboProvincia.Text) = "TODOS" Then
                strSQL = strSQL & " order by P.DESCRIPCION, C.DESCRIPCION"
            Else
                strSQL = strSQL & " where P.PROVINCIA = " & SIFGlobal.fxCodText(cboProvincia.Text) _
                    & " order by P.DESCRIPCION, C.DESCRIPCION "
            End If
                
                
        End If
        
    Else
            strSQL = "select C.CANTON,C.DESCRIPCION as NCANTON,P.PROVINCIA,P.DESCRIPCION as NPROVINCIA,RD.CANTON as ASIGNADO,RD.REGISTRO_FECHA,RD.REGISTRO_USUARIO " _
                & "from CANTONES C inner join PROVINCIAS P on C.PROVINCIA = P.PROVINCIA " _
                & "left join CRD_POLIZAS_REGION_DETALLE RD on C.CANTON = RD.CANTON and C.PROVINCIA = RD.PROVINCIA and RD.COD_POLIZA = '" & lblPoliza.Tag & "' " _
                & "where RD.CANTON is null"
                
            If cboProvincia.Text = "TODOS" Then
                strSQL = strSQL & " order by P.DESCRIPCION, C.DESCRIPCION"
            Else
                strSQL = strSQL & " and  P.PROVINCIA = " & SIFGlobal.fxCodText(cboProvincia.Text) _
                    & " order by P.DESCRIPCION, C.DESCRIPCION "
            End If
    
    End If
    
    Call sbCargaGridLocalCheck(3, strSQL)
    
    
End Sub

Public Sub sbCargaGridLocalCheck(vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

vCarga = True

vGridCantones.Sheet = 3
vGridCantones.MaxCols = vGridMaxCol
vGridCantones.MaxRows = 1
vGridCantones.Row = vGridCantones.MaxRows

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGridCantones.Row = vGridCantones.MaxRows
  For i = 1 To vGridCantones.MaxCols
  
    vGridCantones.Col = i
    Select Case i
       Case 1
       
            If Not IsNull(rs!Asignado) Then
               vGridCantones.Value = 1
            Else
               vGridCantones.Value = 0
            End If
            
       Case 2
            vGridCantones.Value = rs!NCANTON
            vGridCantones.CellTag = rs!canton
       Case 3
            vGridCantones.Value = rs!NPROVINCIA
            vGridCantones.CellTag = rs!provincia
            
            vGridCantones.TextTip = TextTipFixed
            vGridCantones.TextTipDelay = 1000
            vGridCantones.CellNote = "Registro : " & rs!registro_usuario & "[" & rs!registro_fecha & "]"
    End Select
    
  Next i
  vGridCantones.MaxRows = vGridCantones.MaxRows + 1
  rs.MoveNext
   
Loop
rs.Close
vGridCantones.MaxRows = vGridCantones.MaxRows - 1

vCarga = False

End Sub

Private Sub vGridCantones_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String
Dim mCanton As String, mProvincia As Integer
    
    If vCarga = False Then
        vGridCantones.Row = Row
        
        ' Carga el Canton
        vGridCantones.Col = 2
        mCanton = Trim(vGridCantones.CellTag)
        ' Carga la Provincia
        vGridCantones.Col = 3
        mProvincia = vGridCantones.CellTag
        
        If Len(txtRegion) = 0 Then
            Exit Sub
        End If
        
        vGridCantones.Col = 1
        If vGridCantones.Value = 1 Then

            ' Elimina el canton si ya está asignado en la poliza
            strSQL = "DELETE CRD_POLIZAS_REGION_DETALLE WHERE COD_POLIZA = '" & lblPoliza.Tag & "' AND " _
                & "CANTON = '" & mCanton & "' AND " _
                & " PROVINCIA = " & mProvincia
            Call ConectionExecute(strSQL)
            
            ' Inserta el Canton en la Poliza
            strSQL = "insert into CRD_POLIZAS_REGION_DETALLE (COD_POLIZA,COD_REGION,CANTON,PROVINCIA,REGISTRO_USUARIO,REGISTRO_FECHA) values('" _
                & lblPoliza.Tag & "'," & txtRegion & ",'" & mCanton & "'," & mProvincia & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
            Call ConectionExecute(strSQL)
            
        Else
            ' Elimina el canton en la poliza
            strSQL = "DELETE CRD_POLIZAS_REGION_DETALLE WHERE COD_POLIZA = '" & lblPoliza.Tag & "' AND " _
                & "COD_REGION = " & txtRegion & " AND " _
                & "CANTON = '" & mCanton & "' AND " _
                & " PROVINCIA = " & mProvincia
            Call ConectionExecute(strSQL)
        End If
    End If
End Sub
