VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSIFConciliacionSaldos 
   Caption         =   "SIF : Conciliación de Saldos"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSIFConciliacionSaldos.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   9015
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   7440
      TabIndex        =   6
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   53411843
      CurrentDate     =   38839
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   5415
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "Conciliación"
      TabPicture(0)   =   "frmSIFConciliacionSaldos.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tlbConciliar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vGrid"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Carga"
      TabPicture(1)   =   "frmSIFConciliacionSaldos.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DaoControl"
      Tab(1).Control(1)=   "txtArchivo"
      Tab(1).Control(2)=   "tlbBuscar"
      Tab(1).Control(3)=   "lbl"
      Tab(1).Control(4)=   "Image1"
      Tab(1).Control(5)=   "Label3(0)"
      Tab(1).Control(6)=   "Label1(2)"
      Tab(1).Control(7)=   "Label3(1)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Reportes"
      TabPicture(2)   =   "frmSIFConciliacionSaldos.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optRep(3)"
      Tab(2).Control(1)=   "optRep(2)"
      Tab(2).Control(2)=   "optRep(1)"
      Tab(2).Control(3)=   "optRep(0)"
      Tab(2).Control(4)=   "tlbReportes"
      Tab(2).Control(5)=   "Line2"
      Tab(2).ControlCount=   6
      Begin VB.OptionButton optRep 
         Appearance      =   0  'Flat
         Caption         =   "Cambios Aplicados x Acción"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   18
         Top             =   1800
         Width           =   2775
      End
      Begin VB.OptionButton optRep 
         Appearance      =   0  'Flat
         Caption         =   "General x Acción"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   17
         Top             =   1440
         Width           =   2775
      End
      Begin VB.OptionButton optRep 
         Appearance      =   0  'Flat
         Caption         =   "Cambios Aplicados"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   16
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton optRep 
         Appearance      =   0  'Flat
         Caption         =   "Reporte General"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -74400
         TabIndex        =   15
         Top             =   720
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.Data DaoControl 
         Caption         =   "DaoControl"
         Connect         =   "dBASE IV;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   -73560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   ""
         Top             =   2400
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.TextBox txtArchivo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   795
         Left            =   -73560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   4695
      End
      Begin FPSpread.vaSpread vGrid 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8295
         _Version        =   393216
         _ExtentX        =   14631
         _ExtentY        =   4048
         _StockProps     =   64
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   502
         SpreadDesigner  =   "frmSIFConciliacionSaldos.frx":68A6
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   780
         Left            =   -68640
         TabIndex        =   9
         Top             =   1560
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1376
         ButtonWidth     =   1217
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Archivo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cargar"
               Key             =   "Cargar"
               Object.ToolTipText     =   "Cargar Datos del Archivo"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbConciliar 
         Height          =   780
         Left            =   3360
         TabIndex        =   13
         Top             =   2880
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1376
         ButtonWidth     =   1429
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Cierre"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Actualizar"
               Key             =   "Actualizar"
               Object.ToolTipText     =   "Actualiza Estado de SIF"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Procesar"
               Key             =   "Procesar"
               Object.ToolTipText     =   "Procesa Cambios"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbReportes 
         Height          =   780
         Left            =   -70680
         TabIndex        =   14
         Top             =   2280
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1376
         ButtonWidth     =   1217
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Reporte"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74400
         X2              =   -70080
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lbl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73560
         TabIndex        =   12
         Top             =   3000
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   -74760
         Picture         =   "frmSIFConciliacionSaldos.frx":B2D2
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   $"frmSIFConciliacionSaldos.frx":20434
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   -74400
         TabIndex        =   10
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   -74640
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFConciliacionSaldos.frx":204C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFConciliacionSaldos.frx":35637
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFConciliacionSaldos.frx":4A7A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFConciliacionSaldos.frx":5100B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   6720
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Proceso de Conciliación de Saldos SIF vrs Instituciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7800
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmSIFConciliacionSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCodigo As String, mOperacion As Long

Private Sub sbConsultaOperacion(xCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select codigo,id_solicitud from reg_Creditos" _
       & " where cedula ='" & xCedula & "' and estado = 'A'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  mCodigo = ""
  mOperacion = 0
Else
  mCodigo = Trim(rs!codigo)
  mOperacion = rs!id_solicitud
End If
rs.Close

End Sub



Private Sub cboInstitucion_Click()
 vGrid.MaxRows = 0
End Sub

Private Sub dtpCorte_Change()
 vGrid.MaxRows = 0
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_institucion,descripcion,pr_fecha_corte from instituciones"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 cboInstitucion.AddItem rs!descripcion
 cboInstitucion.ItemData(cboInstitucion.NewIndex) = rs!cod_institucion
 rs.MoveNext
Loop
rs.Close

vGrid.MaxCols = 12
ssTab.Tab = 0

End Sub

Private Sub Form_Resize()
 Line1.X2 = Me.Width
 
 ssTab.Width = Me.Width - 320
 ssTab.Height = Me.Height - 2000
 
 vGrid.Width = ssTab.Width - 220
 vGrid.Height = ssTab.Height - 1480
 
 tlbConciliar.Top = vGrid.Top + (vGrid.Height + 100)
 tlbConciliar.Left = ssTab.Left + (ssTab.Width / 2) - 900
 
End Sub


Private Sub sbActualizaCorte()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from SIFConciliacionSaldos" _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and estado = 'P'"

rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
  strSQL = "select S.cedula,S.nombre,S.cod_institucion" _
           & ",sum(R.saldo) as Saldo, Sum(R.cuota) as Cuota, count(*) as Operaciones" _
           & ",min(R.priDeduc) as Prideduc" _
           & " from socios S inner join Reg_creditos R on S.cedula = R.cedula" _
           & " where R.estado = 'A' and S.cedula = '" & rs!cedula & "'" _
           & " group by S.cedula,S.nombre,S.cod_institucion"
  rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  If rsTmp.EOF And rsTmp.BOF Then
     strSQL = "update SIFConciliacionSaldos set cuota_sif = 0, saldo_sif = 0,NOP_SIF = 0" _
            & ",INST_SIF = 0,PRIDEDUC = 0,RECOMENDACION = '02',ACCION = '02'" _
            & " where cod_institucion = " & rs!cod_institucion _
            & " and corte = '" & Format(rs!corte, "yyyy/mm/dd") & "' and cedula = '" & rs!cedula & "'"
     glogon.Conection.Execute strSQL
  
  Else
  
     strSQL = "update SIFConciliacionSaldos set cuota_sif = " & rsTmp!cuota & ",saldo_sif = " _
            & rsTmp!Saldo & ",NOP_SIF = " & rsTmp!Operaciones & ",INST_SIF = " & rsTmp!cod_institucion _
            & ",PRIDEDUC = " & rsTmp!PriDeduc
            
     If Abs(rs!Saldo_inst - rsTmp!Saldo) > 5 Then
     
        If rs!Saldo_inst > rsTmp!Saldo Then
            strSQL = strSQL & ",RECOMENDACION = '02',ACCION = '02'"
        Else
            strSQL = strSQL & ",RECOMENDACION = '03',ACCION = '03'"
        End If
              
     Else
        strSQL = strSQL & ",RECOMENDACION = '00',ACCION = '00'"
     End If
     
     strSQL = strSQL & " where cod_institucion = " & rs!cod_institucion _
            & " and corte = '" & Format(rs!corte, "yyyy/mm/dd") & "' and cedula = '" & rs!cedula & "'"
     glogon.Conection.Execute strSQL
  
  End If
  rsTmp.Close

 rs.MoveNext
Loop
rs.Close

'Indicar todos los casos que estan pendientes en la institucion


strSQL = "insert into dbo.SIFConciliacionSaldos(cod_institucion,corte,cedula,nombre" _
       & ",saldo_inst ,cuota_inst ,saldo_sif ,cuota_sif ,nop_sif ,inst_sif ,prideduc" _
       & ",recomendacion,accion,estado,fecha,usuario) " _
       & " select " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',S.cedula,S.nombre,0,0" _
       & ",sum(R.saldo), Sum(R.cuota), count(*) ,S.cod_institucion" _
       & ",min(R.priDeduc),'03','03','P',getdate(),'" & glogon.Usuario & "'" _
       & "  from socios S inner join Reg_creditos R on S.cedula = R.cedula" _
       & " where R.estado = 'A' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & "   and S.cedula not in(select cedula from SIFConciliacionSaldos" _
       & "   where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & "   and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "')" _
       & " group by S.cedula,S.nombre,S.cod_institucion "
glogon.Conection.Execute strSQL


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub sbCargaArchivo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

lng = 1

strSQL = "delete SIFConciliacionSaldos where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
glogon.Conection.Execute strSQL

With DaoControl.Recordset

Do While Not .EOF

  lbl.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  lbl.Refresh
  
  strSQL = "insert SIFConciliacionSaldos(cod_institucion,corte,cedula,nombre,saldo_Inst,cuota_Inst" _
           & ",Estado,Fecha,Usuario)" _
           & " values(" & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") _
           & "','" & Trim(!cedula) & "','" & Trim(!nombre) & "'," & CCur(!Saldo) & "," & CCur(!cuota) & ",'P',getdate(),'" _
           & glogon.Usuario & "')"
  glogon.Conection.Execute strSQL
  lng = lng + 1
  .MoveNext
Loop

End With

Me.MousePointer = vbDefault


lbl.Caption = "Comparando Datos (Espere)..."

Call sbActualizaCorte

lbl.Caption = ""

MsgBox "Proceso finalizado Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
If ssTab.Tab = 0 Then
 tlbConciliar.Visible = True
Else
 tlbConciliar.Visible = False
End If
End Sub

Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim itmX As ListItem, vArchivo As String
Dim vPasa As Boolean


Select Case Button.Key
  Case "Buscar"

        With frmContenedor.dlg
         .InitDir = "C:\"
         .ShowOpen
         
         If .FileName = "" Then
           MsgBox "Archivo no válido...", vbExclamation
           Exit Sub
         End If
         
         If UCase(Right(.FileName, 3)) <> "DBF" Then
           MsgBox "La Extensión del Archivo no es válido...", vbExclamation
           Exit Sub
         End If
        
         txtArchivo = .FileName
        
        End With
        
        DaoControl.RecordSource = Dir(txtArchivo, vbArchive)
        DaoControl.DatabaseName = Mid(txtArchivo, 1, Len(txtArchivo) - (Len(DaoControl.RecordSource) + 1))
        DaoControl.Refresh
  
  Case "Cargar"
       Call sbCargaArchivo
 End Select


End Sub


Private Sub sbConsultaSaldos()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select cedula,nombre,saldo_inst,cuota_inst,saldo_sif,cuota_sif,nop_sif" _
       & ",inst_sif,prideduc,case recomendacion" _
       & "     when '00' then 'Par'" _
       & "     when '01' then 'Ignorar'" _
       & "     when '02' then 'Ajustar SIF'" _
       & "     when '03' then 'Ajustar Inst'" _
       & "     when '04' then 'Revisar' end," _
       & " Case Accion " _
       & "     when '00' then 'Par'" _
       & "     when '01' then 'Ignorar'" _
       & "     when '02' then 'Ajustar SIF'" _
       & "     when '03' then 'Ajustar Inst'" _
       & "     when '04' then 'Revisar' end," _
       & " Case Estado" _
       & "      when  'P' then 0" _
       & "      when  'G' then 2 end" _
       & " From SIFConciliacionSaldos" _
       & " where cod_institucion =  " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & "  and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"

Call sbCargaGrid(vGrid, 12, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical


End Sub

Private Sub sbProcesaSaldos()
Dim i As Integer, strSQL As String, rsTmp As New ADODB.Recordset
Dim rs As New ADODB.Recordset, vCodigo As String
Dim vFechaSistema As String, vCedula As String, vAccion As String


On Error GoTo vError

Me.MousePointer = vbHourglass

vFechaSistema = fxFechaServidor

'Saca un codigo por Institucion para el credito
strSQL = "Select codigo from catalogo where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
rs.Open strSQL, glogon.Conection, adOpenStatic
 vCodigo = rs!codigo
rs.Close


For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.Col = 1
  vCedula = vGrid.Text
  
  vGrid.Col = 12
  
  If vGrid.Value = 1 Then
     vGrid.Col = 11
     
     Select Case vGrid.Text
       Case "Par"
          vAccion = "00"
       Case "Ignorar"
          vAccion = "01"
       Case "Ajustar SIF"
          vAccion = "02"
       Case "Ajustar Inst"
          vAccion = "03"
       Case "Revisar"
          vAccion = "04"
     End Select
     
     
    strSQL = "update SIFConciliacionSaldos set Accion = '" & vAccion & "',Estado = 'G'" _
           & " where cod_institucion =  " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
           & "  and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and cedula = '" & vCedula & "'"
    glogon.Conection.Execute strSQL
  End If 'Aplica o no
  
  vGrid.Col = 12
  If vAccion = "02" And vGrid.Value = 1 Then
     With rsTmp
        strSQL = "select * from SIFConciliacionSaldos" _
                & " where cod_institucion =  " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                & "  and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and cedula = '" & vCedula & "'"
        .Open strSQL, glogon.Conection, adOpenStatic
     
        strSQL = "select coalesce(count(*),0) as existe from socios where cedula = '" & !cedula & "'"
        rs.Open strSQL, glogon.Conection, adOpenStatic
           If rs!existe = 0 Then
                     strSQL = "Insert Socios(Cedula,Nombre,EstadoActual,FechaIngreso,Fecha_Nac," _
                            & "Sexo,EstadoCivil,Provincia,Canton,Distrito,Direccion,Af_Email,Apto," _
                            & "Cod_sector,cod_profesion,id_promotor,EstadoLaboral,Ultimo_Estado,Cod_Institucion" _
                            & ",Cod_departamento,cod_seccion,boleta,cedulaR,af_npagos,hijos,reg_user,reg_fecha) Values('" _
                            & Trim(!cedula) & "','" & Trim(!nombre) & "'," & "'S','" & Format(vFechaSistema, "yyyy/mm/dd") & "','" _
                            & Format(vFechaSistema, "yyyy/mm/dd") & "','M','S',1,'1','','','','',1,1,1,1,'N'," _
                            & !cod_institucion & ",'','',0,'',2,0,'" & glogon.Usuario & "',getdate())"
                     glogon.Conection.Execute strSQL
                    
                     strSQL = "Insert into Ahorro_Consolidado(Cedula,Aporte,Ahorro,Extra,Capitaliza," _
                            & "FecAporte,FecAhorro,FecExtra,FecCapitaliza,AportAnt,AhorroAnt) Values(" _
                            & "'" & Trim(!cedula) & "',0,0,0,0,getdate(),getdate(),getdate(),getdate(),0,0)"
                     glogon.Conection.Execute strSQL
            End If
            rs.Close
      
         'Ajustes en Credito
         Select Case !nop_sif
           Case 0
            '1. Insertar la Formalización
            strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
                   & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
                   & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
                   & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
                   & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol,cod_destino,cod_grupo) values('" _
                   & vCodigo & "',1,'" & Trim(!cedula) & "'," & !Saldo_inst & "," & !Saldo_inst _
                   & ",0," & !Saldo_inst & ",0,0," & !Saldo_inst & "," & !Cuota_inst & ",0,0," & (!Saldo_inst / !Cuota_inst) _
                   & ",'" & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "'," & "'" & glogon.Usuario & "','" _
                   & Format(vFechaSistema, "yyyy/mm/dd") & "','" & Format(vFechaSistema, "yyyy/mm/dd") & "','" _
                   & Format(vFechaSistema, "yyyy/mm/dd") & "','" & Format(vFechaSistema, "yyyy/mm/dd") & "','" _
                   & Format(vFechaSistema, "yyyy/mm/dd") & "','" & Format(vFechaSistema, "yyyy/mm/dd") & "','N'" _
                   & ",'N','OT','',0,1,0,'Ajuste de Conciliacion de Saldos','A'," & Year(!corte) & Format(Month(!corte), "00") _
                   & "," & Year(!corte) & Format(Month(!corte), "00") & ",'F','','')"
            glogon.Conection.Execute strSQL
            
            Case 1 'Una Operacion
            
                 Call sbConsultaOperacion(vCedula)
            
                 If !Saldo_inst < !Saldo_Sif Then
                    strSQL = "update reg_creditos set saldo = saldo - " & (!Saldo_Sif - !Saldo_inst) _
                           & ",amortiza = amortiza + " & (!Saldo_Sif - !Saldo_inst) _
                           & " where cedula = '" & !cedula & "' and estado = 'A'"
                    glogon.Conection.Execute strSQL
                    
                    'Registrar Nota de Credito de Ajuste
                    strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
                           & "fechas,fechap,tcon,ncon,estado) values('" & mCodigo _
                           & "'," & mOperacion & ",0,0,0," & (!Saldo_Sif - !Saldo_inst) & ",'" _
                           & Format(vFechaSistema, "yyyy/mm/dd") & "'," & Year(vFechaSistema) & Format(Month(vFechaSistema), "00") _
                           & ",7,0,'A')"
                    If mOperacion > 0 Then glogon.Conection.Execute strSQL
                 
                 Else
                    strSQL = "update reg_creditos set saldo = saldo + " & (!Saldo_inst - !Saldo_Sif) _
                           & ",amortiza = amortiza - " & (!Saldo_inst - !Saldo_Sif) _
                           & " where cedula = '" & !cedula & "' and estado = 'A'"
                    glogon.Conection.Execute strSQL
                    
                    'Registrar Nota de Credito de Debito de Ajuste
                    strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
                           & "fechas,fechap,tcon,ncon,estado) values('" & mCodigo _
                           & "'," & mOperacion & ",0,0,0," & (!Saldo_inst - !Saldo_Sif) & ",'" _
                           & Format(vFechaSistema, "yyyy/mm/dd") & "'," & Year(vFechaSistema) & Format(Month(vFechaSistema), "00") _
                           & ",7,0,'A')"
                    If mOperacion > 0 Then glogon.Conection.Execute strSQL
                 
                 End If
                
                 If !Saldo_inst = 0 Then
                  strSQL = "update reg_creditos set estado = 'C'" _
                         & " where cedula = '" & !cedula & "' and estado = 'A'"
                  glogon.Conection.Execute strSQL
                 End If
             
            Case Else
              'Ojo tiene mas de dos operaciones, Poner como Revisar
                strSQL = "update SIFConciliacionSaldos set Accion = '04',Estado = 'G'" _
                       & " where cod_institucion =  " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                       & "  and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and cedula = '" & vCedula & "'"
                glogon.Conection.Execute strSQL
         
         End Select
      
      
      
      
        .Close
     End With
  End If 'Ajuste en SIF
Next i

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub


Private Sub tlbConciliar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbConsultaSaldos
  
  Case "Actualizar"
    Call sbActualizaCorte
    Call sbConsultaSaldos
    
  Case "Procesar"
    Call sbProcesaSaldos
    Call sbConsultaSaldos
End Select

End Sub

Private Sub tlbReportes_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vTitulo As String, vSubTitulo As String
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vTitulo = ""
vSubTitulo = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de SIF Auxiliar de Planillas"
 
 .Connect = glogon.ConectRPT
  
    strSQL = "{SIFConciliacionSaldos.cod_institucion} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    strSQL = strSQL & " and date({SIFConciliacionSaldos.Corte}) = date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
     
    vSubTitulo = cboInstitucion.Text & " [Corte : " & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
    
    
    Select Case True
      Case optRep.Item(0).Value 'General
         vTitulo = "Listado General"
         .ReportFileName = App.Path & "\SIFConciliacion.rpt"
      Case optRep.Item(1).Value 'Aplicado
         vTitulo = "Ajustes Aplicados"
         .ReportFileName = App.Path & "\SIFConciliacion.rpt"
         strSQL = strSQL & " and {SIFConciliacionSaldos.Estado} = 'G'"
      Case optRep.Item(2).Value 'General x Accion
         vTitulo = "Listado General x Acción"
         .ReportFileName = App.Path & "\SIFConciliacionAccion.rpt"
      Case optRep.Item(3).Value 'Aplicado x Accion
         vTitulo = "Ajustes Aplicados x Acción"
         .ReportFileName = App.Path & "\SIFConciliacionAccion.rpt"
         strSQL = strSQL & " and {SIFConciliacionSaldos.Estado} = 'G'"
    End Select
  
  
 
 .Formulas(0) = "Titulo='" & vTitulo & "'"
 .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = strSQL
 
 .PrintReport
 
 
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)
Dim strSQL As String, rs As New ADODB.Recordset

If Col <> 12 Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1

strSQL = "select estado From SIFConciliacionSaldos" _
       & " where cod_institucion =  " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & "  and corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and cedula = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs!estado = "G" Then
  vGrid.Col = 12
  vGrid.Value = 2
  MsgBox "Este registro ya fue ajustado...verifique.!!!", vbExclamation
End If
rs.Close

End Sub
