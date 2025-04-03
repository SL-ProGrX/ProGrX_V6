VERSION 5.00
Begin VB.Form frmPruebaDLL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prueba DLL's Sistema ASE"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmPruebaDLL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPruebas 
      Caption         =   "Prueba Procesos"
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   2280
      Width           =   1455
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Fondo Solidario"
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Conciliacion de Aportes"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   14
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtUsuario 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Text            =   "sa"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Text            =   "perseus"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtDB 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Text            =   "aseccss"
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Conciliacion de Cartera"
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Control de Documentos"
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Traspaso Asientos"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Control de Polizas de Vida"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Excedentes Mensuales"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Traspaso de OP a Tesoreria"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Proceso Mensual Créditos"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.OptionButton optDLL 
      Caption         =   "Proceso Mensual Ahorros"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Servidor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Base de datos"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmPruebaDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection

Private Sub cmdPrueba_Click()
Dim x As New clsProcesosASE_CC

On Error Resume Next
Con.Close

Con.ConnectionString = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" _
        & txtServer & ";UID=sa;PWD=;Database=" & txtDB & ";"
Con.Open


x.Autor = "Pedro baltodano navarro"
Set x.Con = Con
x.RutaReportes = "C:\proyectos\is-ase"
x.User = txtUsuario
Select Case True
 Case optDLL(0) 'ahorro
  x.vID_Proceso = 1
  x.AbreProcesos
 Case optDLL(1) 'credito
  x.vID_Proceso = 2
  x.AbreProcesos
 Case optDLL(2) 'Traspaso Tesoreria
  x.vID_Proceso = 4
  x.AbreProcesos
 Case optDLL(3) 'Excedentes
  x.vID_Proceso = 3
  x.AbreProcesos
 Case optDLL(4) 'Polizas
  x.vID_Proceso = 6
  x.AbreProcesos
 Case optDLL(5) 'Asientos
  x.vID_Proceso = 5
  x.AbreProcesos
 Case optDLL(6) 'Recibos
  x.vID_Proceso = 7
  x.AbreProcesos
 Case optDLL(7) 'Conciliacion de Cartera
  x.vID_Proceso = 8
  x.AbreProcesos
 Case optDLL(8) 'Conciliacion de Aportes
  x.vID_Proceso = 9
  x.AbreProcesos
 Case optDLL(9) 'Fondo Solidario
  x.vID_Proceso = 10
  x.AbreProcesos
  
  
End Select
End Sub


Private Sub cmdPruebas_Click()
Dim strSQL As String, curDiferencia As Currency
Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim vCon As New ADODB.Connection


vCon.ConnectionString = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" _
        & txtServer & ";UID=sa;PWD=;Database=" & txtDB & ";"
vCon.Open


'Ajustando Mora para Evitar Saldos del Mes Negativos
strSQL = "select R.id_solicitud,R.saldo,V.amortiza,C.retencion,C.poliza,V.cuota" _
       & " from reg_creditos R inner join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " inner join catalogo C on R.codigo = C.codigo" _
       & " where R.saldo < V.amortiza and C.retencion = 'N' and C.poliza = 'N'"
rs.Open strSQL, vCon, adOpenStatic
Do While Not rs.EOF
 If rs!Saldo = 0 Then
   'Borra la mora
   strSQL = "delete morosidad where estado = 'A' and id_solicitud = " & rs!Id_solicitud
   vCon.Execute strSQL
 Else
   curDiferencia = rs!amortiza - rs!Saldo
   strSQL = "select id_moro,amortiza from morosidad where estado = 'A'" _
          & "  and id_solicitud = " & rs!Id_solicitud & " order by amortiza desc"
   rsTmp.Open strSQL, vCon, adOpenStatic
   Do While Not rsTmp.EOF
     If rsTmp!amortiza >= curDiferencia Then
       strSQL = "update morosidad set amortiza = amortiza - " & curDiferencia _
              & " where id_moro = " & rsTmp!id_moro
       vCon.Execute strSQL
       curDiferencia = 0
      Else
       strSQL = "update morosidad set amortiza = 0 where id_moro = " & rsTmp!id_moro
       vCon.Execute strSQL
       curDiferencia = curDiferencia - rsTmp!amortiza
      End If
     rsTmp.MoveNext
   Loop
   rsTmp.Close
   
   If curDiferencia > 0 Then
     MsgBox "Iconsistencia en Morosidad vrs Saldos, Revisar Manualmente la Operacion " _
           & rs!Id_solicitud, vbExclamation
   End If
 
 End If
 rs.MoveNext
Loop
rs.Close

vCon.Close
End Sub

Private Sub optDLL_DblClick(Index As Integer)
Me.MousePointer = 11
Call cmdPrueba_Click
Me.MousePointer = 1
End Sub
