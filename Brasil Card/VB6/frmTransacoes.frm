VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ¢ 
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnUltimoMes 
      Caption         =   "Relatório Último Mês"
      Height          =   495
      Left            =   11280
      TabIndex        =   18
      Top             =   7560
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13800
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnRelatorioP 
      Caption         =   "Relatório Período"
      Height          =   495
      Left            =   13200
      TabIndex        =   17
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   16
      Top             =   4560
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   4080
      ScaleHeight     =   135
      ScaleWidth      =   15
      TabIndex        =   15
      Top             =   4680
      Width           =   15
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtDescricao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   13
      Top             =   5520
      Width           =   4215
   End
   Begin VB.TextBox txtValor 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3600
      Width           =   4215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5775
      Left            =   7800
      TabIndex        =   6
      Top             =   1680
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10186
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   4215
   End
   Begin VB.ComboBox cmbCartao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   4215
   End
   Begin VB.CommandButton btnExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      Top             =   6960
      Width           =   1445
   End
   Begin VB.CommandButton btnAtualizar 
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Picture         =   "frmTransacoes.frx":0000
      TabIndex        =   2
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton btnInserir 
      Caption         =   "Inserir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   22
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   7080
      TabIndex        =   21
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   20
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   19
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label descricao 
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   12
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label data 
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label valor 
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label cartao 
      Caption         =   "Cartão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label codigo 
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cadastro de Transações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   6585
   End
End
Attribute VB_Name = "¢"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dao As clsTransacaoDAO

Private Sub Form_Load()
    Set dao = New clsTransacaoDAO
    
    AbrirConexao
    CarregarCartoes
    CarregarGrid
End Sub


Private Sub btnInserir_Click()

    If Not ValidarCampos Then Exit Sub

    dao.Inserir MontarTransacao()

    MsgBox "Inserido com sucesso!", vbInformation
    
    AtualizarTela

End Sub

Private Sub btnAtualizar_Click()

    If txtCod(0).Text = "" Then
        MsgBox "Selecione um registro!", vbExclamation
        Exit Sub
    End If

    If Not ValidarCampos Then Exit Sub

    dao.Atualizar MontarTransacao()

    MsgBox "Atualizado!", vbInformation
    
    AtualizarTela

End Sub

Private Sub btnExcluir_Click(Index As Integer)

    If txtCod(0).Text = "" Then
        MsgBox "Selecione um registro!", vbExclamation
        Exit Sub
    End If

    If MsgBox("Confirma exclusão?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    dao.Excluir CLng(txtCod(0).Text)

    MsgBox "Excluído com sucesso!", vbInformation

    AtualizarTela

End Sub

Private Sub btnBuscar_Click()

    Dim rs As ADODB.Recordset

    Set rs = dao.Buscar(GetCartao(), GetData(), GetValor())

    Set DataGrid1.DataSource = rs

End Sub

Private Sub btnRelatorioP_Click()

    Dim strDataInicial As String
    Dim strDataFinal As String
    Dim dtInicial As Date
    Dim dtFinal As Date

    strDataInicial = InputBox("Informe a Data Inicial (dd/mm/yyyy):", "Relatório")
    If strDataInicial = "" Then Exit Sub

    strDataFinal = InputBox("Informe a Data Final (dd/mm/yyyy):", "Relatório")
    If strDataFinal = "" Then Exit Sub

    If Not IsDate(strDataInicial) Or Not IsDate(strDataFinal) Then
        MsgBox "Datas inválidas!", vbExclamation
        Exit Sub
    End If

    dtInicial = CDate(strDataInicial)
    dtFinal = CDate(strDataFinal)

    If dtInicial > dtFinal Then
        MsgBox "Data inicial maior que final!", vbExclamation
        Exit Sub
    End If

    GerarExcel dtInicial, dtFinal

End Sub

Private Sub btnUltimoMes_Click()
    GerarExcelUltimoMes
End Sub


Private Function ValidarCampos() As Boolean

    If cmbCartao.Text = "" Then
        MsgBox "Selecione um cartão!", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(txtValor.Text) Then
        MsgBox "Valor inválido!", vbExclamation
        Exit Function
    End If

    If Not IsDate(txtData.Text) Then
        MsgBox "Data inválida!", vbExclamation
        Exit Function
    End If

    If txtDescricao(0).Text = "" Then
        MsgBox "Informe a descrição!", vbExclamation
        Exit Function
    End If

    ValidarCampos = True

End Function

Private Function MontarTransacao() As clsTransacao

    Dim t As New clsTransacao

    If txtCod(0).Text <> "" Then
        t.Cod = CLng(txtCod(0).Text)
    End If

    t.NumCartao = GetCartao()
    t.valor = CDbl(txtValor.Text)
    t.DataTransacao = CDate(txtData.Text)
    t.Descricao = txtDescricao(0).Text

    Set MontarTransacao = t

End Function

Private Function GetCartao() As String

    If Trim(cmbCartao.Text) = "" Then
        GetCartao = ""
    Else
        GetCartao = Split(cmbCartao.Text, " - ")(1)
    End If

End Function

Private Function GetData() As Variant

    If Trim(txtData.Text) = "" Then
        GetData = Null
    Else
        GetData = CDate(txtData.Text)
    End If

End Function

Private Function GetValor() As Variant

    If Trim(txtValor.Text) = "" Then
        GetValor = Null
    Else
        GetValor = CDbl(txtValor.Text)
    End If

End Function

Private Sub AtualizarTela()
    CarregarGrid
    LimparCampos
End Sub

Private Sub LimparCampos()
    txtCod(0).Text = ""
    cmbCartao.Text = ""
    txtValor.Text = ""
    txtDescricao(0).Text = ""
    txtData.Text = ""
End Sub

Private Sub CarregarGrid()

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open "SELECT * FROM vw_TransacoesClientes", conn, adOpenStatic, adLockOptimistic

    Set DataGrid1.DataSource = rs

End Sub

Private Sub CarregarCartoes()

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open "SELECT CLI_NOME, CLI_NUM_CARTAO FROM BCA_CLIENTES", conn, adOpenStatic, adLockReadOnly
    
    cmbCartao.Clear
    
    Do While Not rs.EOF
        cmbCartao.AddItem rs!CLI_NOME & " - " & rs!CLI_NUM_CARTAO
        rs.MoveNext
    Loop
    
    rs.Close

End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next
    
    If Not DataGrid1.DataSource Is Nothing Then
        txtCod(0).Text = DataGrid1.Columns("TRA_COD").Value
        txtValor.Text = DataGrid1.Columns("TRA_VALOR").Value
        txtDescricao(0).Text = DataGrid1.Columns("TRA_DESCRICAO").Value
        cmbCartao.Text = DataGrid1.Columns("CLI_NOME").Value & " - " & DataGrid1.Columns("TRA_NUM_CARTAO").Value
        txtData.Text = DataGrid1.Columns("TRA_DATA").Value
    End If

End Sub

Private Sub Form_DblClick()
    LimparCampos
End Sub


Private Sub GerarExcel(dtInicial As Date, dtFinal As Date)

    Dim rs As ADODB.Recordset
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim i As Integer
    Dim caminho As String

    Set rs = dao.BuscarRelatorio(dtInicial, dtFinal)

    If rs.EOF Then
        MsgBox "Nenhum registro encontrado!", vbInformation
        Exit Sub
    End If

    With CommonDialog1
        .DialogTitle = "Salvar Relatório"
        .Filter = "Arquivos Excel (*.xlsx)|*.xlsx"
        .DefaultExt = "xlsx"
        .FileName = "relatorio.xlsx"
        .ShowSave
        caminho = .FileName
    End With

    If Trim(caminho) = "" Then Exit Sub

    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlWS = xlWB.Sheets(1)

    For i = 0 To rs.Fields.Count - 1
        xlWS.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i

    xlWS.Range("A2").CopyFromRecordset rs

    xlWS.Rows(1).Font.Bold = True
    xlWS.Columns.AutoFit

    xlWB.SaveAs caminho

    xlWB.Close
    xlApp.Quit

    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing

    MsgBox "Relatório gerado com sucesso!", vbInformation

End Sub

Private Sub GerarExcelUltimoMes()

    Dim rs As ADODB.Recordset
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim i As Integer
    Dim caminho As String

    Set rs = dao.BuscarUltimoMes()

    If rs.EOF Then
        MsgBox "Nenhuma transação no último mês!", vbInformation
        Exit Sub
    End If

    On Error GoTo Cancelado

    With CommonDialog1
        .DialogTitle = "Salvar Relatório Último Mês"
        .Filter = "Arquivos Excel (*.xlsx)|*.xlsx"
        .DefaultExt = "xlsx"
        .FileName = "relatorio_ultimo_mes.xlsx"
        .CancelError = True
        .ShowSave
        caminho = .FileName
    End With

    If Trim(caminho) = "" Then Exit Sub

    On Error GoTo 0

    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlWS = xlWB.Sheets(1)

    For i = 0 To rs.Fields.Count - 1
        xlWS.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i

    xlWS.Range("A2").CopyFromRecordset rs

    xlWS.Rows(1).Font.Bold = True
    xlWS.Columns.AutoFit

    xlWB.SaveAs caminho

    xlWB.Close
    xlApp.Quit

    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing

    MsgBox "Relatório do último mês gerado!", vbInformation
    
    Exit Sub
    
Cancelado:
    Exit Sub

End Sub


