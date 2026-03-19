VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransacaoDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Inserir(t As clsTransacao)
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "INSERT INTO BCA_TRANSACOES (TRA_NUM_CARTAO, TRA_VALOR, TRA_DATA, TRA_DESCRICAO) VALUES (?, ?, ?, ?)"
        .CommandType = adCmdText
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, t.NumCartao)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , t.valor)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , t.DataTransacao)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, t.descricao)
        
        .Execute
    End With
End Sub

Public Sub Atualizar(t As clsTransacao)
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "UPDATE BCA_TRANSACOES SET TRA_NUM_CARTAO = ?, TRA_VALOR = ?, TRA_DATA = ?, TRA_DESCRICAO = ? WHERE TRA_COD = ?"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, t.NumCartao)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, , t.valor)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , t.DataTransacao)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, t.descricao)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , t.Cod)
        
        .Execute
    End With
End Sub

Public Sub Excluir(Cod As Long)
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "DELETE FROM BCA_TRANSACOES WHERE TRA_COD = ?"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , Cod)
        
        .Execute
    End With
End Sub

Public Function Buscar(num As String, data As Variant, valor As Variant) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim sql As String

    sql = "SELECT * FROM vw_TransacoesClientes WHERE 1=1"

    If Trim(num) <> "" Then
        sql = sql & " AND TRA_NUM_CARTAO = '" & num & "'"
    End If

    If Not IsNull(data) Then
        sql = sql & " AND CONVERT(DATE, TRA_DATA) = '" & Format(data, "yyyy-mm-dd") & "'"
    End If

    If Not IsNull(valor) Then
        sql = sql & " AND TRA_VALOR = " & valor
    End If

    rs.CursorLocation = adUseClient
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    Set Buscar = rs
End Function

Public Function BuscarRelatorio(dtInicial As Date, dtFinal As Date) As ADODB.Recordset

    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset

    With cmd
        Set .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "sp_TotalTransacoesPorPeriodo"

        .Parameters.Append .CreateParameter("@DATA_INICIAL", adDate, adParamInput, , dtInicial)
        .Parameters.Append .CreateParameter("@DATA_FINAL", adDate, adParamInput, , dtFinal)
    End With

    rs.CursorLocation = adUseClient
    rs.Open cmd

    Set BuscarRelatorio = rs

End Function

Public Function BuscarUltimoMes() As ADODB.Recordset

    Dim rs As New ADODB.Recordset
    Dim sql As String

    sql = "SELECT " & _
          "TRA_NUM_CARTAO AS CARTAO, " & _
          "TRA_VALOR AS VALOR, " & _
          "TRA_DATA AS DATA, " & _
          "TRA_DESCRICAO AS DESCRICAO, " & _
          "CATEGORIA " & _
          "FROM vw_TransacoesClientes " & _
          "WHERE TRA_DATA >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0) AND TRA_DATA < DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)"

    rs.CursorLocation = adUseClient
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    Set BuscarUltimoMes = rs

End Function

