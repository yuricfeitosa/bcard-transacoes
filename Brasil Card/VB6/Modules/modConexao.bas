Attribute VB_Name = "modConexao"
Public conn As ADODB.Connection

Public Sub AbrirConexao()
    Set conn = New ADODB.Connection
    
    conn.ConnectionString = "Provider=SQLOLEDB;" & _
                            "Data Source=PC-YURI\MSSQLSERVER001;" & _
                            "Initial Catalog=BCARD;" & _
                            "Integrated Security=SSPI;"
    
    conn.Open
End Sub
