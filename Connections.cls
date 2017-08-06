VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'comando SQL na propria planilha' // ado novo (ACE)
Sub excelSQL()
	Dim cn As Object
	Dim rs As Object
	Dim strCon as String
	Dim sSQL as String
	'------------------------------------  
	strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
	Set cn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	'------------------------------------  
	cn.Open strCon
	sSQL = "select * from [Sheet1$A1:G3]"
	rs.Open sSQL, cn
	Debug.Print rs.GetString
End Sub

'utiliza comandos SQL na própria planilha do Excel // ado antigo
Public Sub activeWorkbookADO()
  Dim q As String
  Dim adoCN As Object
  Dim adoRS As Object
  Set adoCN = CreateObject("ADODB.Connection")
  Set adoRS = CreateObject("ADODB.Recordset")
  '------------------------------------  
  'Declarando localização do arquivo SQL (no caso a própria planilha)
  With adoCN
      .Provider = "Microsoft.Jet.OLEDB.4.0"
      .Properties("Extended Properties") = "Excel 8.0"
      .ConnectionString = "Data Source=" & ActiveWorkbook.FullName
      .Open
  End With
  '------------------------------------  
  'Comando SQL
       q = "select Nome, Sum(Idade) as IdadeCount"
  q = q & " from [Excel 8.0;HDR=YES;DATABASE=" & ActiveWorkbook.FullName & "].[" & ActiveSheet.Name & "$A1:B2]" & " as ExcelQuery"
  q = q & " group by Nome"
  q = q & " having ExcelQuery.Nome = 'Fernando'"
  adoRS.Open q, adoCN, 3 'adOpenStatic
	Debug.Print rs.GetString
End Sub

'Faz o insert no banco de dados de uma planilha de Excel com JOIN em uma tabela do Access
Private Sub insertNoBDComExcelJOIN()
	Dim cn As Object
	Dim q As String
	Dim dbEndereco As String
	Dim activeWorkbooklName As String
	Dim sheetName As String
	Dim sheetNameQuery As String
	'------------------------------------  
	dbEndereco = "'" & ActiveWorkbook.Path & "\db_mercadoAcoes.accdb '" '<---- endereco DB
	activeWorkbooklName = Application.ActiveWorkbook.FullName
	sheetName = "cotacao_data"
	sheetNameQuery = "[" & sheetName & "$]"
	Set cn = CreateObject("ADODB.Connection") 'conecta ao Access Database
	cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbEndereco
	'------------------------------------  
	'Append data from Sheet1 of workbook to Table1 of mydb.mdb:
	     q = "insert into tb_Cotacao"
	q = q & "    (cotacao_Papel, cotacao_Data, cotacao_UltimaCotacao, cotacao_Abertura, cotacao_Maxima, cotacao_Minima, cotacao_Volume,"
	q = q & "     cotacao_AnteriorUltimaCotacao, cotacao_AnteriorAbertura, cotacao_AnteriorMaxima, cotacao_AnteriorMinima, cotacao_AnteriorVolume)"
	q = q & " select"
	q = q & "    Papel, Data, [Ultima Cotação], Abertura, Máxima, Mínima, Volume,"
	q = q & "    [(Anterior) Ultima Cotação], [(Anterior) Abertura], [(Anterior) Maxima], [(Anterior) Mínima], [(Anterior) Volume]"
	q = q & " from"
	q = q & "    [Excel 8.0;HDR=YES;DATABASE=" & activeWorkbooklName & "]." & sheetNameQuery & " AS ExcelSQL"
	q = q & " left join"
	q = q & "    tb_Cotacao AS Cotacao"
	q = q & " on"
	q = q & "    ExcelSQL.Papel = Cotacao.cotacao_Papel and ExcelSQL.Data = Cotacao.cotacao_Data"
	q = q & " where"
	q = q & "    Cotacao.cotacao_Papel IS NULL AND Cotacao.cotacao_Data IS NULL"
	'------------------------------------  
	cn.Execute q
End Sub

'Faz uma query no banco de dados Access e cola os dados em um range específico
Public Sub AccessSQLToExcel()

    Dim adoCN As ADODB.Connection: Set adoCN = New ADODB.Connection
    Dim adoRS As ADODB.Recordset: Set adoRS = New ADODB.Recordset
    Dim sSQL As String
    
    'Local do banco de dados
    With adoCN
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & "C:\Users\Fernando\Documents\myDd.accdb"
        .Open
    End With

    'Comando SQL
    sSQL = "SELECT * FROM tbPessoas"
    adoRS.Open sSQL, adoCN, adOpenStatic, adLockReadOnly
    
    'Range que será copiado os dados
    Range("A1").CopyFromRecordset adoRS

End Sub

'Usa DAO para popular Array usando dados de tabela do Access
Public Sub AccessToArrayFromAccess()

    Dim daoDB As DAO.Database, daoRS As DAO.Recordset
    Dim iLinhas As Long
    Dim sSQL(1 To 2) As String
    Dim i As Long
    Dim ArrDados() As Variant
    
    Set daoDB = DBEngine.OpenDatabase(CurrentProject.FullName)

    sSQL(1) = "SELECT Count(AprCodERP) AS CountOfAprCodERP"
    sSQL(2) = "FROM QRY_APRESENTACAO_UNID"
    
    Set daoRS = daoDB.OpenRecordset(Join(sSQL, Space(1)), dbOpenDynaset)
    
    iLinhas = daoRS.Fields(0).Value 'Tamanho do recordset
        
    Set daoRS = daoDB.OpenRecordset("QRY_APRESENTACAO_UNID", dbOpenDynaset)
    
    ReDim ArrDados(1 To iLinhas, 1 To 7) 'set do tamanho Array
    
    'Passando os dados para o Array
    For i = 1 To iLinhas
        ArrDados(i, 1) = daoRS.Fields![AprCodERP].Value
        ArrDados(i, 2) = daoRS.Fields![BuBU].Value
        ArrDados(i, 3) = daoRS.Fields![FamFamilia].Value
        ArrDados(i, 4) = daoRS.Fields![PrdProduto].Value
        ArrDados(i, 5) = daoRS.Fields![AprApresentacao].Value
        ArrDados(i, 6) = daoRS.Fields![UnidAprDescr].Value
        ArrDados(i, 7) = daoRS.Fields![AprQtdUnid].Value
        
        daoRS.MoveNext
    Next i

End Sub

'Importa os dados da planilha do Excel por ADO
Public Sub ExcelSQLToAccess()

'Insere dados da planilha para a table Cliente_tmp

    Dim dbWb As String
    Dim dbWs As String
    Dim dsh As String
    
    dbWb = EnderecoWb 'Endereço do arquivo
    dbWs = Ws 'Nome da Worksheet
    dsh = "[" & dbWs & "$]"

    Dim SQL As String
    Dim cn As ADODB.Connection: Set cn = CurrentProject.AccessConnection
    Dim rs As ADODB.Recordset: Set rs = New ADODB.Recordset

    'Apaga os dados da temp (antes de inserir os dados do Excel)
    SQL = "DELETE * " _
        & "FROM Cliente_tmp"
    rs.Open SQL, cn, adOpenDynamic
    
    'String SQL na planilha que será importada. IMPORTANTE: Nomes das colunas são do select
    SQL = "INSERT INTO Cliente_tmp([Reg], [Distrito], [Setor], [Tipo Cliente], [Codigo JDE], [Razão Social], [Endereço], [Cidade], [UF], [CEP], [Região], [CNPJ], [Status]) " _
        & "SELECT [Reg], [Distrito], [Setor], [Tipo Cliente], [Codigo JDE], [Razão Social], [Endereço], [Cidade], [UF], [CEP], [Região], [CNPJ], [Status] " _
        & "FROM [Excel 8.0;HDR=YES;DATABASE=" & dbWb & "]." & dsh
        
    rs.Open SQL, cn, adOpenStatic
    
    cn.Close
    Set cn = Nothing
    Set rs = Nothing

End Sub

'Cria conexão ADO e compara dados com o banco de dados criando um dicionário
Public Function PreencheDicionarioPelaQuery() As String
       
'(set da conexão)

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL(1 To 5) As String
    Dim dic As Scripting.Dictionary
    Dim vAchado As String

    Set cn = CurrentProject.AccessConnection
    Set rs = New ADODB.Recordset

'(Compara Clientes_tmp com Clientes)

    sSQL(1) = "SELECT Clientes_tmp.CNPJ, [Codigo JDE] & [CNPJ] AS A, Count([Codigo JDE] & [CNPJ]) AS B"
    sSQL(2) = "FROM Clientes_tmp"
    sSQL(3) = "GROUP BY Clientes_tmp.CNPJ, [Codigo JDE] & [CNPJ]"
    sSQL(4) = "HAVING (((Count([Codigo JDE] & [CNPJ])) > 1))"
    sSQL(5) = "ORDER BY [Codigo JDE] & [CNPJ]"


    rs.Open Join(sSQL, Space(1)), cn, adOpenStatic, adLockReadOnly

    'Loop que monta a string de erros
    Set dic = New Scripting.Dictionary
    
    While Not rs.EOF
    
        vAchado = Format(Left(rs.Fields(0).Value, 14), "00"".""000"".""000""/""0000""-""00")
        
        If dic.Exists(vAchado) = False Then
            dic.Add vAchado, vAchado
        End If
        
        rs.MoveNext
    Wend
    
    ComparaJDE_CNPJ_Clientes = Join(dic.Items, ", ")

End Function

'Cola os dados do array no arquivo em Excel
Public Sub ArrayToExcel(ByRef DataRange As Variant, ByRef xls As Excel.Application)

    Dim iLinhas As Long
    
    iLinhas = UBound(DataRange) + 5
    
    'Vai preencher o arquivo xls
    With xls.Sheets(1)
        .Unprotect
        .Range("G1").Value = "Data: " & Format(Date, "dd/mm/yyyy") 'cola os dados de data no arquivo
        .Range("A6", "G" & iLinhas).Value = DataRange 'começa da linha 6
        .Protect _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            AllowFormattingCells:=True, _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowInsertingColumns:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True

    End With
    
End Sub

'Importa os dados do array para uma tabela do Access
Sub ArrayToAccess(DataRange As Variant)
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim sSQL As String

'(Apaga os dados anteriores de clientes_tmp)
    DoCmd.SetWarnings False
    
    sSQL = "DELETE * " _
        & "FROM Clientes_tmp"
    DoCmd.RunSQL sSQL
    
    DoCmd.SetWarnings True
'(Insere novos dados aqui se a validação de colunas estiver ok)

    Set db = DBEngine.OpenDatabase(CurrentProject.FullName)
    Set rs = db.OpenRecordset("SELECT * FROM Clientes_tmp", dbOpenDynaset)
    
    For i = 2 To UBound(DataRange)
        rs.AddNew
        rs.Fields![REG].Value = DataRange(i, 1)
        rs.Fields![DISTRITO].Value = DataRange(i, 2)
        rs.Fields![SETOR].Value = DataRange(i, 3)
        rs.Fields![TIPO CLIENTE].Value = DataRange(i, 4)
        rs.Fields![Codigo JDE].Value = StringToNumber(DataRange(i, 5))
        rs.Fields![RAZÃO SOCIAL].Value = DataRange(i, 6)
        rs.Fields![ENDEREÇO].Value = DataRange(i, 7)
        rs.Fields![CIDADE].Value = DataRange(i, 8)
        rs.Fields![UF].Value = DataRange(i, 9)
        rs.Fields![CEP].Value = DataRange(i, 10)
        rs.Fields![REGIÃO].Value = DataRange(i, 11)
        rs.Fields![CNPJ].Value = ConvertCNPJ(DataRange(i, 12))
        rs.Fields![Status].Value = DataRange(i, 13)
        rs.Update
    Next i
    
    Set rs = Nothing
 
End Sub

'Captura dados da instância do Excel
Public Sub ExcelToArray(ByVal Caminho As String)

'Retorna variável para valor padrão

    Dim xls As Excel.Application
    Dim Wb As String
       
    Set xls = New Excel.Application
    
    'Abre a instancia do Excel para pegar o arquivo de clientes
    With xls
        .Workbooks.Open Caminho, ReadOnly:=True
        .Visible = False
        .DisplayAlerts = False
        .EnableEvents = False
        Wb = .Workbooks(1).Name
        Ws = .Workbooks(1).Sheets(1).Name
    End With

    '(Capturando range da tabela e montanto o array na memória)

    Dim fimColunas As String: fimColunas = xls.Sheets(Ws).Range("A1").End(xlToRight).Address
    Dim fimLinhas As String: fimLinhas = xls.Sheets(Ws).Cells(linha, xls.Sheets(Ws).Range("A1").End(xlToRight).Column).Address
    Dim linha As Long: linha = xls.Sheets(Ws).Range("A1").End(xlDown).Row
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If xls.Sheets(Ws).ProtectContents _
        Or fimColunas = "$XFD$1" _
        Or fimColunas = "$IV$1" _
        Or fimLinhas = "$XFD$1048576" Then
                
        'Erros de validação
        For j = 1 To UBound(V)
            V(j) = "Não Conforme"
        Next j
        
        ArrayErrosClientes = "Erro"
                
        GoTo erroValida 'Pula toda a validação
    End If
    
    Dim s As String
    Dim DataRange As Variant
    
    s = "A1:" & fimLinhas
    xls.Sheets(Ws).Range(s).NumberFormat = "General" 'Corrigindo bug
    DataRange = xls.Sheets(Ws).Range(s).Value
    j = xls.Sheets(Ws).Range(s).Rows.Count

erroValida:
End Sub

'Cola os dados de array na sheet de Excel com apenas alguns argumentos
Private Sub ColarArraysNaSheet(ByVal DataRange As Variant, _
                               ByRef nomeDaSheet As String, _
                               ByRef linhaInicial As Integer, _
                               ByRef colunaInicial As Integer)

    Dim linhaFinal As Integer: linhaFinal = UBound(DataRange, 1) + linhaInicial - 1
    Dim colunaFinal As Integer: colunaFinal = UBound(DataRange, 2) + colunaInicial - 1
    Dim sheetInicial As String: sheetInicial = ActiveSheet.Name

    'Seleciona o endereco inicial para colar os dados
    Sheets(nomeDaSheet).Select
    Range(Cells(linhaInicial, colunaInicial), Cells(linhaFinal, colunaFinal)).Value = DataRange
    
    'Volta pra sheet anterior
    Sheets(sheetInicial).Select
    
End Sub

'Usa DAO para popular Array usando dados de tabela do Access
Public Sub AccessTableToArrayFromExcel()

    Dim daoDB As DAO.Database, daoRS As DAO.Recordset
    Dim linhasDAO As Long
    Dim q As String
    Dim i As Long
    Dim ArrDados() As Variant
    Dim dbEndereco As String: dbEndereco = ActiveWorkbook.Path & "\db_mercadoAcoes.accdb"
    
    Set daoDB = DBEngine.OpenDatabase(dbEndereco)

        q = "SELECT"
    q = q & "   *"
    q = q & " FROM"
    q = q & "   tb_Cotacao "
    
    Set daoRS = daoDB.OpenRecordset(q, dbOpenDynaset)
    
    'Tamanho do recordset
    Do While Not daoRS.EOF
        linhasDAO = linhasDAO + 1
        daoRS.MoveNext
    Loop
    
    'set do tamanho Array
    ReDim ArrDados(1 To linhasDAO, 1 To 7)
    daoRS.MoveFirst 'volta pra primeira linha
    
    'Passando os dados para o Array
    For i = 1 To linhasDAO
        ArrDados(i, 1) = daoRS.Fields![cotacao_Papel].Value
        ArrDados(i, 2) = daoRS.Fields![cotacao_Data].Value
        ArrDados(i, 3) = daoRS.Fields![cotacao_UltimaCotacao].Value
        ArrDados(i, 4) = daoRS.Fields![cotacao_Abertura].Value
        ArrDados(i, 5) = daoRS.Fields![cotacao_Maxima].Value
        ArrDados(i, 6) = daoRS.Fields![cotacao_Minima].Value
        ArrDados(i, 7) = daoRS.Fields![cotacao_Volume].Value
        
        daoRS.MoveNext
    Next i

End Sub

'Deleta dados e faz um insert no banco de dados Access
Private Sub DeleteEInsertNoBD()

    Dim cn As ADODB.Connection: Set cn = New ADODB.Connection
    Dim rs As ADODB.Recordset: Set rs = New ADODB.Recordset
    Dim q As String
    
    Dim dbEndereco As String: dbEndereco = "'" & ActiveWorkbook.Path & "\db_mercadoAcoes.accdb '"
    Dim activeWorkbooklName As String: activeWorkbooklName = Application.ActiveWorkbook.FullName
    Dim sheetName As String: sheetName = "cotacao_data"
    Dim sheetNameQuery As String: sheetNameQuery = "[" & sheetName & "$]"
    
    ' connect to the Access database
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbEndereco
    
    'Verifica se existem dados na tb Cotação, se sim, deleta os dados
             q = "SELECT"
        q = q & " *"
        q = q & " FROM"
        q = q & "    tb_cotacao AS Cotacao"
        q = q & " WHERE"
        q = q & "    Cotacao.cotacao_Data = #02/07/2014#"
        
    rs.Open q, cn, adOpenStatic, adLockReadOnly
    
    'Existem dados do dia procurado na tb Cotação
    If rs.RecordCount > 0 Then
        
            q = " DELETE"
        q = q & " * "
        q = q & " FROM "
        q = q & "   tb_Cotacao AS Cotacao"
        q = q & " WHERE "
        q = q & " Cotacao.cotacao_Data =#02/07/2014#"
        
    cn.Execute q

    End If

    ' Append data da cotação_data para o banco de dados
             q = "INSERT INTO tb_Cotacao"
        q = q & "    (cotacao_Papel, cotacao_Data, cotacao_UltimaCotacao, cotacao_Abertura, cotacao_Maxima, cotacao_Minima, cotacao_Volume,"
        q = q & "     cotacao_AnteriorUltimaCotacao, cotacao_AnteriorAbertura, cotacao_AnteriorMaxima, cotacao_AnteriorMinima, cotacao_AnteriorVolume)"
        q = q & " SELECT"
        q = q & "    Papel, Data, [Ultima Cotação], Abertura, Máxima, Mínima, Volume,"
        q = q & "    [(Anterior) Ultima Cotação], [(Anterior) Abertura], [(Anterior) Maxima], [(Anterior) Mínima], [(Anterior) Volume]"
        q = q & " FROM"
        q = q & "    [Excel 8.0;HDR=YES;DATABASE=" & activeWorkbooklName & "]." & sheetNameQuery & " AS ExcelSQL"
        q = q & " LEFT JOIN"
        q = q & "    tb_Cotacao AS Cotacao"
        q = q & " ON"
        q = q & "    ExcelSQL.Papel = Cotacao.cotacao_Papel AND ExcelSQL.Data = Cotacao.cotacao_Data"
        q = q & " WHERE"
        q = q & "    Cotacao.cotacao_Papel IS NULL AND Cotacao.cotacao_Data IS NULL"

    cn.Execute q
    
    'Fechando conn
    cn.Close
    
    Set cn = Nothing

End Sub

'Monta a string de conexão de dados da 'tbtbCotacaoAnterior'
Private Sub MontaConnectionString()
    
    Dim conexao As WorkbookConnection
    Dim q As String
    Dim diaAnterior As String: diaAnterior = GetDiaAnterior
    
        q = "SELECT"
    q = q & " Cotacao.cotacao_Papel,"
    q = q & " Cotacao.cotacao_Data,"
    q = q & " Cotacao.cotacao_UltimaCotacao,"
    q = q & " Cotacao.cotacao_Abertura,"
    q = q & " Cotacao.cotacao_Maxima,"
    q = q & " Cotacao.cotacao_Minima,"
    q = q & " Cotacao.cotacao_Volume"
    q = q & " FROM"
    q = q & " tb_Cotacao As Cotacao"
    q = q & " WHERE"
    q = q & " Cotacao.cotacao_Data = #" & diaAnterior & "#"
    
    'Itera pela coleção de conexões
    For Each conexao In ActiveWorkbook.Connections
        
        'Quando achar a conexão 'CotacaoAnterior' muda a string de conexão
        If conexao.Name = "CotacaoAnterior" Then
            conexao.OLEDBConnection.CommandText = q
            Exit For
        End If
    Next conexao
    
End Sub

'Abre o banco de dados Access e transforma consulta em array
Private Sub AccessToArrayFromExcel()
    
    Dim q As String
    Dim enderecoDoBanco As String: enderecoDoBanco = ActiveWorkbook.Path & "\" & "db_MercadoAcoes.accdb"

    'Usa DAO para popular Array usando dados de tabela do Access
    Dim daoDB As DAO.Database, daoRS As DAO.Recordset
    Dim i As Long
    Dim myDiaAnterior() As Variant
    
    'Abre o banco na memória
    Set daoDB = DBEngine.OpenDatabase(enderecoDoBanco)
    
    'Monta a string da query
    Dim calculaDiaAnterior As String: calculaDiaAnterior = GetDiaAnterior
    
        q = "SELECT"
    q = q & " Cotacao.cotacao_Papel,"
    q = q & " Cotacao.cotacao_Data,"
    q = q & " Cotacao.cotacao_UltimaCotacao,"
    q = q & " Cotacao.cotacao_Abertura,"
    q = q & " Cotacao.cotacao_Maxima,"
    q = q & " Cotacao.cotacao_Minima,"
    q = q & " Cotacao.cotacao_Volume"
    q = q & " FROM"
    q = q & " tb_Cotacao As Cotacao"
    q = q & " WHERE"
    q = q & " Cotacao.cotacao_Data = #" & calculaDiaAnterior & "#"
    
    'Abre a query na memória
    Set daoRS = daoDB.OpenRecordset(q, dbOpenDynaset)
    
    'Calcula o Tamanho do recordset
    Dim iLinhas As Long
    Do While Not daoRS.EOF
        iLinhas = iLinhas + 1
        daoRS.MoveNext
    Loop
        
    'Volta pra primeira linha do recordset
    daoRS.MoveFirst
        
    ReDim myDiaAnterior(1 To iLinhas, 1 To 6) 'set do tamanho Array
    
    'Passando os dados para o Array
    For i = 1 To iLinhas
        myDiaAnterior(i, 1) = daoRS.Fields![cotacao_Papel].Value
        myDiaAnterior(i, 2) = daoRS.Fields![cotacao_UltimaCotacao].Value
        myDiaAnterior(i, 3) = daoRS.Fields![cotacao_Abertura].Value
        myDiaAnterior(i, 4) = daoRS.Fields![cotacao_Maxima].Value
        myDiaAnterior(i, 5) = daoRS.Fields![cotacao_Minima].Value
        myDiaAnterior(i, 6) = daoRS.Fields![cotacao_Volume].Value
        daoRS.MoveNext
    Next i
    
    'Seta a propriedade com os dados do dia anterior
    diaAnterior = myDiaAnterior
    
End Sub
