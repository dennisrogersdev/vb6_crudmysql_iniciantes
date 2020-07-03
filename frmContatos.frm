VERSION 5.00
Begin VB.Form frmContatos 
   BackColor       =   &H00FCFFFB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meus Contatos"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2850
      TabIndex        =   7
      Top             =   2400
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1725
      TabIndex        =   6
      Top             =   2400
      Width           =   990
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   990
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2265
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   8640
      Begin VB.ListBox contatos 
         BackColor       =   &H00C0FFFF&
         Height          =   1740
         Left            =   4050
         TabIndex        =   8
         Top             =   225
         Width           =   4440
      End
      Begin VB.TextBox txtFone 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   525
         TabIndex        =   1
         Top             =   1500
         Width           =   3165
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   525
         TabIndex        =   0
         Top             =   750
         Width           =   3165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         Height          =   210
         Left            =   525
         TabIndex        =   4
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome*"
         Height          =   210
         Left            =   525
         TabIndex        =   3
         Top             =   450
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************
'**
'**     MEUS CONTATOS (CRUD MYSQL UTILIZANDO ADODB)
'**
'**************************************************************

'**************************************************************
'**
'**     Sistema criado para quem está iniciando em VB6
'**     tem pouco conhecimento de programação e não sabe
'**     por onde começar
'**
'**************************************************************


Option Explicit 'impede que você use uma variável sem criá-la

'Configurações para Conexão com o Banco de dados Mysql
'Configure de acordo com o seu ambiente
'Constante com o Host/Server do nosso banco de dados Mysql
Const m_DB_Server = "localhost"
'Constante com a Porta de Conexão ao Banco de dados
Const m_DB_Port = 3306
'Constante com o Usuário do Banco de dados
Const m_DB_User = "root"
'Constante com a senha do banco de dados
Const m_DB_Pass = ""
'Constante com o nome do banco de dados
Const m_DB_Database = "vb6_meuscontatos"

'Declaração da variável de conexão
'Utilizar Referência "Microsoft ActiveX Data Objects 2.8 Library" [menu:Project->References])
Dim conexao As ADODB.Connection

'Declaração da Variavel lID do Tipo Long (variavel será utilizada em todo escopo do formulário)
Dim lID As Long

'Evento do Load do Formulário (Executado ao abrir o Formulário)
Private Sub Form_Load()
    'Tratamento de Erros
    On Error GoTo erros
    
    'Declaração da variável query do tipo String (variável poderá ser utilizada somente no escopo do procedimento)
    Dim query As String
    
    'Cria uma instância da variável conexão que está declarado no topo do código do formulário
    Set conexao = New ADODB.Connection
    
    'Alimentando String para Conexão com o Banco de dados Mysql Utilizando o Driver Mysql ODBC 3.51 (https://downloads.mysql.com/archives/get/p/10/file/mysql-connector-odbc-3.51.30-win32.msi)
    'Note que estão sendo utilizadas as constantes configuradas no topo do formulário
    'Fique a vontade caso queira utilizar um driver mais recente
    conexao.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};" & _
                               "Server=" & m_DB_Server & ";" & _
                               "Port=" & m_DB_Port & ";" & _
                               "User=" & m_DB_User & ";" & _
                               "Password=" & m_DB_Pass & ";" & _
                               "Option=3;"
                               
    'Efetivando a conexão com o Banco de dados
    conexao.Open
    
    'Comando para criar banco de dados caso não exista
    query = "CREATE DATABASE IF NOT EXISTS " & m_DB_Database
    
    'Executando comando
    conexao.Execute query
    
    'Comando para usar o banco de dados
    query = "USE " & m_DB_Database
    
    'Executando comando
    conexao.Execute query
    
    'Comando para criar a tabela que será utilizada em nosso exemplo
    query = "CREATE TABLE IF NOT EXISTS pessoa (" & _
            "   id INT NOT NULL AUTO_INCREMENT PRIMARY KEY, " & _
            "   nome VARCHAR(255) DEFAULT '' NOT NULL, " & _
            "   fone VARCHAR(30) DEFAULT '' NOT NULL ) "
    
    'executando comando
    conexao.Execute query
    
    'Chamando lista de contatos
    Call ListarContatos
    
    'Sai do Procedimento
    Exit Sub
erros:
    'Caso ocorra um erro o tratamento de erros é enviado para exibir esta mensagem
    MsgBox "Ocorreu um Erro no Sistema!" & vbNewLine & _
           Err.Number & " => " & Err.Description, vbCritical
End Sub

'Evento Click do Botão (Command) Novo (Executado ao Clicar no Botão)
Private Sub cmdNovo_Click()
    'Atribui o valor zero a variável lID
    lID = 0
    'Atribui Vazio ao Campo Nome (Objeto TextBox)
    txtNome.Text = Empty
    'Atribui Vazio ao Campo Fone (Objeto TextBox)
    txtFone.Text = Empty
End Sub

'Evento Click do Botão (Command) Gravar (Executado ao Clicar no Botão)
Private Sub cmdGravar_Click()
    'Declaração da variável query com o tipo string
    Dim query As String
    'Tratamento de Erros
    On Error GoTo erros
    
    'Verifica se o Campo Nome esta preenchido
    If txtNome.Text = "" Then
        'Exibe mensagem para avisar o usuário que o campo nome não foi preenchido
        MsgBox "Campo nome não preenchido!", vbExclamation, "Atenção"
        'Sai do Procedimento
        Exit Sub
    End If
    
    'Verifica se a variável lID seja igual a zero
    If lID = 0 Then
        'Caso a variável lID seja igual a zero entra para incluir um novo contato no banco de dados
        'Exibe mensagem questionando se o usuário deseja incluir um contato
        If MsgBox("Deseja incluir este contato?", vbQuestion + vbYesNo, "Inclusão de contato") = vbNo Then Exit Sub
        
        'Comando INSERT para inclusão do contato no banco de dados
        query = "INSERT INTO pessoa (nome,fone) VALUES ('" & txtNome.Text & "','" & txtFone.Text & "')"
        
        'Executa o comando
        conexao.Execute query
        
        'Exibe mensagem informando que o contato foi incluído com Sucesso
        MsgBox "Contato incluído com sucesso!", vbInformation, "Inclusão de contato"
    Else
        'Caso a variavél lID não seja igual a zero entra para fazer a alteração do contato existente
        'Exibe mensagem questionando se o usuário deseja alterar um contato
        If MsgBox("Deseja alterar este contato?", vbQuestion + vbYesNo, "Alteração de contato") = vbNo Then Exit Sub
        
        'Comando UPDATE para alteração do contato no banco de dados filtrando o id selecionado
        query = "UPDATE pessoa SET nome='" & txtNome.Text & "', fone='" & txtFone.Text & "' WHERE id=" & lID
        
        'Executa o comando
        conexao.Execute query
        
        'Exibe mensagem informando que o contato foi alterado com Sucesso
        MsgBox "Contato alterado com sucesso!", vbInformation, "Alteração de contato"
    End If
    
    'Chama Procedimento que Lista os Contatos no listbox
    Call ListarContatos
    
    'Sai do Procedimento
    Exit Sub
erros:
    'Caso ocorra um erro o tratamento de erros é enviado para exibir esta mensagem
    MsgBox "Ocorreu um erro no sistema!" & vbNewLine & _
           Err.Number & " => " & Err.Description, vbCritical, "Erro"
End Sub

'Procedimento ListarContatos
Private Sub ListarContatos()
    'Declaração da variável query com o tipo string
    Dim query As String
    'Declara e instancia rs do tipo ADODB.Recordset
    Dim rs As New ADODB.Recordset
    'Tratamento de Erros
    On Error GoTo erros
    
    'Comando SELECT para buscar todos os contatos cadastrados
    query = "SELECT * FROM pessoa"
    
    'Abre o recordset passando o comando (query) e a conexão com o banco de dados (conexao)
    rs.Open query, conexao
    
    'Limpa a lista (listbox name=contatos)
    contatos.Clear
    
    'Repetição enquanto não for o fim do arquivo (fim dos registros)
    Do While Not rs.EOF
        'Adiciona um item a lista (listbox)
        'passamos o nome do campo no objeto rs(recordset) para que ele nos devolva o valor encontrado
        'É feita uma formatação nos itens utilizando as seguintes funções:
        'Space = retorna espaços em branco de acordo com o valor passado (Ex: Space(3)='   ')
        'Right = Retorna uma string contendo o número de caracteres definido em Tamanho do lado direito da String
        'Left = Retorna uma string contendo o número de caracteres definido em Tamanho do lado esquerdo da String
        contatos.AddItem Right(Space(3) & rs("id"), 3) & " | " & Left(rs("nome") & Space(15), 15) & " | " & rs("fone")
        
        'move o recordset para o próximo registro
        rs.MoveNext
    Loop
    
    'Fecha o recordset
    rs.Close
    
    'Sai do procedimento
    Exit Sub
erros:
    'Caso ocorra um erro o tratamento de erros é enviado para exibir esta mensagem
    MsgBox "Ocorreu um erro no sistema!" & vbNewLine & _
           Err.Number & " => " & Err.Description, vbCritical, "Erro"
End Sub

'Evento DblClick do ListBox contatos (Executado ao Clicar Duas Vezes)
Private Sub contatos_DblClick()
    'Tratamento de Erros
    On Error GoTo erros
    
    'Atribui a variavel lID o valor do item selecionando
    'quebrando a string em um vetor e devolvendo a posição 0
    'que seria o id do contato (pessoa)
    lID = Trim(Split(contatos.Text, "|")(0))
    
    'Chama o Procedimento BuscaContato
    Call BuscaContato
    
    'Sai do Procedimento
    Exit Sub
erros:
    'Caso ocorra um erro o tratamento de erros é enviado para exibir esta mensagem
    MsgBox "Ocorreu um erro no sistema!" & vbNewLine & _
           Err.Number & " => " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub BuscaContato()
    'Declaração da variável query com o tipo string
    Dim query As String
    
    'Declara e instancia rs do tipo ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    'Tratamento de Erros
    On Error GoTo erros
    
    'Comando SELECT para buscar dados da pessoa filtrando o id informado na variável lID
    query = "SELECT * FROM pessoa WHERE id=" & lID
    
    'Abre o recordset passando o comando (query) e a conexão com o banco de dados (conexao)
    rs.Open query, conexao
    
    'Verifica se o recordset encontra-se no final
    If rs.EOF Then
        'Caso esteja no final significa que a busca não devolveu nenhum registro
        'Então limpa os campos e atribui zero ao lID
        lID = 0
        txtNome.Text = Empty
        txtFone.Text = Empty
    Else
        'Caso não esteja no final existe registro
        'Então atribui os valores aos seus respectivos campos
        txtNome.Text = rs("nome")
        txtFone.Text = rs("fone")
    End If
    
    'Fecha o Recordset
    rs.Close
    
    'Sai do Procedimento
    Exit Sub
erros:
    'Caso ocorra um erro o tratamento de erros é enviado para exibir esta mensagem
    MsgBox "Ocorreu um erro no sistema!" & vbNewLine & _
           Err.Number & " => " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmdExcluir_Click()
    'Declaração da variável query com o tipo string
    Dim query As String
    
    'Tratamento de Erros
    On Error GoTo erros
    
    'Verifica se a variável lID é igual a zero
    If lID = 0 Then
        'Se a variável lID é igual a zero não temos um registro selecionado para ser excluído
        'então exibe a mensagem abaixo
        MsgBox "Não é possível realizar a exclusão", vbExclamation, "Exclusão de contato"
        Exit Sub
    End If
    
    'Exibe mensagem questionando se o usuário deseja excluir o contato
    If MsgBox("Deseja excluir este contato?", vbQuestion + vbYesNo, "Atenção") = vbNo Then Exit Sub
    
    'Comando DELETE para exclusão do contato filtrando o id
    query = "DELETE FROM pessoa WHERE id=" & lID
        
    'Executa o comando
    conexao.Execute query
    
    'Exibe mensagem informando que o contato foi excluído com sucesso
    MsgBox "Contato excluído com sucesso!", vbInformation, "Exclusão de contato"
    
    'chama o procedimento do evento de clique no botão novo para iniciar um novo registro
    Call cmdNovo_Click
    
    'chama o procedimento para listar os contatos
    Call ListarContatos
    
    'Sai do Procedimento
    Exit Sub
erros:
    'Caso ocorra um erro o tratamento de erros é enviado para exibir esta mensagem
    MsgBox "Ocorreu um erro no sistema!" & vbNewLine & _
           Err.Number & " => " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Finaliza a Aplicação
    End
End Sub
