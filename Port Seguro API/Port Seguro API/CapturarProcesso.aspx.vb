Imports Techway
Imports System.Reflection
Imports IntegracaoTribunais.Estrutura.Fabricas
Imports IntegracaoTribunais.Estrutura
Imports IntegracaoTribunais.Estrutura.Exceptions
Imports IntegracaoTribunais.Utils
Imports System.IO
Imports System.Net
Imports System.Xml

Partial Class Paginas_CapturarProcesso
    Inherits System.Web.UI.Page
    Private usuario As New Usuario
    Private util As New Util
    Private sistemaExterno As New SistemaExterno
    Private opcaoSistemaExterno As New OpcaoSistemaExterno
    Private opcaoSistemaAdicional As New OpcaoSistemaAdicional
    Private opcaoCaptura As New OpcaoCaptura
    Private processo As New Processo
    Private instancia As New Instancia
    Private captura As New CapturaAtiva
    Private chkCabecalho As CheckBox
    Private chkItem As CheckBox
    Private funcoes As String
    Private Const colunaProcessonovo As Integer = 1
    Private Const colunaProcessoExistente As Integer = 2
    Private nvcProcesso As New NameValueCollection
    Private dtgLinha As DataRow
    Private dtgRow As GridViewRow
    Private listaDeProcessos As String
    Private listaRetorno As New List(Of KeyValuePair(Of String, String))
    Private diretorioDocumentos As String = ConfigurationManager.AppSettings("diretorioDocumentos")
    Private diretorioLogomarca As String = ConfigurationManager.AppSettings("diretorioLogomarca").Replace("..", "").Replace("~", "")
    Private lysisUtil As New LysisUtil
    Private idCaptcha As String = ""
    Private cookie As String = ""
    Private nvcEmpresa As NameValueCollection
    Private pastaArquivo As New PastaArquivo
    Private nvcPastaArquivo As NameValueCollection
    Protected Sub Page_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit

        If Page.Request.ServerVariables("http_user_agent").ToLower.Contains("safari") Then
            Page.ClientTarget = "uplevel"
        End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim acesso As Boolean
        Dim processo As New Processo
        Dim instancia As New Instancia
        Dim nvcProcesso As New NameValueCollection
        Dim pertenceEmpresa As Boolean
        Dim identificadorTipoInstanciaPadrao As String = ""
        Dim grupoInstancia As New GrupoInstancia

        Try
            lblMensagem.Text = ""
            lblMensagem.Visible = False

            usuario = CType(Session("usuario"), Usuario)

            'processo.identificador = Request("IsnProcesso").ToString

            'instancia.identificador = Request("Isn").ToString
            pertenceEmpresa = False

            'Verifica se a empresa do usuário é igual a empresa do registro acessado
            'If Request("Isn").ToString <> "" Then
            '    If usuario.empresa.pessoa.identificador = processo.getIdentificadorEmpresa And usuario.empresa.pessoa.identificador = instancia.getIdentificadorEmpresa Then
            '        pertenceEmpresa = True
            '    End If
            'Else
            '    If usuario.empresa.pessoa.identificador = processo.getIdentificadorEmpresa Then
            '        pertenceEmpresa = True
            '    End If
            'End If

        Catch ex As Exception
            Session("erro") = ex
            Response.Redirect("../Paginas/Erro.aspx")
        End Try

        'If Not pertenceEmpresa Then
        '    Session("erro") = Mensagem.naoPertenceEmpresa
        '    Response.Redirect("../Paginas/Erro.aspx")
        'End If

        If Not IsPostBack Then
            Try

                lblMensagem.Text = ""
                lblMensagem.Visible = False
                Session("consultaNumeroOab") = False
                Session("consultaComplementoOab") = False
                'If acesso Then

                If lblIsn.Text <> "" Then
                    exibir()
                Else
                    carregarlistas()
                End If

                util.PrevineDuplaSubmissao(btnCapturar, ClientScript)
                util.PrevineDuplaSubmissao(btnConsultar, ClientScript)

                If Request("IsnEvento") IsNot Nothing Then

                    Dim evento = New Evento

                    txtNumPro.Visible = True
                    lblTipoConsulta.Visible = True
                    txtNumPro.Text = evento.GetNumeroProcesso(Request("isnEvento").ToString)
                    txtNumPro.Enabled = False
                    lblOpcaoSistemaExterno.Visible = True
                    drpSeletorOpcao.Enabled = False
                    drpSeletorOpcao.SelectedIndex = 1
                    drpSistemaExterno.Visible = True
                    btnCapturar.Visible = True
                    btnConsultar.Visible = False

                    'Verificação do número no formato CNJ
                    Dim lysisUtil = New LysisUtil

                    If lysisUtil.verificarNumeroCNJ(txtNumPro.Text) Then

                        'Sistema Externo
                        Dim processoPublicacao = New Processo With {.numeroProcessos = txtNumPro.Text}
                        Dim sistemaInsercaoARL As ArrayList = processoPublicacao.listSistemasCaptura("1", txtNumPro.Text)

                        drpCodOpcao.Items.Clear()
                        drpCodOpcao.Items.Add(New ListItem("", 0))
                        drpCodOpcao.Visible = False
                        lblOpcaoSistema.Visible = False
                        drpSistemaExterno.Items.Clear()
                        drpSistemaExterno.Items.Add(New ListItem("", 0))

                        For Each sistema In sistemaInsercaoARL
                            drpSistemaExterno.Items.Add(New ListItem(sistema("nome"), sistema("identificador")))
                        Next

                        lblMensagem.Visible = False
                        drpSistemaExterno.Focus()

                    Else
                        carregarlistas()
                    End If

                End If

                'End If

            Catch ex As Exception
                Session("erro") = ex
                Response.Redirect("../Paginas/Erro.aspx")
            End Try
        End If

    End Sub
    Private Function obtemSubPasta(ByVal isnDocumento As String, ByVal tabela As String, ByVal indentificadorTabela As String) As String
        Dim subPasta As String = Nothing
        nvcEmpresa = usuario.empresa.show()
        If nvcEmpresa("TIP_SUBPASTA") = 1 Then
            nvcPastaArquivo = criarDiretoriosSubPastas()

            If usuario.empresa.tipoArmazenamentoExterno = 2 And usuario.empresa.idDiretorioDrive <> "0" Then
                'Upload GoogleDrive
                If nvcEmpresa("TIP_SUBPASTA") = 1 Then
                    pastaArquivo.isnPasta = nvcPastaArquivo("ISN_PASTA_ARQUIVO").ToString()
                    pastaArquivo.isnDocumento = isnDocumento
                    pastaArquivo.isnEmpresa = usuario.empresa.pessoa.identificador
                    If pastaArquivo.isnDocumento <> "" And pastaArquivo.isnEmpresa <> "" Then
                        pastaArquivo.updateIsnPasta(tabela, indentificadorTabela)
                    End If
                    subPasta = nvcPastaArquivo("DSC_DIRETORIO_DRIVE")

                Else

                    subPasta = usuario.empresa.idDiretorioDrive
                End If
            Else
                If nvcEmpresa("TIP_SUBPASTA") = 1 Then
                    pastaArquivo.isnPasta = nvcPastaArquivo("ISN_PASTA_ARQUIVO").ToString()
                    pastaArquivo.isnDocumento = isnDocumento
                    pastaArquivo.isnEmpresa = usuario.empresa.pessoa.identificador
                    pastaArquivo.updateIsnPasta(tabela, indentificadorTabela)
                    subPasta = nvcPastaArquivo("NUM_PASTA")
                Else

                    subPasta = Nothing
                End If

            End If

        Else
            subPasta = Nothing
        End If

        Return subPasta
    End Function

    Private Function criarDiretoriosSubPastas() As NameValueCollection

        nvcEmpresa = usuario.empresa.show()
        If nvcEmpresa("TIP_SUBPASTA") = 1 Then

            pastaArquivo.isnEmpresa = usuario.empresa.pessoa.identificador
            nvcPastaArquivo = pastaArquivo.insert()
            pastaArquivo.isnPasta = nvcPastaArquivo("ISN_PASTA_ARQUIVO")
            'Verificar se a pasta retornada pela procedure existe, caso não, criar
            If usuario.empresa.tipoArmazenamentoExterno = 2 Then
                If Not lysisUtil.verificaExistenciaPastaDrive(usuario.empresa.pessoa.identificador, nvcPastaArquivo("NUM_PASTA").ToString, usuario.empresa.idDiretorioDrive) Then
                    'Criando diretório no drive dentro da pasta pai e obtendo o ID da nova pasta criada
                    Dim idPastaCriadaNoDrive = lysisUtil.CriarDiretorioGoogleDrive(usuario.empresa.pessoa.identificador, nvcPastaArquivo("NUM_PASTA").ToString, usuario.empresa.idDiretorioDrive)
                    pastaArquivo.diretorio = idPastaCriadaNoDrive

                    pastaArquivo.updateDiretorioDriver()
                    nvcPastaArquivo.Remove("DSC_DIRETORIO_DRIVE")
                    nvcPastaArquivo.Add("DSC_DIRETORIO_DRIVE", idPastaCriadaNoDrive.ToString)
                End If
            Else
                'verificar se a pasta existe localmente, caso não, cria.
                If Not IO.Directory.Exists(diretorioDocumentos & usuario.empresa.pessoa.identificador & "\" & nvcPastaArquivo("NUM_PASTA").ToString) Then
                    My.Computer.FileSystem.CreateDirectory(diretorioDocumentos & usuario.empresa.pessoa.identificador & "\" & nvcPastaArquivo("NUM_PASTA").ToString)
                End If

            End If

        End If
        Return nvcPastaArquivo

    End Function

    Protected Sub drpSistemaExterno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles drpSistemaExterno.SelectedIndexChanged
        Dim opcaoSistemaExterno As New OpcaoSistemaExterno
        Dim opcaoCaptura As New OpcaoCaptura
        Dim sistemaExterno As New SistemaExterno
        Dim opcaoSistemaAdicional As New OpcaoSistemaAdicional
        Dim nvcSistemaExterno As New NameValueCollection
        Dim lista As New Lista
        Dim tipoDeBibliotecaIntegracao As String = ""
        Try

            drpOpcaoAdicionalCaptura.Items.Clear()
            drpOpcaoAdicionalCaptura.Visible = False
            lblOpcaoAdicionalCaptura.Visible = False

            'Carrega a opção do sistema de acordo com o sistema selecionado
            If drpSistemaExterno.SelectedIndex > 0 Then

                'Opção sistema externo
                tipoDeBibliotecaIntegracao = "3"
                sistemaExterno.identificador = drpSistemaExterno.SelectedValue
                opcaoSistemaExterno.sistemaExterno = sistemaExterno
                opcaoSistemaExterno.carregarlista(drpCodOpcao, tipoDeBibliotecaIntegracao, drpSeletorOpcao.SelectedIndex)
                opcaoCaptura.sistemaExterno = sistemaExterno
                opcaoCaptura.opcaoSistemaExterno = New OpcaoSistemaExterno
                opcaoCaptura.carregarlista(drpOpcaoCaptura)

                If drpCodOpcao.Items.Count > 1 Then
                    'Verificação para que as opções de sistema não sejam chamadas para o TJMG na consulta de processos.
                    If Not sistemaExterno.identificador.Equals("103") Or drpSeletorOpcao.SelectedIndex = 2 Then
                        lblOpcaoSistema.Visible = True
                        drpCodOpcao.Visible = True
                    End If

                    sistemaExterno.identificador = drpSistemaExterno.SelectedValue
                    nvcSistemaExterno = sistemaExterno.show

                    'Nome opção
                    If nvcSistemaExterno("nom_opcao").ToString() <> "" Then
                        lblOpcaoSistema.Text = "Opção do Sistema - " & nvcSistemaExterno("nom_opcao").Trim
                    Else
                        lblOpcaoSistema.Text = "Opção do Sistema"
                    End If

                    drpCodOpcao.Focus()

                Else
                    validacaoDropDownEstado(sistemaExterno, opcaoSistemaExterno)
                    validacaoDropDownComplemento(sistemaExterno, opcaoSistemaExterno)

                    drpCodOpcao.Items.Clear()

                    lblOpcaoSistema.Visible = False
                    drpCodOpcao.Visible = False

                    drpSistemaExterno.Focus()

                End If

                'Carrega a opção de captura de acordo com o sistema selecionado, caso a busca seja po oab
                If drpOpcaoCaptura.Items.Count > 1 And drpSeletorOpcao.SelectedIndex = 2 Then

                    lblOpcaoCaptura.Visible = True
                    drpOpcaoCaptura.Visible = True

                    sistemaExterno.identificador = drpSistemaExterno.SelectedValue
                    nvcSistemaExterno = sistemaExterno.show

                    'Nome opção captura
                    If nvcSistemaExterno("nom_opcao_captura").ToString() <> "" Then
                        lblOpcaoCaptura.Text = "Opção de Captura - " & nvcSistemaExterno("nom_opcao_captura").Trim
                    Else
                        lblOpcaoCaptura.Text = "Opção de Captura"
                    End If

                    Session("consultaOpcaoCaptura") = True
                    drpCodOpcao.Focus()

                Else
                    validacaoDropDownEstado(sistemaExterno, opcaoSistemaExterno)
                    validacaoDropDownComplemento(sistemaExterno, opcaoSistemaExterno)

                    drpOpcaoCaptura.Items.Clear()

                    lblOpcaoCaptura.Visible = False
                    drpOpcaoCaptura.Visible = False

                    Session("consultaOpcaoCaptura") = False

                    drpSistemaExterno.Focus()

                End If

            Else

                drpCodOpcao.Items.Clear()

                lblOpcaoSistema.Visible = False
                drpCodOpcao.Visible = False

                drpOpcaoCaptura.Items.Clear()

                lblOpcaoCaptura.Visible = False
                drpOpcaoCaptura.Visible = False

                drpSistemaExterno.Focus()

            End If

        Catch ex As Exception

            Session("erro") = ex
            Response.Redirect("../Paginas/Erro.aspx")

        End Try

    End Sub

    Private Sub exibir()
        Dim processo As New Processo


        processo.identificador = lblIsn.Text

    End Sub

    Public Shared Function ArrayListToDataTable(ByVal arraylist1 As ArrayList) As System.Data.DataTable
        Dim dt As New System.Data.DataTable()

        For i As Integer = 0 To arraylist1.Count - 1
            Dim GenericObject As Object = arraylist1.Item(i)
            Dim NbrProp As Integer = GenericObject.GetType().GetProperties().Count

            For Each item As PropertyInfo In GenericObject.GetType().GetProperties()
                Try
                    Dim column = New DataColumn()
                    Dim ColName As String = item.Name.ToString

                    column.ColumnName = ColName
                    dt.Columns.Add(column)

                Catch

                End Try
            Next

            Dim row As DataRow = dt.NewRow()

            Dim j As Integer = 0
            For Each item As PropertyInfo In GenericObject.GetType().GetProperties()
                row(j) = item.GetValue(GenericObject, Nothing)
                j += 1
            Next

            dt.Rows.Add(row)

        Next
        Return dt

    End Function

    Private Sub carregarlistas()
        Dim lista As New Lista

        Dim opcaoCaptura = drpSeletorOpcao.SelectedIndex
        If opcaoCaptura = Nothing Or opcaoCaptura = 0 Then
            opcaoCaptura = 1
        End If
        Dim processo As New Processo

        Dim sistemaCapturaARL As ArrayList = processo.listSistemasCaptura(opcaoCaptura, "")

        drpSistemaExterno.Items.Clear()

        drpSistemaExterno.Items.Add(New ListItem("", 0))

        For Each sistema In sistemaCapturaARL
            drpSistemaExterno.Items.Add(New ListItem(sistema("nome"), sistema("identificador")))
        Next

        lista.CarregaCombo(drpIsnEstadoOab, "ISN_ESTADO", "NOM_ESTADO", "EST_ESTADO", True)



    End Sub

    Private Function validarCampos() As Boolean
        Dim validacaoCampos As Boolean = False

        'Número do processo e OAB vazios
        If txtNumOAB.Text = "" And txtNumPro.Text = "" Then
            lblMensagem.Text = "Número do processo ou número OAB deve ser informado."
            lblMensagem.Visible = True
            validacaoCampos = False
        End If
        'Número do processo informado e Sistema vazio
        If txtNumPro.Text <> "" And drpSistemaExterno.SelectedValue = -1 Then
            lblMensagem.Text = "Sistema deve ser informado."
            lblMensagem.Visible = True
            validacaoCampos = False
        End If
        'Número do OAB informado e Sistema vazio
        If txtNumOAB.Text <> "" And drpSistemaExterno.SelectedValue = -1 Then
            lblMensagem.Text = "Sistema deve ser informado."
            lblMensagem.Visible = True
            validacaoCampos = False
        End If


        Return validacaoCampos
    End Function

    Private Function validarFormatoProcesso() As Boolean
        lblMensagem.Text = ""
        lblMensagem.Visible = False
        Dim status As Boolean = False

        Dim numProc As String = txtNumPro.Text

        'Exceções
        Select Case drpSistemaExterno.SelectedValue
            'STF
            Case 7
                Return True
            'TCECE
            Case 37
                Return True
            'PROJUDI
            Case 75
                Select Case drpCodOpcao.SelectedValue
                    'Ceará
                    Case 1
                        Return True
                End Select
        End Select

        For index As Integer = 0 To numProc.Length - 1
            If index = 7 Then
                If Not (numProc(index).ToString.Contains("-") Or numProc(index).ToString.Contains(".")) Then
                    status = False
                    Return status
                Else
                    status = True
                End If
            End If

            If index = 10 Then
                If Not numProc(index).ToString.Contains(".") Then
                    status = False
                    Return status
                Else
                    status = True
                End If
            End If
            If index = 15 Then
                If Not numProc(index).ToString.Contains(".") Then
                    status = False
                    Return status
                Else
                    status = True
                End If
            End If

            If index = 17 Then
                If Not numProc(index).ToString.Contains(".") Then
                    status = False
                    Return status
                Else
                    status = True
                End If
            End If

            If index = 20 Then
                If Not numProc(index).ToString.Contains(".") Then
                    status = False
                    Return status
                Else
                    status = True
                End If
            End If
        Next

        If status And numProc.Length < 23 Then
            status = False
            Return status
        End If

        Return status
    End Function

    Private Function RetornaProcesso(ByVal numProcesso As String) As String
        Dim isnProcesso As String
        isnProcesso = instancia.getNumeroProcessoPorIdentificador(numProcesso, usuario.empresa)
        Return isnProcesso
    End Function

    Private Function RetornaProcessoAnterior(ByVal numProcesso As String) As String
        Dim isnProcesso As String
        numProcesso = numProcesso.Split(" ").First.Trim.ToUpper & " " & numProcesso.Split(" ").Last.Trim
        isnProcesso = instancia.getNumeroProcessoAnteriorPorIdentificador(numProcesso, usuario.empresa)
        Return isnProcesso
    End Function

    Private Function SalvaProcesso(ByVal nvcResultadoIntegracao As NameValueCollection) As String
        Dim Isn As Integer
        Dim processo As New Processo
        Dim materia As New Materia
        Dim materiaPessoa As New MateriaPessoa
        Dim retorno As String = ""
        Dim partePoloAtivo As Parte = New Parte
        Dim partePoloPassivo As Parte = New Parte
        Dim pessoaPartePoloAtivo As New Pessoa
        Dim pessoaPartePoloPassivo As New Pessoa
        Dim tipoPartePoloPassivo As TipoParte
        Dim tipoPartePoloAtivo As TipoParte
        Dim pessoaEmpresa As New Pessoa
        Dim empresa As New Empresa
        Dim pastaVirtual As PastaVirtual = New PastaVirtual
        Dim logAtualizacao As LogAtualizacao = New LogAtualizacao
        Dim pessoaAdvogado As Pessoa = New Pessoa
        Dim advogado As Advogado = New Advogado
        Dim aDvogadosPartesAtivos As String()
        Dim aDvogadosPartesPassivos As String()
        Dim pessoaPartesAtivos As String()
        Dim pessoaPartesPassivos As String()
        Dim andamentos As String()
        Dim tipoPartesAtivos As String()
        Dim tipoPartesPassivos As String()
        Dim tipoAcao As TipoAcao = New TipoAcao
        Dim vara As Vara = New Vara
        Dim tipoVara As TipoVara = New TipoVara
        Dim comarca As Comarca = New Comarca
        Dim grupoInstancia As GrupoInstancia = New GrupoInstancia
        Dim numero As Numero = New Numero
        Dim evento As Evento = New Evento
        Dim orgao As Orgao = New Orgao
        Dim tipoMagistrado As TipoMagistrado = New TipoMagistrado
        Dim pessoaRelator As Pessoa = New Pessoa
        Dim sessionX As New Session
        Dim resultadoIntegracao As New ArrayList
        Dim resultadoIntegracaoEventos As New ArrayList
        Dim contadorIntegracao As String = "0"
        Dim contadorSessao As String = "0"
        Dim contadorPartes As Int32 = 0
        Dim tituloParteAtivo As String = ""
        Dim tituloPartePassivo As String = ""
        Dim rito As New Rito
        Dim nvcRito As New NameValueCollection
        Dim fase As New Fase
        Dim nvcFase As New NameValueCollection
        Dim tipoPartes As String()
        Dim advogadosGeral As String()
        Dim nomePartesGeral As String()
        Dim tipoDeParte As New TipoParte
        Dim advogadosGerais As String()
        Dim parteGeral As New Parte
        Dim pessoaParteGeral As New Pessoa
        Dim pessoaCliente As New Pessoa
        Dim grupoPessoa As New GrupoPessoa
        Dim listaDeProcesso(0) As String
        Dim resultadoIntegracaoAntiga As New ArrayList
        Dim linhaQtdAtivos As Integer = 0
        Dim linhaQtdPassivos As Integer = 0
        Dim erroMensagem As String = ""
        Dim objetoProcesso As New ObjetoProcesso
        Dim objetos As String()
        Dim relatores As String()
        Dim identificadorPessoasAtivas As String()
        Dim identificadorPessoasPassivas As String()
        Dim identificadorAdvAtivo As String()
        Dim identificadorAdvPassivo As String()
        '---------------------------------------
        'Processo
        '---------------------------------------
        processo.titulo = ""

        'Verifica se o processo existe ou não no tribunal escolhido
        erroMensagem = nvcResultadoIntegracao("ErroCaptura")
        If erroMensagem <> "" Then
            lblMensagem.Visible = True
            lblMensagem.Text = erroMensagem
            Return retorno
        End If



        ' materia
        materia.descricao = nvcResultadoIntegracao("Area").Replace(",", "")
        If materia.descricao <> "" Then
            materia.identificador = materiaPessoa.getIdentificadorMateria(nvcResultadoIntegracao("Area"), usuario.empresa.pessoa.identificador)
            If materia.identificador = "" Then 'Se não existe identificador
                materia.empresa = usuario.empresa
                materia.status = "1"
                materia.insert()
                materia.identificador = materia.getIdentificadorMateria
            End If
        End If


        'Valor da causa
        processo.valor = nvcResultadoIntegracao("valoracao")


        'Data do status
        processo.dataStatus = Date.Now

        'Status
        processo.status = 0

        'Indicador de processo eletrônico
        processo.eletronico = IIf(nvcResultadoIntegracao("EEletronico").Equals("1"), nvcResultadoIntegracao("EEletronico"), "0")

        'Indicador de processo estratégico
        processo.estrategico = "0"


        'Empresa
        pessoaEmpresa.identificador = CType(usuario.empresa.pessoa.identificador, String)
        empresa.pessoa = pessoaEmpresa
        processo.empresa = empresa
        'Data distribuicao
        processo.dataDistribuicao = IIf(nvcResultadoIntegracao("dtDistribuicao") <> "", nvcResultadoIntegracao("dtDistribuicao"), Date.Now)

        'Data citação
        processo.dataCitacao = IIf(nvcResultadoIntegracao("dtCitacao") <> "", nvcResultadoIntegracao("dtCitacao"), "")

        'Observação
        processo.observacao = IIf(nvcResultadoIntegracao("Observacoes") <> "", nvcResultadoIntegracao("Observacoes"), "")


        processo.identificador = ""
        instancia.identificador = ""
        instancia.numeroProcesso = IIf(nvcResultadoIntegracao("numeroProcesso") <> "", nvcResultadoIntegracao("numeroProcesso"), txtNumPro.Text.Trim)
        instancia.processo = processo
        processo.identificador = RetornaProcesso(instancia.numeroProcesso)
        If processo.identificador <> "" Then
            lblMensagem.Visible = True
            lblMensagem.Text = "Já existe um registro no Sistema!"
            retorno = ""
            Return retorno
        Else
            processo.identificador = instancia.getProcessoSimilar(txtNumPro.Text, usuario.empresa)
            If processo.identificador <> "" Then
                lblMensagem.Visible = True
                lblMensagem.Text = "Já existe um registro no Sistema!"
                retorno = ""
                Return retorno
            End If
        End If

        'salva o processo
        tipoAcao.descricao = IIf(nvcResultadoIntegracao("Classe") <> "", nvcResultadoIntegracao("Classe"), nvcResultadoIntegracao("Assunto"))
        If tipoAcao.descricao <> "" Then
            tipoAcao.empresa = usuario.empresa
            tipoAcao.identificador = tipoAcao.selecionaTipoAcao()
            If tipoAcao.identificador = "" Then
                tipoAcao.materia = materia
                tipoAcao.status = "1"
                tipoAcao.grupoProcedimento = New GrupoProcedimento
                tipoAcao.identificador = tipoAcao.insert
            End If
        Else
            tipoAcao = New TipoAcao
        End If

        processo.tipoAcao = tipoAcao
        processo.materia = materia
        processo.centroCusto = New CentroCusto
        processo.natureza = New Natureza
        processo.rito = rito
        processo.fase = New Fase
        processo.tipoRisco = New TipoRisco
        processo.escritorio = New Pessoa
        processo.grupoProcesso = New GrupoProcesso
        processo.probabilidadePerda = New Probabilidade
        processo.probabilidade = New Probabilidade
        processo.prioridadeDe = New Pessoa
        processo.valorAtualizado = ""
        processo.valorProvisionado = ""
        processo.valorContingencia = ""
        processo.cliente = New Pessoa

        retorno = processo.insert()

        processo.identificador = retorno

        'Objeto do processo

        If Not nvcResultadoIntegracao("Objetos") Is Nothing Then
            objetos = IIf(nvcResultadoIntegracao("Objetos") <> "", nvcResultadoIntegracao("Objetos").Split("<br>"), Nothing)
            If Not nvcResultadoIntegracao("Objetos").Contains("&") Then
                For Each obj As String In objetos
                    obj = obj.Replace("br>", "").Trim
                    If obj <> "" Then
                        objetoProcesso.processo = processo
                        objetoProcesso.tipoObjeto = New TipoObjeto
                        objetoProcesso.descricao = obj.Trim
                        objetoProcesso.principal = "1"
                        objetoProcesso.insert()
                    End If
                Next
            Else
                objetos = IIf(nvcResultadoIntegracao("Objetos") <> "", nvcResultadoIntegracao("Objetos").Split("&"), Nothing)
                If objetos.Length >= 1 Then
                    For Each obj As String In objetos
                        obj = obj.Replace("br>", "").Trim
                        If obj <> "" Then
                            objetoProcesso.processo = processo
                            objetoProcesso.tipoObjeto = New TipoObjeto
                            objetoProcesso.descricao = obj.Trim
                            objetoProcesso.principal = "1"
                            objetoProcesso.insert()
                        End If
                    Next
                End If
            End If
        End If
        'If usuario.empresa.geraNumPasta = 1 Then

        'Retirado geração de número de processo conforme atualização de erro em 26/01/2018 relativo ao mesmo.
        'pastaVirtual.processo = processo
        'pastaVirtual.identificador = pastaVirtual.getIdentificadorPasta
        'pastaVirtual.nomePasta = ""
        'Gera log do processo
        logAtualizacao.gerarLog(usuario, "1", "1", processo.identificador, "PRO_PROCESSO", processo.identificador, "1")

        'Gera log da pasta virtual do processo
        'logAtualizacao.gerarLog(usuario, "1", "18", pastaVirtual.identificador, "PAV_PASTA_VIRTUAL", processo.identificador, "1")
        'End If
        If processo.identificador <> "" Then

            'Parte e tipo partes
            'Pessoa
            pessoaPartesAtivos = nvcResultadoIntegracao("Requerente").Split("&")
            pessoaPartesPassivos = nvcResultadoIntegracao("Requerido").Split("&")
            'Insere, se houver CPF ou CNPJ
            If Not nvcResultadoIntegracao("identificadorAtivos") Is Nothing And nvcResultadoIntegracao("identificadorAtivos") <> "" Then
                identificadorPessoasAtivas = nvcResultadoIntegracao("identificadorAtivos").Split("&")
            Else
                identificadorPessoasAtivas = Nothing
            End If
            If Not nvcResultadoIntegracao("identificadorPassivos") Is Nothing And nvcResultadoIntegracao("identificadorPassivos") <> "" Then
                identificadorPessoasPassivas = nvcResultadoIntegracao("identificadorPassivos").Split("&")
            Else
                identificadorPessoasPassivas = Nothing
            End If


            If pessoaPartesAtivos.Length >= pessoaPartesPassivos.Length Then
                contadorPartes = pessoaPartesAtivos.Length - 1
            Else
                contadorPartes = pessoaPartesPassivos.Length - 1
            End If
            Dim j As Int32 = 0
            Dim k As Int32 = 0
            For i As Integer = 0 To contadorPartes
                If k < pessoaPartesAtivos.Length Then
                    If Not pessoaPartesAtivos(k) Is Nothing Then
                        pessoaPartePoloAtivo.nome = pessoaPartesAtivos(k).Trim
                    End If
                End If
                If j < pessoaPartesPassivos.Length Then
                    If Not pessoaPartesPassivos(j) Is Nothing Then
                        pessoaPartePoloPassivo.nome = pessoaPartesPassivos(j).Trim
                    End If
                End If

                If pessoaPartePoloAtivo.nome <> "" Or pessoaPartePoloPassivo.nome <> "" Then
                    pessoaPartePoloAtivo.empresa = usuario.empresa
                    tipoPartePoloAtivo = New TipoParte
                    tipoPartePoloPassivo = New TipoParte
                    tipoPartesAtivos = nvcResultadoIntegracao("tipoParteTipoAtivo").Split(",")
                    tipoPartesPassivos = nvcResultadoIntegracao("tipoParteTipoPassivo").Split(",")
                    pessoaPartePoloPassivo.empresa = usuario.empresa
                    For y As Integer = 0 To ((pessoaPartesAtivos.Length - 1))

                        tipoPartePoloAtivo.descricao = tipoPartesAtivos(y)
                        tipoPartePoloAtivo.empresa = usuario.empresa

                        If drpSistemaExterno.SelectedValue = 98 And tipoPartesAtivos.Length > i Then
                            tipoPartePoloAtivo.descricao = tipoPartesAtivos(i)
                            tipoPartePoloAtivo.empresa = usuario.empresa
                        End If

                        If y = 0 Then
                            tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                            If tipoPartePoloAtivo.identificador = "" Then
                                tipoPartePoloAtivo.status = "1"
                                tipoPartePoloAtivo.insert()
                                tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                            End If
                            partePoloAtivo.tipoParte = tipoPartePoloAtivo
                        End If

                        tipoPartePoloPassivo.descricao = tipoPartesPassivos(y)
                        tipoPartePoloPassivo.empresa = usuario.empresa
                        If y = 0 Then

                            tipoPartePoloPassivo.identificador = tipoPartePoloPassivo.getIdentificadorTipoParte()
                            If tipoPartePoloPassivo.identificador = "" Then
                                tipoPartePoloPassivo.status = "1"
                                tipoPartePoloPassivo.insert()
                                tipoPartePoloPassivo.identificador = tipoPartePoloPassivo.getIdentificadorTipoParte()
                            End If
                            partePoloPassivo.tipoParte = tipoPartePoloPassivo
                        End If

                        pessoaPartePoloAtivo.identificador = pessoaPartePoloAtivo.getIdentificadorPessoa()
                        If pessoaPartePoloAtivo.identificador <> "" Then
                            partePoloAtivo.pessoa = pessoaPartePoloAtivo
                        Else
                            pessoaPartePoloAtivo.profissao = New Profissao
                            pessoaPartePoloAtivo.nacionalidade = New Nacionalidade
                            pessoaPartePoloAtivo.estadoOab = New Estado
                            pessoaPartePoloAtivo.estado = New Estado
                            pessoaPartePoloAtivo.pais = New Pais
                            pessoaPartePoloAtivo.grupoPessoa = New GrupoPessoa
                            pessoaPartePoloAtivo.tipoPessoa = New TipoPessoa
                            pessoaPartePoloAtivo.grupoTrabalho = New GrupoTrabalho
                            pessoaPartePoloAtivo.status = "1"
                            pessoaPartePoloAtivo.identificador = pessoaPartePoloAtivo.insert()
                            tipoPartePoloAtivo.descricao = tipoPartesAtivos(y)
                            tipoPartePoloAtivo.empresa = usuario.empresa

                            If drpSistemaExterno.SelectedValue = 98 And tipoPartesAtivos.Length > i Then
                                tipoPartePoloAtivo.descricao = tipoPartesAtivos(i)
                                tipoPartePoloAtivo.empresa = usuario.empresa
                            End If

                            If y = 0 Then
                                tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                                If tipoPartePoloAtivo.identificador = "" Then
                                    tipoPartePoloAtivo.status = "1"
                                    tipoPartePoloAtivo.insert()
                                    tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                                End If
                                partePoloAtivo.tipoParte = tipoPartePoloAtivo
                            Else
                                tipoPartePoloPassivo.descricao = tipoPartesPassivos(j)
                                tipoPartePoloPassivo.empresa = usuario.empresa
                                tipoPartePoloPassivo.identificador = tipoPartePoloPassivo.getIdentificadorTipoParte()
                                If tipoPartePoloPassivo.identificador = "" Then
                                    tipoPartePoloPassivo.status = "1"
                                    tipoPartePoloPassivo.insert()
                                    tipoPartePoloPassivo.identificador = tipoPartePoloPassivo.getIdentificadorTipoParte()
                                End If
                                partePoloPassivo.tipoParte = tipoPartePoloPassivo
                            End If
                        End If

                        pessoaPartePoloPassivo.identificador = pessoaPartePoloPassivo.getIdentificadorPessoa()
                        If pessoaPartePoloPassivo.identificador <> "" Then
                            partePoloPassivo.pessoa = pessoaPartePoloPassivo
                        Else
                            pessoaPartePoloPassivo.profissao = New Profissao
                            pessoaPartePoloPassivo.nacionalidade = New Nacionalidade
                            pessoaPartePoloPassivo.estadoOab = New Estado
                            pessoaPartePoloPassivo.estado = New Estado
                            pessoaPartePoloPassivo.pais = New Pais
                            pessoaPartePoloPassivo.grupoPessoa = New GrupoPessoa
                            pessoaPartePoloPassivo.tipoPessoa = New TipoPessoa
                            pessoaPartePoloPassivo.grupoTrabalho = New GrupoTrabalho
                            pessoaPartePoloPassivo.status = "1"
                            pessoaPartePoloPassivo.identificador = pessoaPartePoloPassivo.insert()
                            tipoPartePoloPassivo.descricao = tipoPartesPassivos(y)
                            tipoPartePoloPassivo.empresa = usuario.empresa
                            If y = 0 Then
                                tipoPartePoloPassivo.identificador = tipoPartePoloPassivo.getIdentificadorTipoParte()
                                If tipoPartePoloPassivo.identificador = "" Then
                                    tipoPartePoloPassivo.status = "1"
                                    tipoPartePoloPassivo.insert()
                                    tipoPartePoloPassivo.identificador = tipoPartePoloPassivo.getIdentificadorTipoParte()
                                End If
                                partePoloPassivo.tipoParte = tipoPartePoloPassivo
                            Else
                                tipoPartePoloAtivo.descricao = tipoPartesAtivos(y)
                                tipoPartePoloAtivo.empresa = usuario.empresa

                                If drpSistemaExterno.SelectedValue = 98 And tipoPartesAtivos.Length > i Then
                                    tipoPartePoloAtivo.descricao = tipoPartesAtivos(i)
                                    tipoPartePoloAtivo.empresa = usuario.empresa
                                End If

                                tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                                If tipoPartePoloAtivo.identificador = "" Then
                                    tipoPartePoloAtivo.status = "1"
                                    tipoPartePoloAtivo.insert()
                                    tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                                End If
                                partePoloPassivo.tipoParte = tipoPartePoloPassivo
                            End If
                        End If

                        Exit For

                    Next

                    If Not tipoPartePoloAtivo.identificador Is Nothing Then
                        partePoloAtivo.tipoPoloAtivo = "1"
                    Else
                        partePoloAtivo.tipoPoloAtivo = "0"
                    End If

                    If Not tipoPartePoloPassivo.identificador Is Nothing Then
                        partePoloPassivo.tipoPoloAtivo = "1"
                    Else
                        partePoloPassivo.tipoPoloAtivo = Nothing
                    End If

                    If pessoaPartePoloAtivo.nome <> "" Then
                        partePoloAtivo.planoPrevidencia = New PlanoPrevidencia
                        If (Not identificadorPessoasAtivas Is Nothing) Then
                            If identificadorPessoasAtivas(k).Trim.Length = 14 Then
                                pessoaPartePoloAtivo.cpf = identificadorPessoasAtivas(k).Trim
                                pessoaPartePoloAtivo.classePessoa = "1"
                                pessoaPartePoloAtivo.tipoPessoa = New TipoPessoa
                                pessoaPartePoloAtivo.pertenceEmpresa = "0"
                                Dim grupoPessoas As New GrupoPessoa
                                Dim tiPessoa As New TipoPessoa
                                Dim pais As New Pais
                                Dim estaOab As New Estado
                                Dim estado As New Estado
                                Dim nacio As New Nacionalidade
                                Dim prof As New Profissao
                                pessoaPartePoloAtivo.grupoPessoa = grupoPessoas
                                pessoaPartePoloAtivo.tipoPessoa = tiPessoa
                                pessoaPartePoloAtivo.pais = pais
                                pessoaPartePoloAtivo.estadoOab = estaOab
                                pessoaPartePoloAtivo.estado = estado
                                pessoaPartePoloAtivo.nacionalidade = nacio
                                pessoaPartePoloAtivo.profissao = prof
                                pessoaPartePoloAtivo.update()
                            Else
                                pessoaPartePoloAtivo.cnpj = identificadorPessoasAtivas(k).Trim
                                pessoaPartePoloAtivo.classePessoa = "2"
                                pessoaPartePoloAtivo.tipoPessoa = New TipoPessoa
                                pessoaPartePoloAtivo.pertenceEmpresa = "0"
                                Dim grupoPessoas As New GrupoPessoa
                                Dim tiPessoa As New TipoPessoa
                                Dim estaOab As New Estado
                                Dim estado As New Estado
                                Dim nacio As New Nacionalidade
                                Dim prof As New Profissao
                                Dim pais As New Pais
                                pessoaPartePoloAtivo.grupoPessoa = grupoPessoas
                                pessoaPartePoloAtivo.tipoPessoa = tiPessoa
                                pessoaPartePoloAtivo.pais = pais
                                pessoaPartePoloAtivo.estadoOab = estaOab
                                pessoaPartePoloAtivo.estado = estado
                                pessoaPartePoloAtivo.nacionalidade = nacio
                                pessoaPartePoloAtivo.profissao = prof
                                pessoaPartePoloAtivo.update()
                            End If
                        End If
                        partePoloAtivo.pessoa = pessoaPartePoloAtivo
                        partePoloAtivo.processo = processo
                        partePoloAtivo.empresa = usuario.empresa
                        'partePoloAtivo.tipoParte = tipoPartePoloAtivo
                        partePoloAtivo.tipoParte = tipoPartePoloAtivo

                        partePoloAtivo.insert()
                        k += 1
                    End If

                    If pessoaPartePoloPassivo.nome <> "" Then
                        partePoloPassivo.planoPrevidencia = New PlanoPrevidencia
                        If (Not identificadorPessoasPassivas Is Nothing) Then
                            If identificadorPessoasPassivas(j).Trim.Length = 14 Then
                                pessoaPartePoloPassivo.cpf = identificadorPessoasPassivas(j).Trim
                                pessoaPartePoloAtivo.pertenceEmpresa = "1"
                                pessoaPartePoloPassivo.tipoPessoa = New TipoPessoa
                                pessoaPartePoloPassivo.pertenceEmpresa = "0"
                                Dim grupoPessoas As New GrupoPessoa
                                Dim tiPessoa As New TipoPessoa
                                Dim pais As New Pais
                                Dim estaOab As New Estado
                                Dim estado As New Estado
                                Dim nacio As New Nacionalidade
                                Dim prof As New Profissao
                                pessoaPartePoloPassivo.grupoPessoa = grupoPessoas
                                pessoaPartePoloPassivo.tipoPessoa = tiPessoa
                                pessoaPartePoloPassivo.pais = pais
                                pessoaPartePoloPassivo.estadoOab = estaOab
                                pessoaPartePoloPassivo.estado = estado
                                pessoaPartePoloPassivo.nacionalidade = nacio
                                pessoaPartePoloPassivo.profissao = prof
                                pessoaPartePoloPassivo.update()
                            Else
                                pessoaPartePoloPassivo.cnpj = identificadorPessoasPassivas(j).Trim
                                pessoaPartePoloAtivo.pertenceEmpresa = "2"
                                pessoaPartePoloPassivo.tipoPessoa = New TipoPessoa
                                pessoaPartePoloPassivo.pertenceEmpresa = "0"
                                Dim grupoPessoas As New GrupoPessoa
                                Dim tiPessoa As New TipoPessoa
                                Dim pais As New Pais
                                Dim estaOab As New Estado
                                Dim estado As New Estado
                                Dim nacio As New Nacionalidade
                                Dim prof As New Profissao
                                pessoaPartePoloPassivo.grupoPessoa = grupoPessoas
                                pessoaPartePoloPassivo.tipoPessoa = tiPessoa
                                pessoaPartePoloPassivo.pais = pais
                                pessoaPartePoloPassivo.estadoOab = estaOab
                                pessoaPartePoloPassivo.estado = estado
                                pessoaPartePoloPassivo.nacionalidade = nacio
                                pessoaPartePoloPassivo.profissao = prof
                                pessoaPartePoloPassivo.update()

                            End If
                        End If
                        partePoloPassivo.pessoa = pessoaPartePoloPassivo
                        partePoloPassivo.processo = processo
                        partePoloPassivo.empresa = usuario.empresa
                        partePoloPassivo.tipoParte = tipoPartePoloPassivo
                        partePoloPassivo.insert()
                        'Gravar evento de inclusão no log de atualizações (1 - Inclusão, 2 - parte)
                        logAtualizacao.gerarLog(usuario, "1", "2", processo.identificador, "PAR_PARTE", processo.identificador, "1", pessoaPartePoloAtivo.identificador)
                        j += 1

                    End If

                End If
            Next

            'Insere advogados e as partes quando não segue o padrão reclamante x requerido
            If Not (pessoaPartePoloAtivo.nome <> "" Or pessoaPartePoloPassivo.nome <> "") And nvcResultadoIntegracao("AdvogadosGeral") IsNot Nothing And nvcResultadoIntegracao("TipoPartesGeral") IsNot Nothing And nvcResultadoIntegracao("NomeGeral") IsNot Nothing Then
                nomePartesGeral = nvcResultadoIntegracao("NomeGeral").Split(",")
                advogadosGeral = nvcResultadoIntegracao("AdvogadosGeral").Split(",")
                tipoPartes = nvcResultadoIntegracao("TipoPartesGeral").Split(",")
                If nvcResultadoIntegracao("NomeGeral") <> "" Then
                    For index = 0 To nomePartesGeral.Length - 1

                        'Tipo de Parte
                        tipoDeParte.descricao = tipoPartes.GetValue(index).trim
                        tipoDeParte.descricao.Trim()
                        If tipoDeParte.descricao <> "" Then
                            tipoDeParte.empresa = usuario.empresa
                            tipoDeParte.identificador = tipoDeParte.getIdentificadorTipoParte()
                            If tipoDeParte.identificador = "" Then
                                tipoDeParte.status = "1"
                                tipoDeParte.insert()
                                tipoDeParte.identificador = tipoDeParte.getIdentificadorTipoParte()
                            End If

                        End If

                        'Partes
                        If nomePartesGeral.GetValue(index).trim <> "" Then
                            pessoaParteGeral.empresa = usuario.empresa
                            pessoaParteGeral.status = "1"
                            pessoaParteGeral.nome = nomePartesGeral.GetValue(index)
                            pessoaParteGeral.nome = pessoaParteGeral.nome.Trim
                            If pessoaParteGeral.getIdentificadorPessoa() <> "" And pessoaParteGeral.nome <> "" Then
                                pessoaParteGeral.identificador = pessoaParteGeral.getIdentificadorPessoa
                                parteGeral.pessoa = pessoaParteGeral
                                parteGeral.planoPrevidencia = New PlanoPrevidencia
                                parteGeral.processo = processo
                                parteGeral.tipoParte = tipoDeParte
                                parteGeral.empresa = usuario.empresa
                                parteGeral.insert()
                            ElseIf pessoaParteGeral.nome <> "" Then
                                pessoaParteGeral.pertenceEmpresa = "0"
                                Dim grupoPessoas As New GrupoPessoa
                                Dim tiPessoa As New TipoPessoa
                                Dim pais As New Pais
                                Dim estaOab As New Estado
                                Dim estado As New Estado
                                Dim nacio As New Nacionalidade
                                Dim prof As New Profissao
                                pessoaParteGeral.grupoPessoa = grupoPessoas
                                pessoaParteGeral.tipoPessoa = tiPessoa
                                pessoaParteGeral.pais = pais
                                pessoaParteGeral.estadoOab = estaOab
                                pessoaParteGeral.estado = estado
                                pessoaParteGeral.nacionalidade = nacio
                                pessoaParteGeral.profissao = prof
                                pessoaParteGeral.identificador = pessoaParteGeral.insert()
                                parteGeral.planoPrevidencia = New PlanoPrevidencia
                                parteGeral.pessoa = pessoaParteGeral
                                parteGeral.processo = processo
                                parteGeral.tipoParte = tipoDeParte
                                parteGeral.empresa = usuario.empresa
                                parteGeral.insert()
                            End If
                        End If

                        'Advogados
                        advogadosGerais = advogadosGeral(index).Trim.Split(",")
                        For Each item As String In advogadosGerais
                            pessoaAdvogado.nome = item.Trim
                            If pessoaAdvogado.nome <> "NÃO CADASTRADO" And pessoaAdvogado.nome <> "" Then
                                pessoaAdvogado.empresa = usuario.empresa
                                pessoaAdvogado.status = "1"
                                If pessoaAdvogado.getIdentificadorPessoa() <> "" Then
                                    advogado.nome = pessoaAdvogado.nome
                                    pessoaAdvogado.identificador = pessoaAdvogado.getIdentificadorPessoa()
                                    advogado.pessoa = pessoaAdvogado
                                    advogado.processo = processo
                                    advogado.processo.identificador = processo.identificador
                                    tipoPartePoloAtivo = New TipoParte
                                    tipoPartePoloAtivo.descricao = tipoPartes.GetValue(index).trim
                                    tipoPartePoloAtivo.empresa = usuario.empresa
                                    tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                                    advogado.tipoParte = tipoPartePoloAtivo
                                    advogado.insert()
                                Else
                                    pessoaAdvogado.pertenceEmpresa = "0"
                                    Dim grupoPessoas As New GrupoPessoa
                                    Dim tiPessoa As New TipoPessoa
                                    Dim pais As New Pais
                                    Dim estaOab As New Estado
                                    Dim estado As New Estado
                                    Dim nacio As New Nacionalidade
                                    Dim prof As New Profissao
                                    pessoaAdvogado.grupoPessoa = grupoPessoas
                                    pessoaAdvogado.tipoPessoa = tiPessoa
                                    pessoaAdvogado.pais = pais
                                    pessoaAdvogado.estadoOab = estaOab
                                    pessoaAdvogado.estado = estado
                                    pessoaAdvogado.nacionalidade = nacio
                                    pessoaAdvogado.profissao = prof
                                    pessoaAdvogado.status = "1"
                                    pessoaAdvogado.empresa = usuario.empresa
                                    pessoaAdvogado.identificador = pessoaAdvogado.insert()
                                    advogado.nome = pessoaAdvogado.nome
                                    advogado.pessoa = pessoaAdvogado
                                    advogado.pessoa.identificador = pessoaAdvogado.identificador
                                    advogado.processo = processo
                                    advogado.empresa = usuario.empresa
                                    tipoPartePoloAtivo = New TipoParte
                                    advogado.tipoParte = tipoDeParte
                                    advogado.tipoParte.status = "1"
                                    advogado.tipoParte.empresa = usuario.empresa
                                    advogado.tipoParte.identificador = advogado.tipoParte.getIdentificadorTipoParte()
                                    advogado.insert()
                                End If

                            End If
                        Next



                    Next
                End If

            End If

            Dim nvcPartesETiposProcess As DataTable
            Dim nvcPartesETiposProcessAtivo As DataTable
            Dim nvcPartesETiposProcessPassivo As DataTable

            If Not partePoloAtivo.tipoPoloAtivo Is Nothing And (drpSistemaExterno.SelectedValue = 98) Then

                nvcPartesETiposProcess = partePoloAtivo.list(usuario).Table
                nvcPartesETiposProcessAtivo = partePoloAtivo.listPartes(usuario, "1").Table
                'Dim partesProcessPassivo = partePoloAtivo.listPartes(usuario, "0").Table
                nvcPartesETiposProcessPassivo = partePoloAtivo.listPartes(usuario, "0").Table
                Dim tipoAtivo As String = ""
                Dim tipoPassivo As String = ""

                For Each value As DataRow In nvcPartesETiposProcessAtivo.Rows
                    Dim tipo As String = value("Tipo").ToString
                    Dim tipoPassivos As String = ""
                    If tipo <> "" Then
                        linhaQtdAtivos += 1
                    End If
                    For Each value2 As DataRow In nvcPartesETiposProcessPassivo.Rows
                        If value2("Identificador").ToString <> "0" Then
                            tipoPassivos = value2("Tipo").ToString
                            If tipo.Equals(tipoPassivos) Then
                                linhaQtdAtivos += 1
                            Else
                                linhaQtdPassivos += 1
                            End If
                        End If

                    Next

                Next

                For Each value As DataRow In nvcPartesETiposProcess.Rows
                    Dim contador As Integer
                    If usuario.empresa.tipoPoloAtivo = "1" Then

                        If value("Pólo Ativo").ToString.Equals("Sim") Then
                            If tituloParteAtivo = "" Then
                                tituloParteAtivo = value("Nome").ToString
                            End If
                            tipoAtivo = value("Tipo").ToString
                        End If

                        If value("Pólo Ativo").ToString.Equals("Não") Then
                            If tituloPartePassivo = "" Then
                                tituloPartePassivo = value("Nome").ToString
                            End If
                            tipoPassivo = value("Tipo").ToString
                        End If
                    Else
                        If contador = 0 Then
                            contador = nvcPartesETiposProcess.Rows.Count
                        End If
                        If contador = nvcPartesETiposProcess.Rows.Count Then
                            tituloParteAtivo = value("Nome").ToString
                        End If
                        If contador < nvcPartesETiposProcess.Rows.Count Then
                            tituloPartePassivo = value("Nome").ToString
                        End If

                    End If
                    contador = contador - 1
                Next

                'Título
                If tituloParteAtivo <> "" And tituloPartePassivo <> "" Then

                    If linhaQtdAtivos > 1 Then
                        tituloParteAtivo &= " e Outros"
                    End If

                    If linhaQtdPassivos > 1 Then
                        tituloPartePassivo &= " e Outros"
                    End If
                    processo.titulo = tituloParteAtivo.Trim & " X " & tituloPartePassivo.Trim
                Else
                    If tituloPartePassivo <> "" Then
                        processo.titulo = tituloPartePassivo.Trim
                    Else
                        processo.titulo = tituloParteAtivo.Trim
                    End If
                End If

            End If

            If nvcResultadoIntegracao("TituloParteAtiva") <> "" Then
                If nvcResultadoIntegracao("TituloPartePassiva") <> "" Then
                    processo.titulo = nvcResultadoIntegracao("TituloParteAtiva") + " X " + nvcResultadoIntegracao("TituloPartePassiva")
                Else
                    processo.titulo = nvcResultadoIntegracao("TituloParteAtiva")
                End If
            ElseIf nvcResultadoIntegracao("TituloPartePassiva") <> "" Then
                processo.titulo = nvcResultadoIntegracao("TituloPartePassiva")
            End If

            processo.tipoResultado = New TipoResultado
            processo.contrato = New Contrato
            processo.servico = New Servico

            'Rito
            rito = New Rito
            If nvcResultadoIntegracao("Rito") <> "" Then
                rito.descricao = nvcResultadoIntegracao("Rito")
                rito.empresa = usuario.empresa
                nvcRito = rito.getRito()
                If nvcRito.Count > 0 Then
                    rito.identificador = nvcRito("isn_rito")
                    rito.empresa = usuario.empresa
                    processo.rito = rito
                Else
                    rito.status = 1
                    rito.empresa = usuario.empresa
                    rito.insert()
                    nvcRito = rito.getRito()
                    rito.identificador = nvcRito("isn_rito")
                    processo.rito = rito
                End If
            End If

            'Fase
            fase = New Fase
            If nvcResultadoIntegracao("Fase") <> "" Then
                fase.descricao = nvcResultadoIntegracao("Fase")
                fase.empresa = usuario.empresa
                nvcFase = fase.getFase()
                If nvcFase.Count > 0 Then
                    fase.identificador = nvcFase("isn_fase")
                    fase.empresa = usuario.empresa
                    processo.fase = fase
                Else
                    fase.status = 1
                    fase.empresa = usuario.empresa
                    fase.insert()
                    nvcFase = fase.getFase()
                    fase.identificador = nvcFase("isn_fase")
                    processo.fase = fase
                End If
            End If

            'Retirado geração de número de pasta virtual conforme erro reportado em 26/01 referente ao mesmo.
            'If usuario.empresa.geraNumPasta = 1 Then
            '    processo.numeroPasta = processo.identificador
            'End If
            'Texto de Observação
            processo.textoObservacao = New Texto
            processo.textoObservacao.descricao = processo.observacao

            processo.update()

            'Advogados
            Dim contadorAdvo = 0
            aDvogadosPartesAtivos = nvcResultadoIntegracao("Advogados").Split(",")
            If Not nvcResultadoIntegracao("identificadorAdvogadosAtivos") Is Nothing And nvcResultadoIntegracao("identificadorAdvogadosAtivos") <> "" Then
                identificadorAdvAtivo = nvcResultadoIntegracao("identificadorAdvogadosAtivos").Split("&")
            End If

            If nvcResultadoIntegracao("Advogados") <> "" Then
                For index = 0 To aDvogadosPartesAtivos.Length - 2
                    pessoaAdvogado.nome = aDvogadosPartesAtivos(index).Trim("&").Trim
                    pessoaAdvogado.empresa = usuario.empresa
                    pessoaAdvogado.status = 1
                    If pessoaAdvogado.nome <> "" Then
                        If pessoaAdvogado.getIdentificadorPessoa() <> "" Then
                            advogado.nome = pessoaAdvogado.nome
                            advogado.pessoa = pessoaAdvogado
                            advogado.pessoa.identificador = pessoaAdvogado.getIdentificadorPessoa
                            advogado.processo = processo
                            advogado.empresa = usuario.empresa
                            tipoPartePoloAtivo = New TipoParte
                            tipoPartePoloAtivo.empresa = usuario.empresa
                            tipoPartesAtivos = nvcResultadoIntegracao("tipoParteTipoAtivo").Split(",")
                            tipoPartePoloAtivo.descricao = tipoPartesAtivos(0)
                            tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                            advogado.tipoParte = tipoPartePoloAtivo
                            advogado.tipoParte.identificador = tipoPartePoloAtivo.identificador
                            advogado.insert()
                        Else
                            pessoaAdvogado.pertenceEmpresa = "0"
                            Dim grupoPessoas As New GrupoPessoa
                            Dim tiPessoa As New TipoPessoa
                            Dim pais As New Pais
                            Dim estaOab As New Estado
                            Dim estado As New Estado
                            Dim nacio As New Nacionalidade
                            Dim prof As New Profissao
                            pessoaAdvogado.grupoPessoa = grupoPessoas
                            pessoaAdvogado.tipoPessoa = tiPessoa
                            pessoaAdvogado.pais = pais
                            pessoaAdvogado.estadoOab = estaOab
                            pessoaAdvogado.estado = estado
                            pessoaAdvogado.nacionalidade = nacio
                            pessoaAdvogado.profissao = prof
                            pessoaAdvogado.insert()
                            advogado.nome = pessoaAdvogado.nome.Trim("&")
                            advogado.pessoa = pessoaAdvogado
                            advogado.pessoa.identificador = pessoaAdvogado.identificador
                            advogado.processo = processo
                            tipoPartePoloAtivo = New TipoParte
                            tipoPartesAtivos = nvcResultadoIntegracao("tipoParteTipoAtivo").Split(",")
                            advogado.tipoParte = tipoPartePoloAtivo
                            advogado.tipoParte.descricao = tipoPartesAtivos(0)
                            advogado.tipoParte.empresa = usuario.empresa
                            advogado.tipoParte.identificador = advogado.tipoParte.getIdentificadorTipoParte()
                            advogado.insert()

                        End If
                        'Insere o número da Oab
                        If Not identificadorAdvAtivo Is Nothing Then
                            pessoaAdvogado.oab = identificadorAdvAtivo(contadorAdvo)
                            pessoaAdvogado.tipoPessoa = New TipoPessoa
                            pessoaAdvogado.pertenceEmpresa = "0"
                            Dim grupoPessoas As New GrupoPessoa
                            Dim tiPessoa As New TipoPessoa
                            Dim pais As New Pais
                            Dim estaOab As New Estado
                            Dim estado As New Estado
                            Dim nacio As New Nacionalidade
                            Dim prof As New Profissao
                            pessoaAdvogado.grupoPessoa = grupoPessoas
                            pessoaAdvogado.tipoPessoa = tiPessoa
                            pessoaAdvogado.pais = pais
                            pessoaAdvogado.estadoOab = estaOab
                            pessoaAdvogado.estado = estado
                            pessoaAdvogado.nacionalidade = nacio
                            pessoaAdvogado.profissao = prof
                            pessoaAdvogado.update()
                            contadorAdvo += 1
                        End If

                    End If

                Next

            End If

            contadorAdvo = 0
            aDvogadosPartesPassivos = nvcResultadoIntegracao("AdvogadosRequerido").Split(",")
            If Not nvcResultadoIntegracao("identificadorAdvogadosPassivos") Is Nothing And nvcResultadoIntegracao("identificadorAdvogadosPassivos") <> "" Then
                identificadorAdvPassivo = nvcResultadoIntegracao("identificadorAdvogadosPassivos").Split("&")
            End If
            If nvcResultadoIntegracao("AdvogadosRequerido") <> "" Then
                For index = 0 To aDvogadosPartesPassivos.Length - 2
                    pessoaAdvogado.nome = aDvogadosPartesPassivos(index).Trim("&").Trim
                    pessoaAdvogado.empresa = usuario.empresa
                    If pessoaAdvogado.nome <> "" Then
                        If pessoaAdvogado.getIdentificadorPessoa() <> "" Then
                            advogado.nome = pessoaAdvogado.nome
                            advogado.pessoa = pessoaAdvogado
                            advogado.pessoa.identificador = pessoaAdvogado.getIdentificadorPessoa
                            advogado.processo = processo
                            advogado.empresa = usuario.empresa
                            tipoPartePoloAtivo = New TipoParte
                            tipoPartePoloAtivo.empresa = usuario.empresa
                            tipoPartesAtivos = nvcResultadoIntegracao("tipoParteTipoAtivo").Split(",")
                            tipoPartePoloAtivo.descricao = tipoPartesAtivos(tipoPartesAtivos.Count - 1)
                            tipoPartePoloAtivo.identificador = tipoPartePoloAtivo.getIdentificadorTipoParte()
                            advogado.tipoParte = tipoPartePoloAtivo
                            advogado.tipoParte.identificador = tipoPartePoloAtivo.identificador
                            advogado.insert()
                        Else
                            pessoaAdvogado.pertenceEmpresa = "0"
                            Dim grupoPessoas As New GrupoPessoa
                            Dim tiPessoa As New TipoPessoa
                            Dim pais As New Pais
                            Dim estaOab As New Estado
                            Dim estado As New Estado
                            Dim nacio As New Nacionalidade
                            Dim prof As New Profissao
                            pessoaAdvogado.grupoPessoa = grupoPessoas
                            pessoaAdvogado.tipoPessoa = tiPessoa
                            pessoaAdvogado.pais = pais
                            pessoaAdvogado.estadoOab = estaOab
                            pessoaAdvogado.estado = estado
                            pessoaAdvogado.nacionalidade = nacio
                            pessoaAdvogado.profissao = prof
                            pessoaAdvogado.empresa = usuario.empresa
                            pessoaAdvogado.insert()
                            advogado.nome = pessoaAdvogado.nome.Trim("&")
                            advogado.pessoa = pessoaAdvogado
                            advogado.pessoa.identificador = pessoaAdvogado.identificador
                            advogado.processo = processo
                            tipoPartePoloAtivo = New TipoParte
                            tipoPartesAtivos = nvcResultadoIntegracao("tipoParteTipoPassivo").Split(",")
                            advogado.tipoParte = tipoPartePoloAtivo
                            advogado.tipoParte.descricao = tipoPartesAtivos(0)
                            advogado.tipoParte.empresa = usuario.empresa
                            advogado.tipoParte.identificador = advogado.tipoParte.getIdentificadorTipoParte()
                            advogado.insert()

                        End If

                        'Insere o número da Oab
                        If Not identificadorAdvPassivo Is Nothing Then
                            pessoaAdvogado.oab = identificadorAdvPassivo(contadorAdvo)
                            pessoaAdvogado.tipoPessoa = New TipoPessoa
                            pessoaAdvogado.pertenceEmpresa = "0"
                            Dim grupoPessoas As New GrupoPessoa
                            Dim tiPessoa As New TipoPessoa
                            Dim pais As New Pais
                            Dim estaOab As New Estado
                            Dim estado As New Estado
                            Dim nacio As New Nacionalidade
                            Dim prof As New Profissao
                            pessoaAdvogado.grupoPessoa = grupoPessoas
                            pessoaAdvogado.tipoPessoa = tiPessoa
                            pessoaAdvogado.pais = pais
                            pessoaAdvogado.estadoOab = estaOab
                            pessoaAdvogado.estado = estado
                            pessoaAdvogado.nacionalidade = nacio
                            pessoaAdvogado.profissao = prof
                            pessoaAdvogado.update()
                            contadorAdvo += 1
                        End If

                    End If

                Next

            End If

        End If

        '---------------------------------------
        'Instância
        '---------------------------------------
        'Comarca
        comarca.descricao = IIf(nvcResultadoIntegracao("Comarca") <> "", nvcResultadoIntegracao("Comarca"), "Fortaleza")
        comarca.empresa = usuario.empresa
        comarca.identificador = comarca.getIdentificadorComarca
        If comarca.identificador <> "" Then
            instancia.comarca = comarca
        Else
            comarca.estado = New Estado
            comarca.status = 1
            comarca.insert()
            comarca.identificador = comarca.getIdentificadorComarca
            instancia.comarca = comarca
        End If


        'Orgao
        orgao.descricao = nvcResultadoIntegracao("Orgao")
        'orgao.identificador = orgao.

        'grupo instancia
        grupoInstancia.empresa = usuario.empresa
        grupoInstancia.identificador = grupoInstancia.getTipoInstanciaPadrao
        instancia.grupoInstancia = grupoInstancia
        instancia.data = processo.dataDistribuicao
        'processo
        instancia.processo = processo

        'numero
        numero.identificador = IIf(nvcResultadoIntegracao("Numero") <> "", nvcResultadoIntegracao("Numero"), "")
        If numero.identificador.Length >= 1 Then
            instancia.numero = numero
        Else
            instancia.numero = New Numero
        End If

        'vara
        vara.descricao = nvcResultadoIntegracao("Vara")
        vara.empresa = usuario.empresa
        vara.identificador = vara.getIdentificadorVara
        If vara.identificador <> "" Then
            instancia.vara = vara
        Else
            vara.status = "1"
            vara.identificador = vara.insert
            instancia.vara = vara
        End If


        'tipo vara
        tipoVara.descricao = nvcResultadoIntegracao("TipoVara")
        If tipoVara.descricao = "" Then
            tipoVara.descricao = nvcResultadoIntegracao("Area")
        End If

        tipoVara.empresa = usuario.empresa
        tipoVara.identificador = tipoVara.getIdentificadorTipoVara
        If tipoVara.identificador <> "" Then
            instancia.tipoVara = tipoVara
        Else
            If tipoVara.descricao <> "" Then
                tipoVara.status = "1"
                tipoVara.identificador = tipoVara.insert
                instancia.tipoVara = tipoVara
            Else
                instancia.tipoVara = tipoVara
            End If
        End If

        'Número processo
        instancia.numeroProcesso = IIf(nvcResultadoIntegracao("numeroProcesso") <> "", nvcResultadoIntegracao("numeroProcesso"), txtNumPro.Text.Trim)

        'Parâmetro de Última Instância
        If usuario.empresa.implementaInstanciaOrigem = "1" Then
            instancia.indicadorInstanciaOrigem = "1"
        End If

        'Parâmetro de Última Instância
        If usuario.empresa.indicadorUltimaInstancia = "1" Then
            instancia.indicadorUltimaInstancia = "1"
        End If

        'Tipo Magistrado
        tipoMagistrado.descricao = nvcResultadoIntegracao("tipoMagistrado")
        tipoMagistrado.empresa = usuario.empresa
        tipoMagistrado.identificador = instancia.getTipoMagistrado(tipoMagistrado.descricao, usuario.empresa.pessoa.identificador)
        'Insere o tipo de magistrado caso não haja no banco e não seja nada
        If tipoMagistrado.identificador <> "" Then
            instancia.tipoMagistrado = tipoMagistrado
        Else
            If tipoMagistrado.descricao <> "" Then
                tipoMagistrado.status = "1"
                tipoMagistrado.insert()
                tipoMagistrado.identificador = instancia.getTipoMagistrado(tipoMagistrado.descricao, usuario.empresa.pessoa.identificador)
                instancia.tipoMagistrado = tipoMagistrado
            Else
                instancia.tipoMagistrado = New TipoMagistrado
            End If
        End If

        'Magistrado
        If instancia.tipoMagistrado.identificador <> "" Then
            pessoaRelator.nome = nvcResultadoIntegracao("Magistrado")
            pessoaRelator.empresa = usuario.empresa
            pessoaRelator.identificador = pessoaRelator.getIdentificadorPessoa()
            If pessoaRelator.identificador = "" And pessoaRelator.nome <> "" Then
                pessoaRelator.profissao = New Profissao
                pessoaRelator.nacionalidade = New Nacionalidade
                pessoaRelator.estadoOab = New Estado
                pessoaRelator.estado = New Estado
                pessoaRelator.pais = New Pais
                pessoaRelator.grupoPessoa = New GrupoPessoa
                pessoaRelator.tipoPessoa = New TipoPessoa
                pessoaRelator.grupoTrabalho = New GrupoTrabalho
                pessoaRelator.status = "1"
                pessoaRelator.identificador = pessoaRelator.insert()
            End If
            If pessoaRelator.identificador <> "" Then
                instancia.magistrado = pessoaRelator
            Else
                instancia.magistrado = New Pessoa
            End If
        Else
            instancia.magistrado = New Pessoa
        End If


        'Sistema externo
        instancia.sistemaExterno = New SistemaExterno
        instancia.sistemaExterno.identificador = drpSistemaExterno.SelectedValue
        instancia.opcaoSistemaExterno = New OpcaoSistemaExterno
        instancia.opcaoSistemaExterno.codigo = drpCodOpcao.SelectedValue

        'Correção de erro referente a processos indisponibilizados para consulta na página de integração
        If instancia.opcaoSistemaExterno.codigo <> "" Then
            instancia.opcaoSistemaExterno.codigo = Nothing
        End If

        instancia.opcaoSistemaAdicional = New OpcaoSistemaAdicional
        instancia.tipoResultado = New TipoResultado
        instancia.tipoParte = New TipoParte
        instancia.orgao = New Orgao
        instancia.tipoRecurso = New TipoRecurso
        instancia.revisor = New Pessoa

        'Insere a opcao do sistema externo pela comarca para ser possível a integração
        If instancia.opcaoSistemaExterno.codigo Is Nothing And instancia.sistemaExterno.identificador = 57 Then
            instancia.opcaoSistemaExterno.codigo = instancia.opcaoSistemaExterno.opcaoSistemaExternoPorComarca(instancia.comarca, instancia.sistemaExterno)
        End If

        If processo.identificador <> "" Then
            instancia.insert()
            If nvcResultadoIntegracao("numeroProcesso") <> "" And Not (nvcResultadoIntegracao("numeroProcessoSegundoGrau") Is Nothing Or Not nvcResultadoIntegracao("numeroProcessoSegundoGrau") <> "") Then
                instancia.numeroProcesso = nvcResultadoIntegracao("numeroProcessoSegundoGrau")
                instancia.data = nvcResultadoIntegracao("dtDistribuicaoSegundoGrau")
                instancia.insert()
            End If
        End If

        '---------------------------------------
        'Andamentos
        '---------------------------------------
        Select Case instancia.sistemaExterno.identificador

            Case 57
                resultadoIntegracao = processo.integrarEventos(usuario.empresa.eventoAFazer, instancia, usuario, "", "", "", sessionX, contadorIntegracao, contadorSessao)

            Case 98
                If instancia.identificador <> "" Then
                    resultadoIntegracaoEventos = processo.integrarEventos(usuario.empresa.eventoAFazer, instancia, usuario, "", "", "", sessionX, contadorIntegracao, contadorSessao)
                End If

            Case 109
                resultadoIntegracao = processo.integrarEventos(usuario.empresa.eventoAFazer, instancia, usuario, "", "", "", sessionX, contadorIntegracao, contadorSessao)

                'Case 133
                '    Dim DataView As DataView = instancia.listIntegracao()

                '    If Not DataView Is Nothing Then
                '        For i As Integer = 0 To DataView.Table.Rows.Count - 1
                '            instancia.identificador = DataView.Table.Rows(i).Item(0).ToString()
                '            resultadoIntegracao = processo.integrarEventos(usuario.empresa.eventoAFazer, instancia, usuario, "", "", "", sessionX, contadorIntegracao, contadorSessao)
                '        Next
                '    End If

        End Select

        retorno = processo.identificador
        Return retorno
    End Function

    Public Shared Function GetParamValor(ByVal texto As String, ByVal inicio As String, ByVal fim As String) As String

        texto = texto.Replace(vbCr, "").Replace(vbLf, "")
        Dim pattern As String = Regex.Escape(inicio) & "(.*?)" & Regex.Escape(fim)
        Dim m As Match = Regex.Match(texto, pattern)

        Return m.Groups(1).ToString 'retorna o primeiro parametro achado

    End Function

    Protected Sub btnCapturar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCapturar.Click

        Dim resultadoIntegracao As New NameValueCollection
        Dim isnProcesso As String = ""
        Dim identificadorTribunal As String = ""
        Dim lysisUtil = New LysisUtil
        Dim numeroProcesso = ""

        'Verifica se o número de processos excede a quantidade contratada
        If Not usuario.verificarNumeroProcesso Then
            lblMensagem.Text = "O número do processos excede a quantidade contratada."
            lblMensagem.Visible = True
            Exit Sub
        End If

        'Verifica formato CNJ se checkbox estiver selecionada
        If ckbMascaraCNJ.Checked.Equals(True) And ckbMascaraCNJ.Visible Then

            If txtNumProCNJ.Text.Trim.Equals("") Then

                lblMensagem.Visible = True
                lblMensagem.Text = "Você deve informar o número do processo"
                Exit Sub

            End If

            If Not lysisUtil.verificarNumeroCNJ(txtNumProCNJ.Text) Then

                lblMensagem.Visible = True
                lblMensagem.Text = "Número do processo não está de acordo com o formato CNJ."
                Exit Sub

            End If

        End If

        If drpSistemaExterno.SelectedIndex > 0 Then

            Try
                ''Verifica formato quando checkbox não estiver selecionada
                'If ckbMascaraCNJ.Checked.Equals(False) Then

                '    If Not validarFormatoProcesso() And drpSistemaExterno.SelectedIndex = 0 Then
                '        lblMensagem.Visible = True
                '        lblMensagem.Text = "Formato do número do processo incorreto, selecione um sistema para integração ou coloque o número do processo em formato correto."
                '        Return
                '    Else

                '        If Not validarFormatoProcesso() Then
                '            lblMensagem.Visible = True
                '            lblMensagem.Text = "Formato do número do processo incorreto, por favor coloque o número do processo em formato correto."
                '            Return
                '        Else
                '            If txtNumPro.Text = "" And drpSistemaExterno.SelectedIndex = 0 Then
                '                lblMensagem.Visible = True
                '                lblMensagem.Text = "Você deve informar o número do processo ou escolher um sistema para integração."
                '                Return
                '            Else
                '                lblMensagem.Visible = False
                '                lblMensagem.Text = ""
                '            End If
                '        End If

                '    End If
                'End If

                If ckbMascaraCNJ.Checked.Equals(True) And ckbMascaraCNJ.Visible Then
                    txtNumPro.Text = txtNumProCNJ.Text
                End If

                'Sistema externo
                sistemaExterno.identificador = drpSistemaExterno.SelectedValue

                'Verificação de existência de processo antigo do STF
                If Char.IsLetter(txtNumPro.Text(0)) Then
                    processo.identificador = RetornaProcessoAnterior(txtNumPro.Text)
                Else
                    processo.identificador = RetornaProcesso(txtNumPro.Text)
                End If

                If processo.identificador <> "" Then
                    lblMensagem.Visible = True
                    lblMensagem.Text = "Já existe um registro no Sistema."
                    Return
                Else
                    isnProcesso = CapturarPorNumProcesso(txtNumPro.Text.Trim)
                    If isnProcesso <> "" Then
                        Response.Redirect("../Paginas/ProcessoEditar.aspx?Isn=" & isnProcesso, False)
                    Else
                        lblMensagem.Visible = True
                        lblMensagem.Text = "Não foi possível capturar o processo."
                        Return
                    End If
                End If

            Catch ex As Exception
                Session("erro") = ex.Message
                Response.Redirect("../Paginas/Erro.aspx")
            End Try

        Else
            lblMensagem.Text = "Selecione um sistema para captura."
            lblMensagem.Visible = True
        End If

    End Sub

    Protected Sub txtNumPro_TextChanged(sender As Object, e As System.EventArgs) Handles txtNumPro.TextChanged
        Dim numProcesso As String

        Try
            numProcesso = txtNumPro.Text.ToString().Trim
            'If numProcesso.Length = 25 Then
            '    numProcesso.ToCharArray()

            '    Select Case numProcesso(16)

            '        'Case "4"
            '        '    Select Case numProcesso(18)
            '        '        Case "0"
            '        '            Select Case numProcesso(19)
            '        '                Case "5"
            '        '                    drpSistemaExterno.SelectedValue = 109
            '        '                    drpSistemaExterno_SelectedIndexChanged(sender, e)
            '        '                    If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '        '                        drpCodOpcao.SelectedIndex = 0
            '        '                    End If
            '        '            End Select
            '        '    End Select

            '        Case "5"
            '            Select Case numProcesso(18)
            '                Case "0"
            '                    Select Case numProcesso(19)
            '                        Case "7"
            '                            drpSistemaExterno.SelectedValue = 2
            '                            drpSistemaExterno_SelectedIndexChanged(sender, e)
            '                            If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '                                drpCodOpcao.SelectedIndex = 1
            '                            End If
            '                    End Select
            '            End Select

            '        Case "8"
            '            Select Case numProcesso(18)

            '                Case "0"
            '                    Select Case numProcesso(19)
            '                        Case "6"
            '                            drpSistemaExterno.SelectedValue = 98
            '                            drpSistemaExterno_SelectedIndexChanged(sender, e)
            '                            If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '                                drpCodOpcao.SelectedIndex = 0
            '                            End If
            '                    End Select

            '                Case "1"
            '                    Select Case numProcesso(19)
            '                        Case "2"
            '                            drpSistemaExterno.SelectedValue = 185
            '                            drpSistemaExterno_SelectedIndexChanged(sender, e)
            '                            If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '                                drpCodOpcao.SelectedIndex = 0
            '                            End If
            '                        Case "8"
            '                            drpSistemaExterno.SelectedValue = 57
            '                            drpSistemaExterno_SelectedIndexChanged(sender, e)
            '                            If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '                                drpCodOpcao.SelectedIndex = 0
            '                            End If
            '                    End Select
            '            End Select
            '    End Select

            'End If
            'Transformado em Select Case
            '    If numProcesso(16) = "5" Then
            '        If numProcesso(18) = "0" Then
            '            If numProcesso(19) = "7" Then
            '                drpSistemaExterno.SelectedValue = 2
            '                drpSistemaExterno_SelectedIndexChanged(sender, e)
            '                If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '                    drpCodOpcao.SelectedIndex = 1
            '                End If

            '            End If
            '        End If
            '    ElseIf numProcesso(16) = "8" Then
            '        If numProcesso(18) = "1" Then
            '            If numProcesso(19) = "8" Then
            '                drpSistemaExterno.SelectedValue = 57
            '                drpSistemaExterno_SelectedIndexChanged(sender, e)
            '                If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '                    drpCodOpcao.SelectedIndex = 0
            '                End If
            '            End If
            '        End If
            '    ElseIf numProcesso(16) = "4" Then
            '        If numProcesso(18) = "0" Then
            '            If numProcesso(19) = "5" Then
            '                drpSistemaExterno.SelectedValue = 109
            '                drpSistemaExterno_SelectedIndexChanged(sender, e)
            '                If Not drpCodOpcao.Items.FindByValue("1") Is Nothing Then
            '                    drpCodOpcao.SelectedIndex = 0
            '                End If
            '            End If
            '        End If
            '    End If
            'End If





        Catch ex As Exception
            Throw

        End Try


    End Sub

    Private Sub CapturarPorOab()

        Dim numProcesso As String = ""
        Dim arrProcessoNovo() As String

        numProcesso = processo.numeroProcessos.Replace("'", "")
        arrProcessoNovo = numProcesso.Split(",")

        For Each numProcessoUnico In arrProcessoNovo

            Try

                Capturar(numProcessoUnico, False)
                listaRetorno.Add(New KeyValuePair(Of String, String)(numProcessoUnico, "Processo capturado com êxito."))

            Catch ex As IntegracaoException

                listaRetorno.Add(New KeyValuePair(Of String, String)(numProcessoUnico, ex.Message))


            Catch ex As Exception

                listaRetorno.Add(New KeyValuePair(Of String, String)(numProcessoUnico, "Erro ao capturar processo."))

            End Try

        Next

        exibeResultados(listaRetorno)

    End Sub

    Private Function CapturarPorNumProcesso(numeroProcesso As String) As String

        Dim logIntegracao As New LogIntegracao
        Dim processo As New Processo
        Dim instancia As New Instancia
        Dim sistemaExterno As New SistemaExterno With {.identificador = drpSistemaExterno.SelectedValue}

        Try

            Return Capturar(txtNumPro.Text, True)

        Catch ex As IntegracaoTribunais.Estrutura.Exceptions.IntegracaoException
            Dim mensagemErroInnerExceptionStackTrace = "Exception | Message: " & ex.Message & " - Stack Trace:" & ex.StackTrace

            lblMensagem.Text = ex.Message
            lblMensagem.Visible = True

            logIntegracao.sistema.identificador = sistemaExterno.identificador
            logIntegracao.processo = New Processo
            logIntegracao.instancia = New Instancia
            logIntegracao.empresa = usuario.empresa
            logIntegracao.empresa.pessoa = usuario.empresa.pessoa
            logIntegracao.empresa.pessoa.identificador = usuario.empresa.pessoa.identificador
            logIntegracao.erro = 11
            logIntegracao.descricao = "Captura de Processo"
            logIntegracao.descricaoComplemento = mensagemErroInnerExceptionStackTrace
            logIntegracao.insert()

        Catch ex As Exception
            lblMensagem.Text = "Ocorreu um erro ao tentar capturar o processo."
            lblMensagem.Visible = True
        End Try

        Return ""
    End Function

    Private Function Capturar(numeroProcesso As String, capturaUnica As Boolean) As String

        Dim numProcRetorno = numeroProcesso
        Dim logAtualizacao As LogAtualizacao = New LogAtualizacao
        Dim logIntegracao As LogIntegracao = New LogIntegracao
        Dim enderecoWebService = System.Configuration.ConfigurationManager.AppSettings("enderecoWebService")

        Dim sistemaExterno As New SistemaExterno With {.identificador = drpSistemaExterno.SelectedValue}
        Dim sistemaExternoCodOpcao = drpCodOpcao.SelectedValue

        Dim processoCaptura = New ProcessoCaptura
        Dim login = System.Text.UTF8Encoding.UTF8.GetBytes("webserviceLysis:%PsQ45R")
        Dim documentoString = ""

        Dim listLoginAdvogado = processo.listAcessoAdvogado(sistemaExterno.identificador, usuario.empresa)

        If listLoginAdvogado.Count > 0 Then

            capturaUnica = False

        End If

        If Not capturaUnica And sistemaExterno.possuiAcessoAdvogado And listLoginAdvogado.Count > 0 Then

            Dim msgErro = ""

            For Each acesso In listLoginAdvogado

                Dim requestAdvogado As HttpWebRequest = Nothing
                Dim responseAdvogado As HttpWebResponse = Nothing

                requestAdvogado = HttpWebRequest.Create(enderecoWebService & "/api/captura?TipoCaptura=Numero&NumeroProcesso=" &
                                                        numeroProcesso & "&SistemaExterno=" & drpSistemaExterno.SelectedValue &
                                                        "&SistemaExternoCodOpcao=" & drpCodOpcao.SelectedValue & "&idCaptcha=" & idCaptcha &
                                                        "&Cookie=" & cookie & "&Login=" & acesso.Item("Login").ToString & "&Senha=" & acesso.Item("Senha").ToString)

                requestAdvogado.Method = "GET"
                requestAdvogado.PreAuthenticate = True
                requestAdvogado.Headers.Add("Authorization", "Bearer " & System.Convert.ToBase64String(login))
                requestAdvogado.Accept = "application/Xml"
                requestAdvogado.Headers.Add(HttpRequestHeader.Cookie, "AspxAutoDetectCookieSupport=1")
                requestAdvogado.Timeout = 600000

                Try

                    responseAdvogado = requestAdvogado.GetResponse

                    Dim rd = New StreamReader(responseAdvogado.GetResponseStream(), Encoding.UTF8)
                    documentoString = rd.ReadToEnd.Trim

                    msgErro = ""
                    Exit For

                Catch ex As WebException

                    Dim rd = New StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8)
                    msgErro = rd.ReadToEnd.Trim

                    If msgErro.Contains("erro no tratamento do Captcha") Or msgErro.Contains("Erro ao tentar requisitar o Captcha") Then
                        cookie = ""
                    End If

                Catch ex As Exception

                    msgErro = ex.Message

                End Try

            Next

            If msgErro <> "" Then

                If msgErro.StartsWith("Integracao Exception") Then

                    Throw New IntegracaoException(IntegracaoUtil.GetParamValor(msgErro, "Message: ", "-").Trim, IntegracaoUtil.GetParamValor(msgErro, "Stack Trace:").Trim)

                Else

                    Throw New Exception(msgErro)

                End If

            End If

        Else

            'Requisição para acesso ao webservice de integração
            Dim request As HttpWebRequest = Nothing
            Dim response As HttpWebResponse = Nothing

            request = HttpWebRequest.Create(enderecoWebService & "/api/captura?TipoCaptura=Numero&NumeroProcesso=" & numeroProcesso & "&SistemaExterno=" & drpSistemaExterno.SelectedValue & "&SistemaExternoCodOpcao=" & drpCodOpcao.SelectedValue & "&idCaptcha=" & idCaptcha & "&Cookie=" & cookie)
            request.Method = "GET"
            request.PreAuthenticate = True
            request.Headers.Add("Authorization", "Bearer " & System.Convert.ToBase64String(login))
            request.Accept = "application/Xml"
            request.Headers.Add(HttpRequestHeader.Cookie, "AspxAutoDetectCookieSupport=1")
            request.Timeout = 600000

            Try

                response = request.GetResponse

            Catch ex As WebException

                Dim reader = New StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8)
                Dim mensagemErro = reader.ReadToEnd.Trim

                If mensagemErro.StartsWith("Integracao Exception") Then

                    Throw New IntegracaoException(IntegracaoUtil.GetParamValor(mensagemErro, "Message: ", "-").Trim, IntegracaoUtil.GetParamValor(mensagemErro, "Stack Trace:").Trim)

                Else

                    Throw New Exception(mensagemErro)

                End If

            End Try

            Dim rd = New StreamReader(response.GetResponseStream(), Encoding.UTF8)
            documentoString = rd.ReadToEnd.Trim
        End If

        If documentoString.StartsWith("<?xml") Or documentoString.StartsWith("<html") Then

            Dim sb As New StringBuilder
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Encoding = Encoding.Unicode
            settings.Indent = True

            Using reader As XmlReader = XmlReader.Create(New StringReader(documentoString))

                While (reader.Read)

                    Select Case (reader.NodeType)
                        Case XmlNodeType.Element
                            Select Case (reader.Name)
                                Case "idCaptcha"
                                    reader.Read()
                                    idCaptcha = reader.Value
                                Case "cookie"
                                    reader.Read()
                                    cookie = reader.Value
                                Case "numeroProcesso"
                                    reader.Read()
                                    processoCaptura.NumeroProcesso = reader.Value.Trim
                                Case "titulo"
                                    reader.Read()
                                    processoCaptura.Titulo = reader.Value.Trim
                                Case "numeroProcessoAdicional"
                                    reader.Read()
                                    processoCaptura.NumeroProcessoAdicional = reader.Value.Trim
                                Case "numeroProcessoAnterior"
                                    reader.Read()
                                    processoCaptura.NumeroProcessoAnterior = reader.Value.Trim
                                Case "dataDistribuicao"
                                    reader.Read()
                                    processoCaptura.DataDistribuicao = reader.Value.Trim
                                Case "classe"
                                    reader.Read()
                                    processoCaptura.Classe = reader.Value.Trim
                                Case "comarca"
                                    reader.Read()
                                    processoCaptura.Comarca = reader.Value.Trim
                                Case "materia"
                                    reader.Read()
                                    processoCaptura.Materia = reader.Value.Trim
                                Case "natureza"
                                    reader.Read()
                                    processoCaptura.Natureza = reader.Value.Trim
                                Case "fase"
                                    reader.Read()
                                    processoCaptura.Fase = reader.Value.Trim
                                Case "tipoMagistrado"
                                    reader.Read()
                                    processoCaptura.TipoMagistrado = reader.Value.Trim
                                Case "nomeMagistrado"
                                    reader.Read()
                                    processoCaptura.NomeMagistrado = reader.Value.Trim
                                Case "numeroVara"
                                    reader.Read()
                                    processoCaptura.NumeroVara = reader.Value.Trim
                                Case "vara"
                                    reader.Read()
                                    processoCaptura.Vara = reader.Value.Trim
                                Case "tipoVara"
                                    reader.Read()
                                    processoCaptura.TipoVara = reader.Value.Trim
                                Case "observacao"
                                    reader.Read()
                                    processoCaptura.Observacao = reader.Value.Trim
                                Case "orgao"
                                    reader.Read()
                                    processoCaptura.Orgao = reader.Value.Trim
                                Case "processoEletronico"
                                    reader.Read()
                                    Dim processoEletronico = reader.Value.Trim

                                    If processoEletronico.Equals("True") Then
                                        processoCaptura.ProcessoEletronico = True
                                    Else
                                        processoCaptura.ProcessoEletronico = False
                                    End If
                                Case "revisor"
                                    reader.Read()
                                    processoCaptura.Revisor = reader.Value.Trim
                                Case "rito"
                                    reader.Read()
                                    processoCaptura.Rito = reader.Value.Trim
                                Case "status"
                                    reader.Read()
                                    Dim status = reader.Value.Trim

                                    If status.Equals("True") Then
                                        processoCaptura.Status = True
                                    Else
                                        processoCaptura.Status = False
                                    End If
                                Case "valorDaCausa"
                                    reader.Read()
                                    processoCaptura.ValorDaCausa = reader.Value.Trim
                                Case "instancia"

                                    Dim instanciaCaptura = New InstanciaCaptura

                                    While (reader.Read)

                                        Select Case (reader.NodeType)
                                            Case XmlNodeType.Element
                                                Select Case (reader.Name)
                                                    Case "numeroProcesso"
                                                        reader.Read()
                                                        instanciaCaptura.NumeroProcesso = reader.Value.Trim
                                                    Case "sistemaExternoOpcao"
                                                        reader.Read()
                                                        instanciaCaptura.SistemaExternoOpcao = reader.Value.Trim
                                                    Case "dataAutuacao"
                                                        reader.Read()
                                                        instanciaCaptura.DataAutuacao = reader.Value.Trim
                                                End Select

                                            Case XmlNodeType.EndElement And reader.Name.Equals("instancia")
                                                processoCaptura.Instancias.Add(instanciaCaptura)
                                                Exit While
                                        End Select
                                    End While

                                Case "parte"

                                    Dim parteCaputra = New ParteCaptura

                                    While (reader.Read)

                                        Select Case (reader.NodeType)
                                            Case XmlNodeType.Element
                                                Select Case (reader.Name)
                                                    Case "nome"
                                                        reader.Read()
                                                        parteCaputra.Nome = reader.Value.Trim
                                                    Case "tipo"
                                                        reader.Read()
                                                        parteCaputra.Tipo = reader.Value.Trim
                                                    Case "cpf"
                                                        reader.Read()
                                                        parteCaputra.CPF = reader.Value.Trim
                                                    Case "rg"
                                                        reader.Read()
                                                        parteCaputra.RG = reader.Value.Trim
                                                    Case "orgaoExpedidor"
                                                        reader.Read()
                                                        parteCaputra.OrgaoExpedidor = reader.Value.Trim
                                                    Case "cnpj"
                                                        reader.Read()
                                                        parteCaputra.CNPJ = reader.Value.Trim
                                                End Select

                                            Case XmlNodeType.EndElement And reader.Name.Equals("parte")
                                                processoCaptura.Partes.Add(parteCaputra)
                                                Exit While
                                        End Select
                                    End While

                                Case "advogado"

                                    Dim advogadoCaptura = New AdvogadoCaptura

                                    While (reader.Read)

                                        Select Case (reader.NodeType)
                                            Case XmlNodeType.Element
                                                Select Case (reader.Name)
                                                    Case "nome"
                                                        reader.Read()
                                                        advogadoCaptura.Nome = reader.Value.Trim
                                                    Case "tipo"
                                                        reader.Read()
                                                        advogadoCaptura.Tipo = reader.Value.Trim
                                                    Case "oab"
                                                        reader.Read()
                                                        advogadoCaptura.OAB = reader.Value.Trim
                                                    Case "cpf"
                                                        reader.Read()
                                                        advogadoCaptura.CPF = reader.Value.Trim
                                                    Case "cnpj"
                                                        reader.Read()
                                                        advogadoCaptura.CNPJ = reader.Value.Trim
                                                End Select

                                            Case XmlNodeType.EndElement And reader.Name.Equals("advogado")
                                                processoCaptura.Advogados.Add(advogadoCaptura)
                                                Exit While
                                        End Select

                                    End While

                                Case "resultado"

                                    Dim resultadoCaptura = New ResultadoCaptura

                                    While (reader.Read)

                                        Select Case (reader.NodeType)
                                            Case XmlNodeType.Element
                                                Select Case (reader.Name)
                                                    Case "data"
                                                        reader.Read()
                                                        resultadoCaptura.Data = reader.Value.Trim
                                                    Case "tipo"
                                                        reader.Read()
                                                        resultadoCaptura.Tipo = reader.Value.Trim
                                                    Case "descricao"
                                                        reader.Read()
                                                        resultadoCaptura.Descricao = reader.Value.Trim
                                                End Select

                                            Case XmlNodeType.EndElement And reader.Name.Equals("resultado")
                                                processoCaptura.Resultado.Add(resultadoCaptura)
                                                Exit While
                                        End Select

                                    End While

                                Case "andamento"

                                    Dim andamentoCaptura = New AndamentoCaptura
                                    Dim anexoCaptura = New AnexoCaptura

                                    While (reader.Read)

                                        Select Case (reader.NodeType)
                                            Case XmlNodeType.Element
                                                Select Case (reader.Name)
                                                    Case "descricao"
                                                        reader.Read()
                                                        andamentoCaptura.Descricao = reader.Value.Trim
                                                    Case "data"
                                                        reader.Read()
                                                        andamentoCaptura.Data = reader.Value.Trim
                                                    Case "hora"
                                                        reader.Read()
                                                        andamentoCaptura.Hora = reader.Value.Trim
                                                    Case "instancia"
                                                        reader.Read()
                                                        andamentoCaptura.Instancia = reader.Value.Trim
                                                    Case "observacao"
                                                        reader.Read()
                                                        andamentoCaptura.Observacao = reader.Value.Trim
                                                    Case "anexo"
                                                        anexoCaptura = New AnexoCaptura
                                                    Case "arquivo"
                                                        reader.Read()
                                                        anexoCaptura.Arquivo = New IO.MemoryStream(Convert.FromBase64String(reader.Value))
                                                    Case "extensao"
                                                        reader.Read()
                                                        anexoCaptura.Extensao = reader.Value
                                                End Select
                                            Case XmlNodeType.EndElement And reader.Name.Equals("andamento")
                                                processoCaptura.Andamentos.Add(andamentoCaptura)
                                                Exit While
                                            Case XmlNodeType.EndElement And reader.Name.Equals("anexo")
                                                andamentoCaptura.Anexos.Add(anexoCaptura)
                                        End Select
                                    End While

                            End Select
                    End Select
                End While
            End Using

        ElseIf documentoString.StartsWith("Integracao Exception") Then

            Throw New IntegracaoException(IntegracaoUtil.GetParamValor(documentoString, "Message: ", "-").Trim, IntegracaoUtil.GetParamValor(documentoString, "Stack Trace:").Trim)

        Else

            Throw New Exception(documentoString)

        End If


        'parsers
        Dim lysisProcesso = ProcessoIntegracaoParser.ParseToProcesso(processoCaptura)
        Dim lysisInstancia As New Instancia
        Dim lysisListaIdentificadorInstancias As New List(Of String)
        Dim lysisPartes As List(Of Parte) = ProcessoIntegracaoParser.ParseToPartes(processoCaptura, usuario.empresa)
        Dim lysisAdvogados As List(Of Advogado) = ProcessoIntegracaoParser.ParseToAdvogados(processoCaptura, usuario.empresa)

        Dim lysisMateria As Materia = ProcessoIntegracaoParser.ParseToMateria(processoCaptura)
        Dim lysisEletronico As String = ProcessoIntegracaoParser.ParseToEletronico(processoCaptura)
        Dim lysisTipoAcao As TipoAcao = ProcessoIntegracaoParser.ParseToTipoAcao(processoCaptura)
        Dim lysisDataDistribuicao As String = ProcessoIntegracaoParser.ParseToDataDistribuicao(processoCaptura)
        Dim lysisValorDaCausa As String = ProcessoIntegracaoParser.ParseToValorDaCausa(processoCaptura)
        Dim lysisDataStatus As Date = ProcessoIntegracaoParser.ParseToDataStatus
        Dim lysisEmpresa As Empresa = usuario.empresa
        lysisEmpresa.geraNumPasta = "0"
        Dim lysisFase As Fase = ProcessoIntegracaoParser.ParseToFase(processoCaptura)
        Dim lysisRito As Rito = ProcessoIntegracaoParser.ParseToRito(processoCaptura)
        Dim lysisStatus As String = ProcessoIntegracaoParser.ParseToStatus(processoCaptura)

        Dim lysisResultado As List(Of Resultado) = ProcessoIntegracaoParser.ParseToResultados(processoCaptura, usuario.empresa)

        Dim lysisNumeroVara As Numero = ProcessoIntegracaoParser.ParseToNumeroVara(processoCaptura)
        Dim lysisVara As Vara = ProcessoIntegracaoParser.ParseToVara(processoCaptura)
        Dim lysisTipoVara As TipoVara = ProcessoIntegracaoParser.ParseToTipoVara(processoCaptura)
        Dim lysisComarca As Comarca = ProcessoIntegracaoParser.ParseToComarca(processoCaptura)
        Dim lysisTipoMagistrado As TipoMagistrado = ProcessoIntegracaoParser.ParseToTipoMagistrado(processoCaptura)
        Dim lysisMagistrado As Pessoa = ProcessoIntegracaoParser.ParseToMagistrado(processoCaptura, usuario.empresa)
        Dim lysisNumeroProcesso As String = ProcessoIntegracaoParser.ParseToNumeroProcesso(processoCaptura, txtNumPro.Text)
        Dim lysisNumeroProcessoAnterior As String = ProcessoIntegracaoParser.ParseToNumeroProcessoAnterior(processoCaptura)
        Dim lysisNumeroProcessoSemFormatoSemPrefixo As String = lysisNumeroProcesso.Replace("-", "").Replace(".", "")
        Dim lysisGrupoInstancia As New GrupoInstancia
        Dim lysisListaDadosInstancias As List(Of InstanciaCaptura) = ProcessoIntegracaoParser.ParseToListaInstanciaCaptura(processoCaptura)
        'tratamento extra
        'lysisInstancia.sistemaExterno = New SistemaExterno With {.identificador = sistemaAndamento}
        Dim lysisOrgao As Orgao = ProcessoIntegracaoParser.ParseToOrgao(processoCaptura)
        Dim lysisObjetos As List(Of ObjetoProcesso) = ProcessoIntegracaoParser.ParseToObjetos(processoCaptura)
        Dim lysisRevisor As Pessoa = ProcessoIntegracaoParser.ParseRevisor(processoCaptura, usuario.empresa)

        'Materia
        lysisMateria.empresa = lysisEmpresa
        Dim existeMateria As Boolean = Not String.IsNullOrEmpty(lysisMateria.descricao) Or Not String.IsNullOrWhiteSpace(lysisMateria.descricao)
        If existeMateria Then
            lysisMateria.identificador = lysisMateria.getIdentificadorMateria.ToString
            If lysisMateria.identificador = "0" Then
                lysisMateria = ProcessoIntegracaoInsercao.InsercaoMateria(lysisMateria, usuario)
            End If
        End If

        'Tipos de Resultados
        Dim resultados As New Resultado
        Dim isnTipoResultado = ""
        For Each resultado In lysisResultado

            resultados.tipo = resultado.tipo
            resultados.descricao = resultado.descricao
            resultados.data = resultado.data
            resultados.empresa = lysisEmpresa

            'resultados.empresa.pessoa.identificador = usuario.empresa.pessoa.identificador

            Dim resultadosarl = resultados.Consult()

            If resultadosarl Is Nothing Then

                resultados.insert()
                isnTipoResultado = resultados.getIdentificadorTipoResultado()

            Else

                isnTipoResultado = resultados.getIdentificadorTipoResultado()

            End If



        Next

        'Tipo Parte
        Dim tipoParte As New TipoParte
        Dim partesArl = tipoParte.listParte(usuario)
        Dim partesAdicionarLista As New List(Of String)
        Dim nomeParte(100)
        Dim posicao = 0
        Dim countNumParte1 = 0
        Dim countNumParte2 = 0
        For Each parte In lysisPartes
            Dim query = From parteInfo In partesArl
                        Where parteInfo("Nome").ToString.ToLower.Equals(parte.tipoParte.descricao.ToLower.Trim) And Not String.IsNullOrWhiteSpace(parte.tipoParte.descricao)
                        Select parteInfo("Nome")

            If query.Count = 0 Then
                partesAdicionarLista.Add(parte.tipoParte.descricao.ToLower.Trim)
            End If

            If lysisPartes.Count < 3 Then

                If query(0).ToString <> "" And nomeParte(posicao) <> query(0).ToString Then

                    If nomeParte(posicao) = "" Or nomeParte(posicao) = query(0).ToString Then

                        nomeParte(posicao) = query(0).ToString
                        countNumParte1 += 1

                    Else

                        nomeParte(posicao + 1) = query(0).ToString
                        countNumParte2 += 1

                    End If

                End If

            End If

        Next

        If (countNumParte1 = 1 And countNumParte2 = 1) Or (countNumParte1 = 1 And countNumParte2 = 0) Or (countNumParte1 = 0 And countNumParte2 = 1) Then

            For Each partes In lysisPartes

                If countNumParte1 = 1 And countNumParte2 = 1 Then

                    partes.tipoPrincipal = "1"

                ElseIf countNumParte1 = 1 Then

                    partes.tipoPrincipal = "1"
                    Exit For

                ElseIf countNumParte2 = 1 Then

                    If Not (lysisPartes(0).pessoa.nome = partes.pessoa.nome) Then

                        partes.tipoPrincipal = "1"
                        Exit For

                    End If

                End If

            Next

        End If

        Dim partesAdicionar = From parteDescricao In partesAdicionarLista
                              Select parteDescricao Distinct
        If partesAdicionar.Count > 0 Then
            ProcessoIntegracaoInsercao.InsercaoTipoParte(partesAdicionar, lysisEmpresa, usuario)
        End If

        'TipoVara
        lysisTipoVara.empresa = usuario.empresa
        Dim existeTipoVara = Not String.IsNullOrEmpty(lysisTipoVara.descricao) Or Not String.IsNullOrWhiteSpace(lysisTipoVara.descricao)
        If existeTipoVara Then
            lysisTipoVara.identificador = lysisTipoVara.getIdentificadorTipoVara
            If lysisTipoVara.identificador = "" Then
                lysisTipoVara = ProcessoIntegracaoInsercao.InsercaoTipoVara(lysisTipoVara, usuario)
            End If
        End If

        'NumeroVara
        lysisNumeroVara.empresa = lysisEmpresa
        If lysisNumeroVara.descricao <> "" Then
            lysisNumeroVara.identificador = lysisNumeroVara.getIdentificadorNumero
        End If

        'Vara
        lysisVara.empresa = usuario.empresa
        Dim existeVara = Not String.IsNullOrEmpty(lysisVara.descricao) Or Not String.IsNullOrWhiteSpace(lysisVara.descricao)
        If existeVara Then
            lysisVara.identificador = lysisVara.getIdentificadorVara
            If lysisVara.identificador = "" And lysisVara.descricao <> "" Then
                lysisVara = ProcessoIntegracaoInsercao.InsercaoVara(lysisVara, usuario)
            End If
        End If

        'Grupo Instancia
        lysisGrupoInstancia.empresa = usuario.empresa
        lysisGrupoInstancia.identificador = lysisGrupoInstancia.getTipoInstanciaPadrao


        'Partes 
        For Each parte In lysisPartes
            parte.tipoParte.empresa = lysisEmpresa
            parte.tipoParte.identificador = parte.tipoParte.getIdentificadorTipoParte
        Next

        'Advogados
        For Each advogado In lysisAdvogados
            advogado.tipoParte.empresa = lysisEmpresa
            advogado.tipoParte.identificador = advogado.tipoParte.getIdentificadorTipoParte
        Next

        'Tipo Acao
        lysisTipoAcao.empresa = lysisEmpresa
        Dim existeTipoAcao = Not String.IsNullOrEmpty(lysisTipoAcao.descricao) Or Not String.IsNullOrWhiteSpace(lysisTipoAcao.descricao)
        If existeTipoAcao Then
            If Not lysisTipoAcao.existeTipoAcao() Then
                lysisTipoAcao.materia = lysisMateria
                lysisTipoAcao = ProcessoIntegracaoInsercao.InsercaoTipoAcao(lysisTipoAcao, usuario)
            End If
        End If

        'Comarca
        lysisComarca.empresa = lysisEmpresa
        Dim existeComarca = Not String.IsNullOrEmpty(lysisComarca.descricao) Or Not String.IsNullOrWhiteSpace(lysisComarca.descricao)
        If existeComarca Then
            lysisComarca.identificador = lysisComarca.getIdentificadorComarca
            If lysisComarca.identificador = "" Then
                lysisComarca = ProcessoIntegracaoInsercao.InsercaoComarca(lysisComarca, usuario)
            End If
        End If

        'Tipo Magistrado
        lysisTipoMagistrado.empresa = lysisEmpresa
        Dim existeTipoMagistrado = Not String.IsNullOrEmpty(lysisTipoMagistrado.descricao) Or Not String.IsNullOrWhiteSpace(lysisTipoMagistrado.descricao)
        If existeTipoMagistrado Then
            lysisTipoMagistrado.identificador = lysisTipoMagistrado.gettipoMagistrado("isn_tipo_magistrado")
            If lysisTipoMagistrado.identificador = "" Then
                lysisTipoMagistrado = ProcessoIntegracaoInsercao.InsertTipoMagistrado(lysisTipoMagistrado, usuario)
            End If
        End If

        'Magistrado
        lysisMagistrado.empresa = lysisEmpresa
        Dim existeNomeMagistrado = Not String.IsNullOrEmpty(lysisMagistrado.nome) Or Not String.IsNullOrWhiteSpace(lysisMagistrado.nome)
        If existeNomeMagistrado Then
            If lysisMagistrado.identificador = "" Then
                lysisMagistrado.identificador = lysisMagistrado.getIdentificadorPessoa
            End If

            If lysisMagistrado.identificador = "" Then
                lysisMagistrado = ProcessoIntegracaoInsercao.InsertMagistrado(lysisMagistrado, usuario)
            End If
        End If

        'Revisor
        lysisRevisor.empresa = lysisEmpresa
        Dim existeNomeRevisor = Not String.IsNullOrEmpty(lysisRevisor.nome) Or Not String.IsNullOrWhiteSpace(lysisRevisor.nome)
        If existeNomeRevisor Then
            If lysisRevisor.identificador = "" Then
                lysisRevisor.identificador = lysisRevisor.getIdentificadorPessoa
            End If

            If lysisRevisor.identificador = "" Then
                lysisRevisor = ProcessoIntegracaoInsercao.InsertRevisor(lysisRevisor, usuario)
            End If
        End If

        'Fase
        lysisFase.empresa = lysisEmpresa
        Dim existeFase = Not String.IsNullOrEmpty(lysisFase.descricao) Or Not String.IsNullOrWhiteSpace(lysisFase.descricao)
        If existeFase Then
            lysisFase.identificador = lysisFase.getIdentificadorFase
            If lysisFase.identificador = "" Then
                lysisFase = ProcessoIntegracaoInsercao.InsercaoFase(lysisFase)
                lysisFase.identificador = lysisFase.getIdentificadorFase
            End If
        End If

        'Natureza
        'lysisNatureza.empresa = lysisEmpresa
        'Dim existeNatureza = Not String.IsNullOrEmpty(lysisNatureza.descricao) Or Not String.IsNullOrWhiteSpace(lysisNatureza.descricao)
        'If existeNatureza Then
        '    If lysisNatureza.identificador = "" Then
        '        lysisNatureza = ProcessoIntegracaoInsercao.InsercaoNatureza(lysisNatureza)
        '    End If
        'End If

        'Rito
        lysisRito.empresa = lysisEmpresa
        Dim existeRito = Not String.IsNullOrEmpty(lysisRito.descricao) Or Not String.IsNullOrWhiteSpace(lysisRito.descricao)
        If existeRito Then
            lysisRito.identificador = lysisRito.getIdentificadorRito
            If lysisRito.identificador = "" Then
                lysisRito = ProcessoIntegracaoInsercao.InsercaoRito(lysisRito)
            End If
        End If

        'Orgão
        lysisOrgao.empresa = lysisEmpresa
        Dim listaTipoOrgao = lysisOrgao.getTiposDeOrgaos(lysisEmpresa.pessoa.identificador)
        Dim existeOrgao = Not String.IsNullOrEmpty(lysisOrgao.descricao) Or Not String.IsNullOrWhiteSpace(lysisOrgao.descricao)
        If existeOrgao And listaTipoOrgao.Count > 0 Then
            lysisOrgao.identificador = lysisOrgao.getIdentificadorOrgao
            If lysisOrgao.identificador = "" Then
                lysisOrgao.tipoOrgao = listaTipoOrgao.First
                lysisOrgao = ProcessoIntegracaoInsercao.InsercaoOrgao(lysisOrgao)
            End If
        End If

        'Processo
        'Verificando se a parte mudou apos troca de sinonimos
        Dim countParte = 0
        Dim countParteContraria = 0
        Dim titulo = ""
        Dim mudouSinonimo As Boolean = False

        For index = 0 To lysisPartes.Count - 1
            If Not lysisPartes(index).pessoa.nome.Equals(processoCaptura.Partes(index).Nome) Then
                mudouSinonimo = True
                Exit For
            End If
        Next

        If lysisPartes.Any Then
            If mudouSinonimo Then 'Equals(processoCaptura.Partes) Then
                For Each parte1 In lysisPartes
                    If lysisPartes(0).tipoParte.Equals(parte1.tipoParte) Then
                        countParte = countParte + 1
                    Else
                        countParteContraria = countParteContraria + 1
                    End If
                Next
                If countParte >= 3 Then
                    titulo = lysisPartes(0).pessoa.nome & " e Outros"
                ElseIf countParte >= 2 Then
                    titulo = lysisPartes(0).pessoa.nome & " e Outro"
                Else
                    titulo = lysisPartes(0).pessoa.nome
                End If
                If countParteContraria >= 3 Then
                    titulo = titulo & " X " & lysisPartes(countParte).pessoa.nome & " e Outros"
                ElseIf countParteContraria >= 2 Then
                    titulo = titulo & " X " & lysisPartes(countParte).pessoa.nome & " e Outro"
                Else
                    titulo = titulo & " X " & lysisPartes(countParte).pessoa.nome
                End If
                lysisProcesso.titulo = titulo
            Else
                lysisProcesso.titulo = lysisProcesso.titulo
            End If
        End If


        lysisProcesso.status = lysisStatus
        lysisProcesso.materia = New Materia
        lysisProcesso.materia = lysisMateria
        lysisProcesso.tipoAcao = New TipoAcao
        lysisProcesso.tipoAcao = lysisTipoAcao
        lysisProcesso.eletronico = lysisEletronico
        lysisProcesso.estrategico = "0"
        lysisProcesso.dataDistribuicao = lysisDataDistribuicao
        lysisProcesso.valor = lysisValorDaCausa
        lysisProcesso.dataStatus = lysisDataStatus
        lysisProcesso.empresa = lysisEmpresa
        lysisProcesso.numeroPasta = ""
        lysisProcesso.centroCusto = New CentroCusto
        lysisProcesso.natureza = New Natureza
        lysisProcesso.rito = lysisRito
        lysisProcesso.fase = lysisFase
        lysisProcesso.tipoRisco = New TipoRisco
        lysisProcesso.escritorio = New Pessoa
        lysisProcesso.grupoProcesso = New GrupoProcesso
        lysisProcesso.probabilidadePerda = New Probabilidade
        lysisProcesso.probabilidade = New Probabilidade
        lysisProcesso.prioridadeDe = New Pessoa
        lysisProcesso.valorAtualizado = ""
        lysisProcesso.valorProvisionado = ""
        lysisProcesso.valorContingencia = ""
        lysisProcesso.cliente = New Pessoa
        lysisProcesso.numeroPasta = Nothing
        lysisProcesso.tipOrigemProcesso = "2"

        lysisProcesso.insert()
        logAtualizacao.gerarLog(usuario, "1", "1", lysisProcesso.identificador, "PRO_PROCESSO", lysisProcesso.identificador, "1")

        'Insere o Tipo de Resultado
        For Each tipoResultado In lysisResultado

            tipoResultado.identificador = isnTipoResultado
            tipoResultado.empresa = lysisEmpresa
            tipoResultado.processo = lysisProcesso
            tipoResultado.update()

        Next

        'Insere Partes
        For Each parte In lysisPartes
            parte.pessoa.empresa = lysisEmpresa
            parte.pessoa.status = "1"
            parte.pessoa.empresa.verificaDuplicidadeCPFCNPJ = "1"
            If parte.pessoa.identificador = "" Then
                parte.pessoa.identificador = parte.pessoa.getIdentificadorPessoa
            End If
            parte.empresa = lysisEmpresa
            parte.processo = lysisProcesso

            parte.planoPrevidencia = New PlanoPrevidencia
            ProcessoIntegracaoInsercao.InsercaoParte(parte, usuario)
        Next

        'Insere Advogados
        For Each advogado In lysisAdvogados
            advogado.pessoa.empresa = lysisEmpresa
            advogado.pessoa.status = "1"
            advogado.pessoa.empresa.verificaDuplicidadeCPFCNPJ = "1"
            If advogado.pessoa.identificador = "" Then
                advogado.pessoa.identificador = advogado.pessoa.getIdentificadorPessoa
            End If
            advogado.empresa = lysisEmpresa
            advogado.processo = lysisProcesso
            advogado.tipoParte.empresa = New Empresa
            advogado.tipoParte.empresa = lysisEmpresa
            advogado.tipoParte.identificador = advogado.tipoParte.getIdentificadorTipoParte

            ProcessoIntegracaoInsercao.InsercaoAdvogado(advogado, usuario)
        Next

        lysisInstancia.tipoVara = lysisTipoVara
        lysisInstancia.vara = lysisVara
        lysisInstancia.comarca = lysisComarca
        lysisInstancia.tipoMagistrado = lysisTipoMagistrado
        lysisInstancia.magistrado = lysisMagistrado
        lysisInstancia.orgao = lysisOrgao
        lysisInstancia.grupoInstancia = lysisGrupoInstancia
        lysisInstancia.data = lysisDataDistribuicao
        lysisInstancia.processo = lysisProcesso
        lysisInstancia.numero = lysisNumeroVara
        lysisInstancia.numeroProcesso = lysisNumeroProcesso
        lysisInstancia.numeroProcessoAnterior = lysisNumeroProcessoAnterior

        If usuario.empresa.implementaInstanciaOrigem = "1" Then
            lysisInstancia.indicadorInstanciaOrigem = "1"
        End If
        If usuario.empresa.indicadorUltimaInstancia = "1" Then
            lysisInstancia.indicadorUltimaInstancia = "1"
        End If

        'Se o tribunal do processo capturado for o TRT-18-PJE, ele recebe o sistema externo do tribunal TRT-18
        If sistemaExterno.identificador = 172 Then
            sistemaExterno.identificador = 88
        End If

        lysisInstancia.sistemaExterno = New SistemaExterno With {.identificador = sistemaExterno.identificador}
        lysisInstancia.opcaoSistemaExterno = New OpcaoSistemaExterno With {.codigo = sistemaExternoCodOpcao}

        If lysisInstancia.opcaoSistemaExterno.codigo.Trim = "" Then
            lysisInstancia.opcaoSistemaExterno.codigo = Nothing
        End If

        lysisInstancia.opcaoSistemaAdicional = New OpcaoSistemaAdicional
        lysisInstancia.tipoResultado = New TipoResultado
        lysisInstancia.tipoParte = New TipoParte
        lysisInstancia.tipoRecurso = New TipoRecurso
        lysisInstancia.revisor = lysisRevisor

        Dim existemInstancias = lysisListaDadosInstancias.Any
        If existemInstancias Then

            Dim dataUltimaInstancia = lysisListaDadosInstancias.OrderByDescending(Function(instancia) Date.Parse(instancia.DataAutuacao)).First.DataAutuacao
            For Each lysisDadoInstancia As InstanciaCaptura In lysisListaDadosInstancias
                If lysisDadoInstancia.DataAutuacao = dataUltimaInstancia And usuario.empresa.indicadorUltimaInstancia = "1" Then
                    lysisInstancia.indicadorUltimaInstancia = "1"
                Else
                    lysisInstancia.indicadorUltimaInstancia = "0"
                End If
                ProcessoIntegracaoInsercao.InsercaoInstancias(lysisInstancia, lysisProcesso, lysisDadoInstancia, usuario)
            Next

        Else
            Dim lysisDadoInstancia As New InstanciaCaptura With {
                        .DataAutuacao = lysisDataDistribuicao,
                        .numeroProcesso = lysisNumeroProcesso
                    }
            ProcessoIntegracaoInsercao.InsercaoInstancias(lysisInstancia, lysisProcesso, lysisDadoInstancia, usuario)
        End If

        'Associa publicações ao processo capturado        
        Dim eventoPublicacao = New Evento
        eventoPublicacao.numeroProcesso = numeroProcesso
        usuario.empresa.identificador = usuario.empresa.pessoa.identificador

        For Each isnPublicacao In eventoPublicacao.listPublicacaoPorNumeroProcesso(usuario)

            eventoPublicacao.identificador = isnPublicacao

            eventoPublicacao.processo = lysisProcesso
            eventoPublicacao.instancia = lysisInstancia

            eventoPublicacao.updatePublicacao()

            logAtualizacao = New LogAtualizacao
            logAtualizacao.gerarLog(usuario, "1", "5", eventoPublicacao.identificador, "ADA_ANDAMENTO", lysisProcesso.identificador, "1")
            'logAtualizacao.atualizaLog(lysisProcesso, usuario, "1", "6", eventoPublicacao.identificador, "")
        Next


        For Each objeto As ObjetoProcesso In lysisObjetos
            objeto.processo = lysisProcesso
            ProcessoIntegracaoInsercao.InsercaoObjeto(objeto)
        Next

        'remove duplicados
        processoCaptura.Andamentos = processoCaptura.Andamentos.Distinct.ToList

        'rotina para inserção no banco de dados
        For Each andamento As AndamentoCaptura In processoCaptura.Andamentos

            Dim evento As New Evento
            Dim textoDescricao As New Texto
            Dim textoComplemento As New Texto
            Dim textoObservacao As New Texto

            If andamento.IsDataValida Then
                evento.data = andamento.Data
            Else
                Continue For 'pula para o próximo andamento
            End If

            If andamento.IsHoraValida Then
                If andamento.Hora.Count > 5 Then
                    evento.hora = andamento.Hora.Trim.Remove(5)
                Else
                    evento.hora = andamento.Hora
                End If

            End If


            If Not String.IsNullOrWhiteSpace(andamento.Complemento) AndAlso andamento.Complemento.Length > 2000 Then
                evento.complemento = andamento.Complemento.Substring(0, 2000)
            End If

            evento.descricao = andamento.Descricao
            textoDescricao.descricao = evento.descricao
            textoDescricao.identificador = ""
            evento.textoDescricao = textoDescricao
            textoComplemento.descricao = ""
            textoComplemento.identificador = ""
            evento.textoComplemento = textoComplemento
            textoObservacao.descricao = ""
            textoObservacao.identificador = ""
            evento.textoObservacao = textoObservacao

            'Origem
            evento.origem = "2"

            'Cancelado
            evento.cancelado = "0"

            'Tarefa individual
            evento.tarefaIndividual = "0"

            'Número do processo
            evento.numeroProcesso = ""

            'Tipo privado
            evento.privado = ""

            'Indica se o evento será incluído no "A Fazer"
            evento.agenda = usuario.empresa.eventoAFazer

            'Tipo andamento
            evento.tipoAndamento = New TipoAndamento With {.empresa = usuario.empresa}

            'evento decorrente
            evento.geraDecorrencia = False

            'habilita execução da rotina de expressões
            evento.geraExpressao = True

            'Cria listas vazias para rotina de eventos decorrentes
            evento.ListaNotificacaoAndamento = New List(Of NotificacaoAndamento)
            evento.ListaResponsavelAndamento = New List(Of ResponsavelAndamento)
            evento.ListaGrupoAndamento = New List(Of GrupoAndamento)
            evento.ListaClienteAndamentoProcesso = New List(Of ClienteAndamentoProcesso)

            'Cliente
            evento.cliente = New Pessoa With {.identificador = ""}

            'Responsável
            evento.pessoa = New Pessoa With {.identificador = ""}

            'Solicitante
            evento.solicitante = New Pessoa With {.identificador = ""}

            'Processo
            evento.processo = lysisProcesso

            'Identificador instância
            'If existemInstancias Then
            '    lysisInstancia.identificador = lysisListaIdentificadorInstancias(andamento.Instancia)
            '    lysisInstancia.numeroProcesso = lysisListaDadosInstancias(andamento.Instancia).NumeroProcesso
            '    lysisInstancia.data = lysisListaDadosInstancias(andamento.Instancia).DataAutuacao
            '    evento.instancia = lysisInstancia
            'Else
            evento.instancia = lysisInstancia
            'End If

            'Exibe evento para o cliente
            If usuario.empresa.selecionaEventos = "1" And usuario.empresa.selecionaEventosMarcado = "1" Then
                evento.exibeCliente = "1"
            Else
                evento.exibeCliente = "0"
            End If

            'Empresa do evento
            evento.empresa = New Empresa With {.pessoa = New Pessoa With {.identificador = Nothing}}

            'Data da integração
            evento.dataIntegracao = Format(Date.Now, "dd/MM/yyyy HH:mm")

            'verifica período de integração
            evento.insert()

            'Insere documento referente ao evento
            Dim tamanhoAnexo As String = ""
            If andamento.Anexos.Count <> 0 Then

                For Each andamentoAnexo As AnexoCaptura In andamento.Anexos
                    Dim anexoEventoProcesso As New AnexoEventoProcesso

                    'Tipo de documento
                    anexoEventoProcesso.tipoDocumento = New TipoDocumento With {.identificador = ""}

                    'Evento
                    anexoEventoProcesso.evento = evento
                    anexoEventoProcesso.nomeAnexo = "AnexoDocumentoIntegracao" & andamentoAnexo.Extensao

                    anexoEventoProcesso.insert()

                    'Realizando tentativa de gravar no disco, caso não tenha êxito, deletar da base o anexo.
                    Try
                        Dim nomeFisicoAnexo As String = anexoEventoProcesso.nomeArquivoFisico()
                        Dim subPasta As String = obtemSubPasta(anexoEventoProcesso.identificador, "aap_anexo_andamento_processo", "isn_anexo_andamento_processo")

                        'Se implementa GoogleDrive, upload do arquivo no GoogleDrive

                        If usuario.empresa.tipoArmazenamentoExterno = 2 And usuario.empresa.idDiretorioDrive <> "0" Then

                            If (subPasta <> Nothing) Then
                                lysisUtil.UploadGoogleDriveIntegracao(andamentoAnexo.Arquivo, nomeFisicoAnexo, usuario.empresa.pessoa.identificador, subPasta)

                            Else
                                lysisUtil.UploadGoogleDriveIntegracao(andamentoAnexo.Arquivo, nomeFisicoAnexo, usuario.empresa.pessoa.identificador, usuario.empresa.idDiretorioDrive)

                            End If
                            tamanhoAnexo = andamentoAnexo.Arquivo.Length

                        Else

                            If (subPasta <> Nothing) Then
                                Using ioStream As New FileStream(diretorioDocumentos & usuario.empresa.pessoa.identificador & "\" & subPasta & "\" & nomeFisicoAnexo, FileMode.Create)
                                    tamanhoAnexo = andamentoAnexo.Arquivo.Length
                                    andamentoAnexo.Arquivo.WriteTo(ioStream)

                                    ioStream.Flush()
                                    ioStream.Close()

                                    andamentoAnexo.Arquivo.Flush()
                                    andamentoAnexo.Arquivo.Close()
                                End Using


                            Else
                                Using ioStream As New FileStream(diretorioDocumentos & usuario.empresa.pessoa.identificador & "\" & nomeFisicoAnexo, FileMode.Create)

                                    andamentoAnexo.Arquivo.WriteTo(ioStream)
                                    tamanhoAnexo = andamentoAnexo.Arquivo.Length
                                    ioStream.Flush()
                                    ioStream.Close()

                                    andamentoAnexo.Arquivo.Flush()
                                    andamentoAnexo.Arquivo.Close()
                                End Using

                            End If

                        End If
                        atualizarTamanhoArquivoEmDisco(anexoEventoProcesso.identificador, "aap_anexo_andamento_processo", "isn_anexo_andamento_processo", tamanhoAnexo)
                    Catch ex As Exception

                        anexoEventoProcesso.delete()

                    End Try
                Next
            End If

        Next

        logIntegracao.geraLogIntegracao(lysisProcesso, lysisInstancia, sistemaExterno, usuario, "Integração em Captura", "", "", processoCaptura.Andamentos.Count, processoCaptura.Andamentos.Count)

        Return lysisProcesso.identificador
    End Function
    Private Sub atualizarTamanhoArquivoEmDisco(ByVal identificadorDocumento As String, ByVal nomeTabela As String, ByVal identificador_entidade_tabela As String, ByVal tamanhoArquivo As String)
        Dim pastaArquivo As New PastaArquivo
        pastaArquivo.isnDocumento = identificadorDocumento
        pastaArquivo.numTamanhoArquivo = tamanhoArquivo
        pastaArquivo.updateNumTamanhoArquivo(nomeTabela, identificador_entidade_tabela)
    End Sub

    Private Sub btnConsultar_Click(sender As Object, e As EventArgs) Handles btnConsultar.Click

        Dim sistemaExterno = drpSistemaExterno.SelectedValue
        Dim sistemaExternoCodOpcao = drpCodOpcao.SelectedValue
        Dim opcaoCaptura = drpOpcaoCaptura.SelectedValue
        Dim numeroOab As String = ""
        Dim pessoaEmpresa As New Pessoa
        Dim empresa As New Empresa
        Dim estadoOab = drpIsnEstadoOab.SelectedItem.Text.Trim
        Dim enderecoWebService = System.Configuration.ConfigurationManager.AppSettings("enderecoWebService")


        Try
            'Empresa
            pessoaEmpresa.identificador = CType(usuario.empresa.pessoa.identificador, String)
            empresa.pessoa = pessoaEmpresa
            processo.empresa = empresa

            'Validação dos campos
            If txtNumOAB.Text = "" Then
                lblMensagem.Text = "Número OAB deve ser informado."
                lblMensagem.Visible = True
            ElseIf drpSistemaExterno.SelectedIndex <= 0 Then
                lblMensagem.Text = "Sistema deve ser informado."
                lblMensagem.Visible = True
            Else

                'Se algum parâmetro de pesquisa foi alterado, a variavel de 
                'Sessão "processosCapturados" recebe Nothing
                If Session("numOab") <> txtNumOAB.Text Or
                   Session("sistemaExterno") <> drpSistemaExterno.SelectedValue Or
                   Session("codOpcao") <> drpCodOpcao.SelectedValue Or
                   Session("estadoOab") <> drpIsnEstadoOab.SelectedValue Or
                   Session("complementoOab") <> drpIsnComplementoOab.SelectedValue Or
                   Session("opcaoCaptura") <> drpOpcaoCaptura.SelectedValue Or
                   Session("opcaoAdicionalCaptura") <> drpOpcaoAdicionalCaptura.SelectedValue Then

                    Session("processosCapturados") = Nothing
                End If

                'Se a variavel de Sessão "processosCapturados" for Nothing, é realizado 
                ' a consulta e a busca no site e o preenchimento da mesma

                If Session("processosCapturados") Is Nothing Then

                    'Requisição para acesso ao webservice de integração
                    Dim request As HttpWebRequest = Nothing
                    Dim response As HttpWebResponse = Nothing
                    Dim complementoOAB = ""

                    If drpIsnComplementoOab.SelectedValue <> "-1" Then
                        complementoOAB = drpIsnComplementoOab.SelectedValue
                    End If
                    Dim listLoginAdvogado = processo.listAcessoAdvogado(sistemaExterno, usuario.empresa)
                    Dim count = 0
                    Dim acesso As Object = Nothing
                    For Each acessoVerificacaoAdv In listLoginAdvogado
                        acesso = acessoVerificacaoAdv
                        If count = 0 Then
                            Exit For
                        End If
                    Next

                    If listLoginAdvogado.Count = 0 Then
                        request = HttpWebRequest.Create(enderecoWebService & "/api/captura?TipoCaptura=OAB&SistemaExterno=" &
                                                    sistemaExterno & "&SistemaExternoCodOpcao=" & sistemaExternoCodOpcao &
                                                    "&NumeroOab=" & txtNumOAB.Text & "&EstadoOab=" & estadoOab & "&ComplementoOab=" &
                                                    complementoOAB & "&OpcaoCaptura=" & opcaoCaptura &
                                                    "&OpcaoAdicionalCaptura=" & drpOpcaoAdicionalCaptura.SelectedValue)
                    Else
                        request = HttpWebRequest.Create(enderecoWebService & "/api/captura?TipoCaptura=OAB&SistemaExterno=" &
                                                    sistemaExterno & "&SistemaExternoCodOpcao=" & sistemaExternoCodOpcao &
                                                    "&NumeroOab=" & txtNumOAB.Text & "&EstadoOab=" & estadoOab & "&ComplementoOab=" &
                                                    complementoOAB & "&OpcaoCaptura=" & opcaoCaptura &
                                                    "&OpcaoAdicionalCaptura=" & drpOpcaoAdicionalCaptura.SelectedValue & "&Login=" &
                                                    acesso.Item("Login").ToString & "&Senha=" & acesso.Item("Senha").ToString)

                    End If

                    request.Method = "GET"
                    request.PreAuthenticate = True
                    Dim login = System.Text.UTF8Encoding.UTF8.GetBytes("webserviceLysis:%PsQ45R")
                    request.Headers.Add("Authorization", "Bearer " & System.Convert.ToBase64String(login))
                    request.Accept = "application/Xml"
                    request.Headers.Add(HttpRequestHeader.Cookie, "AspxAutoDetectCookieSupport=1")
                    request.Timeout = 600000

                    Try

                        response = request.GetResponse
                        Dim processos = New List(Of ProcessoCaptura)

                        Dim rd = New StreamReader(response.GetResponseStream(), Encoding.UTF8)
                        Dim documentoString = rd.ReadToEnd.Trim

                        Dim sb As New StringBuilder
                        Dim settings As XmlWriterSettings = New XmlWriterSettings()
                        settings.Encoding = Encoding.Unicode
                        settings.Indent = True

                        Using reader As XmlReader = XmlReader.Create(New StringReader(documentoString))

                            While (reader.Read)

                                If reader.NodeType.Equals(XmlNodeType.Element) And reader.Name.Equals("numeroProcesso") Then

                                    reader.Read()
                                    processos.Add(New ProcessoCaptura With {.NumeroProcesso = reader.Value})

                                End If

                            End While

                        End Using

                        'Realizando o parse dos processos e os consultando na base
                        Dim lysisProcessos As List(Of Processo) = ProcessoIntegracaoParser.ParseToProcessos(processos)

                        If lysisProcessos.Count = 0 Then Throw New IntegracaoException("O sistema não encontrou nenhum processo disponível.")

                        For Each processosOab In lysisProcessos
                            If listaDeProcessos Is Nothing Then
                                listaDeProcessos &= "'" & processosOab.numeroProcessos & "'"
                            Else
                                listaDeProcessos &= "," & "'" & processosOab.numeroProcessos & "'"
                            End If
                        Next

                        Session.Add("processosCapturados", listaDeProcessos)
                        Session.Add("numOab", txtNumOAB.Text)
                        Session.Add("sistemaExterno", drpSistemaExterno.SelectedValue)
                        Session.Add("codOpcao", drpCodOpcao.SelectedValue)
                        Session.Add("estadoOab", drpIsnEstadoOab.SelectedValue)
                        Session.Add("complementoOab", drpIsnComplementoOab.SelectedValue)
                        Session.Add("opcaoCaptura", drpOpcaoCaptura.SelectedValue)
                        Session.Add("opcaoAdicionalCaptura", drpOpcaoAdicionalCaptura.SelectedValue)

                        'Montagem do grid com os processos capturados do site
                        processo.numeroProcessos = listaDeProcessos
                        dtgProcesso.DataSource = processo.listCapturaOab()
                        dtgProcesso.DataBind()
                        dtgProcesso.Visible = True

                        If dtgProcesso.DataSource Is Nothing Then
                            lkbCapturarProcesso.Visible = False
                        Else
                            lkbCapturarProcesso.Visible = True
                        End If

                    Catch ex As WebException

                        Dim rd = New StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8)
                        Dim mensagemErro = rd.ReadToEnd.Trim

                        Throw New Exception(IntegracaoUtil.GetParamValor(mensagemErro, "Message: ", " -"))
                    End Try

                    'Se a variavel de Sessão "processosCapturados" já estiver preenchida
                    'não é necessário realizar uma nova busca no site
                Else

                    'Montagem do grid com a lista de processos da variavel de sessao
                    processo.numeroProcessos = Session("processosCapturados")
                    dtgProcesso.DataSource = processo.listCapturaOab()
                    dtgProcesso.DataBind()
                    dtgProcesso.Visible = True

                    If dtgProcesso.DataSource Is Nothing Then
                        lkbCapturarProcesso.Visible = False
                    Else
                        lkbCapturarProcesso.Visible = True
                    End If

                End If

            End If

        Catch ex As Exception
            'Caso não haja processos para o número de OAB informado, é exibido a exceção lançada pela biblioteca.
            lkbCapturarProcesso.Visible = False
            dtgProcesso.Visible = False
            lblMensagem.Text = "Não foi possível listar processos vinculados ao número de OAB. Verifique os parâmetros da busca ou tente novamente." 'ex.Message
            lblMensagem.Visible = True
        End Try

    End Sub

    Private Sub dtgProcesso_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles dtgProcesso.RowCreated
        If e.Row.RowType <> DataControlRowType.Pager Then

            e.Row.Cells(3).Visible = False


        End If
    End Sub

    Private Sub dtgProcesso_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles dtgProcesso.RowDataBound



        funcoes = CType(Session("Funcoes"), String)

        If e.Row.RowType <> DataControlRowType.Pager Then
            e.Row.Cells(0).Width = 25
        End If


        If e.Row.RowType = DataControlRowType.DataRow Then

            'Definindo coluna (0) como checkbox
            If e.Row.Cells(0).Text = "1" Then
                chkItem = CType(e.Row.FindControl("chkItem"), CheckBox)
                If Not IsNothing(chkItem) Then
                    chkItem.Checked = True
                End If
            End If

            'Tratamento: Se tal processo já existe no sistema, não é possível capturá-lo
            For coluna As Integer = 0 To 3

                If e.Row.Cells(2).Text = "" Or e.Row.Cells(2).Text = "&nbsp;" Then
                    e.Row.Cells(coluna).Enabled = True
                Else
                    e.Row.Cells(0).Enabled = False
                    e.Row.Attributes.Add("onclick", "javascript:document.location='../Paginas/ProcessoEditar.aspx?Isn=" & e.Row.Cells(3).Text & "';")
                    e.Row.Attributes.Add("onmouseover", "this.style.cursor='pointer'")
                End If

            Next
        End If


        If e.Row.RowType = DataControlRowType.Header Then

            'Tratamento de seleção múltipla do checkbox
            chkCabecalho = CType(e.Row.FindControl("chkCabecalho"), CheckBox)

            If Not IsNothing(chkCabecalho) Then
                chkCabecalho.Attributes.Add("onclick", "selecionarLinhas(this, 'dtgProcesso');")
            End If

        End If



    End Sub

    Private Sub lkbCapturarProcesso_Click(sender As Object, e As EventArgs) Handles lkbCapturarProcesso.Click

        Try

            Dim isnProcesso As String = ""
            Dim qtdeProcessos As Integer = 0
            Dim qtdeProcessosIntegrados As Integer = 1
            Dim listaDeProcessosDoGrid As New List(Of String)
            Dim listaDeProcessosASeremCapturados As New List(Of String)

            'Verifica se o número de processos excede a quantidade contratada
            If Not usuario.verificarNumeroProcesso Then
                lblMensagem.Text = "O número do processos excede a quantidade contratada."
                lblMensagem.Visible = True
                Exit Sub
            End If

            'Contar os processos selecionados para captura, para exibição do progresso
            For Each Me.dtgRow In dtgProcesso.Rows
                If dtgRow.RowType = DataControlRowType.DataRow Then

                    chkItem = DirectCast(dtgRow.FindControl("chkItem"), CheckBox)

                    If chkItem.Checked Then
                        qtdeProcessos += 1
                    End If
                End If

            Next

            'Verificando se existe algum processo selecionado para captura, caso não, é retonado uma crítica
            If qtdeProcessos <= 0 Then
                lblMensagem.Text = "Selecione algum processo para capturar."
                lblMensagem.Visible = True

            Else

                'Pega todos os processos do grid para exibição de resultados e remontagem do grid
                For Each Me.dtgRow In dtgProcesso.Rows

                    If dtgRow.RowType = DataControlRowType.DataRow Then
                        listaDeProcessosDoGrid.Add(dtgRow.Cells(1).Text)
                    End If

                Next

                ' Pega apenas os processos marcados para captura
                ' e os carrega no objeto para serem capturados
                For Each Me.dtgRow In dtgProcesso.Rows
                    If dtgRow.RowType = DataControlRowType.DataRow Then

                        chkItem = DirectCast(dtgRow.FindControl("chkItem"), CheckBox)

                        If chkItem.Checked Then
                            listaDeProcessosASeremCapturados.Add(dtgRow.Cells(1).Text)

                        End If
                    End If

                Next

                For Each listaProcCaptura In listaDeProcessosASeremCapturados
                    If listaDeProcessos Is Nothing Then
                        listaDeProcessos &= "'" & listaProcCaptura & "'"
                    Else
                        listaDeProcessos &= "," & "'" & listaProcCaptura & "'"
                    End If
                Next

                'Pega apenas os processos marcados para captura os atribui
                'ao atributo numeroProcessos do objeto processo
                processo.numeroProcessos = listaDeProcessos
                lblPaginaAtual.Text = dtgProcesso.PageIndex
                'Chamando método para capturar selecionados
                CapturarPorOab()

            End If

        Catch ex As Exception

            lblMensagem.Text = "Ocorreu um erro ao tentar capturar o processo."
            lblMensagem.Visible = True

        End Try

    End Sub

    Private Sub remontaGridProcessosOab(listaDeProcessosDoGrid As List(Of String))
        Dim processo As New Processo
        Dim pessoaEmpresa As New Pessoa
        Dim empresa As New Empresa
        Dim listaDeProcessosAListar As String = ""

        pnlCaptura.Visible = False
        pnlProcesso.Visible = True


        'Empresa
        pessoaEmpresa.identificador = CType(usuario.empresa.pessoa.identificador, String)
        empresa.pessoa = pessoaEmpresa
        processo.empresa = empresa

        For Each listaProc In listaDeProcessosDoGrid
            If listaDeProcessosAListar = "" Then
                listaDeProcessosAListar &= "'" & listaProc & "'"
            Else
                listaDeProcessosAListar &= "," & "'" & listaProc & "'"
            End If
        Next

        processo.numeroProcessos = listaDeProcessosAListar
        dtgProcesso.DataSource = processo.listCapturaOab()
        dtgProcesso.DataBind()
        pnlCaptura.Visible = False
        pnlProcesso.Visible = True
        dtgProcesso.Visible = True
        dtgProcesso.PageIndex = lblPaginaAtual.Text
        btnConsultar_Click(Nothing, Nothing)




    End Sub

    Private Sub exibeResultados(listaDeProcessos As List(Of KeyValuePair(Of String, String)))

        Dim listaDeProcessosASeremCapturados As New List(Of String)
        Dim nvcResultado As New NameValueCollection
        Dim arlResultados As New ArrayList
        Dim lista As New List(Of KeyValuePair(Of String, String))
        nvcProcesso = New NameValueCollection



        pnlProcesso.Visible = False
        pnlCaptura.Visible = True


        ' Pega apenas os processos marcados para captura para realizar exibição do progresso
        For Each Me.dtgRow In dtgProcesso.Rows
            If dtgRow.RowType = DataControlRowType.DataRow Then

                chkItem = DirectCast(dtgRow.FindControl("chkItem"), CheckBox)

                If chkItem.Checked Then

                    listaDeProcessosASeremCapturados.Add(dtgRow.Cells(1).Text)

                End If
            End If

        Next

        For Each numProcessoResult In listaDeProcessos


            nvcResultado = New NameValueCollection
            nvcResultado.Add("Nº Processo", numProcessoResult.Key)
            nvcResultado.Add("Resultado", numProcessoResult.Value)
            arlResultados.Add(nvcResultado)

        Next
        dtgResultado.DataSource = util.ConverteResultado(arlResultados)
        dtgResultado.DataBind()


    End Sub

    Private Sub dtgProcesso_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles dtgProcesso.PageIndexChanging
        Try


            dtgProcesso.PageIndex = e.NewPageIndex
            btnConsultar_Click(Nothing, Nothing)


        Catch ex As Exception

            Session("erro") = ex
            Response.Redirect("../Paginas/Erro.aspx")

        End Try

    End Sub

    Private Sub lkbVoltarInferior_Click(sender As Object, e As EventArgs) Handles lkbVoltarInferior.Click


        Dim listaDeProcessosDoGrid As New List(Of String)


        'Pega todos os processos do grid para exibição de resultados e remontagem do grid
        For Each Me.dtgRow In dtgProcesso.Rows

            If dtgRow.RowType = DataControlRowType.DataRow Then
                listaDeProcessosDoGrid.Add(dtgRow.Cells(1).Text)
            End If

        Next
        remontaGridProcessosOab(listaDeProcessosDoGrid)

    End Sub

    Private Sub drpSeletorOpcao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpSeletorOpcao.SelectedIndexChanged
        'Carrega a opção do sistema de acordo com a opção de consulta escolhida (Nº Processo ou OAB)

        If drpSeletorOpcao.SelectedIndex = 1 Then
            lblTipoConsulta.Text = "Número do Processo"

            lblTipoConsulta.Visible = True
            txtNumPro.Text = ""
            txtNumProCNJ.Text = ""
            drpSistemaExterno.Items.Clear()
            txtNumOAB.Visible = False
            txtNumProCNJ.Visible = True
            ckbMascaraCNJ.Visible = True
            ckbMascaraCNJ.Checked = True
            lblOpcaoSistemaExterno.Visible = True
            drpSistemaExterno.Visible = True
            drpCodOpcao.Visible = False
            lblOpcaoSistema.Visible = False
            btnConsultar.Visible = False
            btnCapturar.Visible = True
            dtgProcesso.Visible = False
            lkbCapturarProcesso.Visible = False
            lblEstadoOab.Visible = False
            drpIsnEstadoOab.Visible = False
            lblComplementoOab.Visible = False
            drpIsnComplementoOab.Visible = False
            lblOpcaoCaptura.Visible = False
            drpOpcaoCaptura.Visible = False
            lblOpcaoAdicionalCaptura.Visible = False
            drpOpcaoAdicionalCaptura.Visible = False
            'carregarlistas()

        ElseIf drpSeletorOpcao.SelectedIndex = 2 Then
            lblTipoConsulta.Text = "Número da OAB"
            txtNumOAB.Text = ""
            'drpSistemaExterno.SelectedIndex = 0
            lblTipoConsulta.Visible = True
            txtNumPro.Visible = False
            txtNumProCNJ.Visible = False
            ckbMascaraCNJ.Visible = False
            txtNumOAB.Visible = True
            lblOpcaoSistemaExterno.Visible = True
            drpSistemaExterno.Visible = True
            lblOpcaoSistema.Visible = False
            drpCodOpcao.Visible = False
            btnCapturar.Visible = False
            btnConsultar.Visible = True
            drpIsnComplementoOab.Visible = False
            lblComplementoOab.Visible = False

            carregarlistas()

        Else
            lblTipoConsulta.Visible = False
            drpCodOpcao.Visible = False
            lblOpcaoSistema.Visible = False
            txtNumPro.Visible = False
            txtNumProCNJ.Visible = False
            ckbMascaraCNJ.Visible = False
            txtNumOAB.Visible = False
            lblOpcaoSistemaExterno.Visible = False
            drpSistemaExterno.Visible = False
            btnCapturar.Visible = False
            btnConsultar.Visible = False
            dtgProcesso.Visible = False
            lkbCapturarProcesso.Visible = False
            lblEstadoOab.Visible = False
            drpIsnEstadoOab.Visible = False
            lblOpcaoCaptura.Visible = False
            drpOpcaoCaptura.Visible = False
            lblOpcaoAdicionalCaptura.Visible = False
            drpOpcaoAdicionalCaptura.Visible = False
            drpIsnComplementoOab.Visible = False
            lblComplementoOab.Visible = False

        End If
    End Sub

    Private Sub drpSistemaExterno_Init(sender As Object, e As EventArgs) Handles drpSistemaExterno.Init

        If drpSeletorOpcao.SelectedIndex = 1 Then

        ElseIf drpSeletorOpcao.SelectedIndex = 2 Then
        End If

    End Sub

    Private Sub drpCodOpcao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpCodOpcao.SelectedIndexChanged

        Dim sisExterno As New SistemaExterno With {.identificador = drpSistemaExterno.SelectedItem.Value}

        Dim opcaoSisExterno As New OpcaoSistemaExterno With {.codigo = drpCodOpcao.SelectedItem.Value}
        Dim opcaoCaptura As New OpcaoCaptura
        Dim nvcSistemaExterno As New NameValueCollection

        If drpCodOpcao.SelectedIndex > 0 Then

            opcaoCaptura.sistemaExterno = sisExterno
            opcaoCaptura.opcaoSistemaExterno = opcaoSisExterno

            opcaoCaptura.carregarlista(drpOpcaoCaptura)

            'Carrega a opção de captura de acordo com o sistema selecionado, caso a busca seja po oab
            If drpOpcaoCaptura.Items.Count > 1 And drpSeletorOpcao.SelectedIndex = 2 Then

                lblOpcaoCaptura.Visible = True
                drpOpcaoCaptura.Visible = True

                nvcSistemaExterno = sisExterno.show

                'Nome opção captura
                If nvcSistemaExterno("nom_opcao_captura").ToString() <> "" Then
                    lblOpcaoCaptura.Text = "Opção de Captura - " & nvcSistemaExterno("nom_opcao_captura").Trim
                Else
                    lblOpcaoCaptura.Text = "Opção de Captura"
                End If

                Session("consultaOpcaoCaptura") = True
                drpOpcaoCaptura.Focus()

            Else

                drpOpcaoCaptura.Items.Clear()
                drpOpcaoCaptura.Visible = False
                lblOpcaoCaptura.Visible = False
                drpOpcaoAdicionalCaptura.Items.Clear()
                drpOpcaoAdicionalCaptura.Visible = False
                lblOpcaoAdicionalCaptura.Visible = False

                drpCodOpcao.Focus()

                Session("consultaOpcaoCaptura") = False

            End If
        Else

            drpOpcaoCaptura.Items.Clear()
            drpOpcaoCaptura.Visible = False
            lblOpcaoCaptura.Visible = False
            drpOpcaoAdicionalCaptura.Items.Clear()
            drpOpcaoAdicionalCaptura.Visible = False
            lblOpcaoAdicionalCaptura.Visible = False

            drpCodOpcao.Focus()

            Session("consultaOpcaoCaptura") = False

        End If

        validacaoDropDownEstado(sisExterno, opcaoSisExterno)
        validacaoDropDownComplemento(sisExterno, opcaoSisExterno)
    End Sub

    Private Sub validacaoDropDownEstado(sisExterno As SistemaExterno, opcaoSisExterno As OpcaoSistemaExterno)
        If processo.verificaBuscaPorEstadoOab(sisExterno, opcaoSisExterno) And drpSeletorOpcao.SelectedIndex = 2 Then

            If sisExterno.identificador <> 146 Then
                lblEstadoOab.Text = "Estado OAB"
                lblEstadoOab.Visible = True
                drpIsnEstadoOab.Visible = True
                drpIsnEstadoOab.SelectedIndex = 0
                Session("consultaNumeroOab") = True
            End If
        Else
            lblEstadoOab.Visible = False
            drpIsnEstadoOab.Visible = False
            drpIsnEstadoOab.SelectedIndex = 0
            Session("consultaNumeroOab") = False
        End If
    End Sub

    Private Sub validacaoDropDownComplemento(sisExterno As SistemaExterno, opcaoSisExterno As OpcaoSistemaExterno)
        If processo.verificaBuscaPorComplementoOab(sisExterno, opcaoSisExterno) And drpSeletorOpcao.SelectedIndex = 2 Then

            lblComplementoOab.Text = "Complemento de Inscrição OAB"
            lblComplementoOab.Visible = True
            drpIsnComplementoOab.Visible = True
            drpIsnComplementoOab.SelectedIndex = 0
            Session("consultaComplementoOab") = True
        Else
            lblComplementoOab.Visible = False
            drpIsnComplementoOab.Visible = False
            drpIsnEstadoOab.SelectedIndex = 0
            Session("consultaComplementoOab") = False
        End If
    End Sub

    'Private Sub drpIsnComplementoOab_Load(sender As Object, e As EventArgs) Handles drpIsnComplementoOab.Load

    '    If drpSeletorOpcao.SelectedValue = 2 Then
    '        drpIsnComplementoOab.Visible = True
    '        lblComplementoOab.Text = "Complemento de Inscrição OAB"
    '        lblComplementoOab.Visible = True
    '    Else
    '        drpIsnComplementoOab.Visible = False
    '        lblComplementoOab.Visible = False
    '    End If

    'End Sub

    Private Sub drpIsnComplementoOab_Init(sender As Object, e As EventArgs) Handles drpIsnComplementoOab.Init
        drpIsnComplementoOab.Items.Clear()

        drpIsnComplementoOab.Items.Add(New ListItem("", "-1"))
        drpIsnComplementoOab.Items.Add(New ListItem("A - Inscrição Suplementar", "A"))
        drpIsnComplementoOab.Items.Add(New ListItem("B - Inscrição por Transferência", "B"))
        drpIsnComplementoOab.Items.Add(New ListItem("E - Inscrição de Estagiário", "E"))
        drpIsnComplementoOab.Items.Add(New ListItem("N - Inscrição de Provisionado", "N"))
        drpIsnComplementoOab.Items.Add(New ListItem("P - Inscrição Provisória", "P"))

    End Sub

    Protected Sub drpOpcaoCaptura_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpOpcaoCaptura.SelectedIndexChanged

        Dim opcaoAdicionalCaptura As New OpcaoAdicionalCaptura
        Dim opcaoCaptura As New OpcaoCaptura
        Dim sistemaExterno As New SistemaExterno
        Dim opcaoSistemaExter As New OpcaoSistemaExterno
        Dim nvcSistemaExterno As New NameValueCollection

        If drpOpcaoCaptura.SelectedIndex > 0 Then

            opcaoCaptura.codigo = drpOpcaoCaptura.SelectedValue
            opcaoAdicionalCaptura.opcaoCaptura = opcaoCaptura
            sistemaExterno.identificador = drpSistemaExterno.SelectedValue

            If drpCodOpcao.Visible.Equals(True) Then
                opcaoSistemaExterno.codigo = drpCodOpcao.SelectedValue
            End If

            opcaoAdicionalCaptura.sistemaExterno = sistemaExterno
            opcaoAdicionalCaptura.opcaoSistemaExterno = opcaoSistemaExterno
            opcaoAdicionalCaptura.carregarlista(drpOpcaoAdicionalCaptura)

            If drpOpcaoAdicionalCaptura.Items.Count > 1 Then

                lblOpcaoAdicionalCaptura.Visible = True
                drpOpcaoAdicionalCaptura.Visible = True

                nvcSistemaExterno = sistemaExterno.show

                'Nome opção
                If nvcSistemaExterno("nom_opcao_adicional_captura").ToString() <> "" Then
                    lblOpcaoAdicionalCaptura.Text = "Opção Adicional de Captura - " & nvcSistemaExterno("nom_opcao_adicional_captura").Trim
                Else
                    lblOpcaoAdicionalCaptura.Text = "Opção Adicional de Captura"
                End If

                drpOpcaoAdicionalCaptura.Focus()

            Else

                drpOpcaoAdicionalCaptura.Items.Clear()

                drpOpcaoAdicionalCaptura.Visible = False
                lblOpcaoAdicionalCaptura.Visible = False

                drpOpcaoCaptura.Focus()


            End If
        End If
    End Sub
    Protected Sub ckbMascaraCNJ_CheckedChanged(sender As Object, e As EventArgs) Handles ckbMascaraCNJ.CheckedChanged

        If ckbMascaraCNJ.Checked.Equals(True) Then
            Dim numeroProcesso = txtNumPro.Text.Replace(".", "").Replace("-", "")

            If numeroProcesso.Length > 6 Then
                numeroProcesso = numeroProcesso.Insert(7, "-")
            End If

            If numeroProcesso.Length > 9 Then
                numeroProcesso = numeroProcesso.Insert(10, ".")
            End If

            If numeroProcesso.Length > 14 Then
                numeroProcesso = numeroProcesso.Insert(15, ".")
            End If

            If numeroProcesso.Length > 16 Then
                numeroProcesso = numeroProcesso.Insert(17, ".")
            End If

            If numeroProcesso.Length > 19 Then
                numeroProcesso = numeroProcesso.Insert(20, ".")
            End If

            'If numeroProcesso.Length > 25 Then
            '    numeroProcesso = numeroProcesso.Substring(0, 25)
            'End If
            txtNumProCNJ.Text = numeroProcesso
            txtNumProCNJ.Visible = True
            txtNumPro.Visible = False
            txtNumPro.Text = ""

            'Verificação do número no formato CNJ
            Dim lysisUtil = New LysisUtil

            If lysisUtil.verificarNumeroCNJ(txtNumProCNJ.Text) Then

                'Sistema Externo

                'Dim processo = New Processo With {.numeroProcessos = txtNumProCNJ.Text}
                Dim sistemaInsercaoARL As ArrayList = processo.listSistemasCaptura("1", txtNumProCNJ.Text)


                drpCodOpcao.Items.Clear()
                drpCodOpcao.Items.Add(New ListItem("", 0))
                drpCodOpcao.Visible = False
                lblOpcaoSistema.Visible = False
                drpSistemaExterno.Items.Clear()
                drpSistemaExterno.Items.Add(New ListItem("", 0))

                For Each sistema In sistemaInsercaoARL
                    drpSistemaExterno.Items.Add(New ListItem(sistema("nome"), sistema("identificador")))
                Next

                lblMensagem.Visible = False
                drpSistemaExterno.Focus()
            Else
                If txtNumProCNJ.Text <> "" Then
                    lblMensagem.Text = "Número do processo não está de acordo com o formato CNJ."
                    lblMensagem.Visible = True
                End If

                drpCodOpcao.Items.Clear()
                drpCodOpcao.Items.Add(New ListItem("", 0))
                drpCodOpcao.Visible = False
                lblOpcaoSistema.Visible = False
                drpSistemaExterno.Items.Clear()
                txtNumProCNJ.Focus()
            End If
        Else
            'Sistema externo
            carregarlistas()
            drpCodOpcao.Items.Clear()
            drpCodOpcao.Items.Add(New ListItem("", 0))
            drpCodOpcao.Visible = False
            lblOpcaoSistema.Visible = False
            'drpSistemaExterno.Items.Clear()

            'Dim lista As New Lista
            'lista.CarregaCombo(drpSistemaExterno, "ISN_SISTEMA_EXTERNO", "DSC_SISTEMA_EXTERNO", "VSE_SISTEMA_EXTERNO", True)

            txtNumPro.Text = txtNumProCNJ.Text
            txtNumProCNJ.Visible = False
            txtNumProCNJ.Text = ""
            txtNumPro.Visible = True
            lblMensagem.Visible = False
            drpSistemaExterno.Focus()
        End If

    End Sub
    Protected Sub txtNumProCNJ_TextChanged(sender As Object, e As EventArgs) Handles txtNumProCNJ.TextChanged

        If ckbMascaraCNJ.Checked.Equals(True) Then

            Dim lysisUtil = New LysisUtil

            If lysisUtil.verificarNumeroCNJ(txtNumProCNJ.Text) Then

                'Sistema Externo

                'Dim processo = New Processo With {.numeroProcessos = txtNumProCNJ.Text}
                Dim sistemaInsercaoARL As ArrayList = processo.listSistemasCaptura("1", txtNumProCNJ.Text)

                drpCodOpcao.Items.Clear()
                drpCodOpcao.Items.Add(New ListItem("", 0))
                drpSistemaExterno.Items.Clear()
                drpSistemaExterno.Items.Add(New ListItem("", 0))

                For Each sistema In sistemaInsercaoARL
                    drpSistemaExterno.Items.Add(New ListItem(sistema("nome"), sistema("identificador")))
                Next

                lblMensagem.Visible = False
                drpSistemaExterno.Focus()
            Else
                drpCodOpcao.Items.Clear()
                drpCodOpcao.Items.Add(New ListItem("", 0))
                drpSistemaExterno.Items.Clear()
                lblMensagem.Text = "Número do processo não está de acordo com o formato CNJ."
                lblMensagem.Visible = True
                txtNumProCNJ.Focus()
            End If
        Else

        End If

    End Sub
End Class
