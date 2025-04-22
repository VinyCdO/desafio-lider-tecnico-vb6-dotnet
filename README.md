# desafio-lider-tecnico-vb6-dotnet

[![Status do Projeto](https://img.shields.io/badge/status-concluído-yellow)](https://github.com/VinyCDO/desafio-lider-tecnico-vb6-dotnet)
![Linguagens](https://img.shields.io/github/languages/count/VinyCDO/desafio-lider-tecnico-vb6-dotnet)
![Última Commmit](https://img.shields.io/github/last-commit/VinyCDO/desafio-lider-tecnico-vb6-dotnet)

Este projeto representa um módulo de negociação de dívidas para um sistema legado que roda em VB6, com banco de dados PostgreSQL e integração com uma API REST em .NET 8.

## Sumário

- [Sobre o Projeto](#sobre-o-projeto)
- [Funcionalidades](#funcionalidades)
- [Tecnologias Utilizadas](#tecnologias-utilizadas)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Como Usar](#como-usar)

## Sobre o Projeto

Este projetofoi criado para atender a necessidade de melhor gestão de dívidas, com armazenamento e pesquisa dos dados de identificação das dívidas por CPF e fácil acesso a cálculos de juros, com histórico de negociações por CPF, Data, Taxa de Juros e Qtde Parcelas. 

## Funcionalidades

- Lançamento de dívidas por CPF
- Pesquisa de dívidas por CPF
- Simulação de negocicações com aplicação de juros compostos

## Tecnologias Utilizadas

Este projeto foi desenvolvido utilizando as seguintes tecnologias:

- **Visual Basic 6 (VB6)**: A linguagem de programação principal para a lógica de negócios e interface gráfica.
  - *Nota:* A documentação oficial da Microsoft para o VB6 não é mais atualizada, mas é possível encontrar diversos recursos e documentações criadas pela comunidade online.
- **.NET Core 8**: A plataforma de desenvolvimento moderna, de código aberto e multi-plataforma da Microsoft, utilizada para implementação de API em C# responsável por operações integradas a aplicação VB6, e acessível para consumo por outras aplicações se necessário.
  - [Documentação do .NET 8](https://learn.microsoft.com/pt-br/dotnet/core/whats-new/dotnet-8)
- **ADO (ActiveX Data Objects)**: Tecnologia da Microsoft utilizada pelo VB6 para acesso e manipulação de dados de diversas fontes, incluindo bancos de dados.
  - [Documentação do ADO](https://learn.microsoft.com/pt-br/sql/mdac/ado/reference/ado-api-reference)
- **Componentes ActiveX (OCX)**: Componentes reutilizáveis que podem ser incorporados na interface gráfica do Visual Basic 6 para funcionalidades específicas.
  - *Nota:* A documentação específica de componentes ActiveX pode variar dependendo do fornecedor do componente. Geralmente, a documentação é fornecida com o próprio componente ou no site do desenvolvedor.
- **PostgreSQL**: Um sistema de gerenciamento de banco de dados relacional (SGBDR) poderoso e de código aberto, utilizado para armazenar e gerenciar os dados da aplicação. Com conexão a partir do VB6 por drivers ODBC, e utilizada no projeto da API também para gravação e consulta de dados.
  - [Documentação do PostgreSQL](https://www.postgresql.org/docs/)

## Pré-requisitos

- [Visual Basic 6.0](https://winworldpc.com/product/microsoft-visual-bas/60) - Para desenvolvimento da aplicação VB6 (Recomenda-se também pela comunidade a instalação do Service Pack 6, é possível encontrar na web para download vários links, porém recomendo utilizar o setup do site oficial da [Microsoft](https://www.microsoft.com/en-in/download/details.aspx?id=7030)
- [Visual Studio](https://visualstudio.microsoft.com/pt-br/) / [Visual Code](https://code.visualstudio.com/download) - Para desenvolvimento da API (utilizado Visual Studio neste projeto, mas pode ser usado Visual Code se preferível)
- [PostgreSQL](https://www.enterprisedb.com/downloads/postgres-postgresql-downloads) - Para administração/desenvolvimento da base de dados (utilizado pgAdmin 4 para implementações deste projeto)
- [psqlodbc-setup](https://www.postgresql.org/ftp/odbc/releases/ - Para permitir a conexão do seu projeeto VB6 com sua base de dados PostgreSQL
- [MSHFlexgrid](https://www.ocxdump.com/download-ocx-files_new.php/ocxfiles/M/MSHFLXGD.OCX/6.00.30050/download.html#google_vignette) - Para utilização de recursos de grids no VB6 (instrução de instalação e configuração no link de download)

## Instalação

Descreva os passos necessários para instalar e configurar o seu projeto no ambiente local. Seja o mais claro e detalhado possível.

1. Clone o repositório:
   ```bash
   git clone https://github.com/VinyCDO/desafio-lider-tecnico-vb6-dotnet.git

2. Prepare ambiente de banco de dados local (ou se preferir crie a base de dados em seu servidor PostgreSQL)
   ```Instalação
   2.1 Instale o PostgreSQL com setup informado no link dos pré-requisitos acima

   2.2 Acesse o pgAdmin com as credenciais informadas durante seu setup para criação do servidor local (ou com as credencias do seu servidor caso esteja utilizando um ambiente já pré existente)

   2.3 Abre o Query Tool Workspace, e execute os scripts localizado no repositório no caminho abaixo:
      [ScriptsBD](https://github.com/VinyCdO/desafio-lider-tecnico-vb6-dotnet/tree/main/ScriptsBD)
       1. Criação BD
       2. Criação das tabelas dividas e negociacoes
       3. Criação da procedure de inserção de dividas
       4. Criação da function de cálculo de juros compostos

   2.4 Instale o driver ODBC (psqlodbc-setup), para habilitar a conexão da sua aplicação VB6 com sua base de dados PostgreSQL, setup informado no link dos pré-requisitos acima

3. Prepare ambiente para execução do Projeto da API
   ```Instalação
   3.1 Instale o Visual Studio ou Visual Code, o que preferir, conforme links informados acima na sessão de pré requisitos

   3.2 Abre o projeto localizado na pasta abaixo dentro do repositório:
      [apiNegociacaoDividas](https://github.com/VinyCdO/desafio-lider-tecnico-vb6-dotnet/tree/main/apiNegociacaoDividas)
       1. Verifique os arquivos appSettings.json do seu projeto estão com o apontamento e as credenciais corretas para sua base de dados PostgreSQL, conforme configurado no passo 2
   
       **⚠️ ATENÇÃO:** NÃO COMMITAR CREDENCIAIS SENSÍVEIS, MANTENHA SOMENTE CREDENCIAL PARA EXECUÇÃO LOCALHOST

   3.3 Execute o seu projeto, caso deseje realizar testes da API, será apresentada a documentação no padrão OpenAPI(Swagger) para validação dos endpoints

4. Prepare ambiente para execução do Projeto da aplicação VB6
   ```Instalação
   4.1 Instale o Visual Studio 6.0, através do setup localizado na web pelo link informado na sessão de pré requisitos

     **⚠️ ATENÇÃO:** Há um ponto de atenção sobre essa instalação, visto que não há mais suporte a muitos anos para Visual Basic 6.0 pela Microsoft, então podem ser encontrados problemas para instalação ou para localizar um setup disponível para download

   4.2 Abra o projeto localizado na pasta abaixo dentro do repositório:
       [appLancamentoDividas](https://github.com/VinyCdO/desafio-lider-tecnico-vb6-dotnet/tree/main/appLancamentoDividas)

   4.3 Cerifique-se de estar com seu arquivo mdlConstants.bas devidamente configurado:
     * CONN_STRING - Connection string da sua base de dados PostgreSQL configurado no passo 2
     * ENDPOINT_API - Url da sua API configurado no passo 3

      **⚠️ ATENÇÃO:** Há um executável da versão mais recente versionado no repostiório no caminho abaixo, é possível rodar o mesmo sem a necessidade da parametrização do projeto, referente aos itens mencionados acima neste tópico 5, contanto será necessário parametrizar seu ambiente (BD/API) para que fique de acordo com a definição abaixo:
        - Public Const CONN_STRING As String = "Driver={PostgreSQL ANSI};Server=localhost;Port=5432;Database=dividasPaschoalloto;Uid=postgres;Pwd=admin;"
        - Public Const ENDPOINT_API As String = "https://localhost:44362"

## Como usar 

Após parametrizar o ambiente, basta executar o projeto (ou o executável disponível), e seguir os passos abaixo:

  1. Acesso a tela inicial, contendo acesso às operações abaixo:
      ```
     * Opção de pesquisa por CPF, para buscar Dívidas cadastradas
     * Opção de acesso a tela de Lançamento de Dívidas
     * Opção de acesso a tela de Negociação de Dívidas
  
  2. Operação de Lançamento de Dívidas
      ```
     * Campos disponíveis para cadastro de informações da dívida a ser cadastrada: CPF, Valor e Data Vencimento
     * Botão para registrar na base de dados as informações da dívida
      
  3. Operação de Negociação de Dívidas
      ```
     * Campos de identificação da dívida consultada na tela de pesquisa
     * Listagem com todas as negociações já cadastradas anteriormente para a dívida consultada
     * Campos para informar Quantidade de Parcelas e Taxa de Juros para cálculo da negociação
     * Botão para Simular Negociação, com resultado em tela contendo o valor total da negociação para os parâmetros informados e uma opção para gravar no histórico ou fechar a operação de simulação
