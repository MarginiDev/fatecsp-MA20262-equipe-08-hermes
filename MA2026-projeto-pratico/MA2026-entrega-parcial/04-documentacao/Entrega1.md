# Documentação entrega 1

## 1. Requisitos de Negócio (RN) e Requisitos Operacionais (RO)

### Requisitos de Negócio (RN)

- **RN01 - Análise de Vendas por Localidade:** O sistema deve mapear e detalhar o desempenho comercial por estado, região e cidade brasileira, respondendo à pergunta central sobre quais localidades concentram maior ou menor volume de consumo de mercadorias.
- **RN02 - Desempenho por Categoria:** O painel analítico deve destacar quais categorias de produtos são mais consumidas pelo público brasileiro de determinadas localidades.
- **RN03 - Enriquecimento Socioeconômico:** Cruzar dados de vendas com indicadores demográficos, de renda e PIB (via API IBGE SIDRA) para entender o comportamento de compra por perfil socioeconômico.

### Requisitos Operacionais (RO)

- **RO01 - Pipeline de Dados Integrado:** Fluxo contínuo de dados estruturado desde os dados brutos e Excel até o Power BI (`Excel -> Power Query -> Power Pivot/DAX -> VBA -> SQL -> API -> Power BI`), com toda a governança e armazenamento versionados no GitHub.
- **RO02 - Automação de Rotinas:** Desenvolvimento de macros personalizadas em VBA para automatizar relatórios operacionais diários ou semanais.
- **RO03 - Persistência Relacional:** Implementação de um banco de dados relacional com scripts DDL (estrutura) e DML (consultas de agregação).
- **RO04 - Modelagem Dimensional:** Criação de um modelo *Star Schema* estruturado com Tabela Fato e Tabelas Dimensão no Power Pivot.
- **RO05 - Organização de Repositório:** Manutenção rigorosa da árvore de pastas exigida pela disciplina no GitHub (`MA2026-entrega-parcial` e `MA2026-entrega-final`).

## 2. Matriz 5W2H

| Dimensão | Pergunta | Descrição / Implementação no Projeto Hermes |
| :--- | :--- | :--- |
| **What** <br> (O quê?) | O que será feito? | Uma plataforma integrada de inteligência e análise de dados para o varejo (e-commerce Olist), combinando Excel, Power Query, Power Pivot, VBA, Banco de Dados SQL, API REST e um Dashboard interativo no Power BI. |
| **Why** <br> (Por quê?) | Qual o objetivo? | Investigar vendas por região, identificar categorias mais vendidas, mapear o desempenho geográfico e correlacionar os dados de consumo com fatores socioeconômicos. |
| **Who** <br> (Quem?) | Quem fará? | **Equipe 08:** Carybé Gonçalves (Programador API), Daniel Victor (Eng. de Dados/Excel/Documentação), Gabriel Felipe (Automação VBA & APIs), Renato Rodrigues (Banco de Dados SQL/Power BI) e Victor Margini (Modelagem Power Query/DAX). |
| **Where** <br> (Onde?) | Onde será executado/armazenado? | Os arquivos e códigos serão versionados no repositório GitHub da equipe e processados nas ferramentas locais (Excel, Power BI, VS Code/IDE SQL). |
| **When** <br> (Quando?) | Qual o cronograma? | Projeto desenvolvido ao longo do semestre 2º/2026, com Defesa Parcial na Aula 8, Checkpoint na Aula 16 e Entrega Final/Defesa ao Vivo na Aula 20. |
| **How** <br> (Como?) | Como será feito? | Estruturando o repositório no GitHub, limpando dados no Excel e Power Query, modelando no Power Pivot, automatizando em VBA, populando banco SQL e consumindo APIs públicas. |
| **How Much** <br> (Quanto custa?) | Qual o custo financeiro/recursos? | Utilização de ferramentas educacionais e datasets públicos gratuitos, gerando custo financeiro zero, mas exigindo esforço técnico e adaptação a ferramentas online gratuitas. |

## 3. Gestão de Riscos

| Risco Identificado | Impacto | Probabilidade | Plano de Mitigação e Ação Preventiva |
| :--- | :---: | :---: | :--- |
| **Indisponibilidade de Plataformas Pagas e Uso de Alternativas Online** | Alto | Alta | Substituir recursos pagos por ferramentas online gratuitas realizando testes para contornar eventuais limitações de desempenho ou instabilidade das plataformas web. |
| **Instabilidade ou Indisponibilidade de APIs Externas (IBGE/Feriados)** | Médio | Média | Salvar cópias locais em formato JSON ou CSV estático dos dados coletados da API, garantindo que o dashboard funcione mesmo se o serviço externo cair no dia da apresentação. |
| **Inconsistências ou Dados Nulos/Corrompidos na Base Bruta do Kaggle** | Alto | Alta | Realizar auditoria e tratamento de dados na fase inicial de Excel Avançado e Power Query (Etapas 1 e 2), documentando todas as regras de limpeza. |
| **Dificuldade na Integração e Sincronia do Pipeline Completo** | Alto | Média | Seguir o cronograma de entregas por aula, validando cada ferramenta de forma modular antes de unir tudo no Power BI. |
| **Divergência de Versões de Código no GitHub entre os Integrantes** | Médio | Média | Estabelecer padrões de commits, uso de branches específicas para cada subsistema e revisões em equipe antes de subir atualizações para a branch principal. |
| **Estouro de Prazo nos Marcos de Validação (Parcial e Final)** | Alto | Baixa | Acompanhar o cronograma oficial de aulas como marcos rígidos de controle interno da equipe. |

## 4. Arquitetura, Pipeline de Dados e Modelagem (Power Pivot)

Para estruturar o ecossistema do Projeto Hermes, o fluxo operacional e técnico foi desenhado e executado de forma sequencial e integrada:

### 4.1. Configuração do Repositório (GitHub)
- **Versionamento Inicial:** Antes de qualquer manipulação de dados, a equipe estruturou e preparou todo o repositório no GitHub seguindo o padrão exigido pela disciplina, garantindo a separação correta entre as pastas de entrega parcial e final (`01-excel`, `02-powerquery`, `03-powerpivot`, etc.).

### 4.2. Preparação e Auditoria no Excel
- **Higienização e Cruzamentos:** O projeto deu início ao tratamento da base bruta no Excel Avançado, utilizando funções como `PROCX` para realizar cruzamentos fundamentais, checagem de integridade e correções iniciais nas tabelas de origem.

### 4.3. Tratamento Avançado via Power Query
- **Limpeza de Estrutura:** Os códigos e tabelas são importados para o Power Query para padronização. O foco desta etapa consiste na remoção de linhas nulas, tratamento de inconsistências e eliminação de registros inválidos ou inexistentes.
- **Filtro de Escopo:** Durante essa etapa, mantivemos ativas apenas as pastas essenciais para o escopo analítico: Lista de categorias, Lista de clientes, Lista de pagamentos, Lista de pedidos e Lista de vendas.

### 4.4. Modelagem Dimensional no Power Pivot (Star Schema)
- **Integração Relacional:** No Power Pivot, consolidamos e correlacionamos as informações de quatro planilhas diferentes. Unificamos o **ID do pedido**, o **ID do produto**, os dados transacionais de **vendas com suas formas de pagamento e valores**, além da segmentação da **categoria de cada produto**.
- **O Modelo na Prática:** Essa construção estruturou um esquema estrela (*Star Schema*), onde as tabelas descritivas (Dimensões de Produtos, Regiões e Pagamentos) alimentam a tabela central de transações (Fato Vendas), permitindo  análises de desempenho sem redundância de dados.

### 4.5. Automação de Rotinas via VBA
- **Geração de Relatórios:** Desenvolvimento e gravação de macros personalizadas em VBA (armazenadas em módulos `.bas` e na planilha habilitada `.xlsm`) para automatizar a extração, formatação e atualização de relatórios operacionais do e-commerce de forma ágil e sem intervenção manual repetitiva.

### 4.6. Persistência e Consultas em Banco de Dados (SQL)
- **Modelagem Relacional e Consultas (DDL e DML):** Implementação de um banco de dados relacional (PostgreSQL/MySQL) para garantir a persistência segura dos dados. Foram estruturados scripts DDL para a criação de tabelas, chaves primárias e estrangeiras, além de scripts DML contemplando consultas avançadas com junções (`JOIN`), filtros e agregações (`GROUP BY`) para cruzamentos analíticos.

### 4.7. Enriquecimento via Consumo de APIs REST
- **Integração com Fontes Externas:** Consumo de dados públicos externos por meio de requisições a APIs REST (como o IBGE SIDRA para indicadores demográficos, de renda e PIB, e APIs de calendários festivos). Os dados serão obtidos em formato JSON foram tratados e integrados para enriquecer a base principal, permitindo correlacionar fatores externos e socioeconômicos ao comportamento de compra do consumidor.
