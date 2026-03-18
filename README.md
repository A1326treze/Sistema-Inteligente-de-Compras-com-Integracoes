# Sistema Inteligente de Gestão de Compras com Integrações
#### Ecossistema integrado para controle de inventário, automação de pesquisa de preços e análise preditiva.
Este projeto integra o ecossistema Google (Sheets, AppScript, AppSheet, Looker Studio) e Microsoft (Power BI) para resolver o problema de oscilação de preços e gestão de estoque doméstico.

## Google Apps Script
Desenvolvimento de automações em GoogleScript _(JavaScript)_ para otimizar a performance e a usabilidade:
- **Menus Personalizados** que facilitam a utilização da planilha e correção de possíveis erros;
- **Visualização Dinâmica** com algoritmo de gradiente condicional e separação de itens para identificação visual imediata de disparidade de preços na Matriz de Pesquisa;
- **Engenharia de Dados** de sistema de transformação de matrizes _(Pivot/Unpivot)_ para registro de histórico de preços em lote ou datas individuais, com funções avançadas de Desfazer/Refazer _(Undo/Redo)_;
- **Topbar flutuante opcional** para facilitar navegação em telas pequenas ou abas divididas;
- **Limpeza da Matriz de Pesquisa** para agilizar a inclusão de novos preços;
- **Ocultação de coluna** para mostrar apenas as lojas essenciais a serem pesquisadas;
- **Tratamento e Limpeza de Dados _(ETL)_** para padronização automatizada de strings que evitam erros de duplicidade e pesquisa (ex: espaços extras ou caracteres invisíveis);
- - **Filtro de setores** para mostrar apenas o essencial a ser visto;
- **Sincronização da Matriz de preço com AppSheet** para integração completa do Sistema ao celular.
## Google Looker Studio
Gráficos que facilitam a visualização de preços e suas flutuações sobre diferentes itens, pacotes e marcas, sendo possível identificar tendências de preços e sazonalidade.
## Google AppSheet
Uma interface mobile otimizada para coleta de dados direto no mercado _(in loco)_, com sincronização offline-first para garantir a entrada de dados mesmo sem conexão estável.
## Power BI
Em desenvolvimento. Migração para uma arquitetura de dados mais robusta, visando fornecer maior granularidade de filtros e modelagem relacional.

## Skills usadas:
_Linguagens_: Google AppScript (JavaScript).
_Ferramentas_: Google Sheets Avançado, Looker Studio, Power BI, AppSheet.
_Conceitos_: ETL, Modelagem de Dados, UI/UX para Planilhas, Gestão de Inventário.
