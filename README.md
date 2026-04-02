# Automação de Relatórios Financeiros - CLS Outlet 🚀

Este projeto foi desenvolvido para automatizar a extração e análise de dados de vendas do **Mercado Livre** para a empresa **CLS Outlet**.

## 🛠️ Problema Resolvido
O processo de fechamento diário era manual e suscetível a erros. O script realiza a varredura automática de planilhas complexas (células mescladas), calcula o faturamento bruto, taxas de marketplace e **faturamento líquido**, gerando um relatório executivo em PDF.

## 🚀 Tecnologias Utilizadas
* **Python 3.x**
* **Pandas**: Manipulação e limpeza de dados (Data Wrangling).
* **FPDF**: Geração de documentos PDF automatizados.
* **Openpyxl**: Engine para leitura de arquivos Excel (.xlsx).

## 📊 Funcionalidades
* **Busca Inteligente**: Localiza automaticamente o cabeçalho dos dados, independente de linhas vazias no topo.
* **Cálculo Financeiro**: Diferenciação entre Receita Bruta e Receita Líquida (Total BRL).
* **Relatório Profissional**: Gera um PDF formatado pronto para apresentação à gerência.

## 📈 Evolução do Projeto (BI Completo)
Na versão mais recente, o sistema evoluiu para um Dashboard de Business Intelligence:
* **Integração Mensal:** Lê relatórios de "Evolução do Negócio" (30 dias).
* **Filtros de Anomalia:** Ignora linhas residuais do Mercado Livre e busca planilhas automaticamente por regex/substrings.
* **KPIs Avançados:** Extrai e calcula Ticket Médio, Preço Médio por Unidade, Vendas Canceladas e Devolvidas, isolando dados da aba 'Negócio'.
  
