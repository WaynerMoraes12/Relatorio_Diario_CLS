import pandas as pd
import os
from fpdf import FPDF

def limpar_valor(valor):
    """Transforma valores do Excel em números reais de forma segura"""
    if pd.isna(valor): return 0
    if isinstance(valor, str):
        valor = valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
    try:
        return float(valor)
    except:
        return 0

def carregar_planilha_vendas():
    """Lê MÚLTIPLAS planilhas de vendas diárias e junta todas em uma só"""
    arquivos = [f for f in os.listdir('.') if 'Vendas_BR' in f and f.endswith('.xlsx')]
    if not arquivos: 
        return None
    
    lista_de_dataframes = []
    
    for arquivo in arquivos:
        for i in range(15):
            try:
                df = pd.read_excel(arquivo, header=i)
                df.columns = [str(c).strip() for c in df.columns]
                if 'Total (BRL)' in df.columns: 
                    lista_de_dataframes.append(df)
                    break # Achou o cabeçalho, vai para o próximo arquivo
            except: 
                continue
                
    # Junta todas as planilhas lidas em um único "Super DataFrame"
    if lista_de_dataframes:
        return pd.concat(lista_de_dataframes, ignore_index=True)
    return None

def carregar_planilha_evolucao():
    arquivo = [f for f in os.listdir('.') if 'evolucao' in f.lower() and f.endswith('.xlsx')]
    if not arquivo: return None
    try:
        df = pd.read_excel(arquivo[0], sheet_name='Negócio', header=5)
        df.columns = [str(c).strip() for c in df.columns]
        
        # FILTRO DE ANOMALIA: Remove linhas vazias ou a linha de "Total" do Mercado Livre
        col_data = next((c for c in df.columns if 'data' in c.lower()), None)
        if col_data:
            df = df[~df[col_data].astype(str).str.contains('Total', case=False, na=False)]
            df = df.dropna(subset=[col_data])
            
        return df
    except Exception as e:
        print(f"Erro ao acessar aba Negócio: {e}")
        return None

def gerar_pdf_consolidado(df_vendas, df_evol):
    pdf = FPDF()
    pdf.add_page()
    
    # --- CABEÇALHO ---
    pdf.set_fill_color(0, 51, 102)
    pdf.rect(0, 0, 210, 40, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 18)
    pdf.cell(190, 15, "RELATÓRIO REPORT ESTRATEGICO - CLS OUTLET", ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(190, 10, "Responsavel Tecnico: Wayner Moraes | Engenheiro de Software", ln=True, align='C')
    
    pdf.ln(25); pdf.set_text_color(0, 0, 0)

    # --- SEÇÃO 1: MÉTRICAS DO DIA ---
    pdf.set_font("Arial", 'B', 14); pdf.set_fill_color(230, 230, 230)
    pdf.cell(190, 10, "1. PERFORMANCE OPERACIONAL (DIA)", ln=True, fill=True)
    pdf.ln(2)
    
    bruta_dia = df_vendas['Receita por produtos (BRL)'].apply(limpar_valor).sum()
    liquida_dia = df_vendas['Total (BRL)'].apply(limpar_valor).sum()
    qtd_dia = df_vendas['Unidades'].apply(limpar_valor).sum()

    pdf.set_font("Arial", size=12)
    pdf.cell(95, 8, f"Unidades Vendidas: {int(qtd_dia)}", ln=0)
    pdf.cell(95, 8, f"Venda Bruta: R$ {bruta_dia:,.2f}", ln=1)
    pdf.set_font("Arial", 'B', 12); pdf.set_text_color(0, 100, 0)
    pdf.cell(190, 8, f"FATURAMENTO LIQUIDO: R$ {liquida_dia:,.2f}", ln=1)
    pdf.set_text_color(0, 0, 0)
    
    # --- NOVIDADE: ANÁLISE DE STATUS (ESTADO) ---
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(190, 8, "Resumo de Status dos Pedidos:", ln=1)
    pdf.set_font("Arial", size=12)
    
    # Procura a coluna Estado ou Status dinamicamente
    col_estado = next((c for c in df_vendas.columns if 'estado' in c.lower() or 'status' in c.lower()), None)
    
    if col_estado:
        # A mágica do Pandas: value_counts() conta quantas vezes cada palavra aparece
        contagem_status = df_vendas[col_estado].value_counts()
        for status, quantidade in contagem_status.items():
            # Destaca palavras críticas em vermelho (Mediação, Reclamação, Cancelado)
            if any(palavra in str(status).lower() for palavra in ['mediação', 'reclamação', 'cancelad']):
                pdf.set_text_color(200, 0, 0)
            else:
                pdf.set_text_color(0, 0, 0)
            
            pdf.cell(190, 6, f"  - {status}: {quantidade} pedido(s)", ln=1)
    else:
        pdf.cell(190, 6, "  - Coluna 'Estado' não localizada na planilha diária.", ln=1)
        
    pdf.set_text_color(0, 0, 0) # Reseta a cor para preto

    # --- SEÇÃO 2: MÉTRICAS DOS ÚLTIMOS 30 DIAS ---
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 14); pdf.set_fill_color(230, 230, 230)
    pdf.cell(190, 10, "2. EVOLUCAO DO NEGOCIO (ULTIMOS 30 DIAS)", ln=True, fill=True)
    pdf.ln(2); pdf.set_font("Arial", size=12)

    if df_evol is not None:
        col_canc = next((c for c in df_evol.columns if 'canceladas' in c.lower() and 'quantidade' in c.lower()), None)
        col_dev = next((c for c in df_evol.columns if 'devolvidas' in c.lower() and 'quantidade' in c.lower()), None)

        visitas = df_evol['Visitas'].apply(limpar_valor).sum()
        vendas_qtd = df_evol['Quantidade de vendas'].apply(limpar_valor).sum()
        vendas_brutas = df_evol['Vendas brutas'].apply(limpar_valor).sum()
        
        canceladas = df_evol[col_canc].apply(limpar_valor).sum() if col_canc else 0
        devolvidas = df_evol[col_dev].apply(limpar_valor).sum() if col_dev else 0
        
        ticket_medio = df_evol['Valor médio por venda'].apply(limpar_valor).mean()
        preco_unidade = df_evol['Preço médio por unidade'].apply(limpar_valor).mean()

        pdf.cell(95, 8, f"Visitas Totais: {int(visitas)}", ln=0)
        pdf.cell(95, 8, f"Quantidade de Vendas: {int(vendas_qtd)}", ln=1)
        pdf.cell(95, 8, f"Vendas Brutas (30d): R$ {vendas_brutas:,.2f}", ln=0)
        pdf.cell(95, 8, f"Ticket Medio: R$ {ticket_medio:,.2f}", ln=1)
        pdf.cell(95, 8, f"Preco Medio Unidade: R$ {preco_unidade:,.2f}", ln=1)
        
        pdf.set_text_color(200, 0, 0)
        pdf.cell(95, 8, f"Vendas Canceladas: {int(canceladas)}", ln=0)
        pdf.cell(95, 8, f"Vendas Devolvidas: {int(devolvidas)}", ln=1)
        pdf.set_text_color(0, 0, 0)
        
        col_data = next((c for c in df_evol.columns if 'data' in c.lower()), None)
        if col_data and not df_evol.empty:
            inicio = df_evol[col_data].iloc[0]
            fim = df_evol[col_data].iloc[-1]
            pdf.ln(5)
            pdf.set_font("Arial", 'I', 10)
            pdf.cell(190, 8, f"Periodo Analisado: {inicio} ate {fim}", ln=1)
    else:
        pdf.cell(190, 8, "Nao foi possivel ler os dados da planilha de evolucao.", ln=1)

    pdf.output("Relatorio_BI_Final_CLS.pdf")
    print("\n--- SUCESSO! Relatorio_BI_Final_CLS.pdf gerado ---")

df_v = carregar_planilha_vendas()
df_e = carregar_planilha_evolucao()

if df_v is not None:
    gerar_pdf_consolidado(df_v, df_e)
else:
    print("Erro: Planilha diaria nao encontrada.") 