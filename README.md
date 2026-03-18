# teste-pywhatkit


import pandas as pd
import os                                                                                                                                                                                                           
import pywhatkit as kit
from datetime import datetime, timedelta
import time

# Mapeamento de Empresas para Grupos do WhatsApp
# Use o nome do grupo EXATAMENTE como aparece no WhatsApp
GRUPOS_WHATSAPP = {
    "Tropea": "Family",      # Nome do grupo para Tropea
    "Colina": "Family"        # Nome do grupo para Colina
}

def enviar_alerta_whatsapp(conteudo, nome_grupo):
    """
    Envia alerta para WhatsApp usando pywhatkit
    
    Args:
        conteudo: Mensagem a enviar
        nome_grupo: Nome do grupo (ex: "Family")
    """
    try:
        # Enviar instantaneamente para o grupo (requer WhatsApp Web aberto)
        kit.sendwhatmsg_instantly(nome_grupo, conteudo, tab_close=False)
        print(f"✅ Mensagem enviada para {nome_grupo}")
        
        # Pequena pausa entre mensagens para evitar bloqueio
        time.sleep(2)
        
    except Exception as e:
        print(f"❌ Erro ao enviar para WhatsApp ({nome_grupo}): {e}")

def processar_planilhas(diretorio):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Mapeamento estrito conforme seus dois modelos
    sinonimos_nome = ['nome', 'cliente']
    sinonimos_faturamento = ['patrimonio liquido total', 'patrimônio xp']
    sinonimos_saldo = ['saldo', 'saldo disponível']
    sinonimos_assessor = ['assessor']

    def encontrar_coluna(colunas, sinonimos):
        """Procura coluna por correspondência parcial (substring)"""
        for col in colunas:
            col_normalizado = str(col).lower().replace('ç', 'c').replace('ã', 'a').replace('á', 'a').replace('ó', 'o').replace('é', 'e').strip()
            for sinonimo in sinonimos:
                sinonimo_normalizado = sinonimo.lower().replace('ç', 'c').replace('ã', 'a').replace('á', 'a').replace('ó', 'o').replace('é', 'e')
                # Verifica se o sinônimo está contido na coluna (substring matching)
                if sinonimo_normalizado in col_normalizado or col_normalizado in sinonimo_normalizado:
                    return col
        return None

    if not os.path.exists(diretorio):
        print(f"❌ Erro: Pasta '{diretorio}' não encontrada.")
        return

    for arquivo in os.listdir(diretorio):
        # Pula arquivos temporários do Excel (~$) que ficam bloqueados
        if arquivo.startswith('~$'):
            print(f"⏭️ Pulando arquivo temporário: {arquivo}")
            continue
            
        if arquivo.endswith(('.xlsx', '.xls', '.csv')):
            print(f"🔍 Analisando: {arquivo}...")
            caminho = os.path.join(diretorio, arquivo)
            
            try:
                df = None
                
                # Estratégia 1: Tenta ler normalmente
                try:
                    df = pd.read_excel(caminho) if not arquivo.endswith('.csv') else pd.read_csv(caminho)
                except:
                    pass
                
                # Estratégia 2: Se não conseguiu ou encontrou colunas vazias, tenta múltiplos skiprows
                if df is None or "unnamed" in str(df.columns).lower() or df.columns[0] == '':
                    print(f"      🔧 Detectado cabeçalho inválido, testando skiprows...")
                    
                    # Tenta pular de 1 até 10 linhas
                    for skip in range(1, 11):
                        try:
                            df_teste = pd.read_excel(caminho, skiprows=skip) if not arquivo.endswith('.csv') else pd.read_csv(caminho, skiprows=skip)
                            
                            # Verifica se as colunas são válidas (não vazias, não unnamed)
                            colunas_str = str(df_teste.columns).lower()
                            if "unnamed" not in colunas_str and df_teste.columns[0] != '' and len(df_teste.columns) >= 3:
                                df = df_teste
                                print(f"      ✅ Cabeçalho válido encontrado pulando {skip} linhas!")
                                break
                        except:
                            continue
                
                # Se ainda não achou, tenta sem cabeçalho e usa primeira linha não-vazia como header
                if df is None or "unnamed" in str(df.columns).lower():
                    print(f"      🔧 Tentando ler sem cabeçalho...")
                    df_raw = pd.read_excel(caminho, header=None) if not arquivo.endswith('.csv') else pd.read_csv(caminho, header=None)
                    
                    # Encontra primeira linha com dados
                    for idx in range(len(df_raw)):
                        if any(pd.notna(v) and str(v).strip() != '' for v in df_raw.iloc[idx]):
                            df = df_raw.iloc[idx:].reset_index(drop=True)
                            df.columns = df.iloc[0]
                            df = df.iloc[1:].reset_index(drop=True)
                            print(f"      ✅ Cabeçalho sem nomes usado da linha {idx}!")
                            break

                # Normaliza nomes das colunas
                df.columns = [str(c).lower().strip() for c in df.columns]

                # Identificação das colunas (com matching melhorado)
                c_nome = encontrar_coluna(df.columns, sinonimos_nome)
                c_faturamento = encontrar_coluna(df.columns, sinonimos_faturamento)
                c_saldo = encontrar_coluna(df.columns, sinonimos_saldo)
                c_assessor = encontrar_coluna(df.columns, sinonimos_assessor)

                if not c_nome or not c_saldo or not c_faturamento:
                    print(f"   ⚠️ Pulei {arquivo}: Colunas não batem com os modelos")
                    print(f"      Nome: {c_nome} | Saldo: {c_saldo} | Fat: {c_faturamento}")
                    print(f"      Colunas encontradas: {list(df.columns)[:5]}")
                    continue

                # Limpeza e conversão numérica
                df[c_saldo] = pd.to_numeric(df[c_saldo], errors='coerce').fillna(0)
                df[c_faturamento] = pd.to_numeric(df[c_faturamento], errors='coerce').fillna(0)

                # Cálculo da relação Saldo / Faturamento (Margem Individual)
                df['pct_calculada'] = df[c_saldo] / df[c_faturamento].replace(0, float('nan'))
                df['pct_calculada'] = df['pct_calculada'].fillna(0) 

                # Filtra >= 0.5%
                vips = df[df['pct_calculada'] >= 0.005].copy()
                
                # Ordena por percentual decrescente (maior % primeiro)
                vips = vips.sort_values('pct_calculada', ascending=False)
                
                print(f"   ✅ Processado | {len(vips)} clientes VIPs encontrados (ordenados por % decrescente).")

                for _, linha in vips.iterrows():
                    # Lógica de Empresa (Tropea ou Colina)
                    assessor_txt = str(linha.get(c_assessor, "")).lower()
                    empresa = "Tropea" if "pedro" in assessor_txt or "tropea" in arquivo.lower() else "Colina"
                    
                    msg = (f"*DATA DA ANÁLISE: {agora}*\n"
                           f"*ALERTA: RELAÇÃO SALDO/FATURAMENTO - {empresa}*\n"
                           f"*CLIENTE:* {str(linha[c_nome]).upper()}\n"
                           f"*PATRIMÔNIO (BASE):* R$ {linha[c_faturamento]:,.2f}\n"
                           f"*SALDO DISPONÍVEL:* R$ {linha[c_saldo]:,.2f}\n"
                           f"*% (Saldo/PL):* {linha['pct_calculada']*100:.2f}%\n"
                           f"*ARQUIVO:* {arquivo}\n"
                           f"--------------------------------------")
                    
                    # Identifica qual grupo enviar baseado na empresa
                    nome_grupo = GRUPOS_WHATSAPP.get(empresa)
                    if nome_grupo:
                        enviar_alerta_whatsapp(msg, nome_grupo)
                    else:
                        print(f"   ⚠️ Empresa '{empresa}' não mapeada em GRUPOS_WHATSAPP")
                    
            except Exception as e:
                print(f"   ❌ Erro ao processar {arquivo}: {e}")

# Execução
processar_planilhas(r".\arquivos_clientes")
