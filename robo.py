from playwright.sync_api import sync_playwright
import time
import os
import calendar
from datetime import datetime
import sys

# =========================================================
# CONFIGURAÇÃO GLOBAL DOS FILTROS E DESCRIÇÕES PARA O LOG
# =========================================================
CONFIG_FILTROS = {
    "Relatório de Viagens por Cliente": {
        "prefixo": "RelFilViagensCliente_",
        "botao_gerar": "filtrarrelFilViagensCliente",
        "arquivo_nome": "relFilViagensCliente.xls",
        "regras": {
            'tipoData': '1', 'emitido': '0', 'tipoPesoChegada': '0', 'tipoFrete': '0',
            'pagtoFrete': 'T', 'tipoCalculoFreteEmpresa': 'C', 'usaICMSFinal': 'N',
            'tipoCte': 'T', 'cteStatus': '2', 'cancelado': '99', 'agrupar': '3',
            'viagemGrupo': 'T', 'mostrarDocAnt': 'N', 'averbado': 'T',
            'mercadoriaExportacao': 'T', 'imprimirFreteEmpresa': 'S',
            'imprimirFreteMotorista': 'S', 'permiteFaturar': 'T',
            'impPPTEmpMot': '1', 'impProp': 'N', 'comComplementar': 'N',
            'possuiPesoChegada': 'T', 'apenasFaturados': 'T', 'possuiPedido': 'T',
            'resumido': 'N', 'somenteQuebra': 'T', 'tipoCalcMargemFrete': '1',
            'mostrarMargemFreteConhec': 'N'
        },
        "descricoes": {
            'tipoData': 'Emissão', 'emitido': 'Todos', 'tipoPesoChegada': 'Todos', 'tipoFrete': 'Todos',
            'pagtoFrete': 'Todos', 'tipoCalculoFreteEmpresa': 'Chegada', 'usaICMSFinal': 'Não',
            'tipoCte': 'Todos', 'cteStatus': 'Autorizado', 'cancelado': 'Todos', 'agrupar': 'Não agrupar',
            'viagemGrupo': 'Todas', 'mostrarDocAnt': 'Não', 'averbado': 'Todos',
            'mercadoriaExportacao': 'Todas', 'imprimirFreteEmpresa': 'Sim',
            'imprimirFreteMotorista': 'Sim', 'permiteFaturar': 'Todos',
            'impPPTEmpMot': 'Todos', 'impProp': 'Não', 'comComplementar': 'Não',
            'possuiPesoChegada': 'Todos', 'apenasFaturados': 'Todos', 'possuiPedido': 'Todos',
            'resumido': 'Não', 'somenteQuebra': 'Todos', 'tipoCalcMargemFrete': 'Margem Real',
            'mostrarMargemFreteConhec': 'Não'
        }
    },
    "Relatório de Despesas Gerais": {
        "prefixo": "RelFilDespesasGerais_",
        "botao_gerar": "formrelFilDespesasGerais:filtrar",
        "arquivo_nome": "relDespesasGerais.xls",
        "regras": {
            'tipoData': '1', 'layout': '1', 'despesa': 'T', 'investimento': 'T',
            'rateio': 'T', 'finalizada': 'T', 'tipoDespesa': '10', 'faturada': 'T',
            'serieRQ': 'T', 'agruparPorGrupoD': 'N', 'resumidoStr': 'N', 'ordem': '1',
            'mostrarItemDet': 'N', 'mostrarObs': 'N', 'mostrarValoresRateados': 'N'
        },
        "descricoes": {
            'tipoData': 'Controle', 'layout': 'Paisagem', 'despesa': 'Todas', 'investimento': 'Todas',
            'rateio': 'Todas', 'finalizada': 'Todas', 'tipoDespesa': 'Geral', 'faturada': 'Todas',
            'serieRQ': 'Todas', 'agruparPorGrupoD': 'Não', 'resumidoStr': 'Não', 'ordem': 'Filial/Item/Nota',
            'mostrarItemDet': 'Não', 'mostrarObs': 'Não', 'mostrarValoresRateados': 'Não'
        }
    }
}

class RoboAutomacao:
    def __init__(self, url, usuario, senha, relatorios, pasta, mes, ano):
        self.url = url
        self.usuario = usuario
        self.senha = senha
        self.relatorios = [r.strip() for r in relatorios.split(",") if r.strip()]
        self.pasta = os.path.abspath(os.path.normpath(pasta))
        self.mes = mes
        self.ano = ano

    def executar(self):
        if not os.path.exists(self.pasta):
            os.makedirs(self.pasta, exist_ok=True)
            print(f" -> Pasta de destino verificada: {self.pasta}")

        # --- CÁLCULO DAS DATAS ---
        meses_map = {
            "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
            "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
        }
        mes_num = meses_map.get(self.mes, 1)
        ano_num = int(self.ano)
        ultimo_dia = calendar.monthrange(ano_num, mes_num)[1]
        str_data_ini = f"01/{mes_num:02d}/{ano_num}"
        str_data_fim = f"{ultimo_dia:02d}/{mes_num:02d}/{ano_num}"

        with sync_playwright() as p:
            # O channel="chrome" força o robô a usar o seu Google Chrome real de forma invisível!
            browser = p.chromium.launch(headless=True, channel="chrome")
            context = browser.new_context(accept_downloads=True)
            page = context.new_page() # LINHA ADICIONADA E CORRIGIDA

            try:
                # --- PASSO 1: LOGIN ---
                print(f"\nAcessando: {self.url}")
                page.goto(self.url)
                page.wait_for_load_state("load")
                page.locator('input[type="text"]').first.fill(self.usuario)
                page.locator('input[type="password"]').first.fill(self.senha)
                page.locator('input[value="Entrar"]').click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)

                # --- PASSO 2: NAVEGAÇÃO ---
                print("Navegando para Cadastro de Exportações...")
                page.locator('text="Exp./Imp."').first.click()
                time.sleep(1)
                page.locator('text="Cadastro de Exportações"').first.click()
                page.wait_for_load_state("networkidle")

                # --- PASSO 3: LOOP DE RELATÓRIOS ---
                for relatorio in self.relatorios:
                    try:
                        print(f"\n{'-'*50}\n>>> Processando Código: {relatorio}")
                        
                        campo_codigo = page.locator('text=Código:').locator('xpath=./following::input[1]')
                        campo_codigo.click()
                        campo_codigo.fill("")
                        campo_codigo.fill(relatorio)
                        page.locator('img[src*="search"], img[src*="lupa"], input[type="image"]').first.click()
                        
                        selector_codigo = f"td:has-text('{relatorio}')"
                        page.wait_for_selector(selector_codigo, timeout=15000)
                        
                        print(f"Abrindo detalhes...")
                        link_resultado = page.locator(selector_codigo).locator("xpath=./following-sibling::td[1]").locator("a")
                        link_resultado.evaluate("node => node.click()")
                        page.wait_for_load_state("networkidle")
                        time.sleep(2)

                        with context.expect_page() as new_page_info:
                            page.locator('text="Exportar Dados"').evaluate("node => node.click()")
                        nova_aba = new_page_info.value
                        nova_aba.wait_for_load_state("networkidle")

                        nova_aba.locator('text="### Link para exportação ###"').first.evaluate("node => node.click()")
                        nova_aba.wait_for_load_state("networkidle")
                        time.sleep(2)
                        
                        header = nova_aba.locator("div.rf-p-hdr").first.inner_text().strip()
                        print(f"-> Tela Identificada: '{header}'")

                        # --- VERIFICAÇÃO DE RELATÓRIO CONFIGURADO ---
                        if header not in CONFIG_FILTROS:
                            print(f"⚠️ ERRO CRÍTICO: Relatório '{header}' (Código {relatorio}) não está mapeado no sistema.")
                            print(" -> Remova este código ou contate o administrador do sistema para configurar os filtros.")
                            browser.close()
                            return # Encerra o processo inteiro
                        
                        # Carrega a configuração correta baseada na tela
                        conf = CONFIG_FILTROS[header]

                        # --- LOG DOS FILTROS ESPERADOS ---
                        print(f"\n📋 Status Alvo (Filtros que serão verificados e ajustados):")
                        for chave, valor_regra in conf["regras"].items():
                            descricao = conf["descricoes"].get(chave, "Sem Descrição")
                            print(f"   - {chave}: {valor_regra} ({descricao})")
                        print("-" * 50)

                        # --- PREENCHIMENTO DE DATAS ---
                        id_data_ini = f'{conf["prefixo"]}dataIniInputDate'
                        id_data_fim = f'{conf["prefixo"]}dataFimInputDate'
                        
                        nova_aba.locator(f'input[id*="{id_data_ini}"]').fill(str_data_ini)
                        nova_aba.locator(f'input[id*="{id_data_fim}"]').fill(str_data_fim)
                        nova_aba.locator(f'input[id*="{id_data_fim}"]').press("Tab")
                        time.sleep(1.5)

                        # --- APLICAÇÃO INTELIGENTE DOS FILTROS ---
                        print("Verificando seleções na tela...")
                        for id_p, val_alvo in conf["regras"].items():
                            campo = nova_aba.locator(f'select[id*="{conf["prefixo"]}{id_p}"]').first
                            if campo.count() > 0:
                                atual = campo.evaluate("n => n.options[n.selectedIndex].value")
                                if atual != val_alvo:
                                    campo.select_option(value=val_alvo)
                                    print(f" -> 🔄 CORRIGIDO na tela: '{id_p}' de '{atual}' para '{val_alvo}'")
                                    time.sleep(1.2)

                        # --- DOWNLOAD COM TRATAMENTO PARA PLANILHAS VAZIAS ---
                        print("\nIniciando processamento (Gerando arquivo...)")
                        nova_aba.locator(f'input[id*="{conf["botao_gerar"]}"]').first.click(force=True, no_wait_after=True)
                        
                        try:
                            # Mantendo o tempo original de 90 segundos (90000 ms)
                            selector_link = nova_aba.get_by_text("Clique aqui para visualizar arquivo", exact=False)
                            selector_link.wait_for(state="visible", timeout=90000)
                            
                            with nova_aba.expect_download(timeout=60000) as download_info:
                                selector_link.click()
                            
                            # --- BACKUP ---
                            nome_base = conf["arquivo_nome"]
                            caminho_completo = os.path.join(self.pasta, nome_base)
                            if os.path.exists(caminho_completo):
                                bkp = os.path.join(self.pasta, f"{nome_base.replace('.xls','')}_bkp_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xls")
                                os.rename(caminho_completo, bkp)
                                print(f" -> Backup antigo criado: {bkp}")
                            
                            download_info.value.save_as(caminho_completo)
                            print(f"✅ SUCESSO: Arquivo atualizado em {caminho_completo}")

                        except Exception as e_download:
                            if "Timeout" in str(e_download):
                                print("\n⚠️ AVISO: O link de download não apareceu no tempo esperado (90s).")
                                print(" -> POSSÍVEIS CAUSAS:")
                                print("    1. O PERÍODO selecionado não possui nenhum dado (relatório vazio).")
                                print("    2. Os filtros selecionados bloquearam todos os resultados.")
                                print("    3. Lentidão extrema no servidor do sistema SAT.")
                                print(" -> Pulando para o próximo...\n")
                            else:
                                print(f"❌ Erro durante o download: {e_download}")

                        finally:
                            # Garante que a aba secundária sempre será fechada, com sucesso ou erro
                            if 'nova_aba' in locals() and not nova_aba.is_closed():
                                nova_aba.close()

                        # --- RETORNO PARA NOVA PESQUISA ---
                        page.bring_to_front()
                        page.locator('text="Exp./Imp."').first.click()
                        page.locator('text="Cadastro de Exportações"').first.click()
                        page.wait_for_load_state("networkidle")

                    except Exception as e_item:
                        print(f"❌ Erro no item {relatorio}: {e_item}")
                        if 'nova_aba' in locals() and not nova_aba.is_closed(): nova_aba.close()
                        continue

            except Exception as e: print(f"ERRO GERAL: {e}")
            finally:
                if 'browser' in locals():
                    print("\nFechando navegador e finalizando lote atual...")
                    browser.close() # FECHAMENTO CORRIGIDO