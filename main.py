import customtkinter as ctk
import urllib.request
import os
import sys
import subprocess
from tkinter import filedialog, messagebox
import threading
import time # Adicionado para controle de tempo
from datetime import datetime, timedelta # Adicionado para cálculos de hora
from robo import RoboAutomacao
from seguranca import GerenciadorDados

# COLOQUE A FUNÇÃO AQUI (Antes de criar a janela e os botões)
def atualizar_sistema():
    # Link oficial atualizado para o seu repositório no GitHub
    url_arquivo_novo = "https://raw.githubusercontent.com/opalmeida2017-art/Robo-Exportacao-Plani/main/dist/main.exe"
    nome_atual = os.path.basename(sys.argv[0]) 
    nome_temp = "robo_atualizado.exe"

    try:
        print("Baixando nova versão do GitHub... Aguarde.")
        urllib.request.urlretrieve(url_arquivo_novo, nome_temp)

        script_bat = f"""@echo off
timeout /t 2 /nobreak > NUL
del "{nome_atual}"
ren "{nome_temp}" "{nome_atual}"
start "" "{nome_atual}"
del "%~f0"
"""
        with open("atualizador.bat", "w") as bat_file:
            bat_file.write(script_bat)

        print("Atualização concluída! Reiniciando o sistema...")
        
        subprocess.Popen("atualizador.bat", shell=True)
        sys.exit()

    except Exception as e:
        print(f"Erro ao tentar atualizar: {e}")

# ... (Classe RedirecionadorLog permanece igual) ...
class RedirecionadorLog:
    def __init__(self, caixa_texto):
        self.caixa_texto = caixa_texto

    def write(self, texto):
        self.caixa_texto.configure(state="normal")
        self.caixa_texto.insert("end", texto)
        self.caixa_texto.configure(state="disabled")
        self.caixa_texto.see("end")

    def flush(self):
        pass

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.seguranca = GerenciadorDados()
        
        self.title("Robô de Exportação SAT v1.5 - Dashboard")
        self.geometry("950x650") 
        ctk.set_appearance_mode("dark")

        # Variável para controlar se o robô deve parar
        self.executando = False

        # --- LAYOUT (Mesma estrutura que você enviou) ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.frame_esquerda = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frame_esquerda.pack(side="left", fill="y", expand=False, padx=(0, 20))

        self.frame_direita = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frame_direita.pack(side="right", fill="both", expand=True)

        self.label_titulo = ctk.CTkLabel(self.frame_esquerda, text="Automação de Planilhas", font=("Roboto", 22, "bold"))
        self.label_titulo.pack(pady=(0, 15))

        self.input_url = self.criar_campo("URL do Sistema:")
        self.input_user = self.criar_campo("Usuário:")
        self.input_pass = self.criar_campo("Senha:", show="*")
        self.input_relatorios = self.criar_campo("Códigos (Ex: 4, 82, 65):")
        
        self.frame_data = ctk.CTkFrame(self.frame_esquerda, fg_color="transparent")
        self.frame_data.pack(pady=5)

        self.meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        self.anos = [str(ano) for ano in range(2024, 2031)]

        self.label_mes = ctk.CTkLabel(self.frame_data, text="Mês de Referência:")
        self.label_mes.grid(row=0, column=0, padx=(0, 10), pady=(2, 0), sticky="w")
        self.input_mes = ctk.CTkComboBox(self.frame_data, width=200, values=self.meses)
        self.input_mes.grid(row=1, column=0, padx=(0, 10), pady=2)

        self.label_ano = ctk.CTkLabel(self.frame_data, text="Ano:")
        self.label_ano.grid(row=0, column=1, padx=(10, 0), pady=(2, 0), sticky="w")
        self.input_ano = ctk.CTkComboBox(self.frame_data, width=130, values=self.anos)
        self.input_ano.grid(row=1, column=1, padx=(10, 0), pady=2)
        
        self.input_minutos = self.criar_campo("Rodar a cada X minutos (0 = uma vez):")
        self.input_minutos.insert(0, "0")
        
        self.label_pasta = ctk.CTkLabel(self.frame_esquerda, text="Salvar Planilhas em:")
        self.label_pasta.pack(pady=(5, 0))
        self.frame_p = ctk.CTkFrame(self.frame_esquerda, fg_color="transparent")
        self.frame_p.pack(pady=2)
        self.entry_pasta = ctk.CTkEntry(self.frame_p, width=300)
        self.entry_pasta.pack(side="left", padx=5)
        self.btn_p = ctk.CTkButton(self.frame_p, text="...", width=40, command=self.selecionar_pasta)
        self.btn_p.pack(side="left")

        # Botão modificado para permitir Parar o robô
        self.btn_iniciar = ctk.CTkButton(self.frame_esquerda, text="INICIAR ROBÔ", fg_color="green", 
                                         hover_color="darkgreen", font=("Roboto", 16, "bold"),
                                         height=40, command=self.alternar_robo)
        self.btn_iniciar.pack(pady=(20, 0))

        self.label_log = ctk.CTkLabel(self.frame_direita, text="Log de Processamento:", font=("Roboto", 16, "bold"))
        self.label_log.pack(anchor="nw", pady=(0, 5))
        
        self.caixa_log = ctk.CTkTextbox(self.frame_direita, fg_color="#1e1e1e", text_color="#00fa9a", font=("Consolas", 12))
        self.caixa_log.pack(fill="both", expand=True)
        self.caixa_log.configure(state="disabled")

        sys.stdout = RedirecionadorLog(self.caixa_log)
        self.preencher_dados_salvos()

    def criar_campo(self, texto, **kwargs):
        label = ctk.CTkLabel(self.frame_esquerda, text=texto)
        label.pack(pady=(2, 0))
        entry = ctk.CTkEntry(self.frame_esquerda, width=350, **kwargs)
        entry.pack(pady=2)
        return entry

    def selecionar_pasta(self):
        caminho = filedialog.askdirectory()
        if caminho:
            self.entry_pasta.delete(0, "end")
            self.entry_pasta.insert(0, caminho)

    def preencher_dados_salvos(self):
        dados = self.seguranca.carregar()
        if dados:
            self.input_url.insert(0, dados.get('url', ''))
            self.input_user.insert(0, dados.get('user', ''))
            self.input_pass.insert(0, dados.get('pass', ''))
            self.input_relatorios.insert(0, dados.get('relt', ''))
            self.entry_pasta.insert(0, dados.get('path', ''))
            if 'mes' in dados: self.input_mes.set(dados['mes'])
            if 'ano' in dados: self.input_ano.set(dados['ano'])
            if 'minutos' in dados: 
                self.input_minutos.delete(0, "end")
                self.input_minutos.insert(0, dados['minutos'])

    # --- NOVA FUNÇÃO PARA ALTERNAR INICIAR/PARAR ---
    def alternar_robo(self):
        if not self.executando:
            self.executando = True
            self.btn_iniciar.configure(text="PARAR ROBÔ", fg_color="red", hover_color="darkred")
            self.iniciar_thread()
        else:
            self.executando = False
            self.btn_iniciar.configure(text="INICIAR ROBÔ", fg_color="green", hover_color="darkgreen")
            print("\n[!] Comando de parada enviado. O robô irá parar após a rodada atual ou espera.")

    def iniciar_thread(self):
        dados = {
            "url": self.input_url.get(),
            "user": self.input_user.get(),
            "pass": self.input_pass.get(),
            "relt": self.input_relatorios.get(),
            "path": self.entry_pasta.get(),
            "mes": self.input_mes.get(),
            "ano": self.input_ano.get(),
            "minutos": self.input_minutos.get()
        }
        
        if not all([dados["url"], dados["user"], dados["pass"], dados["relt"], dados["path"]]):
            messagebox.showwarning("Atenção", "Preencha todos os campos!")
            self.executando = False
            self.btn_iniciar.configure(text="INICIAR ROBÔ", fg_color="green")
            return

        self.seguranca.salvar(dados)
        # Limpa o log ao iniciar
        self.caixa_log.configure(state="normal")
        self.caixa_log.delete("1.0", "end")
        self.caixa_log.configure(state="disabled")
        
        # Dispara o loop em segundo plano
        threading.Thread(target=self.loop_automacao, args=(dados,), daemon=True).start()

    # --- FUNÇÃO DE LOOP INFINITO ---
    def loop_automacao(self, dados):
        minutos_espera = int(dados.get('minutos', 0))

        while self.executando:
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Iniciando rodada...")
            
            # Instancia e executa o robô
            robo = RoboAutomacao(
                url=dados["url"], usuario=dados["user"], senha=dados["pass"], 
                relatorios=dados["relt"], pasta=dados["path"],
                mes=dados["mes"], ano=dados["ano"]
            )
            
            try:
                robo.executar()
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Rodada concluída.")
            except Exception as e:
                print(f"❌ Erro na execução: {e}")

            # Se minutos for 0, para após a primeira vez
            if minutos_espera <= 0:
                print("Modo execução única finalizado.")
                self.executando = False
                self.after(0, lambda: self.btn_iniciar.configure(text="INICIAR ROBÔ", fg_color="green"))
                break

            if self.executando:
                proxima = datetime.now() + timedelta(minutes=minutos_espera)
                print(f"\n⏱️ Aguardando {minutos_espera} min. Próxima: {proxima.strftime('%H:%M:%S')}")
                
                # Contagem regressiva interna
                for i in range(minutos_espera * 60, 0, -1):
                    if not self.executando: break # Para se o usuário clicar em Parar
                    mins, segs = divmod(i, 60)
                    # Use stdout.write direto para evitar que o RedirecionadorLog crie milhares de linhas
                    # Ou apenas use print simples se preferir ver o histórico
                    if i % 10 == 0 or i <= 5: # Loga a cada 10 seg ou nos ultimos 5
                         print(f"Reiniciando em: {mins:02d}:{segs:02d}...", end="\r")
                    time.sleep(1)
                
                if self.executando:
                    print("\n🔄 Reiniciando agora...")

if __name__ == "__main__":
    app = App()
    app.mainloop()