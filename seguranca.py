from cryptography.fernet import Fernet
import os
import json

class GerenciadorDados:
    def __init__(self, arquivo_config="config.json", arquivo_chave="chave.key"):
        self.arquivo_config = arquivo_config
        self.arquivo_chave = arquivo_chave
        self.chave = self._carregar_ou_gerar_chave()
        self.fernet = Fernet(self.chave)

    def _carregar_ou_gerar_chave(self):
        # Cria uma chave única se ela não existir
        if not os.path.exists(self.arquivo_chave):
            chave = Fernet.generate_key()
            with open(self.arquivo_chave, "wb") as f:
                f.write(chave)
            return chave
        with open(self.arquivo_chave, "rb") as f:
            return f.read()

    def criptografar(self, texto):
        if not texto: return ""
        return self.fernet.encrypt(texto.encode()).decode()

    def descriptografar(self, texto_crip):
        if not texto_crip: return ""
        try:
            return self.fernet.decrypt(texto_crip.encode()).decode()
        except:
            return ""

    def salvar(self, dados):
        # Criptografa apenas a senha antes de salvar no JSON
        dados_protegidos = dados.copy()
        dados_protegidos['pass'] = self.criptografar(dados['pass'])
        with open(self.arquivo_config, "w") as f:
            json.dump(dados_protegidos, f, indent=4)

    def carregar(self):
        if os.path.exists(self.arquivo_config):
            with open(self.arquivo_config, "r") as f:
                dados = json.load(f)
                # Descriptografa a senha para preencher na tela
                dados['pass'] = self.descriptografar(dados.get('pass', ''))
                return dados
        return None