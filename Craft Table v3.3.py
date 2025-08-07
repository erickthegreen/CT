import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, font
import webbrowser
import json
import os
from datetime import datetime
import logging
from typing import Dict, List, Optional, Any
import sys
import getpass
import gc
from tkcalendar import Calendar, DateEntry

# Configurar logging
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(log_dir, 'atendimento_app.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ToastNotification:

    def __init__(self, parent, title, message, duration=3000):
        self.parent = parent
        self.duration = duration
        
        self.window = tk.Toplevel(parent)
        self.window.overrideredirect(True)
        
        # Tenta obter o estilo do tema, com um fallback para cores padrão
        try:
            style = ttk.Style()
            bg_color = style.lookup("TFrame", "background")
            fg_color = style.lookup("TLabel", "foreground")
        except tk.TclError:
            bg_color = "#f0f0f0"
            fg_color = "black"

        self.window.configure(background=bg_color)
        self.window.attributes("-alpha", 0.0)

        frame = ttk.Frame(self.window, padding=(15, 10))
        frame.pack(expand=True, fill="both")

        ttk.Label(frame, text=title, font=("Arial", 11, "bold")).pack(anchor="w")
        ttk.Label(frame, text=message, wraplength=280, justify="left").pack(anchor="w", pady=(5, 0))

        self.window.update_idletasks()
        
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        
        win_width = self.window.winfo_width()
        x = parent_x + parent_width - win_width - 20
        y = parent_y + parent.winfo_height() - self.window.winfo_height() - 20

        self.window.geometry(f"+{x}+{y}")
        
        # O __init__ chama o fade_in
        self.fade_in()

    # IMPORTANTE: fade_in precisa estar indentado no mesmo nível do __init__
    def fade_in(self):
        alpha = self.window.attributes("-alpha")
        if alpha < 0.95:
            alpha += 0.08
            self.window.attributes("-alpha", alpha)
            self.window.after(20, self.fade_in)
        else:
            self.window.after(self.duration, self.fade_out)

    # IMPORTANTE: fade_out também precisa estar indentado no mesmo nível
    def fade_out(self):
        alpha = self.window.attributes("-alpha")
        if alpha > 0.0:
            alpha -= 0.08
            self.window.attributes("-alpha", alpha)
            self.window.after(20, self.fade_out)
        else:
            self.window.destroy()

class HistoryManager:
    """Classe para gerenciar o histórico de atendimentos em arquivo."""
    def __init__(self, caminho_arquivo):
        self.caminho_historico = caminho_arquivo
        self.historico_registros = {"Emergenciais": [], "Comerciais": [], "Informação": [], "Reclamações": []}
        self.carregar_historico()

    def carregar_historico(self):
        """Carrega o histórico de registros do arquivo JSON."""
        try:
            if os.path.exists(self.caminho_historico):
                with open(self.caminho_historico, "r", encoding="utf-8") as f:
                    self.historico_registros = json.load(f)
        except (json.JSONDecodeError, Exception) as e:
            logger.error(f"Erro ao carregar ou decodificar o histórico: {e}")
            self.historico_registros = {"Emergenciais": [], "Comerciais": [], "Informação": [], "Reclamações": []}

    def salvar_historico(self):
        """Salva o histórico de registros no arquivo JSON."""
        try:
            with open(self.caminho_historico, "w", encoding="utf-8") as f:
                json.dump(self.historico_registros, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logger.error(f"Erro ao salvar histórico: {e}")

    def adicionar_registro(self, categoria, registro):
        """Adiciona um novo registro ao histórico e salva."""
        if categoria not in self.historico_registros:
            self.historico_registros[categoria] = []
        self.historico_registros[categoria].append(registro)
        self.salvar_historico()

    def reset_historico(self):
        """Limpa todo o histórico e salva o estado vazio."""
        self.historico_registros = {"Emergenciais": [], "Comerciais": [], "Informação": [], "Reclamações": []}
        self.salvar_historico()
        logger.info("Histórico resetado pelo usuário")

class PausaInteligente:
    def __init__(self, parent_app):
        self.app = parent_app
        self.pausas_config = {
            "1ª P 10": {"horario_var": tk.StringVar(), "duracao": 10, "alertado": False, "widgets": {}},
            "P 20": {"horario_var": tk.StringVar(), "duracao": 20, "alertado": False, "widgets": {}},
            "2ª P 10": {"horario_var": tk.StringVar(), "duracao": 10, "alertado": False, "widgets": {}}
        }
        self.cronometro_ativo = None
    
    def criar_painel_pausas(self, parent):
        """Cria o painel de pausas no layout principal (versão mais compacta e centralizada)"""
        # Header com botão de limpar
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill="x", pady=2, padx=5)
        
        ttk.Label(header_frame, text="🕐 Pausas", font=("Arial", 11, "bold")).pack(side="left")
        ttk.Button(header_frame, text="Limpar Pausas", command=self.limpar_pausas).pack(side="left")
        
        # NOVO: Frame para centralizar o conteúdo das pausas
        centering_frame = ttk.Frame(parent)
        centering_frame.pack(fill="x", pady=(0, 5))

        # MUDANÇA: O frame das pausas agora fica dentro do frame de centralização
        pausas_frame = ttk.Frame(centering_frame)
        pausas_frame.pack() # .pack() sem argumentos centraliza o widget
        
        for nome, config in self.pausas_config.items():
           
            pausa_container = ttk.Frame(pausas_frame)
            pausa_container.pack(side="left", padx=5) # Apenas 5 pixels de espaço entre cada um
            
            # Label da pausa
            config["widgets"]["label"] = ttk.Label(pausa_container, text=f"{nome}:", font=("Arial", 10))
            config["widgets"]["label"].pack(side="left", pady=2)
            
            # Entry para horário
            config["widgets"]["entry"] = ttk.Entry(pausa_container, textvariable=config["horario_var"], width=7, justify="center")
            config["widgets"]["entry"].pack(side="left", padx=(0, 5), pady=2) # Padding ajustado
            config["widgets"]["entry"].bind("<KeyRelease>", lambda e, p=nome: self.validar_horario(p))
            
            # Frame para status/botão
            config["widgets"]["status_frame"] = ttk.Frame(pausa_container, height=30, width=85)
            config["widgets"]["status_frame"].pack(side="left", pady=2)
            config["widgets"]["status_frame"].pack_propagate(False)
    
    def validar_horario(self, nome_pausa):
        """Valida formato do horário digitado"""
        horario = self.pausas_config[nome_pausa]["horario_var"].get()
        if self.is_horario_valido(horario):
            self.pausas_config[nome_pausa]["alertado"] = False
    
    def is_horario_valido(self, horario):
        """Verifica se horário está no formato HH:MM válido"""
        import re
        if not re.match(r"^\d{2}:\d{2}$", horario):
            return False
        try:
            hora, minuto = map(int, horario.split(":"))
            return 0 <= hora <= 23 and 0 <= minuto <= 59
        except:
            return False
    
    def verificar_horarios_pausas(self):
        """Verifica se alguma pausa deve ser alertada (chamado a cada minuto)"""
        hora_atual = datetime.now().strftime("%H:%M")
        
        for nome, config in self.pausas_config.items():
            horario_agendado = config["horario_var"].get()
            
            if (horario_agendado == hora_atual and 
                not config["alertado"] and 
                self.is_horario_valido(horario_agendado)):
                
                self.ativar_alerta_pausa(nome)
        
        # Reagenda para próxima verificação
        self.app.root.after(60000, self.verificar_horarios_pausas)
    
    def ativar_alerta_pausa(self, nome_pausa):
        """Ativa alerta visual para uma pausa"""
        config = self.pausas_config[nome_pausa]
        config["alertado"] = True
        
        # Muda cor do label para vermelho
        config["widgets"]["label"].config(foreground="red", font=("Arial", 10, "bold"))
        
        # Limpa frame de status e adiciona botão
        status_frame = config["widgets"]["status_frame"]
        for widget in status_frame.winfo_children():
            widget.destroy()
        
        btn_iniciar = ttk.Button(
            status_frame, 
            text="▶ Iniciar", 
            command=lambda: self.iniciar_pausa_manual(nome_pausa),
            style="Destaque.TButton"
        )
        btn_iniciar.pack()
        
        # Toca som de notificação e mostra popup
        self.app.root.bell()
        messagebox.showinfo("🔔 Hora da Pausa!", f"Atenção! Está na hora da sua {nome_pausa}.\nClique em 'Iniciar' quando estiver pronto!")
        
        logger.info(f"Alerta ativado para {nome_pausa}")
    
    def iniciar_pausa_manual(self, nome_pausa):
        """Inicia cronômetro da pausa quando atendente clicar"""
        if self.cronometro_ativo:
            if not messagebox.askyesno("Pausa Ativa", "Já há uma pausa em andamento. Deseja interrompê-la?"):
                return
            self.parar_cronometro_atual()
        
        config = self.pausas_config[nome_pausa]
        duracao_minutos = config["duracao"]
        
        # Esconde botão iniciar
        status_frame = config["widgets"]["status_frame"]
        for widget in status_frame.winfo_children():
            widget.destroy()
        
        # Cria label de cronômetro
        config["widgets"]["cronometro"] = ttk.Label(
            status_frame, 
            text="", 
            foreground="green", 
            font=("Arial", 10, "bold")
        )
        config["widgets"]["cronometro"].pack()
        
        # Inicia contagem regressiva
        self.cronometro_ativo = nome_pausa
        self.executar_cronometro(nome_pausa, duracao_minutos * 60)
        
        logger.info(f"Pausa {nome_pausa} iniciada - {duracao_minutos} minutos")
    
    def executar_cronometro(self, nome_pausa, segundos_restantes):
        """Executa contagem regressiva"""
        if self.cronometro_ativo != nome_pausa:
            return
        
        config = self.pausas_config[nome_pausa]
        
        if segundos_restantes >= 0:
            # Formata tempo MM:SS
            minutos, segundos = divmod(segundos_restantes, 60)
            tempo_str = f"{minutos:02d}:{segundos:02d}"
            
            # Atualiza display
            if "cronometro" in config["widgets"]:
                config["widgets"]["cronometro"].config(text=tempo_str)
            
            # Agenda próxima atualização
            self.app.root.after(1000, self.executar_cronometro, nome_pausa, segundos_restantes - 1)
        else:
            # Fim da pausa
            self.finalizar_pausa(nome_pausa)
    
    def finalizar_pausa(self, nome_pausa):
        """Finaliza a pausa e notifica o usuário"""
        config = self.pausas_config[nome_pausa]
        
        # Som + popup
        self.app.root.bell()
        messagebox.showinfo("⏰ Fim da Pausa!", f"Sua {nome_pausa} terminou!\nHora de voltar ao trabalho! 💪")
        
        # Reset visual
        config["widgets"]["label"].config(foreground="green", font=("Arial", 10, "bold"))
        
        # Limpa status
        status_frame = config["widgets"]["status_frame"]
        for widget in status_frame.winfo_children():
            widget.destroy()
        
        ttk.Label(status_frame, text="✅ OK", foreground="green", font=("Arial", 8)).pack()
        
        # Reset estado
        self.cronometro_ativo = None
        config["alertado"] = False
        
        logger.info(f"{nome_pausa} finalizada")
    
    def limpar_pausas(self):
        """Limpa todos os horários e alertas"""
        if not messagebox.askyesno("Confirmar", "Limpar todos os horários agendados?"):
            return
        
        for nome, config in self.pausas_config.items():
            # Limpa horário
            config["horario_var"].set("")
            
            # Reset estado
            config["alertado"] = False
            
            # Reset visual
            if "label" in config["widgets"]:
                config["widgets"]["label"].config(foreground="black", font=("Arial", 10, "bold"))
            
            # Limpa status
            if "status_frame" in config["widgets"]:
                status_frame = config["widgets"]["status_frame"]
                for widget in status_frame.winfo_children():
                    widget.destroy()
        
        # Para cronômetro se ativo
        if self.cronometro_ativo:
            self.cronometro_ativo = None
        
        logger.info("Pausas limpas")
    
    def parar_cronometro_atual(self):
        """Para o cronômetro atual"""
        if self.cronometro_ativo:
            nome = self.cronometro_ativo
            self.cronometro_ativo = None
            config = self.pausas_config[nome]
            
            # Limpa status
            if "status_frame" in config["widgets"]:
                status_frame = config["widgets"]["status_frame"]
                for widget in status_frame.winfo_children():
                    widget.destroy()
            
            # Reset visual
            if "label" in config["widgets"]:
                config["widgets"]["label"].config(foreground="black")
            
            logger.info(f"Cronômetro {nome} interrompido")

class AtendimentoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🛠️ CRAFT TABLE v3 🛠️ - A sua bancada de trabalho interativo e funcional")
        self.root.geometry("880x700")
        self.root.resizable(True, True)
        if getattr(sys, 'frozen', False):
            self.base_path = os.path.dirname(sys.executable)
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))


        self.style = ttk.Style()
        self.style.configure("Wide.Vertical.TScrollbar", width=25)

        # A linha que calcula o self.base_path foi REMOVIDA DAQUI, pois a movemos para cima.

        #self.pasta_anexos = os.path.join(self.base_path, "anexos")
        #if not os.path.exists(self.pasta_anexos):
        #    os.makedirs(self.pasta_anexos)

        # Inicialização de variáveis de estado
        self.tamanho_fonte_base = 10
        self.fonte_atual = font.Font(family="Arial", size=self.tamanho_fonte_base, weight="bold") # <-- MUDANÇA AQUI
        self.fonte_bold_atual = font.Font(family="Arial", size=self.tamanho_fonte_base, weight="bold")
        self.registro_usuario = tk.StringVar()
        self.tema_atual = "claro"
        self.cor_atual = "azul"
        self.favoritos = [None, None, None] 
        self.var_Genesys = tk.StringVar(value="FALTA DE ENERGIA INDIVIDUAL; O CLIENTE INFORMA QUE REALIZOU O TESTE DO DISJUNTOR")
        
        # NOVA INICIALIZAÇÃO: Sistema de Pausas Inteligente
        self.pausa_sistema = PausaInteligente(self)
        
        # Dicionários de dados - SERVIÇO 6 ALTERADO PARA RECLAMAÇÕES
        self.servicos = {
            "0": {"nome": "Geral", "categoria": "Informação"},
            "1": {"nome": "Serviços Emergenciais", "categoria": "Emergenciais"},
            "2": {"nome": "Reclamações", "categoria": "Reclamações"},
            "3": {"nome": "Gravação Telefônica", "categoria": "Informação"},
            "4": {"nome": "Desligamento Definitivo", "categoria": "Comerciais"},
            "5": {"nome": "Nível de Tensão", "categoria": "Reclamações"},
            "6": {"nome": "Denúncia", "categoria": "Reclamações"},  
            "7": {"nome": "Mudança de Data Certa", "categoria": "Comerciais"},
            "8": {"nome": "Cadastro Baixa Renda", "categoria": "Comerciais"},
            "9": {"nome": "Danos Elétricos", "categoria": "Reclamações"},
            "10": {"nome": "Religação", "categoria": "Comerciais"},
            "11": {"nome": "Informações", "categoria": "Informação"},
            "12": {"nome": "Serviços Emergenciais no ATC", "categoria": "Emergenciais"},
            "13": {"nome": "Genesys", "categoria": "Emergenciais"},
            "14": {"nome": "Cancelar Fatura por E-mail", "categoria": "Comerciais"},
            "15": {"nome": "Alterar Dados do Parceiro de Negócios", "categoria": "Comerciais"},
            "16": {"nome": "Problema com Equipamento", "categoria": "Reclamações"},
            "17": {"nome": "Mudança de Medidor de Local", "categoria": "Comerciais"},
            "18": {"nome": "Cancelamento de Atividades Acessórias", "categoria": "Comerciais"},
            "19": {"nome": "Danos Materiais", "categoria": "Reclamações"},
            "20": {"nome": "Inativação Baixa Renda", "categoria": "Comerciais"}
        }
        
        self.cor_categoria = {
            "Emergenciais": "#F16060", 
            "Comerciais": "#F0F880",
            "Informação": "#AEF3AE", 
            "Reclamações": "#AFB5FF",
        }
        
        self.descricoes_emergenciais = {
            "FALTA DE ENERGIA GERAL": "CLIENTE INFORMA FALTA DE ENERGIA GERAL; O CLIENTE INFORMA QUE REALIZOU O TESTE DO DISJUNTOR",
            "FALTA DE ENERGIA INDIVIDUAL": "CLIENTE INFORMA FALTA DE ENERGIA INDIVIDUAL; O CLIENTE INFORMA QUE REALIZOU O TESTE DO DISJUNTOR",
            "AVALIAÇÃO TÉCNICA": "CLIENTE INFORMA QUE A SUA ENERGIA ESTÁ OSCILAÇÃO FRACA/FORTE.",
            "RECHAMADA": "RECHAMADA - CLIENTE INFORMADO QUE JÁ EXISTE SOLICITAÇÃO EM ABERTO COM EQUIPE DE OPERAÇÕES.",
            "FALTA DE FASE": "CLIENTE INFORMA QUE ESTÁ COM FALTA DE FASE EM SUA INSTALAÇÃO; CLIENTE TAMBEM INFORMA FALTA DE ENERGIA EM ALGUNS CÔMODOS DA CASA.",
            "CHOQUE/VAZAMENTO:": "CLIENTE INFORMA QUE ESTÁ HAVENDO CHOQUE OU VAZAMENTO DE ENERGIA NA REDE, SOLICITA AGILIDADE.",
            "REDE PARTIDA": "CLIENTE INFORMA QUE ESTÁ COM REDE PARTIDA DE (BT/AT) OFERECENDO PERIGO.",
            "FAISCAMENTO": "FAISCAMENTO NA REDE, GERANDO RISCO A VIDA DE ENERGIA.",
            "RAMAL PARTIDO": "CLIENTE INFORMA QUE RAMAL DE SERVIÇO ESTÁ PARTIDO, SOLICITA AGILIDADE.",
            "INCÊNDIO": "CLIENTE INFORMA QUE ESTÁ HAVENDO INCÊNDIO NA REDE, ONDE COLOCA TERCEIROS EM RISCO, SOLICITA AGILIDADE.",
            "VÃO BAIXO": "CLIENTE INFORMA VÃO BAIXO COM RISCO DE ROMPIMENTO",
            "VAZAMENTO DE ÓLEO": "CLIENTE INFORMA DE ÓLEO NO TRANSFORMADOR.",
            "ABALROAMENTO": "CLIENTE INFORMA ABALROAMENTO DE POSTE OU EQUIPAMENTO.",
            "ILUMINAÇÃO PÚBLICA ACESA DURANTE O DIA": "CLIENTE INFORMA LUZ DO POSTE LIGADA DURANTE O DIA",
            "INTERVENÇÃO DE TERCEIROS NA REDE": "CLIENTE INFORMA QUE HOUVE INTERVENÇÃ DE TERCEIROS NA REDE"
        }
        
        self.opcoes_informacoes = [
            "INFORMAÇÕES CONEXÃO - LIGAÇÃO NOVA", "CADASTRO E CONTRATOS", "BENEFÍCIOS",
            "MEDIÇÃO E EQUIPAMENTOS DE MEDIÇÃO", "LEITURA", "TARIFAS, FATURAS, FATURAMENTO E COBRANÇA",
            "SERVIÇOS COBRÁVEIS", "PAGAMENTO", "SUSPENSÃO DO FORNECIMENTO", "PROCEDIMENTO IRREGULAR",
            "ATENDIMENTO E ESTRUTURA DE ATENDIMENTO", "QUALIDADE DA PRESTAÇÃO DE SERVIÇO",
            "RESSARCIMENTO DE DANOS ELÉTRICOS", "REDE E MANUTENÇÃO", "GERAÇÃO DISTRIBUÍDA",
            "PRAZOS E ACOMPANHAMENTO DE SOLICITAÇÃO", "INSTALAÇÃO", "ILUMINAÇÃO PÚBLICA",
            "LEGISLAÇÃO SETOR ELÉTRICO CORRELATA", "NORMAS E PADRÕES TÉCNICOS DE DISTRIBUIÇÃO",
            "EFICIÊNCIA ENERGÉTICA - RACIONALIZAÇÃO DE CONSUMO", "CAMINHO DO ATENDIMENTO",
            "DESLIGAMENTO DEFINITIVO/TEMPORÁRIO", "VENDAS DE PRODUTOS E SERVIÇOS", "RECLAME AQUI",
            "SOLICITAÇÃO PIX", "INFORMAÇÃO DE INCORPORAÇÃO DE REDE", "OUTROS"
        ]
        
        # NOVOS DADOS PARA OS SERVIÇOS REFORMULADOS
        self.problemas_equipamento = [
            "DISPLAY APAGADO",
            "MEDIDOR QUEIMADO", 
            "DANIFICADO/QUEBRADO",
            "MEDIDOR PARADO",
            "MEDIDOR FURTADO"
        ]

        # MODIFICADO: Dicionário de taxas agora é aninhado por estado
        self.taxas_religacao_por_estado = {
            "Amapá": {
                "Comum": {"Monofásica": "R$ 10,38", "Bifásica": "R$ 14,29", "Trifásica": "R$ 42,92"},
                "Urgência": {"Monofásica": "R$ 53,78", "Bifásica": "R$ 80,70", "Trifásica": "R$ 134,52"}
            },
            "Alagoas": {
                "Comum": {"Monofásica": "R$ 11,21", "Bifásica": "R$ 15,45", "Trifásica": "R$ 46,38"},
                "Urgência": {"Monofásica": "R$ 53,78", "Bifásica": "R$ 80,70", "Trifásico": "R$ 134,52"}
            },
            "Maranhão": {
                "Comum": {"Monofásica": "R$ 10,73", "Bifásica": "R$ 14,79", "Trifásica": "R$ 44,40"},
                "Urgência": {"Monofásica": "R$ 53,78", "Bifásica": "R$ 80,70", "Trifásico": "R$ 134,52"}
            },
            "Piauí": {
                "Comum": {"Monofásica": "R$ 10,86", "Bifásica": "R$ 14,96", "Trifásica": "R$ 44,91"},
                "Urgência": {"Monofásica": "R$ 53,78", "Bifásica": "R$ 80,70", "Trifásico": "R$ 134,52"}
            },
            "Pará": {
                "Comum": {"Monofásica": "R$ 10,72", "Bifásica": "R$ 14,77", "Trifásica": "R$ 44,35"},
                "Urgência": {"Monofásica": "R$ 53,78", "Bifásica": "R$ 80,70", "Trifásica": "R$ 134,52"}
            }
        }
        
        # Dentro da função __init__ da classe AtendimentoApp

        self.codigos_finalizacao = {
            "Emergencial": [
                "10 - Falta de energia",
                "11 - Risco a vida",
                "12 - Outros"
            ],
            "Comercial": [
                "20 - Reclamacao de faturas",
                "21 - Envio 2ª via de fatura",
                "22 - Religacao",
                "23 - Danos eletricos",
                "24 - Outros"
            ],
            "Informacoes": [
                "30 - Prazos",
                "31 - Debitos e faturas",
                "32 - Outros"
            ],
            "Telefonia": [
                "41 - Problemas na chamada",
                "42 - Ligação Caiu"
            ]
        }
        
        # NOVO: Dicionário separado para os prazos
        self.prazos_religacao = {
            "Comum": "24 horas",
            "Urgência": "4 horas"
        }
        
        self.links_sap = {
            "Maranhão": "http://unifica.equatorialenergia.com.br:9204/sap(bD1wdCZjPTQwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm?sap-client=401&sap-language=PT",
            "Pará": "http://unifica.equatorialenergia.com.br:9203/sap(bD1wdCZjPTQwMiZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm?sap-client=402&sap-language=PT",
            "Piauí": "http://epispdccrm02.equatorial.corp:8000/sap(bD1wdCZjPTQwNCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm",
            "Alagoas": "http://ealspdccrm02.equatorial.corp:8000/sap(bD1wdCZjPTQwMyZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm",
            "AMAPA": "https://ap-crm-prd.equatorial.corp:44301/sap/bc/bsp/sap/crm_ui_start?sap-client=405&sap-language=PT"
        }
        
        self.links_atc = {
            "ATC MARANHÃO": "http://10.7.1.20:8090/account/index.rails",
            "ATC PARÁ": "http://10.130.1.7:8083/account/index.rails",
            "ATC PIAUÍ": "http://10.6.10.170:8090/account/index.rails",
            "ATC RS": "http://10.63.79.108:8090/account/index.rails",
            "ATC GOIAS": "http://10.204.10.156:4002/atc/login"
        }
        
        self.links_adicionais = {
            "PORTAL DO COLABORADOR": "http://portaldocolaborador.equatorial.corp/home", 
            "SisFeedback": "https://sisfeedback.equatorialenergia.com.br/login",
            "Somos": "https://somos-al.equatorialenergia.com.br/somos/login.php", 
            "SIGA": "https://equatorialenergia.etadirect.com/",
            "TEAMS": "https://teams.microsoft.com/v2/",
            "Meu Elogio Vale Prêmio": "https://forms.office.com/pages/responsepage.aspx?id=UUVNECctOUiHsoYmSlIQpXVLerS6m75Io99nQig26PxUNjlURkFUSzBCTzRVMUtaQlhTUFNCWFpJOS4u&route=shorturl",
            "Genesys": "https://login.sae1.pure.cloud/?rid=CgQ8vm4NoAevaSgBMUyupcv6NQEex2P8O-8ZTpsXRyc#/splash",
            "Tela Ágil": "https://sistemas-hm.equatorialenergia.com.br/telaagil",
            "BSC": "http://10.6.1.18/reports/browse/EQTL%20Servi%C3%A7os/4.%20Contact%20Center%20-%20EQTL%20SERVI%C3%87OS/7.%20Opera%C3%A7%C3%A3o%20-%20Contact%20Center/CALL%20CENTER/01.BSC",
            "CONSULTAR CPF": "https://servicos.receita.fazenda.gov.br/servicos/cpf/consultasituacao/consultapublica.asp",
            "WEBCONSULTA": "https://webconsulta.equatorialenergia.com.br/religacao/",
            "APRENDE+": "https://equatorialenergia.edusense.app/#/platform/wiki/folder/1432792c0f00f043ed0a2500ce7f9aa11620",
            "APRENDE+RECLAMAÇÃO": "https://equatorialenergia.edusense.app/#/platform/wiki/content/478704c50b84f04d0c0ad51099f9ef0e1f10",
            "APRENDE+DANOS ELETR": "https://equatorialenergia.edusense.app/#/platform/wiki/content/1e01594c0959a047e50a01f0e91e70d8381e",
            "APRENDE+DANOS MATERI": "https://equatorialenergia.edusense.app/#/platform/wiki/content/28e1cdea0a3b0047580c827054c89089c20d",
        }
        
        self.links_backoffice = {
            "Maranhão": "https://backoffice-ma.equatorialenergia.com.br/login/", 
            "Pará": "https://backoffice-pa.equatorialenergia.com.br/login/",
            "Piauí": "https://backoffice-pi.equatorialenergia.com.br/login/", 
            "Alagoas": "https://backoffice-al.equatorialenergia.com.br/login/"
        }
        
        # Dicionários para armazenar widgets e dados dinâmicos
        self.entries = {}
        self.radio_vars = {} 
        self.historico_tree_map = {}
        self.faturas_frames = []
        
        # Caminhos de arquivos de configuração e dados
        self.caminho_config = os.path.join(self.base_path, "config_tema.json")
        self.caminho_historico = os.path.join(self.base_path, "historico_registros.json")
        self.caminho_ultimo_usuario = os.path.join(self.base_path, "ultimo_usuario.tmp")
        
        self.history_manager = HistoryManager(self.caminho_historico)
        
        # Inicialização da aplicação
        self.carregar_configuracoes()
        
        self.criar_interface()
        self._atualizar_botoes_favoritos() 
        self.configurar_atalhos()
        self.aplicar_tema_simples()
        self.verificar_mudanca_usuario()
        logger.info("Aplicação iniciada com sucesso")
        self.atualizar_relogio() # Inicia o loop do relógio
        self.pausa_sistema.verificar_horarios_pausas() # NOVA LINHA: Inicia verificação de pausas
        self.root.bind_all("<MouseWheel>", self._on_global_mousewheel)

    def _on_global_mousewheel(self, event):
            """
            Função 'inteligente' que detecta qual canvas rolar
            baseado na posição do mouse.
            """
            # --- INÍCIO DA CORREÇÃO ---
            widget = event.widget
            # Às vezes, o bind_all pode retornar o nome do widget (string) em vez do objeto.
            # Esta linha garante que sempre tenhamos o objeto do widget.
            if isinstance(widget, str):
                try:
                    # O método nametowidget converte o nome de volta para o objeto
                    widget = self.root.nametowidget(widget)
                except KeyError:
                    # Se o widget não for encontrado (caso raro), apenas ignore o evento.
                    return 
            # --- FIM DA CORREÇÃO ---
            
            canvas = None

            while widget: # 'while widget' é uma forma mais curta de 'while widget is not None'
                if isinstance(widget, tk.Canvas):
                    canvas = widget
                    break
                widget = widget.master # Sobe na hierarquia de widgets

            if canvas:
                # Verifica se o canvas encontrado é um dos nossos canvases roláveis
                if canvas is self.canvas or canvas is self.table_canvas or canvas is self.main_canvas_acessos:
                    # O delta é diferente em alguns sistemas operacionais, a divisão por 120 normaliza isso.
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def carregar_configuracoes(self):
        try:
            if os.path.exists(self.caminho_config):
                with open(self.caminho_config, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.tema_atual = config.get("tema", "claro")
                    self.cor_atual = config.get("cor", "azul")
                    self.tamanho_fonte_base = config.get("tamanho_fonte", 10)
                    
                    favoritos_carregados = config.get("favoritos", [])
                    # Garante que a lista tenha exatamente 3 itens (preenche com None se necessário)
                    self.favoritos = (favoritos_carregados + [None, None, None])[:3]

        except Exception as e:
            logger.error(f"Erro ao carregar configurações: {e}")
            
    def salvar_configuracoes(self):
        try:
            config = {
                "tema": self.tema_atual,
                "cor": self.cor_atual,
                "tamanho_fonte": self.tamanho_fonte_base,
                "favoritos": self.favoritos
            }
            with open(self.caminho_config, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4)
            logger.info("Configurações salvas com sucesso")
        except Exception as e:
            logger.error(f"Erro ao salvar configurações: {e}")
            
    def _on_favorito_click(self, service_id: str):
        """Preenche a busca de serviço e carrega o formulário."""
        if service_id:
            self.entry_servico.delete(0, tk.END)
            self.entry_servico.insert(0, service_id)
            self.carregar_formulario()
            
    def _atualizar_botoes_favoritos(self):
        """Atualiza a aparência dos botões de favoritos com base nos dados salvos."""
        if not hasattr(self, 'botoes_favoritos'):
            return # Sai se os botões ainda não foram criados

        for i, botao in enumerate(self.botoes_favoritos):
            service_id = self.favoritos[i]
            if service_id and service_id in self.servicos:
                nome_servico = self.servicos[service_id]['nome']
                # Usamos lambda com um argumento padrão para capturar o ID correto
                botao.config(
                    text=f"⭐ {nome_servico}",
                    state="normal",
                    command=lambda sid=service_id: self._on_favorito_click(sid)
                )
            else:
                botao.config(
                    text=f"Favorito {i+1} (Vazio)",
                    state="disabled",
                    command=lambda: None)

    def _bind_mousewheel_recursively(self, widget, command):
        """Aplica o evento de rolagem do mouse a um widget e a todos os seus filhos."""
        widget.bind("<MouseWheel>", command)
        for child in widget.winfo_children():
            self._bind_mousewheel_recursively(child, command)

    def limpar_sessao_completa(self):
        try:
            self.history_manager.reset_historico()

            arquivos_para_limpar = [
                os.path.join(self.base_path, "config_acessos.json"),
                self.caminho_config,
            ]
            
            for arquivo in arquivos_para_limpar:
                try:
                    if os.path.exists(arquivo):
                        os.remove(arquivo)
                        logger.info(f"Arquivo de configuração removido: {arquivo}")
                except Exception as e:
                    logger.error(f"Erro ao remover {arquivo}: {e}")
            
            self.tema_atual = "claro"
            self.cor_atual = "azul"
            self.tamanho_fonte_base = 10
            
            self.historico_tree_map.clear()
            
            if hasattr(self, 'anotacoes_textbox'):
                self.anotacoes_textbox.delete(1.0, tk.END)
            
            self.limpar_campos()
            self.registro_usuario.set("")
            
            self.var_Genesys.set("FALTA DE ENERGIA INDIVIDUAL; O CLIENTE INFORMA QUE REALIZOU O TESTE DO DISJUNTOR.")
            
            for categoria in ["Emergenciais", "Comerciais", "Informação", "Reclamações"]:
                tree_attr = f"tree_{categoria.lower()}"
                if hasattr(self, tree_attr):
                    tree = getattr(self, tree_attr)
                    for item in tree.get_children():
                        tree.delete(item)
            
            self.aplicar_tamanho_fonte()
            self.aplicar_tema_simples()
            
            gc.collect()
            
            logger.info("Sessão limpa COMPLETAMENTE - Pronto para novo usuário")
            
        except Exception as e:
            logger.error(f"Erro ao limpar sessão: {e}")
            messagebox.showerror("Erro", f"Erro na limpeza: {str(e)}")

    def verificar_mudanca_usuario(self):
        try:
            usuario_atual = getpass.getuser()
            
            ultimo_usuario = ""
            
            if os.path.exists(self.caminho_ultimo_usuario):
                with open(self.caminho_ultimo_usuario, "r") as f:
                    ultimo_usuario = f.read().strip()
            
            if ultimo_usuario and ultimo_usuario != usuario_atual:
                resposta = messagebox.askyesno(
                    "🔄 Novo Usuário Detectado", 
                    f"Detectamos que um novo usuário fez login:\n\n"
                    f"Usuário anterior: {ultimo_usuario}\n"
                    f"Usuário atual: {usuario_atual}\n\n"
                    "Deseja limpar automaticamente os dados da sessão anterior?"
                )
                
                if resposta:
                    self.limpar_sessao_completa()
                    messagebox.showinfo("✅ Auto-Limpeza", "Dados da sessão anterior foram limpos automaticamente!")
            
            with open(self.caminho_ultimo_usuario, "w") as f:
                f.write(usuario_atual)
                
        except Exception as e:
            logger.error(f"Erro na verificação de usuário: {e}")

    def _get_entry_value(self, key: str) -> str:
        entry = self.entries.get(key)
        if entry:
            return entry.get().strip()
        return ""       

    def aplicar_tema_simples(self):
        # --- MOTOR DE TEMAS COM TODAS AS PALETAS RESTAURADAS + NEON ---
        temas = {
            "azul_padrao": {
                "fundo_janela": "#D6EAF8", "fundo_container": "#EBF5FB",
                "cor_principal": "#2980B9", "cor_destaque": "#5499C7",
                "texto_primario": "#17202A", "texto_secundario": "#FFFFFF"
            },
            "azul": {
                "fundo_janela": "#B3D4FC", "fundo_container": "#EBF5FB",
                "cor_principal": "#2FA6EB", "cor_destaque": "#002EAD",
                "texto_primario": "#001F5B", "texto_secundario": "#FFFFFF"
            },
            "verde": {
                "fundo_janela": "#C1E1C1", "fundo_container": "#E8F5E9",
                "cor_principal": "#006400", "cor_destaque": "#003300",
                "texto_primario": "#002200", "texto_secundario": "#FFFFFF"
            },
            "vermelho": {
                "fundo_janela": "#F5B7B1", "fundo_container": "#FDEDEC",
                "cor_principal": "#B22222", "cor_destaque": "#800000",
                "texto_primario": "#641E16", "texto_secundario": "#FFFFFF"
            },
            "amarelo": {
                "fundo_janela": "#FDEBD0", "fundo_container": "#FEF9E7",
                "cor_principal": "#FFB900", "cor_destaque": "#F7630C",
                "texto_primario": "#6E2C00", "texto_secundario": "#000000"
            },
            "roxo": {
                "fundo_janela": "#D7BDE2", "fundo_container": "#F5EEF8",
                "cor_principal": "#5C2D91", "cor_destaque": "#40207A",
                "texto_primario": "#32145A", "texto_secundario": "#FFFFFF"
            },
            "rosa": {
                "fundo_janela": "#FADBD8", "fundo_container": "#FEF1F0",
                "cor_principal": "#FF69B4", "cor_destaque": "#C71585",
                "texto_primario": "#780048", "texto_secundario": "#FFFFFF"
            },
            "verde escuro": {
                "fundo_janela": "#71A074", "fundo_container": "#D5E8D4",
                "cor_principal": "#003300", "cor_destaque": "#001A00",
                "texto_primario": "#001A00", "texto_secundario": "#FFFFFF"
            },
            "laranja": {
                "fundo_janela": "#FFD580", "fundo_container": "#FFF5E1",
                "cor_principal": "#FF5733", "cor_destaque": "#C70039",
                "texto_primario": "#8D270D", "texto_secundario": "#FFFFFF"
            },
            "cinza": {
                "fundo_janela": "#CACFD2", "fundo_container": "#E5E7E9",
                "cor_principal": "#7A7A7A", "cor_destaque": "#5C5C5C",
                "texto_primario": "#212121", "texto_secundario": "#FFFFFF"
            },
            "branco": {
                "fundo_janela": "#F2F3F4", "fundo_container": "#FFFFFF",
                "cor_principal": "#E0E0E0", "cor_destaque": "#BDBDBD",
                "texto_primario": "#000000", "texto_secundario": "#000000"
            },
            "preto": {
                "fundo_janela": "#1C1C1C", "fundo_container": "#2E2E2E",
                "cor_principal": "#424242", "cor_destaque": "#000000",
                "texto_primario": "#FFFFFF", "texto_secundario": "#FFFFFF"
            },
            "mistura": {
                "fundo_janela": "#FCF3CF", "fundo_container": "#FFF9E6",
                "cor_principal": "#0026FF", "cor_destaque": "#FF0000",
                "texto_primario": "#00137F", "texto_secundario": "#FFFFFF"
            },
            "dourado": {
                "fundo_janela": "#F9E79F", "fundo_container": "#FEF5E7",
                "cor_principal": "#D4AF37", "cor_destaque": "#A68A1E",
                "texto_primario": "#69560E", "texto_secundario": "#000000"
            },
            "azul escuro": {
                "fundo_janela": "#A9CCE3", "fundo_container": "#D4E6F1",
                "cor_principal": "#0D47A1", "cor_destaque": "#0011FD",
                "texto_primario": "#0A245A", "texto_secundario": "#FFFFFF"
            },
            "violeta": {
                "fundo_janela": "#E1BEE7", "fundo_container": "#F3E5F5",
                "cor_principal": "#8E24AA", "cor_destaque": "#6A1B9A",
                "texto_primario": "#4A148C", "texto_secundario": "#FFFFFF"
            },
            "marrom": {
                "fundo_janela": "#D7CCC8", "fundo_container": "#EFEBE9",
                "cor_principal": "#5D4037", "cor_destaque": "#3E2723",
                "texto_primario": "#3E2723", "texto_secundario": "#FFFFFF"
            },
            "turquesa": {
                "fundo_janela": "#A3E4D7", "fundo_container": "#D1F2EB",
                "cor_principal": "#009688", "cor_destaque": "#00C4B4",
                "texto_primario": "#004D40", "texto_secundario": "#FFFFFF"
            },
            "oceano profundo": {
                "fundo_janela": "#A2D9CE", "fundo_container": "#D0ECE7",
                "cor_principal": "#008080", "cor_destaque": "#004D4D",
                "texto_primario": "#003333", "texto_secundario": "#FFFFFF"
            },
            "floresta": {
                "fundo_janela": "#D5F5E3", "fundo_container": "#E8F8F5",
                "cor_principal": "#229954", "cor_destaque": "#28B463",
                "texto_primario": "#145A32", "texto_secundario": "#FFFFFF"
            },
            "café": {
                "fundo_janela": "#E5E0D5", "fundo_container": "#F5F5DC",
                "cor_principal": "#A0522D", "cor_destaque": "#6B3E2E",
                "texto_primario": "#59371E", "texto_secundario": "#FFFFFF"
            },
            "vibrante": {
                "fundo_janela": "#FAD7A0", "fundo_container": "#FEF9E7",
                "cor_principal": "#FF4500", "cor_destaque": "#FF00FF",
                "texto_primario": "#7E2200", "texto_secundario": "#FFFFFF"
            },
            "cyberpunk": {
                "fundo_janela": "#1A1A1A", "fundo_container": "#2A2A2A",
                "cor_principal": "#FF00FF", "cor_destaque": "#00FFFF",
                "texto_primario": "#00FFFF", "texto_secundario": "#000000"
            },
            "neon (ciano)": {
                "fundo_janela": "#101010", "fundo_container": "#181818",
                "cor_principal": "#00FFFF", "cor_destaque": "#7DF9FF",
                "texto_primario": "#00FFFF", "texto_secundario": "#000000"
            },
            "neon (rosa)": {
                "fundo_janela": "#101010", "fundo_container": "#181818",
                "cor_principal": "#FF00FF", "cor_destaque": "#FF77FF",
                "texto_primario": "#FF00FF", "texto_secundario": "#000000"
            },
            "neon (verde)": {
                "fundo_janela": "#101010", "fundo_container": "#181818",
                "cor_principal": "#39FF14", "cor_destaque": "#ADFF2F",
                "texto_primario": "#39FF14", "texto_secundario": "#000000"
            }
        }
        
        # O resto do método continua EXATAMENTE IGUAL
        tema_escolhido = temas.get(self.cor_atual, temas["azul_padrao"])
        if self.tema_atual == "escuro":
            fundo_janela = "#1e1e1e"
            fundo_container = "#2d2d2d"
            cor_principal = tema_escolhido["cor_principal"]
            cor_destaque = tema_escolhido["cor_destaque"]
            texto_primario = tema_escolhido.get("texto_primario", "#FFFFFF")
            texto_secundario = tema_escolhido.get("texto_secundario", "#000000")
            cor_entry_bg = "#3c3c3c"
            cor_entry_fg = texto_primario
            cor_text_bg = "#3c3c3c"
        else:
            fundo_janela = tema_escolhido["fundo_janela"]
            fundo_container = tema_escolhido["fundo_container"]
            cor_principal = tema_escolhido["cor_principal"]
            cor_destaque = tema_escolhido["cor_destaque"]
            texto_primario = tema_escolhido["texto_primario"]
            texto_secundario = tema_escolhido["texto_secundario"]
            cor_entry_bg = "#FFFFFF"
            cor_entry_fg = texto_primario
            cor_text_bg = "#FFFFFF"

        self.root.configure(bg=fundo_janela)
        self.style.theme_use('clam')
        self.style.configure(".", background=fundo_container, foreground=texto_primario, font=self.fonte_atual)
        self.style.configure("TFrame", background=fundo_container)
        self.style.configure("TLabel", background=fundo_container, foreground=texto_primario)
        self.style.configure("Bold.TLabel", background=fundo_container, foreground=texto_primario, font=self.fonte_bold_atual)
        self.style.configure("TLabelframe", background=fundo_container, foreground=texto_primario, font=self.fonte_bold_atual)
        self.style.configure("TLabelframe.Label", background=fundo_container, foreground=texto_primario)
        self.style.configure("TCheckbutton", background=fundo_container, foreground=texto_primario)
        self.style.configure("TRadiobutton", background=fundo_container, foreground=texto_primario)
        self.style.configure("Destaque.TButton", background=cor_principal, foreground=texto_secundario,
                             borderwidth=1, focuscolor='none', font=self.fonte_bold_atual, padding=5)
        self.style.map("Destaque.TButton",
                       background=[('active', cor_destaque), ('pressed', cor_destaque), ('!active', cor_principal)],
                       foreground=[('active', texto_secundario)])
        self.style.configure("TNotebook", background=fundo_janela, borderwidth=0)
        self.style.configure("TNotebook.Tab", background=fundo_container, foreground=texto_primario,
                             padding=[20, 8], font=self.fonte_bold_atual)
        self.style.map("TNotebook.Tab",
                       background=[("selected", cor_destaque), ('!selected', fundo_container)],
                       foreground=[("selected", texto_secundario), ('!selected', texto_primario)])
        self.style.configure("TEntry", fieldbackground=cor_entry_bg, foreground=cor_entry_fg, insertcolor=texto_primario)
        self.style.configure("TCombobox", fieldbackground=cor_entry_bg, foreground=cor_entry_fg)
        self.style.configure("TSpinbox", fieldbackground=cor_entry_bg, foreground=cor_entry_fg)
        self.style.configure("Treeview", background=cor_entry_bg, foreground=cor_entry_fg,
                             fieldbackground=cor_entry_bg, rowheight=int(self.fonte_atual.cget("size") * 2.5))
        self.style.configure("Treeview.Heading", background=cor_principal, foreground=texto_secundario, font=self.fonte_bold_atual)
        self.style.map("Treeview",
                       background=[('selected', cor_destaque)],
                       foreground=[('selected', texto_secundario)])
        self.atualizar_widgets_nao_ttk(self.root, fundo_container, cor_text_bg, texto_primario, cor_destaque)

    def atualizar_widgets_nao_ttk(self, widget, cor_fundo, cor_text_bg, cor_fg, cor_select_bg):
        try:
            if isinstance(widget, (tk.Frame, tk.Canvas)) and not isinstance(widget, (ttk.Frame, ttk.LabelFrame)):
                widget.configure(bg=cor_fundo, highlightthickness=0)
            
            if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                widget.configure(bg=cor_text_bg, fg=cor_fg, insertbackground=cor_fg,
                                   selectbackground=cor_select_bg, selectforeground="white")

            for child in widget.winfo_children():
                self.atualizar_widgets_nao_ttk(child, cor_fundo, cor_text_bg, cor_fg, cor_select_bg)
        except tk.TclError:
            pass

    def configurar_atalhos(self):
        self.root.bind("<Control-plus>", lambda e: self.ajustar_fonte(2))
        self.root.bind("<Control-minus>", lambda e: self.ajustar_fonte(-2))
        self.root.bind("<Control-0>", lambda e: self.resetar_fonte())
        self.root.bind("<F1>", lambda e: self.mostrar_ajuda())
        self.root.bind("<Control-a>", self.selecionar_tudo)
        self.root.bind("<Control-A>", self.selecionar_tudo)
    
    def selecionar_tudo(self, event):
        widget = event.widget
        try:
            if isinstance(widget, ttk.Entry):
                widget.select_range(0, tk.END)
                widget.icursor(tk.END)
            elif isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                widget.tag_add(tk.SEL, "1.0", tk.END)
                widget.mark_set(tk.INSERT, "1.0")
                widget.see(tk.INSERT)
            return "break"
        except:
            return "break"
    
    def ajustar_fonte(self, delta: int):
        novo_tamanho = max(8, min(36, self.tamanho_fonte_base + delta))
        self.tamanho_fonte_base = novo_tamanho
        self.aplicar_tamanho_fonte()
        self.salvar_configuracoes()
        if hasattr(self, 'label_fonte'): 
            self.label_fonte.config(text=f"{self.tamanho_fonte_base}pt")
    
    def resetar_fonte(self):
        self.tamanho_fonte_base = 10
        self.aplicar_tamanho_fonte()
        self.salvar_configuracoes()
        if hasattr(self, 'label_fonte'): 
            self.label_fonte.config(text=f"{self.tamanho_fonte_base}pt")
    
    def aplicar_tamanho_fonte(self):
        try:
            self.fonte_atual.configure(size=self.tamanho_fonte_base, weight="bold") # <-- MUDANÇA AQUI
            self.fonte_bold_atual.configure(size=self.tamanho_fonte_base, weight="bold")
            
            self.aplicar_tema_simples()
            
            if hasattr(self, 'output_text'): self.output_text.configure(font=self.fonte_atual)
            if hasattr(self, 'anotacoes_textbox'): self.anotacoes_textbox.configure(font=self.fonte_atual)
                                                      
        except Exception as e:
            logger.error(f"Erro ao aplicar tamanho de fonte: {e}")
    
    def atalho_religacao(self):
        self.entry_servico.delete(0, tk.END)
        self.entry_servico.insert(0, "10")
        self.carregar_formulario()
        
    def mostrar_ajuda(self):
        ajuda = tk.Toplevel(self.root)
        ajuda.title("Ajuda - Atalhos de Teclado")
        ajuda.geometry("500x380")  # Aumentei um pouco a altura para caber tudo
        
        # Texto de ajuda atualizado com os novos atalhos
        texto_ajuda = """
        ATALHOS DE TECLADO ATUALIZADOS:

        === CONTROLE DE FONTE ===
        Ctrl + (+) : Aumentar a fonte
        Ctrl + (-) : Diminuir a fonte
        Ctrl +  0  : Restaurar fonte padrão

        === EDIÇÃO E NAVEGAÇÃO (Padrão) ===
        Ctrl + A : Selecionar tudo
        Ctrl + C : Copiar
        Ctrl + V : Colar
        Ctrl + X : Recortar
        Alt + Tab: Alternar entre janelas
        """
        
        # O resto da função continua igual...
        frame = ttk.Frame(ajuda)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        text_widget = tk.Text(frame, wrap=tk.WORD, font=self.fonte_atual)
        text_widget.insert(1.0, texto_ajuda)
        text_widget.config(state="disabled")
        scrollbar = ttk.Scrollbar(frame, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        ttk.Button(ajuda, text="Fechar", command=ajuda.destroy).pack(pady=10)

    def criar_interface(self):
        # --- 1. BARRA SUPERIOR COM BOTÕES E RELÓGIO ---
        frame_top = ttk.Frame(self.root)
        frame_top.pack(fill="x", padx=10, pady=5)
        
        # Botões à esquerda
        ttk.Button(frame_top, text="Ajuda (F1)", command=self.mostrar_ajuda).pack(side="left")
        ttk.Button(frame_top, text="Configurações", command=self.abrir_configuracoes).pack(side="left", padx=5)
        
        # Relógio à direita
        self.label_relogio = ttk.Label(frame_top, text="", font=("Arial", 16, "bold"), foreground="#333")
        self.label_relogio.pack(side="right", padx=10)

        # --- 2. PAINEL DE PAUSAS ---
        # CORREÇÃO: Movido para ser criado antes do Notebook para um layout mais estável
        self.painel_pausas = ttk.LabelFrame(self.root, text="Painel de Pausas")
        self.pausa_sistema.criar_painel_pausas(self.painel_pausas)
        
        # --- 3. NOTEBOOK COM AS ABAS ---
        self.notebook = ttk.Notebook(self.root)
        
# --- 3. NOTEBOOK COM AS ABAS ---
        self.notebook = ttk.Notebook(self.root)
        
        # Cria os frames para cada aba
        self.aba_atendimento = ttk.Frame(self.notebook)
        self.aba_tabulacao = ttk.Frame(self.notebook) # Nova aba Tabulação
        self.aba_acessos = ttk.Frame(self.notebook)
        self.aba_historico = ttk.Frame(self.notebook)
        self.aba_pesquisa = ttk.Frame(self.notebook, padding=20) # Aba de Sexta-feira

        # Chama as funções que criam o conteúdo de cada aba
        self.criar_aba_atendimento()
        self.criar_aba_tabulacao() # Chama a nova função
        self.criar_aba_acessos()
        self.criar_aba_historico()
        self.criar_conteudo_aba_pesquisa() 
        
        # Adiciona as abas visíveis ao notebook na ordem desejada
        self.notebook.add(self.aba_atendimento, text="Atendimento")
        self.notebook.add(self.aba_tabulacao, text="Tabulação", state="hidden") # Começa escondida
        self.notebook.add(self.aba_acessos, text="Acessos")
        self.notebook.add(self.aba_historico, text="Histórico")
        self.notebook.add(self.aba_pesquisa, text="Pesquisa", state="hidden") # A de sexta também
        
        # --- 4. EMPACOTAMENTO FINAL ---
        # O painel de pausas é mostrado primeiro
        self.painel_pausas.pack(fill="x", padx=10, pady=(5, 0))
        # O notebook ocupa o resto do espaço
        self.notebook.pack(expand=True, fill="both", padx=10, pady=5)

        # Bind para mostrar/esconder painel baseado na aba selecionada
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        # Chama uma vez para garantir que o painel de pausas esteja visível no início
        self.on_tab_changed()
    
    def on_tab_changed(self, event=None):
        """Mostra/esconde painel de pausas baseado na aba selecionada."""
        # CORREÇÃO: Pega a aba selecionada diretamente do notebook,
        # em vez de depender do 'event'. Isso funciona em ambos os casos.
        try:
            selected_tab_text = self.notebook.tab(self.notebook.select(), "text").strip()
        except tk.TclError:
            # Pode acontecer se nenhuma aba estiver selecionada ainda
            selected_tab_text = "Atendimento"

        if selected_tab_text == "Atendimento":
            # Usamos pack_forget() antes para evitar erros se ele já estiver visível
            self.painel_pausas.pack_forget() 
            self.painel_pausas.pack(fill="x", padx=10, pady=(5, 0), before=self.notebook)
        else:
            self.painel_pausas.pack_forget()

    def criar_aba_atendimento(self):
        frame_registro = ttk.Frame(self.aba_atendimento)
        frame_registro.pack(fill="x", padx=10, pady=10)
        ttk.Label(frame_registro, text="Matricula do atendente:", style="Bold.TLabel").pack(side="left")
        entry_registro = ttk.Entry(frame_registro, textvariable=self.registro_usuario, width=30, font=self.fonte_atual)
        entry_registro.pack(side="left", padx=5)

        main_container = ttk.Frame(self.aba_atendimento)
        main_container.pack(fill="both", expand=True, padx=5, pady=5)
        frame_esquerdo = ttk.Frame(main_container)
        frame_esquerdo.pack(side="left", fill="both", expand=True, padx=(0, 10))
        frame_direito = ttk.Frame(main_container, width=280)
        frame_direito.pack(side="right", fill="y")
        frame_direito.pack_propagate(False)

        frame_top = ttk.Frame(frame_esquerdo)
        frame_top.pack(fill="x", pady=(0, 10))
        ttk.Label(frame_top, text="Buscar Serviço (Nome ou Nº):", style="Bold.TLabel").pack(side="left")

        # --- MUDANÇA PRINCIPAL AQUI ---
        # Prepara a lista de serviços para o Combobox
        servicos_lista = [f"{num} - {info['nome']}" for num, info in sorted(self.servicos.items(), key=lambda item: int(item[0]))]

        # Substitui o ttk.Entry por um ttk.Combobox
        self.entry_servico = ttk.Combobox(frame_top, values=servicos_lista, width=35, font=self.fonte_atual)
        self.entry_servico.pack(side="left", padx=5)
        
        # Mantém a funcionalidade de pressionar Enter e adiciona a de selecionar com o mouse
        self.entry_servico.bind("<Return>", self.carregar_formulario)
        self.entry_servico.bind("<<ComboboxSelected>>", self.carregar_formulario)
        # --- FIM DA MUDANÇA ---

        self.btn_carregar = ttk.Button(frame_top, text="Carregar", command=self.carregar_formulario)
        self.btn_carregar.pack(side="left")
        self.label_servico = ttk.Label(frame_top, text="", font=self.fonte_bold_atual, foreground="blue")
        self.label_servico.pack(side="left", padx=10)
        
        frame_favoritos = ttk.LabelFrame(frame_esquerdo, text="⭐ Favoritos")
        frame_favoritos.pack(fill="x", pady=5, padx=5)
        self.botoes_favoritos = []
        for i in range(3):
            btn = ttk.Button(frame_favoritos, text=f"Favorito {i+1}", style="Destaque.TButton")
            btn.pack(side="left", fill="x", expand=True, padx=3, pady=3)
            self.botoes_favoritos.append(btn)
        
        paned_window = ttk.PanedWindow(frame_esquerdo, orient="vertical")
        paned_window.pack(fill="both", expand=True, pady=(5, 0))

        form_pane = ttk.Frame(paned_window)
        paned_window.add(form_pane, weight=3)

        output_pane = ttk.Frame(paned_window)
        paned_window.add(output_pane, weight=1)

        canvas_frame = ttk.Frame(form_pane) 
        canvas_frame.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        
        frame_buttons = ttk.Frame(form_pane)
        frame_buttons.pack(fill="x", pady=10)
        self.btn_registrar = ttk.Button(frame_buttons, text="🎯 Registrar Atendimento", 
                                          command=self.gerar_texto, state="disabled", style="Destaque.TButton")
        self.btn_registrar.pack(side="left", padx=5)
        self.btn_limpar = ttk.Button(frame_buttons, text="🔄 Limpar", command=self.limpar_campos)
        self.btn_limpar.pack(side="left", padx=5)

        self.output_text = scrolledtext.ScrolledText(output_pane, height=6, font=self.fonte_atual)
        self.output_text.pack(fill="both", expand=True)

        ttk.Label(frame_direito, text="Tabela de Serviços", font=self.fonte_bold_atual).pack(pady=(0, 5))
        
        self.frame_tabela_servicos = ttk.Frame(frame_direito)
        self.frame_tabela_servicos.pack(fill="both", expand=True)

        self.table_canvas = tk.Canvas(self.frame_tabela_servicos, highlightthickness=0)
        table_scrollbar = ttk.Scrollbar(self.frame_tabela_servicos, orient="vertical", command=self.table_canvas.yview)
        table_frame = ttk.Frame(self.table_canvas)
        table_frame.bind("<Configure>", lambda e: self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all")))
        self.table_canvas.create_window((0, 0), window=table_frame, anchor="nw")
        self.table_canvas.configure(yscrollcommand=table_scrollbar.set)

        categorias = {}
        for num, info in self.servicos.items():
            cat = info["categoria"]
            if cat not in categorias:
                categorias[cat] = []
            categorias[cat].append((num, info["nome"]))

        for cat, servicos_cat in sorted(categorias.items()):
            cat_frame = ttk.Frame(table_frame)
            cat_frame.pack(fill="x", pady=(10, 0))
            ttk.Label(cat_frame, text=cat, font=self.fonte_bold_atual, 
                      background=self.cor_categoria[cat]).pack(fill="x", padx=5, pady=2)
            
            for num, nome in servicos_cat:
                service_frame = ttk.Frame(table_frame)
                service_frame.pack(fill="x", pady=1, padx=5)
                
                cor_numero = "purple" if num == "13" else "red"
                
                ttk.Label(service_frame, text=num, 
                          font=("Arial", self.tamanho_fonte_base + 2, "bold"), 
                          foreground=cor_numero, width=3).pack(side="left")
                ttk.Label(service_frame, text=nome + (" 🚀" if num == "13" else ""), 
                          font=("Arial", self.tamanho_fonte_base - 1,"bold"), 
                          wraplength=220, justify="left").pack(side="left", fill="x", expand=True)

        self.table_canvas.pack(side="left", fill="both", expand=True)
        table_scrollbar.pack(side="right", fill="y")
        
    
    def criar_aba_acessos(self):
        main_paned = ttk.PanedWindow(self.aba_acessos, orient="horizontal")
        main_paned.pack(fill="both", expand=True)
        
        left_frame = ttk.Frame(main_paned)
        
        self.main_canvas_acessos = tk.Canvas(left_frame)
        scrollbar_main = ttk.Scrollbar(left_frame, orient="vertical", command=self.main_canvas_acessos.yview)
        scrollable_main = ttk.Frame(self.main_canvas_acessos)
        scrollable_main.bind("<Configure>", lambda e: self.main_canvas_acessos.configure(scrollregion=self.main_canvas_acessos.bbox("all")))
        self.main_canvas_acessos.create_window((0, 0), window=scrollable_main, anchor="nw")
        self.main_canvas_acessos.configure(yscrollcommand=scrollbar_main.set)
        
        frame_sap = ttk.LabelFrame(scrollable_main, text="🌐 SAP por Estado")
        frame_sap.pack(fill="x", padx=10, pady=5)
        for estado, link in self.links_sap.items():
            linha_sap = ttk.Frame(frame_sap)
            linha_sap.pack(fill="x", pady=2, padx=10)
            ttk.Label(linha_sap, text=f"🔹 {estado}:", width=20).pack(side="left")
            ttk.Button(linha_sap, text="Acessar", command=lambda l=link: webbrowser.open(l), style='Destaque.TButton').pack(side="left", padx=5)

        frame_atc = ttk.LabelFrame(scrollable_main, text="🌐 ATCs")
        frame_atc.pack(fill="x", padx=10, pady=5)
        for nome, link in self.links_atc.items():
            linha_atc = ttk.Frame(frame_atc)
            linha_atc.pack(fill="x", pady=2, padx=10)
            ttk.Label(linha_atc, text=f"🔹 {nome}:", width=20).pack(side="left")
            ttk.Button(linha_atc, text="Acessar", command=lambda l=link: webbrowser.open(l), style='Destaque.TButton').pack(side="left", padx=5)

        frame_backoffice = ttk.LabelFrame(scrollable_main, text="🏢 Backoffice por Estado")
        frame_backoffice.pack(fill="x", padx=10, pady=5)
        info_back = ttk.Label(frame_backoffice, text="🔑 Login: U******* | 🔐 Senha: (senha de rede)", 
                                font=self.fonte_atual, foreground="gray")
        info_back.pack(pady=5)
        for estado, link in self.links_backoffice.items():
            btn_back = ttk.Frame(frame_backoffice)
            btn_back.pack(fill="x", pady=2, padx=10)
            ttk.Label(btn_back, text=f"🔹 {estado}:", width=20).pack(side="left")
            ttk.Button(btn_back, text="Acessar", command=lambda l=link: webbrowser.open(l), style='Destaque.TButton').pack(side="left", padx=5)

        frame_outros = ttk.LabelFrame(scrollable_main, text="⚙️ Outros Sistemas")
        frame_outros.pack(fill="x", padx=10, pady=5)
        
        sistemas_ordenados = [
            "PORTAL DO COLABORADOR", "SisFeedback", "Somos", "SIGA", "TEAMS", 
            "Tela Ágil", "Genesys", "Meu Elogio Vale Prêmio", "BSC", "CONSULTAR CPF", "WEBCONSULTA", "APRENDE+", "APRENDE+RECLAMAÇÃO", "APRENDE+DANOS ELETR", "APRENDE+DANOS MATERI"
        ]
        
        for nome_sistema in sistemas_ordenados:
            if nome_sistema in self.links_adicionais:
                link = self.links_adicionais[nome_sistema]
                self.criar_sistema_com_credenciais(frame_outros, f"🔹 {nome_sistema}", link)

        self.main_canvas_acessos.pack(side="left", fill="both", expand=True)
        scrollbar_main.pack(side="right", fill="y")
        
        
        right_frame = ttk.Frame(main_paned)
        
        ttk.Label(right_frame, text="📝 Anotações", font=self.fonte_bold_atual).pack(pady=(10, 5))
        
        self.anotacoes_textbox = scrolledtext.ScrolledText(right_frame, height=20, width=40, font=self.fonte_atual)
        self.anotacoes_textbox.pack(fill="both", expand=True, padx=10, pady=5)
        
        btn_frame_anotacoes = ttk.Frame(right_frame)
        btn_frame_anotacoes.pack(fill="x", padx=10, pady=10)
        
        self.btn_limpar_anotacoes = ttk.Button(btn_frame_anotacoes, text="Limpar Anotações", 
                                              command=self.limpar_anotacoes, style='Destaque.TButton')
        self.btn_limpar_anotacoes.pack(side="left", padx=5)
        
        self.btn_salvar_anotacoes = ttk.Button(btn_frame_anotacoes, text="Salvar Anotações", 
                                              command=self.salvar_anotacoes_arquivo, style='Destaque.TButton')
        self.btn_salvar_anotacoes.pack(side="left", padx=5)
        
        main_paned.add(left_frame, weight=3)
        main_paned.add(right_frame, weight=1)
        
    def limpar_anotacoes(self):
        if messagebox.askyesno("Confirmar", "Deseja realmente limpar todas as anotações?"):
            self.anotacoes_textbox.delete(1.0, tk.END)
            
    def abrir_configuracoes(self):
        self.config_window = tk.Toplevel(self.root)
        self.config_window.title("Configurações")
        self.config_window.geometry("500x450")
        self.config_window.transient(self.root)
        self.config_window.grab_set()
        
        config_notebook = ttk.Notebook(self.config_window)
        config_notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        aba_aparencia = ttk.Frame(config_notebook)
        self.criar_config_aparencia(aba_aparencia)
        config_notebook.add(aba_aparencia, text="Aparência")
        
        aba_acessibilidade = ttk.Frame(config_notebook)
        self.criar_config_acessibilidade(aba_acessibilidade)
        config_notebook.add(aba_acessibilidade, text="Acessibilidade")

        aba_favoritos = ttk.Frame(config_notebook)
        self.criar_config_favoritos(aba_favoritos)
        config_notebook.add(aba_favoritos, text="⭐ Favoritos")
    
    def criar_config_favoritos(self, parent):
        """Cria a aba de configuração para os serviços favoritos."""
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ttk.Label(main_frame, text="Escolha seus 3 serviços de acesso rápido:", wraplength=400).pack(pady=(0, 15))

        self.servico_nome_para_id = {info["nome"]: num for num, info in self.servicos.items()}
        nomes_servicos = ["(Nenhum)"] + sorted(list(self.servico_nome_para_id.keys()))
        
        self.comboboxes_favoritos = []

        for i in range(3):
            fav_frame = ttk.Frame(main_frame)
            fav_frame.pack(fill="x", pady=5)
            
            ttk.Label(fav_frame, text=f"Favorito {i+1}:", width=10).pack(side="left")
            
            combo = ttk.Combobox(fav_frame, values=nomes_servicos, state="readonly", width=40)
            
            # Tenta preencher com o valor salvo
            service_id_salvo = self.favoritos[i]
            if service_id_salvo and service_id_salvo in self.servicos:
                nome_servico_salvo = self.servicos[service_id_salvo]["nome"]
                combo.set(nome_servico_salvo)
            else:
                combo.set("(Nenhum)")

            combo.pack(side="left", fill="x", expand=True)
            self.comboboxes_favoritos.append(combo)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side="bottom", pady=20)
        ttk.Button(btn_frame, text="Aplicar", command=self.aplicar_configuracoes, style="Destaque.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=self.fechar_configuracoes).pack(side="left")


    def criar_conteudo_aba_pesquisa(self):
        """Cria o conteúdo da aba de Pesquisa de Satisfação."""
        # Variáveis para controlar a animação (código existente)
        self.piscando = False
        self.id_loop_piscar = None

        label_aviso = ttk.Label(
            self.aba_pesquisa,
            text="LEMBRETE DE SEXTA-FEIRA!",
            font=("Segoe UI", 22, "bold"),
            foreground="red",  # Defini a cor diretamente como "vermelho"
            wraplength=500,
        )
        label_aviso.pack(pady=(20, 10))

        # --- INÍCIO DA MUDANÇA ---

        # 1. Instrução principal com fonte grande
        label_instrucao = ttk.Label(
            self.aba_pesquisa,
            text="Ao finalizar, transfira o cliente para a Pesquisa de Satisfação!",
            font=("Segoe UI", 18, "bold"), # Fonte grande e em negrito
            wraplength=500,
            justify="center"
        )
        label_instrucao.pack(pady=(10, 25)) # Aumentei o espaço abaixo

        # 2. Script para o atendente com fonte menor e em itálico
        script_para_falar = (
            "Script sugerido para o final da ligação:\n\n"
            "\"Para finalizar, peço que o(a) senhor(a) permaneça na linha para responder à nossa pesquisa de satisfação, é bem rápido. "
            "A Equatorial Energia agradece a sua ligação!\""
        )
        label_script = ttk.Label(
            self.aba_pesquisa,
            text=script_para_falar,
            font=("Segoe UI", 13, "italic"), # Fonte menor e em itálico
            wraplength=500,
            justify="center"
        )
        label_script.pack(pady=10)

        # --- FIM DA MUDANÇA ---

        # Botão para parar a animação (código existente)
        btn_ok = ttk.Button(
            self.aba_pesquisa,
            text="✅ OK, Entendido!",
            style="Accent.TButton",
            command=lambda: self.parar_piscar_generico(self.aba_pesquisa, "Pesquisa")
        )
        btn_ok.pack(pady=40, ipady=10)

# --- INÍCIO DO NOVO SISTEMA DE ANIMAÇÃO UNIFICADO ---

    def iniciar_piscar_generico(self, aba_alvo, texto_piscando):
        """Inicia a animação de piscar para qualquer aba especificada."""
        if hasattr(self, 'piscando') and self.piscando:
            # Se uma animação já estiver ativa, para a anterior primeiro
            self.parar_piscar_generico(self.aba_piscando_atualmente, self.texto_original_piscando)

        self.piscando = True
        self.aba_piscando_atualmente = aba_alvo
        self.texto_piscando = texto_piscando
        self.texto_original_piscando = self.notebook.tab(aba_alvo, "text")

        self.notebook.tab(aba_alvo, state="normal")
        self.notebook.select(aba_alvo)
        self._loop_piscar_generico()

    def _loop_piscar_generico(self):
        """O loop que efetivamente faz a aba piscar."""
        if not self.piscando:
            return

        texto_atual = self.notebook.tab(self.aba_piscando_atualmente, "text")
        novo_texto = self.texto_original_piscando if self.texto_piscando in texto_atual else self.texto_piscando
        self.notebook.tab(self.aba_piscando_atualmente, text=novo_texto)

        self.id_loop_piscar = self.root.after(600, self._loop_piscar_generico)

    def parar_piscar_generico(self, aba_alvo, texto_original):
        """Para a animação, restaura o texto e esconde a aba."""
        if hasattr(self, 'id_loop_piscar') and self.id_loop_piscar:
            self.root.after_cancel(self.id_loop_piscar)
        
        self.piscando = False
        self.notebook.tab(aba_alvo, text=texto_original)
        self.notebook.tab(aba_alvo, state="hidden")
        self.notebook.select(self.aba_atendimento)

    # --- FIM DO NOVO SISTEMA DE ANIMAÇÃO UNIFICADO ---

    def salvar_anotacoes_arquivo(self):
        try:
            from tkinter import filedialog
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")],
                title="Salvar Anotações"
            )
            
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    conteudo = self.anotacoes_textbox.get(1.0, tk.END)
                    f.write(f"=== ANOTAÇÕES CRAFT TABLE ===\n")
                    f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n")
                    f.write(f"Atendente: {self.registro_usuario.get()}\n")
                    f.write(f"{'='*40}\n\n")
                    f.write(conteudo)
                
                messagebox.showinfo("Sucesso", f"Anotações salvas em:\n{filename}")
                logger.info(f"Anotações salvas em: {filename}")
                
        except Exception as e:
            logger.error(f"Erro ao salvar anotações: {e}")
            messagebox.showerror("Erro", f"Erro ao salvar anotações: {str(e)}")

    def criar_sistema_com_credenciais(self, parent, nome, link):
        sistema_frame = ttk.Frame(parent)
        sistema_frame.pack(fill="x", pady=3, padx=10)
        label = ttk.Label(sistema_frame, text=nome, anchor="w", wraplength=250, justify=tk.LEFT)
        label.pack(side="left", padx=(0, 10))
        
        btn_acessar = ttk.Button(sistema_frame, text="Acessar", style="Destaque.TButton", 
                                 command=lambda: webbrowser.open(link))
        btn_acessar.pack(side="right")

    def criar_aba_historico(self):
        hist_notebook = ttk.Notebook(self.aba_historico)
        hist_notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        for categoria in ["Emergenciais", "Comerciais", "Informação", "Reclamações"]:
            frame_cat = ttk.Frame(hist_notebook)
            tree = ttk.Treeview(frame_cat, columns=("Data", "Serviço", "Nome", "Protocolo", "Atendente"), 
                                show="headings", height=15)
            tree.heading("Data", text="Data/Hora")
            tree.heading("Serviço", text="Serviço")
            tree.heading("Nome", text="Nome")
            tree.heading("Protocolo", text="Protocolo")
            tree.heading("Atendente", text="Atendente")
            tree.column("Data", width=150)
            tree.column("Serviço", width=180)
            tree.column("Nome", width=120)
            tree.column("Protocolo", width=120)
            tree.column("Atendente", width=120)
            
            scrollbar = ttk.Scrollbar(frame_cat, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            self.atualizar_historico_tree(tree, categoria)
            setattr(self, f"tree_{categoria.lower()}", tree)
            tree.bind("<Double-Button-1>", lambda e, cat=categoria: self.visualizar_detalhes_historico(cat))
            
            hist_scroll_command = lambda e, t=tree: t.yview_scroll(int(-1*(e.delta/120)), "units")
            tree.bind("<MouseWheel>", hist_scroll_command)

            hist_notebook.add(frame_cat, text=categoria)
        
        btn_frame = ttk.Frame(self.aba_historico)
        btn_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(btn_frame, text="Reset Histórico", command=self.reset_historico, style="Destaque.TButton").pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Exportar TXT", command=self.exportar_historico_txt).pack(side="right", padx=5)
            
    def carregar_formulario(self, event=None):
        entrada_bruta = self.entry_servico.get().strip()
        servico_id = None

        if not entrada_bruta:
            messagebox.showwarning("Aviso", "Por favor, digite ou selecione um serviço.")
            return
        
        try:
            entrada = entrada_bruta.split(' - ')[0].strip()
        except:
            entrada = entrada_bruta

        if entrada in self.servicos:
            servico_id = entrada
        else:
            entrada_lower = entrada_bruta.lower()
            for num, info in self.servicos.items():
                if entrada_lower in info["nome"].lower():
                    servico_id = num
                    break

        if not servico_id:
            messagebox.showerror("Erro", f"Serviço '{self.entry_servico.get()}' não encontrado!")
            self.servico_id_atual = None # <-- MUDANÇA: Limpa a memória em caso de erro
            return

        self.servico_id_atual = servico_id # <-- MUDANÇA: Salva o ID do serviço na memória

        servico_selecionado_formatado = f"{servico_id} - {self.servicos[servico_id]['nome']}"
        self.entry_servico.set(servico_selecionado_formatado)

        self.label_servico.config(text=self.servicos[servico_id]["nome"])
        campos_basicos = ["NOME", "TELEFONE", "CC/CPF/CNPJ/UC", "PROTOCOLO"]
        valores_preservados = {campo: self.entries[campo].get() for campo in campos_basicos if campo in self.entries and self.entries[campo]}

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.entries, self.radio_vars = {}, {}
        self.combo_descricao, self.combo_informacoes, self.combo_reclamacao = None, None, None
        
        if servico_id == "13":
            self.criar_form_Genesys()
            self.btn_registrar.config(state="normal")
            # CORREÇÃO: Usando a nova função genérica
            self.iniciar_piscar_generico(self.aba_tabulacao, "!! TABULAR !!")
            return
        elif servico_id == "6":
            self.criar_form_servico_6()
        else:
            self.adicionar_secao("Dados do Cliente")
            self.adicionar_campos(campos_basicos)

        for campo, valor in valores_preservados.items():
            if campo in self.entries and self.entries[campo]:
                self.entries[campo].insert(0, valor)

        if servico_id != "6":
            form_builder_name = f"criar_form_servico_{servico_id}"
            form_builder = getattr(self, form_builder_name, self.criar_form_servico_padrao)
            form_builder()

        self.canvas.yview_moveto(0)
        self.btn_registrar.config(state="normal")
        if self.servico_id_atual != "0":
            self.iniciar_piscar_generico(self.aba_tabulacao, "!! TABULAR !!")

        if self.entries:
            try:
                campos_foco = [k for k in self.entries.keys() if k not in campos_basicos]
                if campos_foco:
                    self.entries[campos_foco[0]].focus_set()
                else:
                    next(iter(self.entries.values())).focus_set()
            except (StopIteration, IndexError):
                pass

    def criar_form_servico_padrao(self):
        pass
            
    def criar_form_servico_1(self):
        self.adicionar_campo("PONTO DE REFERENCIA")
        self.adicionar_combobox_descricao()
        self.adicionar_campo("OBSERVAÇÃO DA OCORRÊNCIA")
        
    def criar_form_servico_2(self):
        self.criar_form_reclamacao_unificado()
    
    def criar_form_servico_3(self):
        self.criar_form_gravacao_telefonica()
        
    def criar_form_servico_4(self):
        self.criar_form_desligamento()

    def criar_form_servico_8(self):
        self.criar_form_baixa_renda()

    def criar_form_servico_11(self):
        self.adicionar_combobox_informacoes()

    def criar_form_servico_12(self):
        self.adicionar_campo("CIDADE")
        self.adicionar_campo("BAIRRO")
        self.adicionar_campo("LOGRADOURO")
        self.adicionar_campo("PONTO DE REFERENCIA")
        self.adicionar_combobox_descricao()
        self.adicionar_campo("OBSERVAÇÃO DA OCORRÊNCIA")
        
    def criar_form_servico_14(self):
        self.criar_form_cancelar_fatura_email()

    def criar_form_servico_15(self):
        self.adicionar_secao("Alterar Dados do Parceiro de Negócios")
        self.adicionar_campo("CAMPO A SER ALTERADO")
        self.adicionar_campo("VALOR ANTIGO")
        self.adicionar_campo("NOVO VALOR")

    def criar_form_servico_17(self):
        self.adicionar_secao("Mudança de Medidor de Local")
        self.adicionar_campo("CPF")
        self.adicionar_campo("MOTIVO DA MUDANÇA")

    def criar_form_servico_19(self):
        self.criar_form_danos_materiais()

    def criar_form_servico_20(self):
        self.criar_form_inativacao_baixa_renda()
        
    def criar_form_Genesys(self):
        self.adicionar_secao("Serviço Genesys - 🚀 REGISTRO DIRETO")
        
        frame_aviso = ttk.Frame(self.scrollable_frame)
        frame_aviso.pack(fill="x", pady=10, padx=20)
        ttk.Label(frame_aviso, 
                  text="🚀 Selecione uma opção e clique em 'Registrar' diretamente.", 
                  font=("Arial", self.tamanho_fonte_base, "italic"), 
                  foreground="blue").pack()
        
        frame = ttk.Frame(self.scrollable_frame)
        frame.pack(fill="x", pady=10, padx=20)
        
        opcoes = [
            ("Falta de Energia Individual", "FALTA DE ENERGIA INDIVIDUAL; O CLIENTE INFORMA QUE REALIZOU O TESTE DO DISJUNTOR."),
            ("Falta de Energia Coletiva", "FALTA DE ENERGIA GERAL; O CLIENTE INFORMA QUE REALIZOU O TESTE DO DISJUNTOR"),
        ]
        
        for texto, valor in opcoes:
            rb = ttk.Radiobutton(frame, text=texto, variable=self.var_Genesys, value=valor, 
                                 style="TRadiobutton")
            rb.pack(anchor="w", pady=5)
            
        self.canvas.yview_moveto(0)

    # Formulários auxiliares (mantidos do código original)
    def criar_form_gravacao_telefonica(self):
        self.adicionar_secao("Solicitação de Gravação Telefônica")
        self.adicionar_campos(["CPF", "PROTOCOLO", "DATA", "HORA", "E-MAIL"])

    def criar_form_desligamento(self):
        self.adicionar_secao("Desligamento Definitivo")
        self.adicionar_campos(["CPF", "PONTO DE REFERENCIA"])
        self.adicionar_campo_desc("DESCRIÇÃO", "SOLICITA O DESLIGAMENTO DEFINITIVO DA SUA CC")
        self.adicionar_campo_desc("AVISO", "INFORMADO QUE SE NÃO EFETUAR PAGAMENTO DOS DÉBITOS SEU NOME SERÁ NEGATIVADO.")
        self.adicionar_campo("MOTIVO")
        self.adicionar_radio_buttons("LEITURA ATUAL OU MÉDIA", 
                                     [("Por Média", "MEDIA"), ("Com Leitura (informar abaixo)", "LEITURA")])
        self.adicionar_campo("VALOR DA LEITURA ATUAL")
    
    def criar_form_baixa_renda(self):
        self.adicionar_secao("Cadastro Tarifa Social (Baixa Renda)")
        self.adicionar_campos(["CPF", "NIS", "CÓDIGO FAMILIAR"])
        self.adicionar_campo_desc("DESCRIÇÃO", "SOLICITA O CADASTRO BAIXA RENDA NESSA CONTA CONTRATO.")

    def criar_form_inativacao_baixa_renda(self):
        self.adicionar_secao("Inativação Tarifa Social (Baixa Renda)")
        self.adicionar_campos(["CPF", "NIS"])
        self.adicionar_campo_desc("DESCRIÇÃO", "SOLICITA A INATIVAÇÃO DO BAIXA RENDA NESSA CONTA CONTRATO.")
    
    def criar_form_cancelar_fatura_email(self):
        self.adicionar_secao("Cancelar Fatura por E-mail")
        self.adicionar_campos(["CPF"])
        self.adicionar_campo_desc("DESCRIÇÃO", "CLIENTE SOLICITA O CANCELAMENTO DE ENVIO DE FATURA POR E-MAIL E DESEJA RECEBER NOVAMENTE EM SUA UNIDADE.")
        self.adicionar_campo_desc("STATUS", "REALIZADO CONFORME PEDIDO DO CLIENTE.")

    def criar_form_danos_materiais(self):
        self.adicionar_secao("Danos Materiais")
        self.adicionar_campos([
            "DESCRIÇÃO DA OCORRÊNCIA COM DATA E HORA",
            "RELATO DO MOTIVO POR QUE SUPÕE QUE A RESPONSABILIDADE SEJA DA EMPRESA",
            "DESCRIÇÃO DO PRODUTO PERDIDO - ITEM 1", "DESCRIÇÃO DO PRODUTO PERDIDO - ITEM 2",
            "SOLUÇÃO PRETENDIDA", "MEIO DE COMUNICAÇÃO ESCOLHIDO PELO CLIENTE (E-MAIL)",
            "AUTORIZAÇÃO DE OUTRA PESSOA PARA RECEBER A RESPOSTA",
            "MEIO DE RESSARCIMENTO CASO A RECLAMAÇÃO SEJA PROCEDENTE",
            "INFORMAÇÕES ADICIONAIS", "OBSERVAÇÃO DA OCORRÊNCIA"
        ])

    def criar_form_reclamacao_geral(self):
        self.adicionar_secao("Tipo de Reclamação")
        frame = ttk.Frame(self.scrollable_frame)
        frame.pack(fill="x", pady=2)
        ttk.Label(frame, text="TIPO DE RECLAMAÇÃO:", width=25, anchor="w", style="Bold.TLabel").pack(side="left")
        
        opcoes = ["VALOR DA FATURA", "ERRO DE LEITURA", "PRAZOS", "APRESENTAÇÃO E ENTREGA DE FATURAS", "Outros"]
        self.combo_reclamacao = ttk.Combobox(frame, values=opcoes, state="readonly", width=40, font=self.fonte_atual)
        self.combo_reclamacao.pack(side="left", padx=(0, 5))
        self.combo_reclamacao.bind("<<ComboboxSelected>>", self.on_reclamacao_selected)

        self.frame_reclamacao_especifica = ttk.Frame(self.scrollable_frame)
        self.frame_reclamacao_especifica.pack(fill="x", pady=5)

    def on_reclamacao_selected(self, event=None):
        for widget in self.frame_reclamacao_especifica.winfo_children():
            widget.destroy()
        
        keys_to_remove = [k for k in self.entries if k.startswith("RECL_")]
        for key in keys_to_remove:
            del self.entries[key]
        keys_to_remove_radio = [k for k in self.radio_vars if k.startswith("RECL_")]
        for key in keys_to_remove_radio:
            del self.radio_vars[key]

        selecao = self.combo_reclamacao.get()
        if selecao == "VALOR DA FATURA":
            self.criar_form_reclamacao_fatura()
        elif selecao == "ERRO DE LEITURA":
            self.criar_form_reclamacao_leitura()
        elif selecao == "PRAZOS":
            self.criar_form_reclamacao_prazos()
        elif selecao == "APRESENTAÇÃO E ENTREGA DE FATURAS":
            self.criar_form_reclamacao_entrega()
        elif selecao == "OUTROS":
            self.criar_form_reclamacao_outros()

    def criar_form_reclamacao_fatura(self):
        parent = self.frame_reclamacao_especifica
        self.adicionar_campo("RECL_DESCRIÇÃO DA RECLAMAÇÃO", parent=parent)
        self.adicionar_campo("RECL_REMEDIAÇÃO PRETENDIDA", parent=parent)
        self.adicionar_campo("RECL_ANÁLISE DO ATENDENTE", parent=parent)
        self.adicionar_radio_buttons("RECL_MEIO DE RESPOSTA", [("Telefone", "TELEFONE"), ("E-mail", "EMAIL"), ("Carta", "CARTA")], parent=parent)
        self.adicionar_campo("RECL_INFORME O CONTATO (E-MAIL/TELEFONE)", parent=parent)
        self.adicionar_radio_buttons("RECL_AUTORIZA TERCEIROS?", [("Sim", "SIM"), ("Não", "NAO")], parent=parent)
        self.adicionar_campo("RECL_GRAU DE PARENTESCO", parent=parent)
        self.adicionar_campo("RECL_MEIO DE CONTATO DO TERCEiro", parent=parent)
        self.adicionar_campo("RECL_ACEITA RECEBIMENTO DE FATURA REFATURADA VIA WHATSAPP", parent=parent)
        self.adicionar_campo("RECL_OBSERVAÇÕES", parent=parent)

    def criar_form_reclamacao_leitura(self):
        parent = self.frame_reclamacao_especifica
        self.adicionar_campo("RECL_DESCRIÇÃO", parent=parent)
        self.adicionar_campo("RECL_MÊS DE REFERÊNCIA E VALOR DA FATURA", parent=parent)
        self.adicionar_campo("RECL_LEITURA ATUAL", parent=parent)
        self.adicionar_campo("RECL_REMEDIAÇÃO PRETENDIDA", parent=parent)
        self.adicionar_campo("RECL_ANÁLISE DO ATENDENTE", parent=parent)
        self.adicionar_radio_buttons("RECL_MEIO DE RESPOSTA", [("Telefone", "TELEFONE"), ("E-mail", "EMAIL"), ("Carta", "CARTA")], parent=parent)
        self.adicionar_campo("RECL_INFORME O CONTATO (E-MAIL/TELEFONE)", parent=parent)
        self.adicionar_radio_buttons("RECL_AUTORIZA TERCEIROS?", [("Sim", "SIM"), ("Não", "NAO")], parent=parent)
        self.adicionar_campo("RECL_NOME E PARENTESCO DO TERCEIRO", parent=parent)
        self.adicionar_campo("RECL_ACEITA RECEBIMENTO DE FATURA REFATURADA VIA WHATSAPP", parent=parent)
        self.adicionar_campo("RECL_TELEFONE DO TERCEIRO", parent=parent)
        self.adicionar_campo("RECL_INFORMAÇÕES ADICIONAIS", parent=parent)

    def criar_form_reclamacao_prazos(self):
        parent = self.frame_reclamacao_especifica
        self.adicionar_campo("RECL_QUAL SERVIÇO?", parent=parent)
        self.adicionar_campo("RECL_NÚMERO DA NOTA DE SERVIÇO", parent=parent)
        self.adicionar_campo("RECL_PRAZO DE EXECUÇÃO", parent=parent)
        self.adicionar_campo("RECL_DESCRIÇÃO", parent=parent)
        self.adicionar_campo("RECL_ANÁLISE DO ATENDENTE", parent=parent)
        self.adicionar_radio_buttons("RECL_FORMA DE RESPOSTA", [("Telefone", "TELEFONE"), ("E-mail", "EMAIL"), ("Carta", "CARTA")], parent=parent)
        self.adicionar_campo("RECL_INFORME O CONTATO (E-MAIL/TELEFONE)", parent=parent)
        self.adicionar_radio_buttons("RECL_AUTORIZA TERCEIROS?", [("Sim", "SIM"), ("Não", "NAO")], parent=parent)

    def criar_form_reclamacao_entrega(self):
        parent = self.frame_reclamacao_especifica
        self.adicionar_campo("RECL_MÊS DE REFERÊNCIA DA FATURA", parent=parent)
        self.adicionar_campo("RECL_DATA DE LEITURA DA ÚLTIMA FATURA", parent=parent)
        self.adicionar_campo("RECL_DESCRIÇÃO", parent=parent)
        self.adicionar_campo("RECL_REMEDIAÇÃO PRETENDIDA", parent=parent)
        self.adicionar_radio_buttons("RECL_RESPOSTA DA RECLAMAÇÃO", [("Telefone", "TELEFONE"), ("E-mail", "EMAIL"), ("Carta", "CARTA")], parent=parent)
        self.adicionar_campo("RECL_INFORME O CONTATO (E-MAIL/TELEFONE)", parent=parent)
        self.adicionar_radio_buttons("RECL_AUTORIZA TERCEIROS?", [("Sim", "SIM"), ("Não", "NAO")], parent=parent)
        self.adicionar_campo("RECL_NOME, PARENTESCO E CONTATO DO TERCEIRO", parent=parent)
        self.adicionar_campo("RECL_INFORMAÇÕES ADICIONAIS", parent=parent)
    
    def criar_form_reclamacao_outros(self):
        parent = self.frame_reclamacao_especifica
        self.adicionar_campo("RECL_DESCRIÇÃO E INFORMAÇÕES COMPLEMENTARES", parent=parent)
        self.adicionar_campo("RECL_REMEDIAÇÃO PRETENDIDA", parent=parent)
        self.adicionar_campo("RECL_ANÁLISE DO ATENDENTE", parent=parent)
        self.adicionar_radio_buttons("RECL_RESPOSTA DA RECLAMAÇÃO", [("Telefone", "TELEFONE"), ("E-mail", "EMAIL"), ("Carta", "CARTA")], parent=parent)
        self.adicionar_campo("RECL_INFORME O CONTATO (E-MAIL/TELEFONE)", parent=parent)
        self.adicionar_radio_buttons("RECL_AUTORIZA TERCEIROS?", [("Sim", "SIM"), ("Não", "NAO")], parent=parent)
    # === NOVOS FORMULÁRIOS REFORMULADOS ===
    
    def criar_form_servico_5(self):
        """NOVO: Serviço 5 - Nível de Tensão Reformulado"""
        self.adicionar_campo("PONTO DE REFERENCIA")
        
        self.adicionar_secao("Questionário de Nível de Tensão")
        
        # Novas perguntas técnicas específicas
        perguntas_tensao = [
            "Houve acréscimo de carga instalada?",
            "Lâmpadas queimam com frequência?", 
            "Lâmpadas piscam continuadamente?",
            "Lâmpadas perdem a luminosidade?",
            "Eletrodomésticos se auto desligam?",
            "A energia está oscilando (forte ou fraca)?",
            "Existe vizinho que utiliza motor, solda, bomba ou máquina de raio-x?"
        ]
        
        for pergunta in perguntas_tensao:
            self.adicionar_radio_buttons(pergunta, [("Sim", "SIM"), ("Não", "NÃO")])
        
        self.adicionar_campo("OBSERVAÇÃO DA OCORRÊNCIA")
        
        self.adicionar_secao("Resposta da Reclamação")
        self.adicionar_radio_buttons("MEIO DE RESPOSTA", [("Telefone", "TELEFONE"), ("Carta", "CARTA"), ("E-mail", "EMAIL")])
        self.adicionar_campo("QUAL E-MAIL/TELEFONE?")

    def criar_form_servico_6(self):
        """NOVO: Serviço 6 - Denúncia Anônima"""
        # SEM campos de cliente - Formulário anônimo
        self.adicionar_secao("Denúncia Anônima de Fraude")
        
        self.adicionar_campo("NOME DO RESPONSÁVEL DA FRAUDE")
        self.adicionar_campo("RUA") 
        self.adicionar_campo("BAIRRO")
        self.adicionar_campo("CIDADE")
        self.adicionar_campo("PONTO DE REFERÊNCIA")
        self.adicionar_campo("COR DA CASA")
        
        # Perguntas Sim/Não específicas
        self.adicionar_radio_buttons("CASA MURADA", [("Sim", "SIM"), ("Não", "NÃO")])
        self.adicionar_radio_buttons("TEM MEDIDOR", [("Sim", "SIM"), ("Não", "NÃO")])
        self.adicionar_campo("TEM HORÁRIO ESPECÍFICO PARA A FRAUDE")

    def criar_form_servico_7(self):
        """NOVO: Serviço 7 - Data Certa Simplificado"""
        self.adicionar_secao("Mudança de Data Certa")
        self.adicionar_campo("CPF")
        self.adicionar_campo("SOLICITA ALTERAÇÃO PARA DATA FIXA NO DIA")

    def criar_form_servico_10(self):
        """MODIFICADO: Serviço 10 - Religação com Seletor de Estado"""
        self.adicionar_campo("PONTO DE REFERENCIA")
        
        self.adicionar_secao("Seletor Inteligente de Religação")

        frame_estado = ttk.Frame(self.scrollable_frame)
        frame_estado.pack(fill="x", pady=2, padx=5)
        ttk.Label(frame_estado, text="ESTADO:", style="Bold.TLabel", anchor="w", width=35).pack(side="left")
        
        self.combo_estado_religacao = ttk.Combobox(
            frame_estado, 
            values=list(self.taxas_religacao_por_estado.keys()), 
            state="readonly"
        )
        self.combo_estado_religacao.pack(side="left", padx=10)
        self.combo_estado_religacao.set("Maranhão") 
        self.combo_estado_religacao.bind("<<ComboboxSelected>>", self.atualizar_valores_religacao)

        self.adicionar_radio_buttons("TIPO DE INSTALAÇÃO", [
            ("Monofásica", "Monofásica"), ("Bifásica", "Bifásica"), ("Trifásica", "Trifásica")
        ], command=self.atualizar_valores_religacao)

        self.adicionar_radio_buttons("TIPO DE RELIGAÇÃO", [
            ("Comum", "Comum"), ("Urgência", "Urgência")
        ], command=self.atualizar_valores_religacao)

        self.adicionar_campo_desc("VALOR DE SERVIÇO TAXADO", "Selecione as opções acima")
        self.adicionar_campo_desc("PRAZO DO SERVIÇO", "Selecione as opções acima")

        self.adicionar_radio_buttons("POP-UP DE RELIGAÇÃO DE CONFIANÇA FOI VERIFICADO?", [("Sim", "SIM"), ("Não", "NAO")])

        # MUDANÇA: Usando o novo sistema de faturas
        self.adicionar_secao("Dados do Pagamento")
        self.adicionar_controle_faturas() # <-- Chamando a nova função aqui!
        self.adicionar_campo("NOME LOCAL AGENTE ARRECADADOR") # Mantemos este campo

        self.adicionar_secao("Verificação de Débitos")
        self.adicionar_radio_buttons("FATURA EM ABERTO E VENCIDA?", [("Sim", "SIM"), ("Não", "NAO")])
        self.adicionar_radio_buttons("ENTRADA DE PARCELAMENTO?", [("Sim", "SIM"), ("Não", "NAO")])
        self.adicionar_radio_buttons("FATURA CNR?", [("Sim", "SIM"), ("Não", "NAO")])
        self.adicionar_radio_buttons("FATURA BLOQUEADA?", [("Sim", "SIM"), ("Não", "NAO")])
        
        self.atualizar_valores_religacao()

    def atualizar_valores_religacao(self, event=None):
        """MODIFICADO: Atualiza valor e prazo baseado no estado, instalação e tipo."""
        try:
            # Pega os valores selecionados dos 3 controles
            estado_selecionado = self.combo_estado_religacao.get()
            tipo_instalacao = self.radio_vars.get("TIPO DE INSTALAÇÃO", tk.StringVar()).get()
            tipo_religacao = self.radio_vars.get("TIPO DE RELIGAÇÃO", tk.StringVar()).get()

            # Verifica se todas as opções foram selecionadas
            if estado_selecionado and tipo_instalacao and tipo_religacao:
                # Busca a taxa no dicionário aninhado
                taxa = self.taxas_religacao_por_estado[estado_selecionado][tipo_religacao][tipo_instalacao]
                
                # Busca o prazo no dicionário de prazos
                prazo = self.prazos_religacao[tipo_religacao]
                
                # Lógica especial para urgência (apenas Pará)
                if tipo_religacao == "Urgência" and estado_selecionado != "Pará":
                    taxa = "N/A" # ou "Indisponível"
                    prazo = "Apenas no Pará"

                # Atualiza o campo de VALOR
                if "VALOR DE SERVIÇO TAXADO" in self.entries:
                    entry = self.entries["VALOR DE SERVIÇO TAXADO"]
                    entry.config(state="normal")
                    entry.delete(0, tk.END)
                    entry.insert(0, taxa)
                    entry.config(state="readonly")
                
                # Atualiza o campo de PRAZO
                if "PRAZO DO SERVIÇO" in self.entries:
                    entry = self.entries["PRAZO DO SERVIÇO"]
                    entry.config(state="normal")
                    entry.delete(0, tk.END)
                    entry.insert(0, prazo)
                    entry.config(state="readonly")
                    
        except (KeyError, AttributeError) as e:
            # Este bloco evita erros se algum widget ainda não foi criado ou uma chave não foi encontrada
            logger.warning(f"Ainda não é possível atualizar os valores de religação: {e}")
        except Exception as e:
            logger.error(f"Erro inesperado ao atualizar valores de religação: {e}")

    def criar_form_servico_16(self):
        """NOVO: Serviço 16 - Problema com Equipamento (Guia + Combobox Híbrido)"""
        self.adicionar_secao("Guia Rápido - Problema com Equipamento")
        
        # Informações de orientação
        info_frame = ttk.Frame(self.scrollable_frame)
        info_frame.pack(fill="x", pady=10, padx=5)
        
        # Quando Usar
        quando_usar_frame = ttk.LabelFrame(info_frame, text="✅ Quando Usar")
        quando_usar_frame.pack(fill="x", pady=2)
        ttk.Label(quando_usar_frame, text="• Medidor com display apagado\n• Medidor visivelmente danificado\n• Medidor não registra consumo", 
                  wraplength=400, justify="left").pack(padx=5, pady=5)
        
        # Atenção
        atencao_frame = ttk.LabelFrame(info_frame, text="⚠️ Atenção")
        atencao_frame.pack(fill="x", pady=2)
        ttk.Label(atencao_frame, text="• Verificar se não é problema interno\n• Confirmar se disjuntores estão ligados", 
                  wraplength=400, justify="left").pack(padx=5, pady=5)
        
        # Quando NÃO Usar
        nao_usar_frame = ttk.LabelFrame(info_frame, text="❌ Quando NÃO Usar")
        nao_usar_frame.pack(fill="x", pady=2)
        ttk.Label(nao_usar_frame, text="• Para problemas de falta de energia\n• Para questões de faturamento", 
                  wraplength=400, justify="left").pack(padx=5, pady=5)
        
        # Combobox híbrido
        self.adicionar_secao("Descrição do Problema")
        combo_frame = ttk.Frame(self.scrollable_frame)
        combo_frame.pack(fill="x", pady=2, padx=5)
        
        ttk.Label(combo_frame, text="PROBLEMA:", style="Bold.TLabel").pack(side="left")
        
        self.combo_problema = ttk.Combobox(combo_frame, values=self.problemas_equipamento, 
                                          state="readonly", width=20, font=self.fonte_atual)
        self.combo_problema.pack(side="left", padx=5)
        self.combo_problema.bind("<<ComboboxSelected>>", self.on_problema_selected)
        
        # Campo de texto editável que será preenchido pelo combobox
        self.entry_problema_desc = ttk.Entry(combo_frame, width=50, font=self.fonte_atual)
        self.entry_problema_desc.pack(side="left", fill="x", expand=True, padx=5)
        self.entries["DESCRIÇÃO DO PROBLEMA"] = self.entry_problema_desc

    def on_problema_selected(self, event=None):
        """Preenche o campo de texto com a opção selecionada no combobox"""
        if hasattr(self, 'combo_problema') and hasattr(self, 'entry_problema_desc'):
            selecao = self.combo_problema.get()
            self.entry_problema_desc.delete(0, tk.END)
            self.entry_problema_desc.insert(0, selecao)

    def criar_form_servico_18(self):
        """MODIFICADO: Serviço 18 - Substituir VALOR por MOTIVO"""
        self.adicionar_secao("Cancelar Atividades Acessórias")
        self.adicionar_campos(["CPF"])
        self.adicionar_campo_desc("INFO", "CLIENTE SOLICITA O CANCELAMENTO DA SEGUINTE ATIVIDADE ACESSÓRIA:")
        self.adicionar_campo("QUAL ATIVIDADE?")
        self.adicionar_campo("MOTIVO?")  # MUDOU DE "VALOR?" PARA "MOTIVO?"
        self.adicionar_campo_desc("STATUS", "REALIZADO CONFORME PEDIDO DO CLIENTE.")

    def criar_form_servico_9(self):
        """NOVO: Serviço 9 - Danos Elétricos Reformulado"""
        self.adicionar_secao("Dados do Titular")
        # Usa campos básicos automáticos
        
        self.adicionar_secao("Descrição da Ocorrência")
        self.adicionar_radio_buttons("TIPO DE EVENTO", [
            ("FALTA DE ENERGIA", "FALTA DE ENERGIA"),
            ("OSCILAÇÃO", "OSCILAÇÃO"), 
            ("SOBRETENSÃO", "SOBRETENSÃO"),
        ])
        self.adicionar_campo("DATA DA OCORRÊNCIA")
        self.adicionar_campo("HORA DA OCORRÊNCIA")
        
        
        self.adicionar_secao("Equipamentos Danificados")
        self.adicionar_controle_equipamentos()
        
        self.adicionar_secao("Checklist de Verificação")
        checklist_perguntas = [
            "Atingiu outras residências/instalações",
            "Faltou energia ou a energia oscilava antes da queima", 
            "Havia funcionário da empresa no local executando algum serviço",
            "Possui telefone fixo/antena parabólica",
            "Chovia no dia da ocorrência"
        ]
        
        for pergunta in checklist_perguntas:
            self.adicionar_radio_buttons(pergunta, [("SIM", "SIM"), ("NÃO", "NÃO")])
        
        self.adicionar_secao("Dados para Resposta")
        self.adicionar_radio_buttons("MEIO DE COMUNICAÇÃO ESCOLHIDO PELO CLIENTE", [
            ("Telefone", "TELEFONE"), 
            ("Carta", "CARTA"), 
            ("E-mail", "EMAIL")
        ])
        self.adicionar_campo("TELEFONE EXTRA PARA CONTATO")
        self.adicionar_campo("AUTORIZAÇÃO DE TERCEIRO (Nome, parentesco e telefone)")
        self.adicionar_radio_buttons("FORMA DE RESSARCIMENTO CASO A RECLAMAÇÃO SEJA PROCEDENTE", [
            ("Conta Poupança", "CONTA POUPANÇA"),
            ("Conta Corrente", "CONTA CORRENTE"),
        ])

    # === PROCESSADORES DE TEXTO REFORMULADOS ===
    
    def processar_servico_5(self, servico_id):
        """NOVO: Processador para Nível de Tensão reformulado"""
        texto = self.get_dados_basicos()
        texto += f"PONTO DE REFERENCIA: {self._get_entry_value('PONTO DE REFERENCIA')}\n\n"
        
        texto += "QUESTIONÁRIO DE NÍVEL DE TENSÃO:\n"
        
        perguntas_tensao = [
            "Houve acréscimo de carga instalada?",
            "Lâmpadas queimam com frequência?", 
            "Lâmpadas piscam continuadamente?",
            "Lâmpadas perdem a luminosidade?",
            "Eletrodomésticos se auto desligam?",
            "A energia está oscilando (forte ou fraca)?",
            "Existe vizinho que utiliza motor, solda, bomba ou máquina de raio-x?"
        ]
        
        for pergunta in perguntas_tensao:
            var = self.radio_vars.get(pergunta)
            if var:
                resposta = var.get()
                texto += f"- {pergunta} {resposta}\n"
        
        observacao = self._get_entry_value("OBSERVAÇÃO DA OCORRÊNCIA")
        if observacao:
            texto += f"\nOBSERVAÇÃO DA OCORRÊNCIA: {observacao}\n"
        
        meio_resposta = self.radio_vars.get("MEIO DE RESPOSTA")
        if meio_resposta:
            texto += f"\nMEIO DE RESPOSTA: {meio_resposta.get()}\n"
        
        contato = self._get_entry_value("QUAL E-MAIL/TELEFONE?")
        if contato:
            texto += f"CONTATO: {contato}\n"
        
        return texto

    def processar_servico_6(self, servico_id):
        """NOVO: Processador para Denúncia Anônima"""
        texto = "CLIENTE ANÔNIMO\n"
        texto += "DESCRIÇÃO DA DENUNCIA: INFORMA POSSÍVEL FURTO DE ENERGIA OCORRENDO NO LOCAL INDICADO.\n\n"
        
        campos_denuncia = [
            "NOME DO RESPONSÁVEL DA FRAUDE",
            "RUA", 
            "BAIRRO",
            "CIDADE", 
            "PONTO DE REFERÊNCIA",
            "APARENCIA DA CASA"
        ]
        
        for campo in campos_denuncia:
            valor = self._get_entry_value(campo)
            texto += f"{campo}: {valor}\n"
        
        # Perguntas Sim/Não
        perguntas_radio = ["CASA MURADA", "TEM MEDIDOR", "TEM HORÁRIO ESPECÍFICO PARA A FRAUDE"]
        for pergunta in perguntas_radio:
            var = self.radio_vars.get(pergunta)
            if var:
                texto += f"{pergunta}: {var.get()}\n"
        
        return texto

    def processar_servico_7(self, servico_id):
        """NOVO: Processador para Data Certa simplificado"""
        texto = self.get_dados_basicos()
        cpf = self._get_entry_value("CPF")
        dia = self._get_entry_value("SOLICITA ALTERAÇÃO PARA DATA FIXA NO DIA")
        
        texto += f"CPF: {cpf}\n"
        texto += f"CLIENTE SOLICITA ALTERAÇÃO DA DATA DE VENCIMENTO DA CONTA PARA O DIA {dia}.\n"
        texto += "DADOS CONFIRMADOS DURANTE A LIGAÇÃO\n"
        
        return texto
    
    def processar_servico_9(self, servico_id):
        """NOVO: Processador para o serviço 9 - Danos Elétricos."""
        texto = self.get_dados_basicos()
        texto += "\n--- DESCRIÇÃO DA OCORRÊNCIA ---\n"
        
        # Coleta o tipo de evento
        tipo_evento_var = self.radio_vars.get("TIPO DE EVENTO")
        if tipo_evento_var:
            texto += f"TIPO DE EVENTO: {tipo_evento_var.get()}\n"
            
        texto += f"DATA DA OCORRÊNCIA: {self._get_entry_value('DATA DA OCORRÊNCIA')}\n"
        texto += f"HORA DA OCORRÊNCIA: {self._get_entry_value('HORA DA OCORRÊNCIA')}\n"
        
        # Coleta os equipamentos danificados
        texto += "\n--- EQUIPAMENTOS DANIFICADOS ---\n"
        try:
            num_equipamentos = int(self.spin_equipamentos.get())
            for i in range(1, num_equipamentos + 1):
                aparelho = self._get_entry_value(f"APARELHO_{i}")
                marca = self._get_entry_value(f"MARCA_{i}")
                modelo = self._get_entry_value(f"MODELO_{i}")
                tempo_uso = self._get_entry_value(f"TEMPO DE USO_{i}")
                texto += f"Equipamento {i}: {aparelho} - Marca: {marca} - Modelo: {modelo} - Tempo de Uso: {tempo_uso}\n"
        except (ValueError, AttributeError):
            texto += "Nenhum equipamento informado.\n"

        # --- A PARTE MAIS IMPORTANTE: LER O CHECKLIST ---
        texto += "\n--- CHECKLIST DE VERIFICAÇÃO ---\n"
        checklist_perguntas = [
            "Atingiu outras residências/instalações",
            "Faltou energia ou a energia oscilava antes da queima",
            "Havia funcionário da empresa no local executando algum serviço",
            "Possui telefone fixo/antena parabólica",
            "Chovia no dia da ocorrência"
        ]
        
        for pergunta in checklist_perguntas:
            var = self.radio_vars.get(pergunta)
            if var:
                resposta = var.get() if var.get() else "NÃO INFORMADO"
                texto += f"- {pergunta}? {resposta}\n"
        
        # Coleta os dados para resposta
        texto += "\n--- DADOS PARA RESPOSTA ---\n"
        meio_comunicacao_var = self.radio_vars.get("MEIO DE COMUNICAÇÃO ESCOLHIDO PELO CLIENTE")
        if meio_comunicacao_var:
            texto += f"MEIO DE COMUNICAÇÃO: {meio_comunicacao_var.get()}\n"
            
        texto += f"TELEFONE EXTRA PARA CONTATO: {self._get_entry_value('TELEFONE EXTRA PARA CONTATO')}\n"
        texto += f"AUTORIZAÇÃO DE TERCEIRO: {self._get_entry_value('AUTORIZAÇÃO DE TERCEIRO (Nome, parentesco e telefone)')}\n"

        ressarcimento_var = self.radio_vars.get("FORMA DE RESSARCIMENTO CASO A RECLAMAÇÃO SEJA PROCEDENTE")
        if ressarcimento_var:
            texto += f"FORMA DE RESSARCIMENTO: {ressarcimento_var.get()}\n"
            
        return texto

    def processar_servico_10(self, servico_id):
        """MODIFICADO: Processador para Religação com lista dinâmica de faturas"""
        texto = "SOLICITA RELIGAÇÃO:\n"
        texto += self.get_dados_basicos()
        texto += f"PONTO DE REFERENCIA: {self._get_entry_value('PONTO DE REFERENCIA')}\n"
        
        estado = self.combo_estado_religacao.get()
        tipo_instalacao = self.radio_vars.get("TIPO DE INSTALAÇÃO", tk.StringVar()).get()
        tipo_religacao = self.radio_vars.get("TIPO DE RELIGAÇÃO", tk.StringVar()).get()
        
        texto += f"ESTADO: {estado}\n"
        texto += f"TIPO DE INSTALAÇÃO: {tipo_instalacao}\n"
        texto += f"TIPO DE RELIGAÇÃO: {tipo_religacao}\n"
        
        valor = self._get_entry_value("VALOR DE SERVIÇO TAXADO")
        prazo = self._get_entry_value("PRAZO DO SERVIÇO")
        texto += f"VALOR DE SERVIÇO TAXADO: {valor}\n"
        texto += f"PRAZO DO SERVIÇO: {prazo}\n"
        
        popup_verif = self.radio_vars.get("POP-UP DE RELIGAÇÃO DE CONFIANÇA FOI VERIFICADO?", tk.StringVar()).get()
        texto += f"POP-UP DE RELIGAÇÃO DE CONFIANÇA FOI VERIFICADO? {popup_verif}\n"
        
        texto += f"\n--- Dados do Pagamento ---\n"
        try:
            num_faturas = int(self.spin_faturas.get())
            
            # MUDANÇA: Verifica se a quantidade de faturas é 0
            if num_faturas == 0:
                texto += "Nenhuma fatura informada para esta religação (ex: religação de confiança).\n"
            else:
                for i in range(1, num_faturas + 1):
                    texto += f"--- Fatura {i} ---\n"
                    # Em processar_servico_10 (CORRIGIDO)
                    mes_referente = self._get_entry_value(f"MES_REFERENTE_{i}")
                    data_pgto = self._get_entry_value(f"DATA_PGTO_{i}")
                    hora_pgto = self._get_entry_value(f"HORA_PGTO_{i}")
                    valor_fatura = self._get_entry_value(f"VALOR_{i}")

                    texto += f"  MÊS REFERENTE: {mes_referente}\n"
                    texto += f"  DATA PGTO: {data_pgto}\n"
                    texto += f"  HORA PGTO: {hora_pgto}\n"
                    texto += f"  VALOR: {valor_fatura}\n"
        except (ValueError, AttributeError):
            texto += "Nenhuma fatura detalhada informada.\n"

        texto += f"NOME LOCAL AGENTE ARRECADADOR: {self._get_entry_value('NOME LOCAL AGENTE ARRECADADOR')}\n"
        
        texto += f"\n--- Verificação de Débitos ---\n"
        verificacoes = ["FATURA EM ABERTO E VENCIDA?", "ENTRADA DE PARCELAMENTO?", "FATURA CNR?", "FATURA BLOQUEADA?"]
        for verificacao in verificacoes:
            var = self.radio_vars.get(verificacao)
            if var:
                valor_verif = var.get()
                texto += f"{verificacao} {valor_verif}\n"
                
        return texto
    
    # SUBSTITUA SUA FUNÇÃO ANTIGA POR ESTA

    def adicionar_controle_faturas(self):
        """Cria o seletor e PRÉ-CRIA todos os campos de fatura (escondidos)."""
        MAX_FATURAS = 10  # Define um máximo razoável de faturas

        secao_frame = ttk.Frame(self.scrollable_frame)
        secao_frame.pack(fill="x", pady=5, padx=5)

        control_frame = ttk.Frame(secao_frame)
        control_frame.pack(fill="x")

        ttk.Label(control_frame, text="QUANTIDADE DE FATURAS:").pack(side="left")

        self.spin_faturas = ttk.Spinbox(control_frame, from_=0, to=MAX_FATURAS, width=5, command=self.atualizar_faturas)
        self.spin_faturas.set(1)
        self.spin_faturas.pack(side="left", padx=5)

        # Container onde os campos de fatura vão aparecer/desaparecer
        self.frame_container_faturas = ttk.Frame(secao_frame)
        self.frame_container_faturas.pack(fill="x", pady=5)

        # Limpa a lista de frames de faturas da sessão anterior
        self.faturas_frames.clear()

        # CRIA TODOS OS CAMPOS DE UMA VEZ E OS ESCONDE
        for i in range(MAX_FATURAS):
            # O widget "pai" é o novo container que criamos acima
            fatura_frame = ttk.LabelFrame(self.frame_container_faturas, text=f"Detalhes da Fatura {i+1}")

            campos_fatura = [
                ("MÊS REFERENTE", f"MES_REFERENTE_{i + 1}"),
                ("DATA PGTO", f"DATA_PGTO_{i + 1}"),
                ("HORA PGTO", f"HORA_PGTO_{i + 1}"),
                ("VALOR", f"VALOR_{i + 1}")
            ]

            for nome_label, chave_entry in campos_fatura:
                field_frame = ttk.Frame(fatura_frame)
                field_frame.pack(fill="x", pady=2, padx=5)
                ttk.Label(field_frame, text=f"{nome_label}:", width=20).pack(side="left")
                entry = ttk.Entry(field_frame, width=30)
                entry.pack(side="left", fill="x", expand=True)
                self.entries[chave_entry] = entry

            # Guarda a referência do frame na nossa lista
            self.faturas_frames.append(fatura_frame)

        # Chama a função para mostrar o número correto de faturas
        self.atualizar_faturas()

    def atualizar_faturas(self):
        """Apenas mostra ou esconde os frames de fatura pré-criados. É MUITO RÁPIDO."""
        try:
            num_a_mostrar = int(self.spin_faturas.get())
        except (ValueError, tk.TclError):
            num_a_mostrar = 1

        # Itera por todos os frames pré-criados na lista self.faturas_frames
        for i, frame in enumerate(self.faturas_frames):
            if i < num_a_mostrar:
                # Mostra o frame se ele estiver dentro da quantidade desejada
                frame.pack(fill="x", pady=3, padx=2, ipady=5)
            else:
                # Esconde o frame se ele não for mais necessário
                frame.pack_forget()
    
    def processar_servico_15(self, servico_id):
        texto = self.get_dados_basicos()
        texto += "CLIENTE SOLICITA ALTERAÇÃO DE DADOS CADASTRAIS:\n"
        texto += f"CAMPO A SER ALTERADO: {self._get_entry_value('CAMPO A SER ALTERADO')}\n"
        texto += f"VALOR ANTIGO: {self._get_entry_value('VALOR ANTIGO')}\n"
        texto += f"NOVO VALOR: {self._get_entry_value('NOVO VALOR')}\n"
        return texto

    def processar_servico_16(self, servico_id):
        """NOVO: Processador para Problema com Equipamento"""
        texto = self.get_dados_basicos()
        
        descricao_problema = self._get_entry_value("DESCRIÇÃO DO PROBLEMA")
        texto += f"DESCRIÇÃO DO PROBLEMA: {descricao_problema}\n"
        
        return texto

    def processar_servico_18(self, servico_id):
        """MODIFICADO: Processador para Cancelamento com MOTIVO"""
        texto = self.get_dados_basicos()
        campos = ["CPF"]
        for campo in campos: 
            texto += f"{campo.upper()}: {self._get_entry_value(campo)}\n"
        
        texto += "CLIENTE SOLICITA O CANCELAMENTO DA SEGUINTE ATIVIDADE ACESSÓRIA\n"
        texto += f"QUAL ATIVIDADE? {self._get_entry_value('QUAL ATIVIDADE?')}\n"
        texto += f"MOTIVO? {self._get_entry_value('MOTIVO?')}\n"  # MUDOU DE VALOR PARA MOTIVO
        texto += "REALIZADO CONFORME PEDIDO DO CLIENTE.\n"
        
        return texto
        
    def adicionar_campos(self, lista_campos, parent=None):
        if parent is None:
            parent = self.scrollable_frame
        for campo in lista_campos:
            self.adicionar_campo(campo, parent=parent)
            
    def on_descricao_selected(self, event=None):
        if hasattr(self, 'combo_descricao') and self.combo_descricao:
            selecao = self.combo_descricao.get()
            is_custom = selecao == "PERSONALIZADA"
            if hasattr(self, 'entry_descricao_custom'):
                self.entry_descricao_custom.config(state="normal" if is_custom else "disabled")
                if is_custom:
                    self.entry_descricao_custom.delete(0, tk.END)
                    self.entry_descricao_custom.focus()
                else:
                    self.entry_descricao_custom.delete(0, tk.END)

    def obter_descricao(self):
        if not hasattr(self, 'combo_descricao') or not self.combo_descricao:
            return ""
        selecao = self.combo_descricao.get()
        if selecao == "PERSONALIZADA":
            return self.entry_descricao_custom.get().strip()
        elif selecao in self.descricoes_emergenciais:
            return self.descricoes_emergenciais[selecao]
        else:
            return ""

    def on_informacao_selected(self, event=None):
        if hasattr(self, 'combo_informacoes') and self.combo_informacoes:
            selecao = self.combo_informacoes.get()
            is_custom = selecao == "PERSONALIZADA"
            if hasattr(self, 'entry_informacao_custom'):
                self.entry_informacao_custom.config(state="normal" if is_custom else "disabled")
                if is_custom:
                    self.entry_informacao_custom.delete(0, tk.END)
                    self.entry_informacao_custom.focus()
                else:
                    self.entry_informacao_custom.delete(0, tk.END)

    def obter_informacao(self):
        if not hasattr(self, 'combo_informacoes') or not self.combo_informacoes:
            return ""
        selecao = self.combo_informacoes.get()
        if selecao == "PERSONALIZADA":
            return self.entry_informacao_custom.get().strip()
        elif selecao in self.opcoes_informacoes:
            return selecao
        else:
            return ""

    def adicionar_secao(self, titulo, parent=None):
        if parent is None:
            parent = self.scrollable_frame
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(10, 5))
        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=(0, 5))
        ttk.Label(frame, text=titulo, font=("Arial", self.tamanho_fonte_base + 2, "bold")).pack()
            
    def adicionar_campo(self, nome_campo, parent=None):
        if parent is None:
            parent = self.scrollable_frame
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=2, padx=5)
        label_text = f"{nome_campo.replace('RECL_', '').replace('_', ' ').upper()}:"
        label = ttk.Label(frame, text=label_text, style="Bold.TLabel", anchor="w", 
                          wraplength=350, justify=tk.LEFT)
        label.pack(side="left", padx=(0, 5))
        
        entry = ttk.Entry(frame, width=50, font=self.fonte_atual)
        entry.pack(side="left", fill="x", expand=True)
        self.entries[nome_campo] = entry

    def adicionar_campo_desc(self, nome_campo, texto_preenchido, parent=None):
        if parent is None:
            parent = self.scrollable_frame
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=2, padx=5)
        label_text = f"{nome_campo.upper()}:"
        
        label = ttk.Label(frame, text=label_text, style="Bold.TLabel", anchor="w",
                          wraplength=350, justify=tk.LEFT)
        label.pack(side="left", padx=(0, 5))
        
        entry = ttk.Entry(frame, width=50, font=self.fonte_atual)
        entry.insert(0, texto_preenchido)
        entry.config(state="readonly")
        entry.pack(side="left", fill="x", expand=True)
        self.entries[nome_campo] = entry

    def adicionar_radio_buttons(self, nome_grupo, opcoes, parent=None, command=None, var=None):
        if parent is None:
            parent = self.scrollable_frame
        group_frame = ttk.Frame(parent)
        group_frame.pack(fill="x", pady=2, padx=5)
        ttk.Label(group_frame, text=f"{nome_grupo.replace('RECL_', '').replace('_', ' ').upper()}:", style="Bold.TLabel", anchor="w", width=35).pack(side="left")

        # Usa a variável passada como argumento ou cria uma nova
        variable = var if var is not None else tk.StringVar(value=None)
        self.radio_vars[nome_grupo] = variable

        for texto, valor in opcoes:
            rb = ttk.Radiobutton(group_frame, text=texto, variable=variable, value=valor, command=command)
            rb.pack(side="left", padx=10, pady=2)

    def adicionar_controle_equipamentos(self):
        control_frame = ttk.Frame(self.scrollable_frame)
        control_frame.pack(fill="x", pady=5)
        ttk.Label(control_frame, text="Número de equipamentos:", font=self.fonte_bold_atual).pack(side="left")
        self.spin_equipamentos = ttk.Spinbox(control_frame, from_=1, to=10, width=5, font=self.fonte_atual, 
                                             command=self.atualizar_equipamentos)
        self.spin_equipamentos.set(1)
        self.spin_equipamentos.pack(side="left", padx=5)
        self.frame_equipamentos = ttk.Frame(self.scrollable_frame)
        self.frame_equipamentos.pack(fill="x", pady=5)
        self.atualizar_equipamentos()

    def atualizar_equipamentos(self):
        if hasattr(self, 'equipamentos_frames'):
            for frame in self.equipamentos_frames:
                frame.destroy()
        self.equipamentos_frames = []

        keys_to_remove = [k for k in self.entries.keys() if k.startswith(("APARELHO_", "MARCA_", "MODELO_", "TEMPO DE USO_"))]
        for key in keys_to_remove:
            if key in self.entries:
                del self.entries[key]
        try:
            num = int(self.spin_equipamentos.get())
        except (ValueError, tk.TclError):
            num = 1
        
        for i in range(num):
            equip_frame = ttk.LabelFrame(self.frame_equipamentos, text=f"Equipamento {i+1}")
            equip_frame.pack(fill="x", pady=2)
            campos_equip = ["APARELHO", "MARCA", "MODELO", "TEMPO DE USO"]
            for campo in campos_equip:
                field_frame = ttk.Frame(equip_frame)
                field_frame.pack(fill="x", pady=1)
                ttk.Label(field_frame, text=f"{campo}:", width=15, font=self.fonte_bold_atual).pack(side="left")
                entry = ttk.Entry(field_frame, width=30, font=self.fonte_atual)
                entry.pack(side="left", fill="x", expand=True)
                self.entries[f"{campo}_{i+1}"] = entry
            self.equipamentos_frames.append(equip_frame)

    def adicionar_combobox_descricao(self):
        frame = ttk.Frame(self.scrollable_frame)
        frame.pack(fill="x", pady=2)
        ttk.Label(frame, text="DESCRIÇÃO:", width=25, anchor="w", style="Bold.TLabel").pack(side="left")
        combo_frame = ttk.Frame(frame)
        combo_frame.pack(side="left", fill="x", expand=True)
        opcoes = [""] + list(self.descricoes_emergenciais.keys()) + ["PERSONALIZADA"]
        self.combo_descricao = ttk.Combobox(combo_frame, values=opcoes, state="readonly", width=30, font=self.fonte_atual)
        self.combo_descricao.pack(side="left", padx=(0, 5))
        self.combo_descricao.bind("<<ComboboxSelected>>", self.on_descricao_selected)
        self.entry_descricao_custom = ttk.Entry(combo_frame, width=50, state="disabled", font=self.fonte_atual)
        self.entry_descricao_custom.pack(side="left", fill="x", expand=True)

    def adicionar_combobox_informacoes(self):
        frame = ttk.Frame(self.scrollable_frame)
        frame.pack(fill="x", pady=2)
        ttk.Label(frame, text="CLIENTE INFORMADO SOBRE:", width=25, anchor="w", style="Bold.TLabel").pack(side="left")
        combo_frame = ttk.Frame(frame)
        combo_frame.pack(side="left", fill="x", expand=True)
        opcoes = [""] + self.opcoes_informacoes + ["PERSONALIZADA"]
        self.combo_informacoes = ttk.Combobox(combo_frame, values=opcoes, state="readonly", width=40, font=self.fonte_atual)
        self.combo_informacoes.pack(side="left", padx=(0, 5))
        self.combo_informacoes.bind("<<ComboboxSelected>>", self.on_informacao_selected)
        self.entry_informacao_custom = ttk.Entry(combo_frame, width=50, state="disabled", font=self.fonte_atual)
        self.entry_informacao_custom.pack(side="left", fill="x", expand=True)

    # === PROCESSADORES DE SERVIÇOS ORIGINAIS MANTIDOS ===
    
    def processar_servico_padrao(self, servico_id):
        texto = self.get_dados_basicos()
        if texto:
            texto += "\n"

        campos_cliente = ["NOME", "TELEFONE", "CC/CPF/CNPJ/UC", "PROTOCOLO"]
        campos_servico = [k for k in self.entries.keys() if k not in campos_cliente and not k.startswith("RECL_")]

        for campo in campos_servico:
            valor = self._get_entry_value(campo)
            if valor:
                texto += f"{campo.replace('_', ' ').upper()}: {valor}\n"
        return texto
    
    def processar_servico_1(self, servico_id):
        texto = self.get_dados_basicos()
        texto += f"PONTO DE REFERENCIA: {self._get_entry_value('PONTO DE REFERENCIA')}\n"
        
        descricao = self.obter_descricao()
        if descricao:
            texto += f"DESCRIÇÃO: {descricao}\n"

        texto += f"OBSERVAÇÃO DA OCORRÊNCIA: {self._get_entry_value('OBSERVAÇÃO DA OCORRÊNCIA')}\n"
        return texto

    def processar_servico_2(self, servico_id):
        return self.processar_reclamacao_unificada()

    def processar_servico_3(self, servico_id):
        texto = self.get_dados_basicos()
        campos = ["CPF", "PROTOCOLO", "DATA", "HORA", "E-MAIL"]
        for campo in campos: texto += f"{campo.upper()}: {self._get_entry_value(campo)}\n"
        return texto

    def processar_servico_4(self, servico_id):
        texto = self.get_dados_basicos()
        campos = ["CPF", "PONTO DE REFERENCIA"]
        for campo in campos: texto += f"{campo.upper()}: {self._get_entry_value(campo)}\n"
        texto += "DESCRIÇÃO: SOLICITA O DESLIGAMENTO DEFINITIVO DA SUA CC\n"
        texto += "INFORMADO QUE SE NÃO EFETUAR PAGAMENTO DOS DÉBITOS SEU NOME SERÁ NEGATIVADO.\n"
        texto += f"MOTIVO: {self._get_entry_value('MOTIVO')}\n"
        var_leitura = self.radio_vars.get("LEITURA ATUAL OU MÉDIA")
        leitura_tipo = var_leitura.get() if var_leitura else "" # Pega o valor se a variável existir

        if leitura_tipo == "MEDIA":
            texto += "LEITURA ATUAL OU MÉDIA: POR MEDIA\n"
        else:
            texto += f"LEITURA ATUAL OU MÉDIA: COM LEITURA ({self._get_entry_value('VALOR DA LEITURA ATUAL')})\n"
        return texto

    def processar_servico_8(self, servico_id):
        texto = self.get_dados_basicos()
        texto += "SOLICITA O CADASTRO BAIXA RENDA NESSA CONTA CONTRATO.\n"
        campos = ["CPF", "NIS", "CÓDIGO FAMILIAR"]
        for campo in campos: texto += f"{campo.upper()}: {self._get_entry_value(campo)}\n"
        return texto
    
    def processar_servico_11(self, servico_id):
        texto = self.get_dados_basicos()
        
        informacao = self.obter_informacao()
        if informacao:
            texto += f"CLIENTE INFORMADO SOBRE: {informacao}\n"
        
        return texto
    
    def processar_servico_12(self, servico_id):
        texto = self.get_dados_basicos()
        campos_atc = ["CIDADE", "BAIRRO", "LOGRADOURO", "PONTO DE REFERENCIA"]
        for campo in campos_atc:
            valor = self._get_entry_value(campo)
            if valor:
                texto += f"{campo.upper()}: {valor}\n"

        descricao = self.obter_descricao()
        if descricao:
            texto += f"DESCRIÇÃO: {descricao}\n"

        texto += f"OBSERVAÇÃO DA OCORRÊNCIA: {self._get_entry_value('OBSERVAÇÃO DA OCORRÊNCIA')}\n"
        return texto

    def processar_servico_14(self, servico_id):
        texto = self.get_dados_basicos()
        texto += f"CPF: {self._get_entry_value('CPF')}\n"
        texto += "CLIENTE SOLICITA O CANCELAMENTO DE ENVIO DE FATURA POR E-MAIL\nDESEJA RECEBER NOVAMENTE EM SUA UNIDADE.\nREALIZADO CONFORME PEDIDO DO CLIENTE.\n"
        return texto

    def processar_servico_19(self, servico_id):
        texto = self.get_dados_basicos()
        campos = [
            "DESCRIÇÃO DA OCORRÊNCIA COM DATA E HORA",
            "RELATO DO MOTIVO POR QUE SUPÕE QUE A RESPONSABILIDADE SEJA DA EMPRESA",
            "DESCRIÇÃO DO PRODUTO PERDIDO - ITEM 1", "DESCRIÇÃO DO PRODUTO PERDIDO - ITEM 2",
            "SOLUÇÃO PRETENDIDA", "MEIO DE COMUNICAÇÃO ESCOLHIDO PELO CLIENTE (E-MAIL)",
            "AUTORIZAÇÃO DE OUTRA PESSOA PARA RECEBER A RESPOSTA",
            "MEIO DE RESSARCIMENTO CASO A RECLAMAÇÃO SEJA PROCEDENTE",
            "INFORMAÇÕES ADICIONAIS", "OBSERVAÇÃO DA OCORRÊNCIA"
        ]
        for campo in campos: texto += f"{campo.upper()}: {self._get_entry_value(campo)}\n"
        return texto
    
    def processar_servico_20(self, servico_id):
        texto = self.get_dados_basicos()
        texto += "SOLICITA A INATIVAÇÃO DO BAIXA RENDA NESSA CONTA CONTRATO.\n"
        campos = ["CPF", "NIS"]
        for campo in campos: texto += f"{campo.upper()}: {self._get_entry_value(campo)}\n"
        return texto

    # === MÉTODOS PRINCIPAIS DE OPERAÇÃO ===

    def get_dados_basicos(self):
        texto = ""
        campos_cliente = ["NOME", "TELEFONE", "CC/CPF/CNPJ/UC", "PROTOCOLO"]
        for campo in campos_cliente:
            valor = self._get_entry_value(campo)
            if valor:
                texto += f"{campo.upper()}: {valor}\n"
        return texto

    def gerar_texto(self):
        try:
            # CORREÇÃO: Usa o ID que foi salvo na memória, em vez de ler o combobox.
            servico_id = self.servico_id_atual
            
            # A verificação de erro agora é mais simples e confiável.
            if not servico_id:
                messagebox.showerror("Erro", "Nenhum serviço foi carregado. Por favor, selecione um serviço e carregue o formulário antes de registrar.")
                return
            
            if not self.registro_usuario.get():
                messagebox.showerror("Erro", "Informe a Sua Matricula!")
                return
            
            if servico_id != "6" and servico_id != "13":
                campos_basicos = ["NOME", "PROTOCOLO"]
                for campo in campos_basicos:
                    if not self._get_entry_value(campo):
                        messagebox.showerror("Erro", f"O campo '{campo}' é obrigatório e não foi preenchido!")
                        return

            if servico_id == "13":
                if not self.var_Genesys.get():
                    messagebox.showerror("Erro", "Selecione uma opção do Genesys antes de registrar!")
                    self.iniciar_piscar_tabulacao()
                    return
                texto = self.var_Genesys.get()
                self.salvar_e_copiar_texto(servico_id, texto)
                messagebox.showinfo(
                    "✅ Genesys Registrado", 
                    "Genesys registrado com sucesso!\n\n"
                    "✅ Texto copiado para área de transferência!"
                )
                return
                                      
            texto = f"SERVIÇO: {self.servicos[servico_id]['nome']}\n"
            
            process_function = getattr(self, f"processar_servico_{servico_id}", self.processar_servico_padrao)
            texto += process_function(servico_id)

            texto += f"\n{self.registro_usuario.get()}"
            
            self.salvar_e_copiar_texto(servico_id, texto)
            
            ToastNotification(
                self.root, 
                "✅ Atendimento Concluído", 
                "Registro salvo e texto copiado para a área de transferência!"
            )
            if datetime.now().weekday() == 1: # 4 = Sexta-feira
                self.iniciar_piscar_generico(self.aba_pesquisa, "!! PESQUISA !!")

        except Exception as e:
            logger.error(f"Erro ao gerar texto: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao gerar texto: {str(e)}")

    

    def salvar_e_copiar_texto(self, servico_id, texto):
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, texto)
        self.root.clipboard_clear()
        self.root.clipboard_append(texto)
        self.salvar_registro_historico(servico_id, texto)
        logger.info(f"Texto gerado para serviço '{self.servicos[servico_id]['nome']}': {texto.splitlines()[0]}")
        
    def salvar_registro_historico(self, servico, texto):
        info_servico = self.servicos[servico]
        categoria = info_servico["categoria"]
        
        nome_val = self._get_entry_value("NOME") if servico != "6" else "Denúncia"
        protocolo = self._get_entry_value("PROTOCOLO")
        
        registro = {
            "data": datetime.now().strftime('%d/%m/%Y %H:%M'), 
            "servico": info_servico["nome"],
            "nome": nome_val, 
            "protocolo": protocolo, 
            "atendente": self.registro_usuario.get(), 
            "texto_completo": texto
        }
        
        self.history_manager.adicionar_registro(categoria, registro)

        tree = getattr(self, f"tree_{categoria.lower()}", None)
        if tree:
            self.atualizar_historico_tree(tree, categoria)

    def limpar_campos(self):
        self.btn_registrar.config(state="disabled")

        if hasattr(self, 'scrollable_frame'):
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()

        self.entries, self.radio_vars = {}, {}
        self.combo_descricao, self.combo_informacoes = None, None
        
        if hasattr(self, 'output_text'): self.output_text.delete(1.0, tk.END)
        if hasattr(self, 'label_servico'): self.label_servico.config(text="")
        if hasattr(self, 'entry_servico'): 
            self.entry_servico.set("") # Usa .set() para limpar o combobox
            self.entry_servico.focus_set()
        
        self.servico_id_atual = None # <-- ADICIONE ESTA LINHA
        logger.info("Campos limpos.")
        
    def atualizar_relogio(self):
        """Pega a hora atual e atualiza o label do relógio a cada segundo."""
        hora_atual = datetime.now().strftime('%H:%M:%S')
        self.label_relogio.config(text=hora_atual)
        self.root.after(1000, self.atualizar_relogio)

    # === MÉTODOS DE CONFIGURAÇÃO E HISTÓRICO ===

    def visualizar_detalhes_historico(self, categoria):
        tree = getattr(self, f"tree_{categoria.lower()}", None)
        if not tree or not tree.selection():
            return
        
        selected_item_id = tree.selection()[0]
        
        registro = self.historico_tree_map.get(selected_item_id)
        
        if registro:
            detalhes_window = tk.Toplevel(self.root)
            detalhes_window.title("Detalhes do Registro")
            detalhes_window.geometry("600x400")
            text_widget = scrolledtext.ScrolledText(detalhes_window, wrap=tk.WORD, font=self.fonte_atual)
            text_widget.pack(fill="both", expand=True, padx=10, pady=10)
            text_widget.insert(1.0, registro.get("texto_completo", ""))
            text_widget.config(state="disabled")
            ttk.Button(detalhes_window, text="Fechar", command=detalhes_window.destroy).pack(pady=5)

    def exportar_historico(self, tipo):
        try:
            from tkinter import filedialog
            
            ext = f".{tipo.lower()}"
            filename = filedialog.asksaveasfilename(
                defaultextension=ext, 
                filetypes=[(f"{tipo.upper()} files", f"*{ext}"), ("All files", "*.*")]
            )
            if not filename:
                return

            if tipo == 'txt':
                with open(filename, 'w', encoding='utf-8') as file:
                    file.write(f"=== HISTÓRICO DE ATENDIMENTOS ===\nExportado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n")
                    for categoria, registros in self.history_manager.historico_registros.items():
                        if registros:
                            file.write(f"--- {categoria.upper()} ---\n\n")
                            for i, reg in enumerate(registros, 1):
                                # Usando file.write() diretamente
                                file.write(f"Registro #{i}\n" + "-"*20 + f"\n{reg.get('texto_completo', 'N/A')}\n" + "-"*20 + "\n\n")

            messagebox.showinfo("Sucesso", f"Histórico exportado com sucesso para {filename}!")
            logger.info(f"Histórico exportado ({tipo}) para: {filename}")
                
        except Exception as e:
            logger.error(f"Erro ao exportar histórico ({tipo}): {e}")
            messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")
        
    def exportar_historico_txt(self):
        self.exportar_historico('txt')

    def atualizar_historico_tree(self, tree, categoria):
        for item in tree.get_children():
            tree.delete(item)
        
        registros = self.history_manager.historico_registros.get(categoria, [])

        for reg in registros:
            iid = tree.insert("", "end", values=(
                reg.get("data", ""), reg.get("servico", ""), reg.get("nome", ""), 
                reg.get("protocolo", ""), reg.get("atendente", "")
            ))
            self.historico_tree_map[iid] = reg

    def reset_historico(self):
        if not messagebox.askyesno("Confirmar Reset", 
                                      "⚠️ ATENÇÃO: Esta ação irá APAGAR TODO o histórico!\n\n"
                                      "Esta ação não pode ser desfeita.\n\nDeseja continuar?"):
            return
        if not messagebox.askyesno("Confirmação Final", 
                                      "🚨 ÚLTIMA CHANCE!\n\n"
                                      "Você tem certeza ABSOLUTA que deseja APAGAR PERMANENTEMENTE todo o histórico?"):
            return
        
        self.history_manager.reset_historico()

        for categoria in self.history_manager.historico_registros.keys():
            tree = getattr(self, f"tree_{categoria.lower()}", None)
            if tree:
                self.atualizar_historico_tree(tree, categoria)
        messagebox.showinfo("Reset Concluído", "✅ Histórico resetado com sucesso!")

    def criar_config_aparencia(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        tema_frame = ttk.LabelFrame(main_frame, text="Tema")
        tema_frame.pack(fill="x", pady=(0, 20))
        self.tema_var_config = tk.StringVar(value=self.tema_atual)
        ttk.Radiobutton(tema_frame, text="Claro", variable=self.tema_var_config, value="claro").pack(side="left", padx=10, pady=5)
        ttk.Radiobutton(tema_frame, text="Escuro", variable=self.tema_var_config, value="escuro").pack(side="left", padx=10, pady=5)
        
        cores_frame = ttk.LabelFrame(main_frame, text="Paleta de Cores")
        cores_frame.pack(fill="x", pady=(0, 20))
        
        # Lista final e completa de todos os temas
        cores = [
            "azul_padrao", "azul", "verde", "vermelho", "amarelo", "roxo", "rosa",
            "verde escuro", "laranja", "cinza", "branco", "preto", "mistura",
            "dourado", "azul escuro", "violeta", "marrom", "turquesa",
            "oceano profundo", "floresta", "café", "vibrante", "cyberpunk",
            "neon (ciano)", "neon (rosa)", "neon (verde)"
        ]
        cores.sort() # Organiza a lista em ordem alfabética
        
        self.cor_var_config = tk.StringVar(value=self.cor_atual)
        ttk.Combobox(cores_frame, textvariable=self.cor_var_config, values=cores, state="readonly", width=25).pack(pady=10)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side="bottom", pady=20)
        ttk.Button(btn_frame, text="Aplicar", command=self.aplicar_configuracoes, style="Destaque.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=self.fechar_configuracoes).pack(side="left")

    def criar_config_acessibilidade(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        fonte_frame = ttk.LabelFrame(main_frame, text="Tamanho da Fonte")
        fonte_frame.pack(fill="x", pady=(0, 20))
        self.fonte_var_config = tk.IntVar(value=self.tamanho_fonte_base)
        
        def update_label(v):
            self.label_preview_fonte.config(text=f"{int(float(v))}pt")

        fonte_scale = ttk.Scale(fonte_frame, from_=8, to=36, variable=self.fonte_var_config, 
                                orient="horizontal", length=200, command=update_label)
        fonte_scale.pack(pady=10, padx=10, side="left")
        self.label_preview_fonte = ttk.Label(fonte_frame, text=f"{self.tamanho_fonte_base}pt")
        self.label_preview_fonte.pack(side="left", padx=10)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side="bottom", pady=20)
        ttk.Button(btn_frame, text="Aplicar", command=self.aplicar_configuracoes, style="Destaque.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=self.fechar_configuracoes).pack(side="left")

    def aplicar_configuracoes(self):
        # Aplica configurações de aparência e acessibilidade
        self.tema_atual = self.tema_var_config.get()
        self.cor_atual = self.cor_var_config.get()
        self.tamanho_fonte_base = self.fonte_var_config.get()
        
        # Atualiza a lista de favoritos a partir dos comboboxes da janela de config
        if hasattr(self, 'comboboxes_favoritos'):
            novos_favoritos = []
            for combo in self.comboboxes_favoritos:
                nome_servico = combo.get()
                if nome_servico == "(Nenhum)":
                    novos_favoritos.append(None)
                else:
                    service_id = self.servico_nome_para_id.get(nome_servico)
                    novos_favoritos.append(service_id)
            self.favoritos = novos_favoritos
        
        self.salvar_configuracoes()
        self.aplicar_tamanho_fonte()
        self._atualizar_botoes_favoritos() # Atualiza os botões na tela principal
        
        if hasattr(self, 'label_fonte'):
            self.label_fonte.config(text=f"{self.tamanho_fonte_base}pt")
            
        self.fechar_configuracoes()

    def fechar_configuracoes(self):
        if hasattr(self, 'config_window') and self.config_window.winfo_exists():
            self.config_window.destroy()
            
    # --- INÍCIO DO BLOCO DE CÓDIGO NOVO ---

    def criar_aba_tabulacao(self):
        """Cria a interface da aba de Lembrete de Tabulação."""
        main_frame = ttk.Frame(self.aba_tabulacao, padding=20)
        main_frame.pack(fill="both", expand=True)

        style = ttk.Style()
        try:
            cor_destaque = style.lookup("Accent.TButton", "background")
        except tk.TclError:
            cor_destaque = "red"

        label_aviso = ttk.Label(
            main_frame,
            text="LEMBRETE DE TABULAÇÃO",
            font=("Segoe UI", 26, "bold"),  
            foreground="red"                 
        )
        label_aviso.pack(pady=20)

        label_instrucao = ttk.Label(
            main_frame,
            text="Agora que já sabe qual é o serviço que o cliente deseja, faça a tabulação no GENESYS.",
            font=("Segoe UI", 14),
            wraplength=500,
            justify="center"
        )
        label_instrucao.pack(pady=10)

        btn_ok = ttk.Button(
            main_frame,
            text="✅ OK, Ciente!",
            style="Accent.TButton",
            command=lambda: self.parar_piscar_generico(self.aba_tabulacao, "Tabulação")
        )
        btn_ok.pack(pady=40, ipady=10)

    def _filtrar_codigos(self, event=None):
        """Função chamada a cada tecla digitada na barra de busca."""
        texto_filtro = self.entry_busca_tabulacao.get()
        self.popular_arvore_tabulacao(texto_filtro)

    def limpar_tabulacao(self):
        """Limpa a seleção e o campo de busca."""
        if self.tree_tabulacao.selection():
            self.tree_tabulacao.selection_remove(self.tree_tabulacao.selection()[0])
        self.entry_busca_tabulacao.delete(0, "end")
        self.popular_arvore_tabulacao() # Repopula a árvore com todos os itens
        
    def _toggle_campos_terceiros(self):
        """Mostra ou esconde os campos de dados de terceiros baseado na seleção do RadioButton."""
        if self.autoriza_terceiros_var.get() == "SIM":
            self.frame_dados_terceiros.pack(fill="x", padx=5, pady=5)
        else:
            self.frame_dados_terceiros.pack_forget()

    def criar_form_reclamacao_unificado(self):
        """Cria o formulário unificado para todas as reclamações."""
        # --- a) DESCRIÇÃO DA RECLAMAÇÃO ---
        self.adicionar_secao("a) Descrição da Reclamação")
        self.desc_reclamacao_text = scrolledtext.ScrolledText(self.scrollable_frame, height=5, font=("Segoe UI", 11))
        self.desc_reclamacao_text.pack(fill="x", padx=5, pady=2)

        # --- b) SOLUÇÃO PRETENDIDA ---
        self.adicionar_secao("b) Solução Pretendida pelo Cliente")
        self.adicionar_campo("SOLUCAO_PRETENDIDA")

        # --- c) ANÁLISE DO ATENDENTE ---
        self.adicionar_secao("c) Análise do Atendente")
        self.analise_atendente_text = scrolledtext.ScrolledText(self.scrollable_frame, height=5, font=("Segoe UI", 11))
        self.analise_atendente_text.pack(fill="x", padx=5, pady=2)

        # --- Contato e Permissões ---
        self.adicionar_secao("Dados de Contato e Permissões")
        
        # d) MEIO DE RESPOSTA
        self.meio_resposta_var = tk.StringVar(value=None)
        self.adicionar_radio_buttons("MEIO DE RESPOSTA", [("Telefone", "TELEFONE"), ("Carta", "CARTA"), ("E-mail", "EMAIL")], var=self.meio_resposta_var)
        self.adicionar_campo("CONTATO_PARA_RESPOSTA")

        # e) ACEITA WHATSAPP
        self.adicionar_radio_buttons("ACEITA RESPOSTA/FATURA VIA WHATSAPP", [("Sim", "SIM"), ("Não", "NAO")])

        # f) TELEFONE PARA CONTATO
        self.adicionar_campo("TELEFONE_CONTATO")

        # g) MELHOR HORÁRIO
        self.adicionar_radio_buttons("MELHOR HORÁRIO PARA CONTATO", [("Manhã", "MANHA"), ("Tarde", "TARDE")])

        # h) E-MAIL
        self.adicionar_campo("EMAIL_CONTATO")

        # i) AUTORIZA TERCEIROS (com campos dinâmicos)
        self.autoriza_terceiros_var = tk.StringVar(value="NÃO")
        self.adicionar_radio_buttons("AUTORIZA TERCEIROS A RECEBER A RESPOSTA", [("Sim", "SIM"), ("Não", "NAO")], var=self.autoriza_terceiros_var, command=self._toggle_campos_terceiros)

        # Frame que aparece/desaparece
        self.frame_dados_terceiros = ttk.Frame(self.scrollable_frame)
        self.adicionar_campo("NOME_E_VINCULO_TERCEIRO", parent=self.frame_dados_terceiros)
        self.adicionar_campo("CONTATO_TERCEIRO", parent=self.frame_dados_terceiros)
        
        # --- j) INFORMAÇÕES COMPLEMENTARES ---
        self.adicionar_secao("j) Informações Complementares")
        self.info_comp_text = scrolledtext.ScrolledText(self.scrollable_frame, height=4, font=("Segoe UI", 11))
        self.info_comp_text.pack(fill="x", padx=5, pady=2)

    def processar_reclamacao_unificada(self):
        """Coleta os dados do formulário unificado e formata o texto final."""
        texto = self.get_dados_basicos()
        texto += "\n"

        # a) Descrição
        desc = self.desc_reclamacao_text.get("1.0", tk.END).strip()
        texto += f"a) DESCRIÇÃO DA RECLAMAÇÃO: {desc}\n"

        # b) Solução
        solucao = self._get_entry_value("SOLUCAO_PRETENDIDA")
        texto += f"b) SOLUÇÃO PRETENDIDA: {solucao}\n"

        # c) Análise
        analise = self.analise_atendente_text.get("1.0", tk.END).strip()
        texto += f"c) ANÁLISE DO ATENDENTE: {analise}\n"

        # d) Meio de Resposta
        meio_resp = self.radio_vars.get("MEIO DE RESPOSTA", tk.StringVar()).get() or "NÃO SELECIONADO"
        contato_resp = self._get_entry_value("CONTATO_PARA_RESPOSTA")
        texto += f"d) MEIO DE RESPOSTA DA RECLAMAÇÃO: {meio_resp} - Contato: {contato_resp}\n"
        
        # e) WhatsApp
        aceita_wpp = self.radio_vars.get("ACEITA RESPOSTA/FATURA VIA WHATSAPP", tk.StringVar()).get() or "NÃO SELECIONADO"
        texto += f"e) ACEITA RECEBER RESPOSTA / FATURA VIA WHATSAPP: {aceita_wpp}\n"

        # f) Telefone
        tel_contato = self._get_entry_value("TELEFONE_CONTATO")
        texto += f"f) TELEFONE PARA CONTATO: {tel_contato}\n"

        # g) Horário
        horario = self.radio_vars.get("MELHOR HORÁRIO PARA CONTATO", tk.StringVar()).get() or "NÃO SELECIONADO"
        texto += f"g) MELHOR HORÁRIO PARA CONTATO: {horario}\n"
        
        # h) E-mail
        email = self._get_entry_value("EMAIL_CONTATO") or "não informado"
        texto += f"h) E-MAIL: {email}\n"

        # i) Terceiros
        autoriza = self.radio_vars.get("AUTORIZA TERCEIROS A RECEBER A RESPOSTA", tk.StringVar()).get() or "NÃO"
        texto += f"i) AUTORIZA TERCEIROS: {autoriza}\n"
        if autoriza == "SIM":
            nome_vinc = self._get_entry_value("NOME_E_VINCULO_TERCEIRO")
            cont_terc = self._get_entry_value("CONTATO_TERCEIRO")
            texto += f"   - NOME E VÍNCULO: {nome_vinc}\n"
            texto += f"   - CONTATO: {cont_terc}\n"

        # j) Informações Complementares
        info_comp = self.info_comp_text.get("1.0", tk.END).strip()
        texto += f"j) INFORMAÇÕES COMPLEMENTARES: {info_comp}\n"

        return texto

def main():
    try:
        root = tk.Tk()
        app = AtendimentoApp(root)
        root.mainloop()
    except Exception as e:
        logger.critical(f"Erro fatal ao iniciar aplicação: {e}", exc_info=True)
        messagebox.showerror("Erro Fatal", f"Um erro crítico ocorreu.\n\nDetalhes: {str(e)}\n\nConsulte 'atendimento_app.log' para mais informações.")

if __name__ == "__main__":
    main()