import customtkinter
import tkinter
from tkinter import filedialog, messagebox
import threading
import time
import os
import re
import shutil
from datetime import datetime, timedelta
import pyautogui
from openpyxl import Workbook, load_workbook
import pyperclip

# --- CONFIGURAÇÕES GLOBAIS ---
MODO_SIMULACAO = True  # False = Roda de verdade | True = Apenas finge que clica
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# --- CADASTRO DE OPERADORES ---
LISTA_OPERADORES = [
    "Selecione o Operador", 
    "Ana Silva",
    "Bruno Santos",
    "Carlos Oliveira",
    "Fernanda Souza",
    "Admin/Teste"
]

# Configurações de Tema
customtkinter.set_appearance_mode("Dark") 
customtkinter.set_default_color_theme("blue")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        titulo = "Automação Salus [SIMULAÇÃO]" if MODO_SIMULACAO else "Automação Salus [PROD]"
        self.title(titulo)
        self.geometry("720x880")
        self.grid_columnconfigure(0, weight=1)

        self.txt_log_path = None
        self.excel_log_path = None
        
        # Variáveis de Controle de Thread
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set() # Começa liberado

        self.setup_ui()

    def setup_ui(self):
        # --- 1. HEADER (CABEÇALHO) ---
        self.frame_header = customtkinter.CTkFrame(self, fg_color="transparent")
        self.frame_header.grid(row=0, column=0, pady=(20, 10))
        
        self.label_title = customtkinter.CTkLabel(self.frame_header, text="ROBÔ SALUS", font=("Roboto", 28, "bold"), text_color="#3B8ED0")
        self.label_title.pack()
        
        self.label_subtitle = customtkinter.CTkLabel(self.frame_header, text="Painel de Controle de Automação", font=("Roboto", 12))
        self.label_subtitle.pack()

        if MODO_SIMULACAO:
            self.label_aviso = customtkinter.CTkLabel(self, text="⚠ MODO TESTE ATIVADO ⚠", text_color="#E67E22", font=("Arial", 14, "bold"))
            self.label_aviso.grid(row=1, column=0)

        # --- 2. CONFIGURAÇÕES (CARD CINZA) ---
        self.frame_config = customtkinter.CTkFrame(self, corner_radius=15)
        self.frame_config.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.frame_config.grid_columnconfigure(1, weight=1)

        # Seleção de Pasta
        self.btn_folder = customtkinter.CTkButton(self.frame_config, text="📂 Escolher Pasta", command=self.select_folder, fg_color="#555", hover_color="#444")
        self.btn_folder.grid(row=0, column=0, padx=15, pady=15)
        
        self.entry_folder = customtkinter.CTkEntry(self.frame_config, placeholder_text="Caminho da pasta com os PDFs...")
        self.entry_folder.grid(row=0, column=1, padx=15, pady=15, sticky="ew")

        # Operador
        self.label_op = customtkinter.CTkLabel(self.frame_config, text="Operador:", font=("Arial", 12, "bold"))
        self.label_op.grid(row=1, column=0, padx=15, pady=(0, 15))
        
        self.menu_operador = customtkinter.CTkOptionMenu(self.frame_config, values=LISTA_OPERADORES, fg_color="#2B2B2B", button_color="#3B8ED0")
        self.menu_operador.grid(row=1, column=1, padx=15, pady=(0, 15), sticky="w")

        # Seleção de Sistema (Radio Buttons)
        self.frame_radios = customtkinter.CTkFrame(self.frame_config, fg_color="transparent")
        self.frame_radios.grid(row=2, column=0, columnspan=2, pady=(0, 15))
        
        self.radio_var = tkinter.IntVar(value=0)
        self.radio_biocroma = customtkinter.CTkRadioButton(self.frame_radios, text="BIOCROMA", variable=self.radio_var, value=0)
        self.radio_biocroma.pack(side="left", padx=20)
        self.radio_biovida = customtkinter.CTkRadioButton(self.frame_radios, text="BIOVIDA", variable=self.radio_var, value=1)
        self.radio_biovida.pack(side="left", padx=20)

        # --- 3. ÁREA DE LOGS E STATUS ---
        self.textbox_log = customtkinter.CTkTextbox(self, height=180, font=("Consolas", 11), corner_radius=10, fg_color="#1E1E1E")
        self.textbox_log.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        self.textbox_log.insert("0.0", ">> Sistema pronto. Selecione a pasta e o operador.\n")
        self.textbox_log.configure(state="disabled")

        # Painel de Status (Texto Grande e ETA)
        self.frame_status = customtkinter.CTkFrame(self, fg_color="transparent")
        self.frame_status.grid(row=4, column=0, padx=30, pady=(5, 0), sticky="ew")
        
        self.label_status = customtkinter.CTkLabel(self.frame_status, text="AGUARDANDO", font=("Arial", 16, "bold"), text_color="gray")
        self.label_status.pack(side="left")
        
        self.label_eta = customtkinter.CTkLabel(self.frame_status, text="--:--:--", font=("Arial", 14))
        self.label_eta.pack(side="right")

        # Barra de Progresso
        self.progressbar = customtkinter.CTkProgressBar(self, orientation="horizontal", height=15, progress_color="#2CC985")
        self.progressbar.grid(row=5, column=0, padx=20, pady=(5, 20), sticky="ew")
        self.progressbar.set(0)

        # --- 4. BOTÕES DE CONTROLE (GRANDE DESTAQUE) ---
        self.frame_controls = customtkinter.CTkFrame(self, fg_color="transparent")
        self.frame_controls.grid(row=6, column=0, pady=10, sticky="ew")
        self.frame_controls.grid_columnconfigure(0, weight=1)
        self.frame_controls.grid_columnconfigure(1, weight=1)

        # Botão INICIAR (Verde Gigante)
        self.btn_start = customtkinter.CTkButton(
            self.frame_controls, 
            text="INICIAR PROCESSAMENTO", 
            command=self.start_process_thread, 
            fg_color="#27AE60",       # Verde
            hover_color="#2ECC71",
            width=400, 
            height=50,
            font=("Arial", 16, "bold"),
            corner_radius=25
        )
        self.btn_start.grid(row=0, column=0, columnspan=2, pady=(0, 15))

        # Botão PAUSAR (Laranja)
        self.btn_pause = customtkinter.CTkButton(
            self.frame_controls, 
            text="⏸ PAUSAR", 
            command=self.toggle_pause, 
            fg_color="#F39C12",       # Laranja
            hover_color="#F1C40F",
            width=180, 
            height=40,
            state="disabled",
            font=("Arial", 12, "bold")
        )
        self.btn_pause.grid(row=1, column=0, padx=10, sticky="e")

        # Botão PARAR (Vermelho)
        self.btn_stop = customtkinter.CTkButton(
            self.frame_controls, 
            text="⏹ PARAR TUDO", 
            command=self.stop_process, 
            fg_color="#C0392B",       # Vermelho
            hover_color="#E74C3C",
            width=180, 
            height=40,
            state="disabled",
            font=("Arial", 12, "bold")
        )
        self.btn_stop.grid(row=1, column=1, padx=10, sticky="w")
        
        # Rodapé
        self.label_footer = customtkinter.CTkLabel(self, text="v3.0 - Automação Segura Salus", text_color="gray60", font=("Arial", 10))
        self.label_footer.grid(row=7, column=0, pady=20)

    # --- FUNÇÕES AUXILIARES (LOGS E ARQUIVOS) ---
    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_folder.delete(0, "end")
            self.entry_folder.insert(0, folder)
            self.log_visual("Pasta definida.")

    def log_visual(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.textbox_log.configure(state="normal")
        self.textbox_log.insert("end", f"[{timestamp}] {message}\n")
        self.textbox_log.see("end")
        self.textbox_log.configure(state="disabled")
        self.log_technical(message)

    def log_technical(self, message):
        if self.txt_log_path:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                with open(self.txt_log_path, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] {message}\n")
            except: pass

    def log_excel(self, id_arquivo, nome_arquivo, status, motivo=""):
        if not self.excel_log_path: return
        try:
            wb = load_workbook(self.excel_log_path)
            ws = wb.active
        except:
            wb = Workbook()
            ws = wb.active
            ws.append(["Data", "Hora", "Operador", "ID", "Arquivo", "Status", "Detalhes"])
        try:
            ws.append([
                datetime.now().strftime("%d/%m/%Y"), 
                datetime.now().strftime("%H:%M:%S"), 
                self.menu_operador.get(), 
                id_arquivo, 
                nome_arquivo, 
                status, 
                motivo
            ])
            wb.save(self.excel_log_path)
        except PermissionError: 
            self.log_visual("⚠ ERRO CRÍTICO: Feche o Excel 'RELATORIO_GERENCIAL' agora!")

    def extract_id(self, filename, sistema):
        pattern = r'-(\d{5,})' if sistema == "BIOVIDA" else r'(\d{5,})'
        match = re.search(pattern, filename)
        return match.group(1) if match else None

    # --- FUNÇÕES CORE DO RPA ---

    def take_screenshot(self, prefix="ERRO"):
        """Tira foto da tela para debug."""
        pasta = self.entry_folder.get()
        if not pasta: return
        pasta_erros = os.path.join(pasta, "erros_print")
        os.makedirs(pasta_erros, exist_ok=True)
        try: 
            pyautogui.screenshot(os.path.join(pasta_erros, f"{prefix}_{datetime.now().strftime('%H-%M-%S')}.png"))
            self.log_technical(f"Screenshot salvo em erros_print: {prefix}")
        except: pass

    def find_and_click(self, image_name, confidence=0.85, wait=7):
        """Busca imagem e clica. Tira print se falhar."""
        if MODO_SIMULACAO:
            time.sleep(0.1)
            return True
        path = os.path.join("assets", image_name)
        if not os.path.exists(path):
            self.log_visual(f"❌ Asset ausente: {image_name}")
            return False
        
        start = time.time()
        while time.time() - start < wait:
            if self.stop_event.is_set(): return False
            try:
                pos = pyautogui.locateCenterOnScreen(path, confidence=confidence)
                if pos:
                    pyautogui.click(pos)
                    return True
            except: pass
            time.sleep(0.5)
        
        self.log_technical(f"Timeout (não encontrou): {image_name}")
        self.take_screenshot(f"NAO_ACHOU_{image_name.replace('.png','')}")
        return False

    def safe_move_file(self, origem, destino_pasta, max_tentativas=10):
        """
        Estratégia Blindada: COPIAR -> DELETAR.
        Resolve problemas de permissão onde o Windows bloqueia o 'move', mas permite o 'copy'.
        """
        if MODO_SIMULACAO: return True
        
        nome_arquivo = os.path.basename(origem)
        destino_final = os.path.join(destino_pasta, nome_arquivo)

        # 1. Renomear se já existe no destino
        if os.path.exists(destino_final):
            base, ext = os.path.splitext(nome_arquivo)
            timestamp = datetime.now().strftime("%H%M%S")
            destino_final = os.path.join(destino_pasta, f"{base}_{timestamp}{ext}")

        # 2. Tentar COPIAR
        copiou = False
        for i in range(1, max_tentativas + 1):
            try:
                shutil.copy2(origem, destino_final) # Copy2 preserva metadados
                copiou = True
                break
            except Exception:
                self.log_technical(f"Tentando copiar... ({i}/{max_tentativas})")
                time.sleep(1.5)
        
        if not copiou:
            self.log_visual("❌ Falha crítica: Não consegui copiar o arquivo para 'Processados'.")
            return False

        # 3. Tentar DELETAR o original
        for i in range(1, 6):
            try:
                os.remove(origem)
                self.log_technical("Arquivo original deletado com sucesso.")
                return True
            except Exception:
                time.sleep(2)
        
        self.log_visual("⚠ ALERTA: Arquivo copiado, mas não consegui deletar o original (Travado).")
        return True # Retorna True pois o processo principal (upload + backup) funcionou

    def automation_step_by_step(self, id_alvo, arquivo_caminho):
        """Fluxo de cliques no Salus."""
        if MODO_SIMULACAO:
            time.sleep(1)
            return True, "Simulado"
        
        # 1. Pesquisa
        if not self.find_and_click("1_btn_pesquisar.png"): return False, "Btn Pesquisar ñ achado"
        time.sleep(0.5)
        
        # Clica no campo (opcional)
        if not self.find_and_click("campo_busca.png", wait=3): pass
        
        # Limpa e Cola ID
        pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace'); time.sleep(0.2)
        pyperclip.copy(id_alvo); pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        
        if not self.find_and_click("4_btn_lupa.png"): return False, "Btn Lupa ñ achado"
        time.sleep(3.0) # Tempo de carregamento da grid
        
        # 2. Selecionar Paciente
        pos_header = pyautogui.locateCenterOnScreen(os.path.join("assets", "cabecalho_id.png"), confidence=0.8)
        if not pos_header:
            self.take_screenshot("ERRO_GRID")
            return False, "Grid não carregou"
            
        pyautogui.click(pos_header); time.sleep(1.0) # Ordena
        
        # Lógica de Cancelado (Linha 1 ou 2)
        roi = (int(pos_header.x - 100), int(pos_header.y + 10), 1000, 60)
        cancelado = pyautogui.locateOnScreen(os.path.join("assets", "status_cance_normal.png"), region=roi, confidence=0.8)
        
        offset_y = 45 if cancelado else 25
        pyautogui.click(pos_header.x, pos_header.y + offset_y)
        time.sleep(0.5)

        # 3. Anexar
        if not self.find_and_click("6_btn_anexar.png"): return False, "Btn Anexar ñ achado"
        time.sleep(2.0) # Janela do Windows abrir
        
        pyperclip.copy(os.path.abspath(arquivo_caminho))
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.8)
        pyautogui.press('enter')
        
        self.log_technical("Aguardando upload do Salus...")
        time.sleep(4.5) # Aumentei para garantir
        
        return True, "Sucesso"

    # --- CONTROLE DE FLUXO ---
    def toggle_pause(self):
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.btn_pause.configure(text="▶ CONTINUAR", fg_color="#27AE60")
            self.label_status.configure(text="PAUSADO", text_color="#F39C12")
            self.log_visual("⏸ Sistema PAUSADO. Termine o café ☕")
        else:
            self.pause_event.set()
            self.btn_pause.configure(text="⏸ PAUSAR", fg_color="#F39C12")
            self.label_status.configure(text="EXECUTANDO", text_color="#27AE60")
            self.log_visual("▶ Sistema RETOMADO.")

    def start_process_thread(self):
        if not self.entry_folder.get():
            messagebox.showwarning("Atenção", "Selecione a pasta primeiro!")
            return
        if self.menu_operador.get() == LISTA_OPERADORES[0]:
            messagebox.showwarning("Atenção", "Selecione o Operador!")
            return

        self.stop_event.clear(); self.pause_event.set()
        
        # Trava UI
        self.btn_start.configure(state="disabled", fg_color="gray")
        self.btn_stop.configure(state="normal")
        self.btn_pause.configure(state="normal", text="⏸ PAUSAR", fg_color="#F39C12")
        self.label_status.configure(text="INICIANDO...", text_color="#27AE60")
        self.menu_operador.configure(state="disabled")
        
        threading.Thread(target=self.run_automation, daemon=True).start()

    def stop_process(self):
        if messagebox.askyesno("Parar", "Deseja parar após finalizar o arquivo atual?"):
            self.stop_event.set(); self.pause_event.set()
            self.label_status.configure(text="PARANDO...", text_color="#C0392B")
            self.btn_stop.configure(state="disabled"); self.btn_pause.configure(state="disabled")

    def run_automation(self):
        pasta = self.entry_folder.get()
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
        
        # Cria pastas
        pasta_proc = os.path.join(pasta, "Processados")
        if not os.path.exists(pasta_proc): os.makedirs(pasta_proc)
        
        self.txt_log_path = os.path.join(pasta, f"LOG_TEC_{datetime.now().strftime('%Y-%m-%d')}.txt")
        self.excel_log_path = os.path.join(pasta, "RELATORIO_GERENCIAL.xlsx")

        total = len(arquivos)
        inicio_geral = time.time()
        self.log_visual(f"Iniciando lote de {total} arquivos.")
        
        sistema = "BIOVIDA" if self.radio_var.get() == 1 else "BIOCROMA"

        for i, arquivo in enumerate(arquivos):
            # Controle de Pausa
            if not self.pause_event.is_set(): self.pause_event.wait()
            
            # Controle de Parada
            if self.stop_event.is_set(): 
                self.log_visual("🛑 Parada solicitada pelo usuário.")
                break
            
            # Cálculo de ETA (Tempo Restante)
            if i > 0:
                tempo_decorrido = time.time() - inicio_geral
                tempo_medio = tempo_decorrido / i
                restantes = total - i
                eta = str(timedelta(seconds=int(tempo_medio * restantes)))
                self.label_eta.configure(text=f"Faltam: {eta}")
            
            self.log_visual(f"[{i+1}/{total}] {arquivo}")
            path_full = os.path.join(pasta, arquivo)
            id_val = self.extract_id(arquivo, sistema)
            
            status, motivo = "ERRO", "ID n/a"

            if id_val:
                sucesso_rpa, msg_rpa = self.automation_step_by_step(id_val, path_full)
                motivo = msg_rpa
                
                if sucesso_rpa:
                    # Nova função de mover blindada
                    if self.safe_move_file(path_full, pasta_proc):
                        status = "SUCESSO"
                    else:
                        status = "SUCESSO (FALHA MOVE)"
                        motivo = "Upload OK, mas não moveu arquivo"
            
            self.log_excel(id_val, arquivo, status, motivo)
            self.progressbar.set((i + 1) / total)
        
        # Finalização
        self.label_status.configure(text="CONCLUÍDO", text_color="#3B8ED0")
        self.label_eta.configure(text="--:--:--")
        self.btn_start.configure(state="normal", fg_color="#27AE60")
        self.btn_stop.configure(state="disabled"); self.btn_pause.configure(state="disabled")
        self.menu_operador.configure(state="normal")
        self.log_visual("=== FIM DO PROCESSO ===")
        messagebox.showinfo("Fim", "Processamento finalizado!")

if __name__ == "__main__":
    app = App()
    app.mainloop()