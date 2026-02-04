import customtkinter
import tkinter
from tkinter import filedialog
import threading
import time
import os
import re
import shutil  # <--- NOVA BIBLIOTECA PARA MOVER ARQUIVOS
from datetime import datetime
import pyautogui
from openpyxl import Workbook, load_workbook

# --- CONFIGURAÇÕES ---
MODO_SIMULACAO = True  # Mude para False para usar o mouse/teclado real
pyautogui.FAILSAFE = True  
pyautogui.PAUSE = 0.6      

customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        titulo = "Automação Salus - MODO SIMULAÇÃO" if MODO_SIMULACAO else "Automação Salus - PROD"
        self.title(titulo)
        self.geometry("650x700")
        self.grid_columnconfigure(0, weight=1)

        self.txt_log_path = None
        self.excel_log_path = None
        self.stop_event = threading.Event()

        # Layout
        self.label_title = customtkinter.CTkLabel(self, text="Painel de Automação", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.label_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        if MODO_SIMULACAO:
            self.label_aviso = customtkinter.CTkLabel(self, text="⚠ MODO TESTE (SEM CLIQUES) ⚠", text_color="orange")
            self.label_aviso.grid(row=1, column=0, sticky="s")

        self.frame_config = customtkinter.CTkFrame(self)
        self.frame_config.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.frame_config.grid_columnconfigure(1, weight=1)

        self.btn_folder = customtkinter.CTkButton(self.frame_config, text="1. Selecionar Pasta", command=self.select_folder)
        self.btn_folder.grid(row=0, column=0, padx=10, pady=10)
        self.entry_folder = customtkinter.CTkEntry(self.frame_config, placeholder_text="Caminho da pasta...")
        self.entry_folder.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.label_operador = customtkinter.CTkLabel(self.frame_config, text="Nome Operador:")
        self.label_operador.grid(row=1, column=0, padx=10, pady=10)
        self.entry_operador = customtkinter.CTkEntry(self.frame_config, placeholder_text="Digite seu nome")
        self.entry_operador.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.label_system = customtkinter.CTkLabel(self.frame_config, text="Sistema:")
        self.label_system.grid(row=2, column=0, padx=10, pady=10)
        self.radio_var = tkinter.IntVar(value=0)
        self.radio_biocroma = customtkinter.CTkRadioButton(self.frame_config, text="BIOCROMA", variable=self.radio_var, value=0)
        self.radio_biocroma.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        self.radio_biovida = customtkinter.CTkRadioButton(self.frame_config, text="BIOVIDA", variable=self.radio_var, value=1)
        self.radio_biovida.grid(row=2, column=2, padx=10, pady=10, sticky="w")

        self.textbox_log = customtkinter.CTkTextbox(self, width=250, height=200)
        self.textbox_log.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        self.textbox_log.insert("0.0", "Aguardando início...\n")
        self.textbox_log.configure(state="disabled")

        self.progressbar = customtkinter.CTkProgressBar(self, orientation="horizontal")
        self.progressbar.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        self.progressbar.set(0)

        self.btn_start = customtkinter.CTkButton(self, text="INICIAR", command=self.start_process_thread, fg_color="green", height=50)
        self.btn_start.grid(row=5, column=0, padx=20, pady=20)

    # --- LOGS ---
    def log_technical(self, message):
        if self.txt_log_path:
            timestamp = datetime.now().strftime("%H:%M:%S")
            try:
                with open(self.txt_log_path, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] {message}\n")
            except: pass

    def log_visual(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.textbox_log.configure(state="normal")
        self.textbox_log.insert("end", f"[{timestamp}] {message}\n")
        self.textbox_log.see("end")
        self.textbox_log.configure(state="disabled")
        self.log_technical(message)

    def log_excel(self, id_arquivo, nome_arquivo, status, motivo=""):
        if not self.excel_log_path: return
        try:
            try:
                wb = load_workbook(self.excel_log_path)
                ws = wb.active
            except FileNotFoundError:
                wb = Workbook()
                ws = wb.active
                ws.title = "Relatorio"
                ws.append(["Data", "Hora", "Operador", "ID", "Arquivo", "Status", "Obs"])
            
            data = datetime.now().strftime("%d/%m/%Y")
            hora = datetime.now().strftime("%H:%M:%S")
            operador = self.entry_operador.get() or "Desconhecido"
            ws.append([data, hora, operador, id_arquivo, nome_arquivo, status, motivo])
            wb.save(self.excel_log_path)
        except PermissionError:
            self.log_visual("ERRO: Feche o Excel para salvar!")

    # --- CORE ---
    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_folder.delete(0, "end")
            self.entry_folder.insert(0, folder)
            self.log_visual("Pasta selecionada.")

    def extract_id(self, filename, sistema):
        pattern = r'-(\d{5})' if sistema == "BIOVIDA" else r'(\d+)'
        match = re.search(pattern, filename)
        return match.group(1) if match else None

    def find_and_click(self, image_name, confidence=0.85, wait=5):
        if MODO_SIMULACAO:
            time.sleep(0.1)
            return True

        path = os.path.join("assets", image_name)
        if not os.path.exists(path):
            self.log_technical(f"IMG OFF: {image_name}")
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
        
        self.log_technical(f"NÃO ENCONTRADO: {image_name}")
        return False

    def automation_step_by_step(self, id_alvo, arquivo_caminho):
        if MODO_SIMULACAO:
            self.log_technical(f"Simulando ID {id_alvo}...")
            time.sleep(1)
            return True, "Simulado"

        # 1. Sequência
        if not self.find_and_click("1_btn_pesquisar.png"): return False, "Erro Btn Pesquisar"
        if not self.find_and_click("2_btn_alternar.png"): return False, "Erro Btn Alternar"
        if not self.find_and_click("3_btn_comecar.png"): return False, "Erro Btn Começar"
        if not self.find_and_click("campo_busca.png"): return False, "Erro Campo Busca"
        
        time.sleep(0.3)
        pyautogui.write(id_alvo)
        time.sleep(0.3)
        
        if not self.find_and_click("4_btn_lupa.png"): return False, "Erro Btn Lupa"
        time.sleep(2.0) 
        
        # 2. Verificação
        pos_header = pyautogui.locateCenterOnScreen(os.path.join("assets", "cabecalho_id.png"), confidence=0.8)
        if not pos_header: return False, "Erro Cabeçalho ID"
            
        pyautogui.click(pos_header) # Ordena
        time.sleep(1.0)
        
        roi = (int(pos_header.x), int(pos_header.y + 10), 800, 50)
        cancelado = pyautogui.locateOnScreen(os.path.join("assets", "status_cance_normal.png"), region=roi, confidence=0.8)
        
        if cancelado:
            self.log_technical(f"ID {id_alvo}: Linha 1 Cancelada. Indo para Linha 2.")
            pyautogui.click(pos_header.x, pos_header.y + 45) 
        else:
            pyautogui.click(pos_header.x, pos_header.y + 25) 
            
        time.sleep(0.5)

        # 3. Anexar
        if not self.find_and_click("6_btn_anexar.png"): return False, "Erro Btn Anexar"
        
        time.sleep(1.5)
        pyautogui.write(arquivo_caminho) 
        time.sleep(0.5)
        pyautogui.press('enter') 
        
        return True, "Sucesso"

    def start_process_thread(self):
        self.stop_event.clear()
        threading.Thread(target=self.run_automation).start()

    def run_automation(self):
        pasta = self.entry_folder.get()
        operador = self.entry_operador.get()

        if not pasta or not operador: 
            return self.log_visual("ERRO: Preencha pasta e operador.")

        # Cria pasta "Processados" se não existir
        pasta_processados = os.path.join(pasta, "Processados")
        if not os.path.exists(pasta_processados):
            os.makedirs(pasta_processados)
            self.log_visual("Pasta 'Processados' criada.")

        self.txt_log_path = os.path.join(pasta, f"LOG_TEC_{datetime.now().strftime('%Y-%m-%d')}.txt")
        self.excel_log_path = os.path.join(pasta, "RELATORIO_GERENCIAL.xlsx")

        self.btn_start.configure(state="disabled", text="PROCESSANDO...")
        self.log_technical(f"=== INÍCIO: {operador} ===")

        sistema = "BIOCROMA" if self.radio_var.get() == 0 else "BIOVIDA"
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
        
        self.log_visual(f"Arquivos: {len(arquivos)}")
        if not MODO_SIMULACAO: time.sleep(3)

        for i, arquivo in enumerate(arquivos):
            if self.stop_event.is_set(): break
            
            path_full = os.path.join(pasta, arquivo)
            id_val = self.extract_id(arquivo, sistema)

            if id_val:
                sucesso, motivo = self.automation_step_by_step(id_val, path_full)
                status = "SUCESSO" if sucesso else "ERRO"
                
                self.log_excel(id_val, arquivo, status, motivo)
                self.log_visual(f"ID {id_val}: {status}")
                
                # SE DEU SUCESSO, MOVE O ARQUIVO
                if sucesso:
                    try:
                        destino = os.path.join(pasta_processados, arquivo)
                        shutil.move(path_full, destino)
                        self.log_technical(f"Arquivo movido para: {destino}")
                    except Exception as e:
                        self.log_visual(f"ERRO AO MOVER: {e}")

                    time.sleep(0.5 if MODO_SIMULACAO else 1)
            else:
                self.log_visual(f"Pulado: {arquivo}")
                self.log_excel("N/A", arquivo, "IGNORADO", "Sem ID")
            
            self.progressbar.set((i + 1) / len(arquivos))

        self.btn_start.configure(state="normal", text="INICIAR")
        self.log_visual("=== FIM ===")
        self.txt_log_path = None
        self.excel_log_path = None

if __name__ == "__main__":
    app = App()
    app.mainloop()