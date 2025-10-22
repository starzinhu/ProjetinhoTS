import customtkinter
import threading
from analise_completa import executar_analise_completa, exportar_para_excel_existente
from tkinter import filedialog

class App(customtkinter.CTk):
    LOCAL_ID_MAP = {"AL": "10453", "BA": "10426", "CE": "10542", "SE": "10465", "DF": "10346", "MT": "10619"}

    def __init__(self):
        super().__init__()

        self.title("Analisador de Aulas e-Pratika")
        self.geometry("800x650") # Aumentei a altura para caber o novo botão

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # Ajustado o configure da linha

        # --- Frame de Entradas ---
        input_frame = customtkinter.CTkFrame(self)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        # URL
        self.url_label = customtkinter.CTkLabel(input_frame, text="URL da Consulta:")
        self.url_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.url_entry = customtkinter.CTkEntry(input_frame, placeholder_text="Cole o link completo da consulta do relatório")
        self.url_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Usuário
        self.user_label = customtkinter.CTkLabel(input_frame, text="Usuário:")
        self.user_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.user_entry = customtkinter.CTkEntry(input_frame, placeholder_text="seu.usuario@criar")
        self.user_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.user_entry.insert(0, "gabriel.secco@criar") # Valor padrão

        # Senha
        self.password_label = customtkinter.CTkLabel(input_frame, text="Senha:")
        self.password_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.password_entry = customtkinter.CTkEntry(input_frame, show="*", placeholder_text="Sua senha")
        self.password_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.password_entry.insert(0, "aa123456") # Valor padrão

        # Local ID (ComboBox)
        self.local_id_label = customtkinter.CTkLabel(input_frame, text="Local:")
        self.local_id_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.local_id_options = list(self.LOCAL_ID_MAP.keys())
        self.local_id_combobox = customtkinter.CTkComboBox(input_frame, values=self.local_id_options)
        self.local_id_combobox.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.local_id_combobox.set("CE") # Valor padrão

        # --- Frame de Filtros ---
        filter_frame = customtkinter.CTkFrame(self)
        filter_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        filter_frame.grid_columnconfigure(1, weight=1)
        filter_frame.grid_columnconfigure(3, weight=1)

        # Nome
        self.name_label = customtkinter.CTkLabel(filter_frame, text="Nome:")
        self.name_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.name_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="Nome do Aluno")
        self.name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # RENACH
        self.renach_label = customtkinter.CTkLabel(filter_frame, text="RENACH:")
        self.renach_label.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.renach_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="RENACH do Aluno")
        self.renach_entry.grid(row=0, column=3, padx=10, pady=5, sticky="ew")

        # Instrutor
        self.instructor_label = customtkinter.CTkLabel(filter_frame, text="Instrutor:")
        self.instructor_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.instructor_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="Nome do Instrutor")
        self.instructor_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Tablet
        self.tablet_label = customtkinter.CTkLabel(filter_frame, text="Tablet:")
        self.tablet_label.grid(row=1, column=2, padx=10, pady=5, sticky="w")
        self.tablet_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="ID do Tablet")
        self.tablet_entry.grid(row=1, column=3, padx=10, pady=5, sticky="ew")

        # Veículo
        self.vehicle_label = customtkinter.CTkLabel(filter_frame, text="Veículo:")
        self.vehicle_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.vehicle_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="Placa do Veículo")
        self.vehicle_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Data Início
        self.start_date_label = customtkinter.CTkLabel(filter_frame, text="Data Início (DD/MM/AA):")
        self.start_date_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.start_date_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="Ex: 01/01/23")
        self.start_date_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # Data Fim
        self.end_date_label = customtkinter.CTkLabel(filter_frame, text="Data Fim (DD/MM/AA):")
        self.end_date_label.grid(row=3, column=2, padx=10, pady=5, sticky="w")
        self.end_date_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="Ex: 31/12/23")
        self.end_date_entry.grid(row=3, column=3, padx=10, pady=5, sticky="ew")

        # --- Frame de Saída ---
        output_frame = customtkinter.CTkFrame(self)
        output_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        output_frame.grid_rowconfigure(0, weight=1)
        output_frame.grid_columnconfigure(0, weight=1)

        self.log_textbox = customtkinter.CTkTextbox(output_frame, state="disabled", font=("Consolas", 12))
        self.log_textbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # --- Botões de Ação ---
        action_frame = customtkinter.CTkFrame(self)
        action_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        action_frame.grid_columnconfigure((0, 1), weight=1) # Configura as colunas para terem o mesmo peso

        self.run_button = customtkinter.CTkButton(action_frame, text="Iniciar Análise", command=self.start_analysis_thread)
        self.run_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.export_button = customtkinter.CTkButton(action_frame, text="Exportar para Excel", command=self.export_to_excel)
        self.export_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    def log(self, message):
        self.after(0, self._log_update, message)

    def _log_update(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message)
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")

    def start_analysis_thread(self):
        self.run_button.configure(state="disabled", text="Analisando...")
        self.export_button.configure(state="disabled")
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

        thread = threading.Thread(target=self.run_analysis_in_background)
        thread.start()

    def run_analysis_in_background(self):
        url = self.url_entry.get()
        user = self.user_entry.get()
        password = self.password_entry.get()
        selected_local_name = self.local_id_combobox.get()
        local_id = self.LOCAL_ID_MAP.get(selected_local_name, "10542")

        filters = {
            'nome': self.name_entry.get(),
            'renach': self.renach_entry.get(),
            'instrutor': self.instructor_entry.get(),
            'tablet': self.tablet_entry.get(),
            'veiculo': self.vehicle_entry.get(),
            'data_inicio': self.start_date_entry.get(),
            'data_fim': self.end_date_entry.get()
        }

        try:
            executar_analise_completa(url, user, password, local_id, self.log, filters)
        except Exception as e:
            self.log(f"Ocorreu um erro inesperado na thread: {e}\n")
        finally:
            self.run_button.configure(state="normal", text="Iniciar Análise")
            self.export_button.configure(state="normal")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All files", "*.*")],
            title="Salvar arquivo Excel",
            initialfile="resultado_analise.xlsx"
        )
        if not file_path:
            self.log("Exportação cancelada pelo usuário.\n")
            return

        self.export_button.configure(state="disabled", text="Exportando...")
        self.run_button.configure(state="disabled")

        thread = threading.Thread(target=self.run_export_in_background, args=(file_path,))
        thread.start()

    def run_export_in_background(self, file_path):
        try:
            exportar_para_excel_existente(file_path, self.log)
        except Exception as e:
            self.log(f"Ocorreu um erro inesperado durante a exportação: {e}\n")
        finally:
            self.export_button.configure(state="normal", text="Exportar para Excel")
            self.run_button.configure(state="normal", text="Iniciar Análise")

if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("blue")
    app = App()
    app.mainloop()