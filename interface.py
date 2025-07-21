import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sem_parar
import Gerar_Recibo
import threading
import os
import sys

# --- Redirecionamento de Saída para Evitar Erros no Executável ---
# PyInstaller em modo --windowed não tem um console (stdout/stderr),
# o que causa um erro se qualquer parte do código tentar usar 'print'
# ou escrever para a saída padrão.
if getattr(sys, 'frozen', False):  # Verifica se está rodando como um .exe
    class DummyStream:
        def write(self, text):
            pass
        def flush(self):
            pass

    sys.stdout = DummyStream()
    sys.stderr = DummyStream()
# --- Fim do Redirecionamento ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação de Relatórios")
        self.root.geometry("500x400")

        self.sem_parar_path = tk.StringVar()
        self.vickos_path = tk.StringVar()
        self.modelo = None
        self.lista_rotas = None

        # --- Widgets ---
        # Frame para seleção de arquivos
        file_frame = tk.Frame(root, padx=10, pady=10)
        file_frame.pack(fill='x', padx=10, pady=5)

        # Planilha Sem Parar
        tk.Label(file_frame, text="Planilha Sem Parar:").grid(row=0, column=0, sticky='w', pady=2)
        self.sem_parar_entry = tk.Entry(file_frame, textvariable=self.sem_parar_path, width=50)
        self.sem_parar_entry.grid(row=1, column=0, sticky='we')
        tk.Button(file_frame, text="Selecionar", command=lambda: self.select_file(self.sem_parar_path)).grid(row=1, column=1, padx=5)

        # Planilha Vickos
        tk.Label(file_frame, text="Planilha Vickos:").grid(row=2, column=0, sticky='w', pady=(10,2))
        self.vickos_entry = tk.Entry(file_frame, textvariable=self.vickos_path, width=50)
        self.vickos_entry.grid(row=3, column=0, sticky='we')
        tk.Button(file_frame, text="Selecionar", command=lambda: self.select_file(self.vickos_path)).grid(row=3, column=1, padx=5)
        
        file_frame.columnconfigure(0, weight=1)

        # Frame para botões de ação
        action_frame = tk.Frame(root, padx=10, pady=10)
        action_frame.pack(fill='x', padx=10, pady=5)

        self.btn_iniciar = tk.Button(action_frame, text="Iniciar Relatório", command=self.run_report, height=2, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.btn_iniciar.pack(fill='x', pady=5)

        self.btn_gerar_pdf = tk.Button(action_frame, text="Gerar PDF", command=self.generate_pdfs, height=2, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), state='disabled')
        self.btn_gerar_pdf.pack(fill='x', pady=5)

        # Frame de progresso
        progress_frame = tk.Frame(root, padx=10, pady=5)
        progress_frame.pack(fill='x', side='bottom')

        self.status_label = tk.Label(progress_frame, text="Pronto.", bd=1, relief=tk.SUNKEN, anchor='w')
        self.status_label.pack(fill='x')

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate")
        self.progress.pack(fill='x', pady=(5,0))

    def select_file(self, path_var):
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*"))
        )
        if filepath:
            path_var.set(filepath)

    def update_progress(self, value, text):
        """Atualiza a barra de progresso e o rótulo de status a partir de qualquer thread."""
        self.root.after(0, self._update_progress_ui, value, text)

    def _update_progress_ui(self, value, text):
        """Função auxiliar para ser executada na thread principal da GUI."""
        self.progress['value'] = value
        self.status_label.config(text=text)
        self.root.update_idletasks()

    def run_report(self):
        sem_parar_file = self.sem_parar_path.get()
        vickos_file = self.vickos_path.get()

        if not sem_parar_file or not vickos_file:
            messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos de planilha.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Salvar relatório como...",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if not output_path:
            return

        self.btn_iniciar.config(state='disabled')
        self.btn_gerar_pdf.config(state='disabled')
        
        # Rodar em uma thread para não travar a GUI
        threading.Thread(target=self._run_report_thread, args=(sem_parar_file, vickos_file, output_path)).start()

    def _run_report_thread(self, sem_parar_file, vickos_file, output_path):
        try:
            self.modelo, self.lista_rotas = sem_parar.fluxo_principal(
                sem_parar_file, vickos_file, output_path, self.update_progress
            )
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", "Relatório gerado com sucesso!"))
        except Exception as e:
            self.update_progress(0, "Erro ao gerar relatório.")
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}"))
        finally:
            self.root.after(0, self.btn_iniciar.config, {'state': 'normal'})
            if self.modelo is not None:
                self.root.after(0, self.btn_gerar_pdf.config, {'state': 'normal'})

    def show_info_and_quit(self, title, message):
        """Exibe uma caixa de informação e depois fecha a aplicação."""
        messagebox.showinfo(title, message)
        self.root.destroy()

    def generate_pdfs(self):
        if self.modelo is None:
            messagebox.showerror("Erro", "Você precisa gerar o relatório primeiro.")
            return
        
        output_dir = filedialog.askdirectory(title="Selecione a pasta para salvar os PDFs")
        if not output_dir:
            return

        self.btn_iniciar.config(state='disabled')
        self.btn_gerar_pdf.config(state='disabled')
        
        threading.Thread(target=self._generate_pdfs_thread, args=(output_dir,)).start()

    def _generate_pdfs_thread(self, output_dir):
        try:
            Gerar_Recibo.gerar_recibos(self.modelo, self.lista_rotas, output_dir, self.update_progress)
            self.update_progress(100, f"PDFs salvos com sucesso!")
            success_message = f"PDFs gerados com sucesso na pasta:\n{output_dir}"
            # Exibe a mensagem de sucesso e fecha a aplicação.
            self.root.after(0, self.show_info_and_quit, "Sucesso", success_message)
        except Exception as e:
            self.update_progress(0, "Erro ao gerar PDFs.")
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}"))
            # Reativa os botões em caso de erro para permitir nova tentativa.
            self.root.after(0, self.btn_iniciar.config, {'state': 'normal'})
            self.root.after(0, self.btn_gerar_pdf.config, {'state': 'normal'})


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop() 