import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, YES, X, LEFT, RIGHT

from tkinter import filedialog, messagebox
from pathlib import Path
from .utils import (
    identificar_extensao_arquivo,
    carregar_planilha,
    obter_aba_exemplo,
    obter_dados_aba,
    criar_dict_referencia,
    obter_aba_atual,
    preencher_dados
)


def processar_planilha(caminho_arquivo):
    try:
        caminho_arquivo = Path(caminho_arquivo)

        extensao = identificar_extensao_arquivo(caminho_arquivo)
        if extensao not in [".xlsm", ".xlsx"]:
            messagebox.showerror(
                "Erro", f"Extensão '{extensao}' não suportada. Use .xlsm ou .xlsx.")
            return

        wb = carregar_planilha(caminho_arquivo)
        abas_planilhas = wb.sheetnames

        aba_exemplo_nome = obter_aba_exemplo(abas_planilhas)
        if not aba_exemplo_nome:
            messagebox.showerror(
                "Erro", "Não foi encontrada uma aba com 'exemplo' no nome.")
            return

        aba_referencia = wb[aba_exemplo_nome]
        dados_aba_referencia = obter_dados_aba(aba_referencia)

        dict_referencia = criar_dict_referencia(
            dados_aba_referencia, aba_referencia)

        aba_atual_nome = obter_aba_atual(abas_planilhas)
        if not aba_atual_nome:
            messagebox.showerror(
                "Erro", "Não foi encontrada uma aba de extrato para o ano de 2026.")
            return

        aba_atual = wb[aba_atual_nome]
        dados_aba_atual = obter_dados_aba(aba_atual)

        preencher_dados(dados_aba_atual, dict_referencia, aba_atual)

        wb.save(caminho_arquivo)
        messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")

    except Exception as e:
        messagebox.showerror(
            "Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}")


class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Processador de Planilhas")
        self.geometry("600x400")

        self.caminho_arquivo = ttk.StringVar()

        self.container = ttk.Frame(self, padding=20)
        self.container.pack(fill=BOTH, expand=YES)

        # Header
        hdr_label = ttk.Label(
            self.container,
            text="Completar Planilhas Eliane",
            font=("Helvetica", 18, "bold"),
            bootstyle="primary"
        )
        hdr_label.pack(pady=(0, 20))

        # File Selection Section
        file_frame = ttk.Frame(self.container)
        file_frame.pack(fill=X, pady=10)

        self.file_entry = ttk.Entry(
            file_frame, textvariable=self.caminho_arquivo)
        self.file_entry.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))

        self.browse_btn = ttk.Button(
            file_frame,
            text="Selecionar Arquivo",
            command=self.selecionar_arquivo,
            bootstyle="outline-primary"
        )
        self.browse_btn.pack(side=RIGHT)

        # Info Label
        self.info_label = ttk.Label(
            self.container,
            text="Selecione um arquivo .xlsm ou .xlsx para começar",
            font=("Helvetica", 10),
            bootstyle="secondary"
        )
        self.info_label.pack(pady=10)

        # Process Button
        self.process_btn = ttk.Button(
            self.container,
            text="Processar Planilha",
            command=self.executar_processamento,
            bootstyle="success",
            width=20
        )
        self.process_btn.pack(pady=20)

    def selecionar_arquivo(self):
        file_path = filedialog.askopenfilename(
            title="Selecionar Planilha",
            filetypes=[("Excel Files", "*.xlsm *.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.caminho_arquivo.set(file_path)
            self.info_label.config(
                text=f"Selecionado: {Path(file_path).name}", bootstyle="info")

    def executar_processamento(self):
        caminho = self.caminho_arquivo.get()
        if not caminho:
            messagebox.showwarning(
                "Aviso", "Por favor, selecione um arquivo primeiro.")
            return

        processar_planilha(caminho)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
