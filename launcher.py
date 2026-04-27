"""
launcher.py — Iniciador do Diário de Bordo (para PyInstaller)
==============================================================
Compila com:
    pyinstaller --onefile --windowed launcher.py

Ao executar o .exe gerado, o Streamlit sobe na porta 8501 e o
browser abre automaticamente na aplicação.
"""

import os
import sys
import time
import socket
import webbrowser
import subprocess
import tkinter as tk
from tkinter import messagebox

PORT = 8501
URL  = f"http://localhost:{PORT}"


def porta_em_uso(porta: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("localhost", porta)) == 0


def caminho_app() -> str:
    """Retorna o caminho para o diario_bordo_entrada.py mesmo dentro do bundle PyInstaller."""
    if getattr(sys, "_MEIPASS", None):
        return os.path.join(sys._MEIPASS, "diario_bordo_entrada.py")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "diario_bordo_entrada.py")


def iniciar_streamlit(app_path: str) -> subprocess.Popen:
    python = sys.executable
    cmd = [
        python, "-m", "streamlit", "run", app_path,
        "--server.port", str(PORT),
        "--server.headless", "true",
        "--server.runOnSave", "false",
        "--browser.gatherUsageStats", "false",
        "--client.showErrorDetails", "false",
    ]
    return subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )


def main():
    app_path = caminho_app()

    if not os.path.exists(app_path):
        messagebox.showerror(
            "Erro",
            f"Arquivo da aplicação não encontrado:\n{app_path}"
        )
        return

    # Se já estiver rodando, apenas abre o browser e sai
    if porta_em_uso(PORT):
        webbrowser.open(URL)
        return

    processo = iniciar_streamlit(app_path)

    # Janela de status mínima
    root = tk.Tk()
    root.title("Diário de Bordo — Entrada")
    root.resizable(False, False)
    root.geometry("360x120")

    lbl = tk.Label(root, text="Iniciando aplicação...\nAguarde.",
                   padx=20, pady=20, font=("Segoe UI", 11))
    lbl.pack()

    browser_aberto = [False]

    def verificar_e_abrir():
        """Polling da porta via tkinter; abre o browser UMA vez quando o servidor subir."""
        if not browser_aberto[0]:
            if porta_em_uso(PORT):
                browser_aberto[0] = True
                webbrowser.open(URL)
                lbl.config(text="Aplicação em execução.\nFeche esta janela para encerrar.")
            else:
                root.after(1000, verificar_e_abrir)  # tenta novamente em 1 segundo

    root.after(1500, verificar_e_abrir)  # primeira verificação após 1,5 s

    def ao_fechar():
        processo.terminate()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", ao_fechar)
    root.mainloop()

    # Garante encerramento do processo filho
    if processo.poll() is None:
        processo.terminate()


if __name__ == "__main__":
    main()



# pyinstaller --onefile --windowed --add-data "diario_bordo_entrada.py;." launcher.py