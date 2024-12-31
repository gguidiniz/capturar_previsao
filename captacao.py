import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Função para capturar dados e salvar na planilha
def capturar_dados():
    try:
        # Configuração do Selenium
        driver = webdriver.Chrome()  # Certifique-se de que o ChromeDriver está instalado e configurado no PATH
        driver.get("https://www.google.com")

        # Pesquisar a temperatura no Google
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys("temperatura")
        search_box.send_keys(Keys.RETURN)
        time.sleep(3)  # Aguarde o carregamento da página

        # Capturar a temperatura e status da umidade
        temperatura = driver.find_element(By.ID, "wob_tm").text  # Exemplo para o elemento da temperatura
        umidade = driver.find_element(By.ID, "wob_hm").text  # Exemplo para o elemento da umidade

        # Obter a data e hora atual
        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Fechar o navegador
        driver.quit()

        # Salvar os dados na planilha
        arquivo = "historico_temperatura.xlsx"
        wb = load_workbook(arquivo)
        ws = wb.active

        ws.append([data_hora, temperatura, umidade])
        wb.save(arquivo)

        # Exibir mensagem de sucesso
        messagebox.showinfo("Sucesso", "Dados capturados e salvos com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao capturar dados: {e}")

# Criar a interface com Tkinter
def criar_interface():
    janela = tk.Tk()
    janela.title("Atualizar previsão")

    # Título na interface
    titulo = tk.Label(janela, text="Atualizar previsão na planilha", font=("Arial", 14, "bold"))
    titulo.pack(pady=10)

    # Botão para executar a captação
    btn_capturar = tk.Button(janela, text="Buscar previsão", command=capturar_dados, width=20, height=2, bg="lightblue")
    btn_capturar.pack(pady=20)

    # Loop principal da interface
    janela.mainloop()

if __name__ == "__main__":
    criar_interface()
