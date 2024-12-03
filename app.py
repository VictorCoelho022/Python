"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/MEUS CLIENTES """
# Descrever os passos manuais e dps transformar isso em código
# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open("https://web.whatsapp.com/")
sleep(20)

workbook = openpyxl.load_workbook("MKT.xlsx")
pagina_cliente = workbook["Teste"]

for linha in pagina_cliente.iter_rows(min_row=2):
    # nome, telefone
    nome = linha[0].value
    telefone = linha[1].value

    mensagem = f"""Olá, boa tarde! Tudo bem? 
Meu nome é Victor Freire, sou representante da empresa XXX e estou disponível para te ajudar a encontrar a melhor opção do produto XXX para você e seu negócio. 

Benefícios ao adquirir a seu produto XXX diretamente com o meu atendimento
individual:
✅Acesso a descontos exclusivos e menores taxas do mercado
✅Frete grátis
✅Suporte pós compra

Você já possui o produto XXX?

Me diga por gentileza para que possa entender a sua necessidade e indicar a melhor para você✅💚"""
   


# Criar links personalizados do whatsapp e enviar mensagens para cada cliente
# com base nos dados da planilha
    # https://web.whatsapp.com/send?phone=number&text="como vc ta?"
    try:
        link_mensagem_whatsapp = f"https://web.whatsapp.com/send?phone={
            telefone}&text={quote(mensagem)}"
        webbrowser.open(link_mensagem_whatsapp)
        sleep(15)
        pyautogui.press("tab")
        sleep(6)
        pyautogui.press("enter")
        sleep(6)
        pyautogui.hotkey("ctrl", "w")
        sleep(6)
    except:
        print(f"Não foi possível fazer contato")
        with open("erros.csv", "a", newline="", encoding="utf-8") as arquivo:
            arquivo.write(f"{nome},{telefone}")

