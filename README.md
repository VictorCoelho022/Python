Como funciona? 🔧
- O sistema lê os dados dos clientes (nome e telefone) a partir de uma planilha.
- Gera uma mensagem e envia automaticamente via WhatsApp.
- Caso haja algum erro, a informação é registrada para acompanhamento.

🧩 Código Explicado:

1. Importação das bibliotecas

📝 openpyxl: Para trabalhar com arquivos Excel, neste caso, para ler a planilha de contatos

📝urllib.parse.quote: Para garantir que o texto enviado via WhatsApp seja corretamente codificado em formato de URL.

📝webbrowser: Para abrir URLs (neste caso, o link do WhatsApp).

📝time.sleep: Para pausar a execução do código por alguns segundos, simulando o tempo que o usuário levaria para fazer ações manualmente. 

📝pyautogui: Para simular ações do teclado e do mouse, permitindo que o código interaja com a interface do WhatsApp Web.

2. Abrindo o WhatsApp Web

📝webbrowser.open("https://web.whatsapp.com/")
📝sleep(20)

Aqui, o código usa a função open() da biblioteca webbrowser para abrir o WhatsApp Web no navegador. A função sleep(20) pausa o script por 20 segundos para garantir que o WhatsApp Web tenha tempo suficiente para carregar e o usuário possa se autenticar.

3. Carregando a planilha de contatos

📝workbook = openpyxl.load_workbook("MKT.xlsx")
📝pagina_cliente = workbook["Teste"]

Aqui o código carrega o arquivo Excel chamado "MKT.xlsx" usando a função load_workbook da biblioteca openpyxl. Depois, acessa a aba chamada "Teste", onde você tem os dados dos clientes.

4. Iterando sobre os dados da planilha

📝for linha in pagina_cliente.iter_rows(min_row=2):
 # nome, telefone
 nome = linha[0].value
 telefone = linha[1].value

Neste trecho, o código percorre as linhas da planilha, começando pela segunda linha (min_row=2) — a primeira linha contém os cabeçalhos das colunas. 



1. Construção do link para enviar a mensagem no WhatsApp
python
Copiar código
link_mensagem_whatsapp = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
Aqui, a variável link_mensagem_whatsapp está sendo construída com o formato de URL que o WhatsApp Web usa para enviar mensagens automaticamente.
https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}: Este link é o formato que o WhatsApp Web reconhece. Ele permite enviar uma mensagem para um número específico e incluir um texto na URL.
{telefone}: A variável telefone contém o número do telefone do cliente, e será inserida no link.
{quote(mensagem)}: A função quote() da biblioteca urllib.parse é usada para codificar a mensagem em um formato seguro para a URL. Isso significa que caracteres especiais (como espaços, acentos, etc.) serão convertidos em uma forma compatível com o protocolo de URLs (por exemplo, espaços se tornam %20).
2. Abrir o link no navegador
python
Copiar código
webbrowser.open(link_mensagem_whatsapp)
Essa linha usa a função open() da biblioteca webbrowser para abrir a URL gerada. Esse comando vai abrir o WhatsApp Web no navegador, redirecionando para a conversa com o número de telefone fornecido, já com a mensagem pronta para ser enviada.
3. Esperar o carregamento da página
python
Copiar código
sleep(15)
O sleep(15) faz o código "esperar" 15 segundos antes de prosseguir para o próximo passo. Esse tempo é dado para garantir que a página do WhatsApp Web seja carregada corretamente e que o botão de enviar esteja disponível.
4. Simular pressionamento da tecla "Tab"
python
Copiar código
pyautogui.press("tab")
A função pyautogui.press("tab") simula o pressionamento da tecla Tab, que normalmente move o foco para o próximo campo ou botão. No caso do WhatsApp Web, o foco vai para o campo de texto onde a mensagem aparece, tornando-o pronto para o envio.
5. Esperar um pouco mais
python
Copiar código
sleep(6)
Após pressionar "Tab", o código aguarda 6 segundos para garantir que o campo de texto tenha sido selecionado corretamente antes de passar para o próximo passo.
6. Simular pressionamento da tecla "Enter"
python
Copiar código
pyautogui.press("enter")
A função pyautogui.press("enter") simula o pressionamento da tecla Enter, que no WhatsApp Web é responsável por enviar a mensagem.
7. Fechar a aba do navegador
python
Copiar código
pyautogui.hotkey("ctrl", "w")
A função pyautogui.hotkey("ctrl", "w") simula o atalho de teclado Ctrl + W, que fecha a aba atual do navegador. Nesse caso, a aba do WhatsApp Web que foi aberta para enviar a mensagem será fechada após o envio.
8. Esperar novamente antes de terminar
python
Copiar código
sleep(6)
O último sleep(6) espera 6 segundos antes de terminar a execução do código. Isso garante que o processo de fechar a aba seja feito de forma controlada, dando tempo para o envio da mensagem e o fechamento da aba.
Resumo do Processo:
O link do WhatsApp Web é gerado com o número de telefone e a mensagem.
O WhatsApp Web é aberto no navegador.
O script espera o carregamento da página.
O foco é movido para o campo de mensagem e a tecla "Enter" é pressionada para enviar a mensagem.
A aba do navegador é fechada automaticamente.
O script finaliza aguardando um tempo extra para garantir que o processo seja bem executado.
