# Disparo de Mensagens no WhatsApp

Este projeto implementa uma automação para envio de mensagens no WhatsApp utilizando Python. O código lê dados de clientes a partir de uma planilha Excel, gera mensagens personalizadas e as envia via WhatsApp Web.

## Funcionalidades

- Abre o navegador Chrome e acessa o WhatsApp Web.
- Lê dados dos clientes a partir de uma planilha Excel (`clientes.xlsx`).
- Gera mensagens personalizadas incluindo o nome do cliente.
- Codifica a mensagem em uma URL e a abre no navegador.
- Localiza e clica no botão de envio usando `pyautogui`.
- Fecha a aba do navegador e repete o processo para o próximo cliente.

## Tecnologias Utilizadas

- `webbrowser` - Para abrir o navegador e acessar URLs.
- `openpyxl` - Para ler dados da planilha Excel.
- `urllib.parse` - Para codificar mensagens em URLs.
- `time` - Para adicionar delays no processo de automação.
- `pyautogui` - Para localizar e clicar em elementos na tela.

## Como Usar

1. **Instale as dependências necessárias**:
    ```sh
    pip install openpyxl pyautogui
    ```

2. **Prepare sua planilha Excel**:
    - Crie um arquivo `clientes.xlsx` com os dados dos clientes.
    - A planilha deve ter pelo menos duas colunas: `Nome` e `Telefone`.

3. **Execute o script**:
    - Certifique-se de que o caminho do Chrome está correto no script.
    - Execute o script Python para iniciar o processo de envio de mensagens.

