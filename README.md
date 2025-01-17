# Documentação do Projeto: Análise de E-mails com Gemini e PST

## Descrição
Este projeto tem como objetivo processar arquivos PST do Outlook para extrair e-mails e realizar consultas inteligentes utilizando o modelo generativo Gemini da API Google Generative AI. Ele permite analisar o conteúdo dos e-mails e responder perguntas baseadas nas informações extraídas.

## Funcionalidades
- **Leitura de arquivos PST do Outlook:** Extrai os e-mails, incluindo assunto, corpo, remetente e data de recebimento.
- **Integração com Gemini:** Permite realizar consultas inteligentes com base nos e-mails extraídos.
- **Interface interativa:** Usuário pode fazer perguntas sobre os e-mails diretamente no terminal.

## Estrutura do Código

### 1. **Função `ler_pst(pst_path)`**
Lê e processa o arquivo PST no Outlook, retornando uma lista de dicionários com informações dos e-mails.

**Parâmetros:**
- `pst_path` (str): Caminho do arquivo PST.

**Retorno:**
- `emails` (list): Lista de dicionários contendo os e-mails extraídos.

### 2. **Função `consultar_gemini(pergunta, emails)`**
Envia uma pergunta para o modelo Gemini e processa a resposta com base nos e-mails fornecidos.

**Parâmetros:**
- `pergunta` (str): A pergunta a ser respondida.
- `emails` (list): Lista de e-mails extraídos do arquivo PST.

**Retorno:**
- Resposta gerada pelo Gemini ou mensagem de erro.

### 3. **Função `main()`**
Função principal que:
1. Lê os e-mails do arquivo PST.
2. Solicita perguntas ao usuário.
3. Exibe as respostas geradas pelo Gemini.

### 4. **Execução**
O código é executado no terminal, iniciando pela função `main()`.

## Pré-requisitos

### Ferramentas e Bibliotecas
- Python 3.8 ou superior
- Microsoft Outlook instalado (para manipulação do arquivo PST)
- Bibliotecas Python:
  - `win32com.client`
  - `re`
  - `google.generativeai`

### Instalação de Bibliotecas
Para instalar as bibliotecas necessárias, use o seguinte comando:
```bash
pip install pywin32 google-generativeai

# API Gemini

Configure a API Gemini com sua chave de acesso:

```python
import genai

genai.configure(api_key="API_KEY")
```

## Configuração e Execução

1. **Baixe o arquivo PST**: Salve o arquivo PST a ser processado no local desejado e atualize o caminho na variável `pst_path` dentro da função `main()`.

2. **Execute o script**: No terminal, execute o arquivo Python:

   ```bash
   python script.py
   ```

3. **Faça perguntas sobre os e-mails**: Digite suas perguntas no terminal e receba as respostas do modelo Gemini. Para encerrar, digite `sair`.

## Estrutura do Projeto

```bash
/
├── script.py           # Código principal
├── requirements.txt    # Dependências do projeto
```

### Exemplo do arquivo `requirements.txt`:

```
pywin32
google-generativeai
```

## Considerações

- Certifique-se de que o arquivo PST está acessível e o Outlook está configurado corretamente.
- Substitua `"API_KEY"` pela chave de acesso válida da API Gemini.
- Para grandes volumes de e-mails, o tempo de processamento pode aumentar.

## Erros Comuns e Soluções

**Erro**: `"Erro ao ler mensagem"`

- Isso pode ocorrer devido a problemas de formatação em um e-mail específico. O código ignora e continua processando os demais e-mails.

**Erro ao configurar a API Gemini**

- Verifique se a chave da API foi configurada corretamente e se há conexão com a internet.

## Melhorias Futuras

- Adicionar suporte a arquivos PST criptografados.
- Implementar uma interface gráfica para facilitar o uso.
- Melhorar o processamento de erros para maior robustez.

---

**Autor**
Desenvolvido por Luana Victoria.
