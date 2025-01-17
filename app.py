import win32com.client
import re
import google.generativeai as genai

# Função para ler o arquivo PST no Outlook
def ler_pst(pst_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    namespace.AddStore(pst_path)
    inbox = namespace.GetDefaultFolder(6)  # "Caixa de Entrada"
    messages = inbox.Items

    emails = []
    for message in messages:
        try:
            subject = message.Subject
            body = message.Body
            sender = message.SenderName
            emails.append({'subject': subject, 'body': body, 'date': message.ReceivedTime, 'sender': sender})
        except Exception as e:
            print(f"Erro ao ler mensagem: {e}")

    return emails

# Função para enviar a pergunta para o Gemini e processar a resposta
def consultar_gemini(pergunta, emails):
    # Cria o prompt com base na pergunta do usuário e nos e-mails
    e_mail_texts = "\n".join([f"Assunto: {email['subject']}\nCorpo: {email['body'][:150]}..." for email in emails])

    prompt = f"Você tem acesso a um conjunto de e-mails com os seguintes detalhes:\n{e_mail_texts}\n\nAgora, responda à seguinte pergunta: {pergunta}"

    # Configuração da API Gemini
    try:
        genai.configure(api_key="API_KEY")
        model = genai.GenerativeModel(model_name="gemini-1.5-flash") #versão da api
        response = model.generate_content(contents=[{"text": prompt}])
        return response.text.strip()
    except Exception as e:
        return f"Erro ao consultar Gemini: {e}"

# Função principal
def main():
    # Caminho do arquivo PST
    pst_path = r"Caminho_do_seu_arquivo_PST"

    # Lê os e-mails do arquivo PST
    emails = ler_pst(pst_path)

    print(f"{len(emails)} e-mails encontrados no arquivo PST.")

    while True:
        # Solicita uma pergunta ao usuário
        pergunta = input("Qual pergunta você gostaria de fazer sobre os e-mails? (Digite 'sair' para encerrar) ")

        if pergunta.lower() == 'sair':
            print("Saindo do programa...")
            break

        # Responde à pergunta usando o Gemini
        resposta = consultar_gemini(pergunta, emails)
        print(resposta)

if __name__ == "__main__":
    main()
