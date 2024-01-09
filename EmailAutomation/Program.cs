using MimeKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using Newtonsoft.Json;
using System.Globalization;
using System.Text.RegularExpressions;
//using System.Net.Mail;

namespace EmailAutomation
{
    internal class Program
    {       
        static void Main(string[] args)
        {
            // Lendo o arquivo appsettings.json pra definir os parâmetros de configuração
            string jsonFilePath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
            string settingsFile = File.ReadAllText(jsonFilePath);

            // Convertendo os dados do JSON em uma estrutura de objeto do qual vamos usar ao longo do código
            ApplicationConfiguration configData = JsonConvert.DeserializeObject<ApplicationConfiguration>(settingsFile);

            // Criando algumas pastas padrões que vão auxiliar na parte de resultado do envio do arquivo
            string mainFolder = configData.MainFolder;
            string processedFolder = Path.Combine(mainFolder, "Processado");
            string failedFolder = Path.Combine(mainFolder, "Falha");

            if (Directory.Exists(processedFolder) == false) 
            {
                Directory.CreateDirectory(processedFolder); // Cria a pasta que vai indicar se os arquivos foram enviados com sucesso
                Console.WriteLine("[INFO] Pasta de processados foi criada!");
            }

            if (Directory.Exists(failedFolder) == false)
            {
                Directory.CreateDirectory(failedFolder); // Cria a pasta que vai indicar se os arquivos tiveram falha no envio
                Console.WriteLine("[INFO] Pasta de falhados foi criada!");
            }

            // Loop que vai executar as funções indefinitivamente
            while (true)
            {
                // Marcando a hora inicial do envio dos e-mails
                Console.WriteLine($"[INFO] {DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")} Iniciando envio do lote de e-mails");

                string[] attachmentFilesList = Directory.GetFiles(mainFolder); // Array contendo todos os arquivos .pdf (nome e caminho até o arquivo)

                if (attachmentFilesList.Length == 0) 
                {
                    Console.WriteLine($"[INFO] Nenhum arquivo foi encontrado na pasta, logo nada será feito");
                }
                else 
                {
                    // Enviando por email todos os arquivos da pasta
                    SendEmailWithAttachment(configData, attachmentFilesList, processedFolder, failedFolder);
                }             

                // Definição do tempo em que o programa ficará em stand-by a cada execução
                int timeIntervalInMinutes = configData.StandByTimeInMinutes;
                int minuteInMiliseconds = timeIntervalInMinutes * 60000;

                // Coloca o programa em stand-by antes de rodar novamente
                Console.WriteLine($"\n\nProcesso em stand-by, próxima vez que vai executar será em {DateTime.Now.Add(new TimeSpan(0, 0, timeIntervalInMinutes, 0))}\n\n");
                Thread.Sleep(minuteInMiliseconds);
                Console.WriteLine("---------------------------------------------------------------------------------\n\n");
            }
        }

        private static void SendFileToFinalDestination(string processedFolder, string failedFolder, string file, string fileName, int result)
        {
            // Concatena algumas informações adicionais ao arquivo enviado para evitar duplicidades
            string fileDestination;
            string dateTimeFormat = DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss");
            string fileWithDateTimeSent = $"{fileName}_{dateTimeFormat}.pdf";

            // Aqui será definido o destino final do arquivo que foi enviado
            if (result == 0)
            {
                fileDestination = Path.Combine(processedFolder, fileWithDateTimeSent);
                File.Move(file, fileDestination);
            }
            else if (result == 1)
            {
                fileDestination = Path.Combine(failedFolder, fileWithDateTimeSent);
                File.Move(file, fileDestination);
            }
            else
            {
                Environment.Exit(0);
            }
        }

        // Referência: https://learn.microsoft.com/en-us/dotnet/standard/base-types/how-to-verify-that-strings-are-in-valid-email-format
        private static bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try
            {
                // Normalize the domain
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(200));

                // Examines the domain part of the email and normalizes it.
                string DomainMapper(Match match)
                {
                    // Use IdnMapping class to convert Unicode domain names.
                    var idn = new IdnMapping();

                    // Pull out and process domain name (throws ArgumentException on invalid)
                    string domainName = idn.GetAscii(match.Groups[2].Value);

                    return match.Groups[1].Value + domainName;
                }
            }
            catch (RegexMatchTimeoutException e)
            {
                return false;
            }
            catch (ArgumentException e)
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }

        private static void SendEmailWithAttachment(ApplicationConfiguration configData, string[] filesList, string processedFolder, string failedFolder)
        {
            // Criando os objectos que vão armazenar os dados do e-mail e do corpo do e-mail
            MimeMessage message = new();            
            BodyBuilder bodyBuilder = new();

            // Já definindo alguns aspectos do e-mail do qual serão iguais a todos envios
            message.From.Add(new MailboxAddress(configData.FromEmailName, configData.FromEmailAddress));
            message.Bcc.Add(new MailboxAddress(configData.HiddenRedirectEmail, configData.HiddenRedirectEmail));
            message.Subject = configData.EmailSubject;
            bodyBuilder.HtmlBody = string.Format(configData.EmailTextBody);

            using (SmtpClient client = new())
            {
                string serverHost = configData.ServerHost;
                int serverPort = configData.ServerPort;
                SecureSocketOptions securityOptions = SecureSocketOptions.StartTls;               

                // Fazendo a conexão ao servidor de e-mail
                try
                {
                    client.Connect(serverHost, serverPort, securityOptions);
                    Console.WriteLine("[INFO] Conectado ao servidor!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERRO] Um erro crítico na conexão com o host ocorreu, verificar mensagem abaixo: \nMENSAGEM -> {ex.Message}");                    
                }                

                // Realizando a autenticação com o servidor
                try
                {
                    client.Authenticate(configData.FromEmailAddress, configData.EmailPassword);
                    Console.WriteLine("[INFO] Autenticado com o servidor!");
                }
                catch (AuthenticationException ex)
                {
                    Console.WriteLine($"[ERRO] Usuário ou senha inválidos para autenticação: \nMENSAGEM -> {ex.Message}");                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERRO] Um erro crítico na conexão com o host ocorreu, verificar mensagem abaixo: \nMENSAGEM -> {ex.Message}");
                }                

                // Aqui será analisado cada arquivo e enviado de acordo com o e-mail escrito como nome do arquivo
                foreach (string file in filesList)
                {
                    // Aqui será definido o código do resultado final do envio
                    // 0 = Sucesso / 1 = Falha / 2 = Erro Crítico
                    int emailStatusCode = 0;

                    // Vai criar uma cópia do arquivo a ser enviado, porém ele vai possuir um nome pré-definido ao invés do e-mail na nomenclatura
                    string tempFile = Path.Combine(configData.MainFolder, configData.TempFileName);
                    File.Copy(file, tempFile);

                    // Salva o nome do arquivo sem a extensão, pois é o e-mail a ser enviado o arquivo
                    string fileName = Path.GetFileNameWithoutExtension(file);

                    Console.WriteLine($"[INFO] Preparando o envio do e-mail {fileName}");

                    // Verificando se o e-mail é valido antes de enviar o arquivo
                    if (IsValidEmail(fileName) == false)
                    {
                        Console.WriteLine($"[ALERTA] O e-mail {fileName} está incorreto e o e-mail não será enviado");
                        emailStatusCode = 1;
                    }
                    else 
                    {
                        // Definindo pra quem vamos enviar o arquivo e qual o seu anexo
                        message.To.Add(new MailboxAddress(fileName, fileName));
                        bodyBuilder.Attachments.Add(tempFile);

                        // Atribui o conteúdo do corpo do e-mail ao objeto que contém os dados da mensagem
                        message.Body = bodyBuilder.ToMessageBody();

                        // Enviando o e-mail
                        try
                        {
                            client.Send(message);
                            Console.WriteLine("[INFO] E-mail enviado!");
                        }
                        catch (SmtpCommandException ex)
                        {
                            Console.WriteLine($"[ERRO] Falha ao enviar a mensagem: \nSTATUS CODE -> {ex.ErrorCode} \nMENSAGEM -> ");

                            switch (ex.ErrorCode)
                            {
                                case SmtpErrorCode.RecipientNotAccepted:
                                    Console.WriteLine($"Recebedor não aceito: {ex.Mailbox}");
                                    emailStatusCode = 1;
                                    break;
                                case SmtpErrorCode.SenderNotAccepted:
                                    Console.WriteLine($"Remetente não aceito: {ex.Mailbox}");
                                    emailStatusCode = 1;
                                    break;
                                case SmtpErrorCode.MessageNotAccepted:
                                    Console.WriteLine($"Mensagem não aceita: {ex.Mailbox}");
                                    emailStatusCode = 1;
                                    break;
                                default:
                                    emailStatusCode = 2;
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[ERRO] Um erro crítico no momento de enviar o e-mail ocorreu, verificar mensagem abaixo: \nMENSAGEM -> {ex.Message}");                            
                        }                        
                    }

                    // Deleta o arquivo temporário pois ele só serve para o arquivo atual
                    File.Delete(tempFile);

                    SendFileToFinalDestination(processedFolder, failedFolder, file, fileName, emailStatusCode);
                }

                client.Disconnect(true);
                Console.WriteLine("[INFO] Disconectado do servidor de e-mails");
            }
        }
    }   

    public class ApplicationConfiguration
    {
        public string MainFolder { get; set; }        
        public string TempFileName { get; set; }
        public string FromEmailAddress { get; set; }
        public string FromEmailName { get; set; }
        public string EmailSubject { get; set; }
        public string EmailTextBody { get; set; }        
        public string HiddenRedirectEmail { get; set; }        
        public string ServerHost { get; set; }
        public int ServerPort { get; set; }
        public string EmailPassword { get; set; }
        public int StandByTimeInMinutes {  get; set; }
    }
}
