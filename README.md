Este programa tem como utilidade automatizar o envio de arquivos por e-mail do qual são do mesmo tipo de conteúdo.

Após definir o caminho principal, só é necessário colocar o arquivo na pasta e renomear ele com o e-mail a ser enviado o arquivo e o programa vai fazer a coleta e depois o envio automático através do servidor de e-mail configurado.

É possível definir um texto padrão para o corpo de e-mail e definir também um e-mail do qual vai receber uma cópia oculta caso seja necessário.

O programa vai funcionar indefinidamente onde é possível configurar os intervalos do qual ficará em <i>stand-by</i> aguardando a próxima execução.

## Dependências

1. [MailKit](https://github.com/jstedfast/MailKit)
2. [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json/)

## JSON

Alguns parâmetros principais do programa são definidos através de um arquivo <b>appsettings.json</b> do qual deverá ser criado dentro da pasta do executável.

Abaixo segue a estrutura dele:

```json
{
  "mainFolder": string,  
  "tempFileName": string,
  "fromEmailAddress": string,
  "fromEmailName": string,
  "emailSubject": string,
  "emailTextBody": string,  
  "hiddenRedirectEmail": string,  
  "serverHost": string,
  "serverPort": int,
  "emailPassword": string,
  "standByTimeInMinutes": int
}
```
