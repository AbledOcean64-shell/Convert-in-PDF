# Conversor de Arquivos para PDF (PowerShell)
Script em PowerShell desenvolvido para automatizar a conversão de arquivos do Microsoft Office (Word, Excel, PowerPoint e RTF) para o formato PDF de forma rápida, prática e confiável.
A ferramenta realiza a leitura de uma pasta definida, identifica automaticamente os arquivos suportados e executa a conversão utilizando automação COM do Microsoft Office. Para garantir maior estabilidade, o script conta com um sistema de retry (tentativas automáticas) em caso de falhas durante o processo.

No script, a variável abaixo define o caminho dos arquivos:
$folder = "."

Basta substituir o valor pelo caminho exato de onde estão os arquivos que vão ser convertidos, exemplo:
$folder = "C:\Users\SeuUsuario\Documents\Arquivos"

# Como utilizar

1. Salve o arquivo com extensão .ps1, por exemplo:
   converter_pdf.ps1

2.Preparar os arquivos, escolha uma das opções:
Colocar os arquivos na mesma pasta do script ou Definir um caminho específico na variável $folder

3. Executar o script com modo administrador;
   Obs: O código funciona sem o modo admin também, mas pode dar erro e parar a conversão.

4. Durante a execução, o script exibirá:

- Quantidade de arquivos encontrados;
- Status de cada conversão:
- OK → Conversão realizada com sucesso
- Erro → Tentativa automática até 3 vezes

Obs: Caso dê erro em todos ou o mesmo não converta corretamente, analise a extensão dos arquivos que está tentando converter, na estrutura do script está apenas "(doc|docx|xls|xlsx|ppt|pptx|rtf)", mas pode acrescentar outras extensões compatíveis com o pacote office, como por exemplo "docm ou xlsm";

5. No fim da conversão, estará todos os arquivos na mesma pasta tanto o original e o PDF de cada um.

# NOTA: Quaisquer dúvidas, enviar para o e-mail de dúvidas: abled9912@cyberdude.com
