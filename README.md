# Correção do FGTS com indice INPC
Com a desvalorização da TR como índice de correção do FGTS está sendo proposto o INPC como índice para correção do FGTS.
Para poder facilitar o entendimento de quanto é o diferencial esse código consegue a partir dos PDFs extraidos do aplicativo de Celular do FGTS ele pega todas as informações do documento e transforma em um arquivo JSON. Caso deseje usar de outra forma.

Aproveita o arquivo JSON e escreve um arquivo Excel Formatado.

O arquivo Excel só apresenta sucesso em um Microsoft Excel Desktop(Não testei no ambiente Apple mas no Windows está funcionando. Ambiente Web os calculos do Excel apresentam inconsistência)

## Mode de usar
Tenha uma instalação Python na maquina.

Clone esse projeto
git clone https://github.com/marianoaloi/correcao_FGTS_INPC.git

Entre na pasta do projeto "correcao_FGTS_INPC"
Com o executável Python em seu conhecimento e dentro da pasta do projeto execute o comando :
python -m venv venv
