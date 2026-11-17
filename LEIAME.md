# Envio de Contracheques

Roda direto na máquina do RH, em Python puro: lê os PDFs de uma pasta, confere código + nome
do recibo contra a planilha, protege cada PDF com os 4 últimos dígitos do CPF (AES-256) e envia
por e-mail (Outlook desta máquina ou servidor SMTP com login e senha).

## Arquivos principais

| Arquivo | Para quê |
|---|---|
| **Instalar.pyw** | Duplo clique, uma vez. Confere tudo e cria o atalho. |
| **Executar.pyw** | Abre o programa (o atalho aponta para ele). |
| `configuracoes.ini` | Criado pela aba Configurações. Não editar à mão. |
| `credencial.dat` | Senha do e-mail criptografada pelo Windows (só se usar SMTP). |
| `config.ini` | Regras avançadas: padrão do nome do arquivo, layout do recibo. |

## Instalação na máquina do RH

1. Instalar o **Python 3.12** do python.org (Windows installer 64-bit) → "Install Now".
   Instalação por usuário, sem administrador. Não usar o Python da Microsoft Store.
2. Extrair o zip em `C:\RH\EnvioContracheques` (pasta com acesso restrito ao RH).
3. Duplo clique em **Instalar.pyw** → Instalar → Abrir o programa.
4. Na primeira abertura, a aba **Configurações** abre sozinha: pastas, forma de envio, conta
   ou login/senha, e-mail de teste e senha do PDF. Salvar.

## Por que o Windows não bloqueia

Não existe nenhum `.exe`, `.dll` ou `.pyd` próprio. Quem roda é o `pythonw.exe` oficial
(assinado pela Python Software Foundation); o resto é código `.py`:
`aes_puro.py` substitui o pacote cryptography, `comtypes` substitui o pywin32, e o cofre de senha
usa a DPAPI do próprio Windows. A pasta `lib/` já vem com as bibliotecas — sem pip, sem internet.

## PDF único da contabilidade

Aba Envio → Origem: **PDF único da contabilidade** → selecionar o arquivo → **Separar e renomear**.
Cada página é atribuída pelo código + nome do cabeçalho do recibo; página sem cabeçalho é
continuação da anterior. Os arquivos saem como `Maria Silva_Recibo de Salário 09.2026.pdf`
numa pasta ao lado do PDF original, e a conferência roda em seguida.
Página com dois colaboradores diferentes **não é separada** e aparece no aviso final.

## Tipo de folha e competência

Lidos do próprio recibo (a linha acima de "Setembro de 2026" é o tipo: Folha Mensal, Férias,
13º Salário...). Com a competência em **Automática**, cada recibo usa a sua; arquivo de mês
diferente da maioria vai para Revisar. Os dois aparecem na tabela, no assunto e no e-mail.

## Conferência visual

Duplo clique em qualquer linha abre o PDF. Se a linha estiver em Revisar, o programa pergunta
se pode liberar. Botão direito: abrir PDF, liberar, mostrar na pasta.

## Mensagem do e-mail

Aba **Mensagem do e-mail**: assunto e texto editáveis, botões para inserir campos
({primeiro_nome}, {tipo_folha}, {mes_extenso}, {dica_senha}...), pré-visualização no navegador.
As cores da tela e do e-mail ficam na seção `[tema]` do `config.ini`.

## Logo e cores da empresa

Coloque a logo como **`logo.png`** na pasta do programa (PNG com fundo transparente; a versão
colorida, para fundo claro). Ela aparece no topo do programa, no instalador e no ícone da janela.
Com `cores_da_logo = sim` (seção `[tema]` do `config.ini`), a cor principal da tela e do e-mail é
extraída da própria logo e ajustada para o texto branco ficar legível. Sem logo, vale a paleta
da seção `[tema]` e o nome da empresa aparece no lugar da logo.

O **Instalar.pyw** também gera o `icone.ico` a partir da logo e aplica nos atalhos (Área de
Trabalho, Menu Iniciar e um atalho dentro da pasta do programa). Se houver um `icone.png`
(só o símbolo, quadrado), ele tem prioridade — fica mais nítido em tamanho pequeno.
Trocou a logo? Rode o Instalar.pyw de novo.

## Primeiro uso

1. Configurações → **Simulado** → Salvar. Envio → Conferir → Enviar. Os PDFs protegidos e os
   e-mails ficam em `logs\simulacao`. Abra alguns com a senha.
2. Configurações → **Outlook** (ou SMTP) → Salvar. Envio com **Modo teste** para o e-mail do RH.
3. Desmarque o modo teste, Conferir e enviar de verdade.

## Travas de segurança

- Código **e** nome impressos no recibo precisam bater com a planilha; senão bloqueia.
- Competência do arquivo e do recibo precisa bater com a selecionada.
- Homônimo, código duplicado, PDF sem cadastro ou com mais de um colaborador → bloqueia.
- O PDF protegido é reaberto com a senha antes de anexar.
- Senha do PDF nunca vai para log, histórico ou e-mail. Senha do e-mail só no cofre do Windows.
- Desempenho: ~3 s por PDF. 200 recibos ≈ 15 a 20 minutos.

## Problemas

- Instalador falhou: `instalar.log`.  Programa não abre: `logs\erro_inicializacao.log`.
