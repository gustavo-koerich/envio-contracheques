# Envio de Contracheques

[![testes](https://github.com/gustavo-koerich/envio-contracheques/actions/workflows/testes.yml/badge.svg)](https://github.com/gustavo-koerich/envio-contracheques/actions/workflows/testes.yml)

Programa de desktop para o RH: lê os recibos de salário em PDF, confere cada recibo contra o cadastro de colaboradores, **protege cada PDF com senha** (AES-256) e envia por e-mail, um por pessoa. Antes, o envio era feito à mão, um arquivo por vez.

> Repositório de portfólio. A planilha de exemplo e os nomes usados no código são **fictícios**.

## Como funciona

1. **Origem**: uma pasta com um PDF por colaborador, ou o **PDF único da contabilidade**. No segundo caso, o programa separa as páginas pelo código e nome impressos no cabeçalho de cada recibo e renomeia os arquivos.
2. **Conferência** com várias travas:
   - o código **e** o nome impressos no recibo têm de bater com a planilha;
   - a competência ("Setembro de 2026") e o tipo de folha (mensal, férias, 13º) são lidos do próprio PDF;
   - homônimos, códigos duplicados, PDFs sem cadastro ou com mais de um colaborador → **bloqueia**.
3. **Proteção**: a senha de abertura é montada a partir de dados do colaborador (ex.: 4 últimos dígitos do CPF). O PDF protegido é **reaberto com a senha** antes de ser anexado, para garantir que abre.
4. **Envio** pelo Outlook da máquina (COM) ou por SMTP, com modo **simulado** e modo **teste**, pausa para respeitar o limite do Exchange e histórico de envios.

## Destaques técnicos

- **Sem nenhum binário próprio**, para não esbarrar em antivírus nem em políticas corporativas:
  - roda no `pythonw.exe` oficial;
  - AES-256 implementado em Python puro ([`aes_puro.py`](aes_puro.py)) no lugar do pacote `cryptography`;
  - automação do Outlook via `comtypes` no lugar do `pywin32`;
  - todas as dependências são Python puro, carregadas de `lib/` (sem pip nem internet na máquina do RH).
- **Senha do e-mail no cofre do Windows** (DPAPI, criptografia por usuário e máquina, em [`cofre.py`](cofre.py)). A senha do PDF nunca vai para log, histórico ou e-mail.
- **Interface Tkinter** com conferência visual (duplo clique abre o PDF), editor da mensagem com campos (`{primeiro_nome}`, `{tipo_folha}`, `{dica_senha}`…) e pré-visualização.
- **Identidade visual automática**: as cores da tela e do e-mail são extraídas da logo da empresa. O instalador gera o ícone e os atalhos.

## Rodando

Requer Windows e Python 3.12.

```powershell
pip install --target lib -r requirements.txt   # bibliotecas em Python puro na pasta lib/
pythonw Instalar.pyw                            # confere o ambiente e cria os atalhos
pythonw Executar.pyw
```

Para experimentar sem enviar nada:

1. Em **Configurações**, escolha o modo **Simulado** e use a planilha de exemplo [`modelo_colaboradores.xlsx`](modelo_colaboradores.xlsx).
2. Os PDFs protegidos e os e-mails gerados ficam em `logs\simulacao`.

As regras avançadas ficam em [`config.ini`](config.ini): padrão do nome do arquivo, layout do cabeçalho do recibo e formato da senha. O manual de uso para o RH está em [LEIAME.md](LEIAME.md).

## Testes

```powershell
python -m unittest discover -s tests
```

Cobrem a montagem da senha, o bloqueio quando falta dado e a proteção AES-256 (o PDF só abre com a senha certa).

## Estrutura

```
app.py           interface (Tkinter)
nucleo.py        conferência, separação do PDF único, proteção e envio
aes_puro.py      AES-256 em Python puro (criptografia dos PDFs)
cofre.py         senha do e-mail no cofre do Windows (DPAPI)
identidade.py    cores e ícone a partir da logo
Instalar.pyw     instalador (conferência do ambiente + atalhos)
```

## Stack

Python 3.12 · Tkinter · pypdf · openpyxl · comtypes (Outlook COM) · SMTP · DPAPI

## Licença

Código disponibilizado publicamente para fins de portfólio e demonstração técnica. Todos os direitos reservados.
