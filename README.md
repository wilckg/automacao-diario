# Automação do Diário Eletrônico SENAI

Automação em Python para auxiliar no registro de aulas e frequência de alunos no **Diário Eletrônico SENAI**, utilizando **Playwright** para interação com o sistema web e **OpenPyXL** para leitura de uma planilha Excel.

O projeto foi criado para reduzir o trabalho manual de registrar várias aulas de uma turma, mantendo uma etapa de confirmação antes do processamento em lote e validações para evitar registros inconsistentes ou duplicados.

## O que a automação faz

O fluxo atual permite:

- solicitar NIF e senha no terminal, sem armazenar a senha no código;
- realizar login no Diário Eletrônico com até 3 tentativas;
- listar as unidades disponíveis e permitir escolher a unidade desejada;
- listar as turmas disponíveis e permitir escolher a turma;
- ler as aulas definidas na aba `Aulas` do Excel;
- identificar as aulas já registradas no Diário e removê-las do lote;
- pedir uma confirmação única antes de iniciar os registros;
- abrir o formulário de cada aula;
- preencher automaticamente data, carga horária e conteúdo;
- comparar os alunos do Excel com os alunos do Diário pelo nome normalizado;
- interpretar `.` como presença e `I` como falta;
- registrar faltas como `Não compensado`, incluindo o período da falta;
- salvar a aula automaticamente;
- repetir o processo para todas as aulas pendentes;
- interromper o lote no primeiro erro, evitando continuar em um estado desconhecido.

## Tecnologias utilizadas

- Python
- Playwright
- OpenPyXL
- Chromium
- Excel (`.xlsx`)

## Requisitos

Recomenda-se utilizar **Python 3.10 ou superior**.

Verifique a instalação:

```bash
python --version
```

No Windows, dependendo da instalação, também pode ser necessário usar:

```bash
py --version
```

## Instalação

### 1. Clone ou copie o projeto

Exemplo de estrutura:

```text
automacao-diario/
├── main.py
├── lista_presenca_com_frequencia.xlsx
└── README.md
```

O código atual espera encontrar o arquivo:

```text
lista_presenca_com_frequencia.xlsx
```

na mesma pasta do `main.py`.

Caso queira utilizar outro nome, altere no `main()`:

```python
ARQUIVO_PRESENCA = "lista_presenca_com_frequencia.xlsx"
```

### 2. Criar um ambiente virtual

No Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

No Prompt de Comando:

```cmd
python -m venv .venv
.venv\Scripts\activate
```

### 3. Instalar as dependências

```bash
pip install playwright openpyxl
```

### 4. Instalar o navegador usado pelo Playwright

```bash
playwright install chromium
```

Se o comando `playwright` não estiver disponível diretamente:

```bash
python -m playwright install chromium
```

## Configuração da planilha

A automação utiliza duas abas principais:

```text
Lista de Presença
Aulas
```

A antiga aba `Legenda` não é necessária para a execução.

### Aba `Lista de Presença`

Essa aba contém os alunos e a frequência de cada dia.

No formato usado atualmente:

- linha 2: dias das aulas (`07`, `08`, `13`, etc.);
- coluna B: nome do aluno;
- linha 3 em diante: alunos;
- `.` = presente;
- `I` = falta.

Exemplo simplificado:

| Nome | 07 | 08 | 13 |
|---|---|---|---|
| AIRTON GABRIEL MENDES ASSUNÇÃO | . | . | I |
| FLORENCE MANIUS | I | . | . |
| RODRIGO VINICIUS MARTINS | I | . | . |

> A automação faz a associação com o Diário pelo nome. Os nomes são normalizados, removendo diferenças de maiúsculas/minúsculas, acentos e espaços extras, mas o ideal é manter os nomes o mais próximos possível dos exibidos no Diário.

### Aba `Aulas`

A aba `Aulas` define quais aulas devem ser registradas e qual conteúdo deve ser lançado.

Utilize as seguintes colunas:

| Data | Horas | Conteúdo | Registrar |
|---|---|---|---|
| 07/07/2026 | 03:00 | Introdução ao Marketing Digital e Inteligência Artificial | SIM |
| 08/07/2026 | 03:00 | Fundamentos de Marketing Digital | SIM |
| 13/07/2026 | 03:00 | Inteligência Artificial aplicada ao Marketing | SIM |
| 14/07/2026 | 03:00 | Criação de conteúdo com IA | NÃO |

#### Data

Preferencialmente utilize uma data real do Excel no formato visual:

```text
07/07/2026
```

Internamente o programa converte para o formato exigido pelo campo HTML:

```text
2026-07-07
```

#### Horas

Use valores como:

```text
03:00
02:00
04:00
```

A mesma duração é usada para:

- carga horária da aula;
- período de falta integral do aluno ausente.

#### Conteúdo

Informe exatamente o conteúdo que deverá aparecer no campo **Conteúdo da Aula** do Diário.

Exemplo:

```text
Introdução ao Marketing Digital e Inteligência Artificial.
```

#### Registrar

São considerados valores habilitados:

```text
SIM
S
1
```

Qualquer outro valor faz a linha ser ignorada.

Isso permite preparar várias aulas na planilha e controlar quais serão processadas em determinada execução.

## Como executar

Com o ambiente virtual ativo:

```bash
python main.py
```

O programa irá solicitar:

```text
=== Diário Eletrônico SENAI ===

Digite seu NIF:
Digite sua senha:
```

A senha é digitada de forma oculta no terminal.

Depois do login, será necessário escolher a unidade:

```text
=== UNIDADES ===

1. CFP ...
2. CFP ...

Escolha a unidade: 2
```

Em seguida, escolher a turma:

```text
=== TURMAS DISPONÍVEIS ===

1. MDIA-BOL-2N-26 | Marketing Digital com Inteligência Artificial | ...
2. ...

Escolha a turma: 1
```

## Proteção contra aulas duplicadas

Antes de iniciar o lote, o programa procura datas já registradas nas tabelas da tela da turma.

Exemplo:

```text
=== VERIFICAÇÃO DE DUPLICIDADE ===

IGNORAR - 07/07/2026 já está registrada.
IGNORAR - 08/07/2026 já está registrada.
PROCESSAR - 13/07/2026
PROCESSAR - 14/07/2026

Pendentes: 2
Ignoradas por duplicidade: 2
```

Somente as aulas pendentes entram na confirmação do lote.

A data também é verificada novamente imediatamente antes do processamento de cada aula.

## Confirmação antes do lote

Antes de qualquer gravação automática, o programa apresenta as aulas que pretende registrar:

```text
=== AULAS QUE SERÃO REGISTRADAS ===

1. 13/07/2026 | 03:00 | Inteligência Artificial aplicada ao Marketing
2. 14/07/2026 | 03:00 | Criação de conteúdo com IA

Total de aulas: 2

Deseja registrar TODAS essas aulas? [S/N]:
```

Somente após responder `S` ou `SIM` o processamento começa.

## Como a frequência é processada

Para cada aula, o programa usa a data da aba `Aulas` para localizar automaticamente a coluna correspondente na aba `Lista de Presença`.

Exemplo:

```text
Aula: 07/07/2026
        ↓
Dia: 07
        ↓
Coluna 07 da Lista de Presença
```

Os símbolos são interpretados assim:

```text
. -> PRESENTE
I -> FALTA
```

Antes de aplicar a frequência, o programa compara os alunos encontrados no Excel com os alunos carregados no Diário.

Exemplo:

```text
Encontrados nos dois: 14
Presentes: 11
Faltas: 3
Sem registro/desconhecido: 0
Somente no Excel: 4
Somente no Diário: 0
```

Alunos que existem apenas no Excel podem ser ignorados. Porém, se existir um aluno no Diário sem correspondência no Excel, ou se houver um símbolo de frequência inválido, a aula não é registrada.

## Registro de faltas

Quando o Excel contém `I`, o programa:

1. desmarca o controle de presença;
2. define o detalhamento como `Não compensado`;
3. informa o período da falta com a mesma duração da aula;
4. valida se os valores foram efetivamente aplicados.

Exemplo:

```text
FLORENCE MANIUS -> FALTA | detalhamento=1 | período=03:00
```

Atualmente os códigos de detalhamento conhecidos são:

```text
1 = Não compensado
2 = Dispensado
3 = Compensação
4 = Tolerância por atraso
```

O fluxo automático utiliza `1 - Não compensado` para as faltas da planilha.

## Salvamento automático

Após preencher os dados e aplicar a frequência, o programa aciona automaticamente o botão **Salvar**.

Ele considera a gravação concluída quando:

- o formulário atual deixa de ser exibido;
- o botão para registrar uma nova aula volta a aparecer.

Se isso não ocorrer dentro do tempo esperado, a execução daquela aula é considerada um erro e o lote é interrompido.

## Tratamento de erros

O processamento em lote para no primeiro erro.

Exemplo:

```text
[1/4] Iniciando 07/07/2026
OK

[2/4] Iniciando 08/07/2026
OK

[3/4] Iniciando 13/07/2026
ERRO
```

Resumo:

```text
=== RESULTADO DO PROCESSAMENTO ===

Aulas salvas: 2
OK - 07/07/2026
OK - 08/07/2026

Erros: 1
ERRO - 13/07/2026 | descrição do erro
```

Essa abordagem evita continuar cadastrando aulas quando a página estiver em um estado inesperado.

## Primeiro uso recomendado

Antes de processar todo o curso, faça um teste controlado:

1. faça uma cópia de segurança da planilha;
2. deixe apenas uma aula com `Registrar = SIM`;
3. execute a automação;
4. confira o resumo apresentado no terminal;
5. confirme o processamento;
6. acompanhe o navegador durante o preenchimento;
7. confirme no Diário se data, horas, conteúdo e frequência foram registrados corretamente;
8. somente depois habilite várias aulas no Excel.

## Observações importantes

- Utilize a automação apenas com credenciais e turmas para as quais você tenha autorização de registro.
- Revise a planilha antes de confirmar um lote, pois o processo altera registros reais do Diário.
- A associação dos alunos é feita por nome, não por matrícula.
- Mantenha o navegador visível durante os primeiros testes (`headless=False`).
- O código utiliza `page.pause()` ao final, permitindo inspecionar o estado final do navegador antes de encerrá-lo manualmente.
- Alterações futuras na interface do Diário podem exigir atualização dos seletores utilizados pelo Playwright.

## Problemas comuns

### Arquivo Excel não encontrado

Mensagem semelhante a:

```text
Arquivo Excel não encontrado.
```

Confirme se o arquivo possui exatamente o nome configurado no código e está na mesma pasta do `main.py`.

### Navegador do Playwright não instalado

Execute:

```bash
python -m playwright install chromium
```

### Dia não encontrado na frequência

Exemplo:

```text
Dia 13 não encontrado na planilha de presença.
```

Confirme se o dia correspondente está presente na linha 2 da aba `Lista de Presença`.

### Aluno do Diário não encontrado no Excel

O programa bloqueia o registro para evitar marcar uma frequência incorreta.

Confira principalmente:

- nome do aluno;
- linha correta na planilha;
- presença de espaços ou caracteres diferentes;
- se o aluno realmente pertence à lista utilizada para a turma.

### A aula já existe

A aula deverá aparecer como:

```text
IGNORAR - 07/07/2026 já está registrada.
```

Ela não será incluída no lote.

## Estrutura sugerida do projeto

```text
automacao-diario/
│
├── main.py
├── lista_presenca_com_frequencia.xlsx
├── README.md
│
└── .venv/
```

Para versionamento com Git, recomenda-se não enviar `.venv` ao repositório.

Exemplo de `.gitignore`:

```gitignore
.venv/
__pycache__/
*.pyc
```

Caso a planilha contenha dados reais de alunos, também é recomendável não versioná-la em um repositório público:

```gitignore
*.xlsx
```

## Possíveis melhorias futuras

Algumas evoluções possíveis para o projeto:

- registrar no Excel o status `REGISTRADA`, `IGNORADA` ou `ERRO`;
- gerar arquivo de log de cada execução;
- criar modo `--dry-run` que faça todas as validações sem salvar;
- selecionar unidade e turma por configuração, em vez de interação manual;
- validar o período da turma antes de registrar cada data;
- gerar screenshot automático quando ocorrer erro;
- permitir regras diferentes para faltas parciais, compensação e dispensa;
- separar o projeto em módulos (`diario.py`, `excel.py`, `main.py`).

## Aviso

Este projeto automatiza o preenchimento de informações em um sistema institucional. Antes de utilizar em lote, valide o comportamento com uma aula de teste e confira os registros gerados. A automação não substitui a conferência do responsável pelos lançamentos.
