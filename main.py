import time
import unicodedata
from datetime import datetime
from getpass import getpass

from openpyxl import load_workbook

from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError
)


DIARIO_URL = "https://diariofic.sp.senai.br/"

def normalizar_nome(nome):
    if nome is None:
        return ""

    nome = str(nome).strip().upper()

    nome = unicodedata.normalize(
        "NFKD",
        nome
    )

    nome = "".join(
        caractere
        for caractere in nome
        if not unicodedata.combining(caractere)
    )

    return " ".join(
        nome.split()
    )

def obter_credenciais():
    print("\n=== Diário Eletrônico SENAI ===\n")

    nif = input("Digite seu NIF: ").strip()
    senha = getpass("Digite sua senha: ")

    if not nif:
        raise ValueError("O NIF não pode estar vazio.")

    if not senha:
        raise ValueError("A senha não pode estar vazia.")

    return nif, senha


def login_realizado(page):
    try:
        # Primeiro valida se chegou na página /home
        page.wait_for_url(
            "**/home",
            timeout=6000
        )

        # Depois valida um elemento que só existe após o login
        page.get_by_text(
            "Usuário:",
            exact=False
        ).wait_for(
            state="visible",
            timeout=5000
        )

        return True

    except PlaywrightTimeoutError:
        return False


def realizar_login(page, nif, senha, max_tentativas=3):
    print("\nAbrindo Diário Eletrônico...")

    for tentativa in range(1, max_tentativas + 1):

        print(
            f"\nTentativa de login "
            f"{tentativa}/{max_tentativas}"
        )

        # Volta para a página de login a cada tentativa
        page.goto(
            DIARIO_URL,
            wait_until="domcontentloaded",
            timeout=30000
        )

        campo_nif = page.locator(
            'input[name="aIdentificacao"]'
        )

        campo_senha = page.locator(
            'input[name="aSenha"]'
        )

        # Aguarda os campos realmente aparecerem
        campo_nif.wait_for(
            state="visible",
            timeout=10000
        )

        campo_senha.wait_for(
            state="visible",
            timeout=10000
        )

        # Limpa antes de preencher
        campo_nif.fill("")
        campo_senha.fill("")

        campo_nif.fill(nif)
        campo_senha.fill(senha)

        # Validação temporária
        print(
            "NIF preenchido:",
            campo_nif.input_value()
        )

        page.get_by_role(
            "button",
            name="Entrar"
        ).click()

        print("Aguardando resposta do sistema...")

        if login_realizado(page):
            print("\nLogin realizado com sucesso.")
            return True

        print(
            "Não foi possível confirmar o login."
        )

        # Podemos verificar se existe alguma mensagem
        # contendo "incorret"
        mensagem_erro = page.get_by_text(
            "incorret",
            exact=False
        )

        if mensagem_erro.count() > 0:
            try:
                if mensagem_erro.first.is_visible():
                    print(
                        "O sistema informou NIF ou senha incorretos."
                    )
            except Exception:
                pass

        if tentativa < max_tentativas:
            espera = tentativa * 2

            print(
                f"Tentando novamente em "
                f"{espera} segundos..."
            )

            time.sleep(espera)

    print(
        "\nNão foi possível realizar o login "
        "após todas as tentativas."
    )

    return False


def selecionar_unidade_e_abrir_turmas(page):
    print("\nCarregando unidades...")

    select = page.locator("select").first

    select.wait_for(
        state="visible",
        timeout=10000
    )

    opcoes = select.locator("option")

    unidades = []

    for i in range(opcoes.count()):

        opcao = opcoes.nth(i)

        nome = (
            opcao.text_content()
            or ""
        ).strip()

        valor = opcao.get_attribute(
            "value"
        )

        if valor and nome:
            unidades.append({
                "nome": nome,
                "valor": valor
            })

    if not unidades:
        raise RuntimeError(
            "Nenhuma unidade encontrada."
        )

    print("\n=== UNIDADES ===\n")

    for i, unidade in enumerate(
        unidades,
        start=1
    ):
        print(
            f"{i}. {unidade['nome']}"
        )

    while True:

        try:
            escolha = int(
                input(
                    "\nEscolha a unidade: "
                )
            )

            if 1 <= escolha <= len(unidades):
                break

            print(
                "Escolha uma opção válida."
            )

        except ValueError:

            print(
                "Digite apenas o número da unidade."
            )

    unidade = unidades[
        escolha - 1
    ]

    print(
        f"\nSelecionando: {unidade['nome']}"
    )

    select.select_option(
        value=unidade["valor"]
    )

    print("Unidade selecionada.")

    page.get_by_role(
        "link",
        name="Turmas"
    ).click()

    page.wait_for_load_state(
        "domcontentloaded"
    )

    print("Página de turmas aberta.")

def selecionar_turma(page):
    print("\nCarregando turmas...")

    page.wait_for_load_state("networkidle")

    tabela = page.locator("#tbDados")

    tabela.wait_for(
        state="visible",
        timeout=15000
    )

    linhas = tabela.locator("tbody tr")

    # linhas.first.wait_for(
    #     state="visible",
    #     timeout=15000
    # )

    quantidade_linhas = linhas.count()

    print(f"Turmas encontradas na tabela: {quantidade_linhas}")

    if quantidade_linhas == 0:
        raise RuntimeError(
            "Nenhuma turma encontrada na tabela."
        )

    turmas = []

    for i in range(quantidade_linhas):
        linha = linhas.nth(i)

        colunas = linha.locator("td")

        quantidade_colunas = colunas.count()

        if quantidade_colunas < 5:
            continue

        tipo_curso = (
            colunas.nth(0).inner_text()
            or ""
        ).strip()

        span_curso = colunas.nth(1).locator(
            "span[title]"
        )

        if span_curso.count() > 0:
            curso = (
                span_curso.first.get_attribute("title")
                or ""
            ).strip()
        else:
            curso = (
                colunas.nth(1).inner_text()
                or ""
            ).strip()

        codigo_turma = (
            colunas.nth(2).inner_text()
            or ""
        ).strip()

        periodo = (
            colunas.nth(3).inner_text()
            or ""
        ).strip()

        botao_editar = linha.locator(
            "i.fa-edit.edit"
        )

        if botao_editar.count() == 0:
            print(
                f"Aviso: turma {codigo_turma} "
                "não possui botão de editar."
            )
            continue

        id_turma = botao_editar.get_attribute(
            "data-id"
        )

        turmas.append({
            "tipo_curso": tipo_curso,
            "curso": curso,
            "codigo": codigo_turma,
            "periodo": periodo,
            "id": id_turma,
            "linha": linha
        })

    if not turmas:
        raise RuntimeError(
            "Nenhuma turma válida encontrada."
        )

    print("\n=== TURMAS DISPONÍVEIS ===\n")

    for i, turma in enumerate(
        turmas,
        start=1
    ):
        print(
            f"{i}. "
            f"{turma['codigo']} | "
            f"{turma['curso']} | "
            f"{turma['periodo']}"
        )

    while True:
        try:
            escolha = int(
                input(
                    "\nEscolha a turma: "
                )
            )

            if 1 <= escolha <= len(turmas):
                break

            print(
                "Escolha uma opção válida."
            )

        except ValueError:
            print(
                "Digite apenas o número da turma."
            )

    turma_selecionada = turmas[
        escolha - 1
    ]

    print(
        "\nTurma selecionada:"
    )

    print(
        f"Código: {turma_selecionada['codigo']}"
    )

    print(
        f"Curso: {turma_selecionada['curso']}"
    )

    print(
        f"Período: {turma_selecionada['periodo']}"
    )

    print(
        f"ID interno: {turma_selecionada['id']}"
    )

    return turma_selecionada

def abrir_turma_para_edicao(page, turma):
    print(
        f"\nAbrindo turma "
        f"{turma['codigo']}..."
    )

    botao_editar = turma["linha"].locator(
        "i.fa-edit.edit"
    )

    botao_editar.click()

    page.wait_for_load_state(
        "domcontentloaded"
    )

    print(
        "Tela de edição da turma aberta."
    )

def abrir_registro_aula(page):
    print("\nAbrindo registro de aula...")

    botao = page.locator("#NovaAula")

    botao.wait_for(
        state="visible",
        timeout=10000
    )

    botao.click()

    print("Aguardando formulário de registro...")

    page.get_by_text(
        "Diário de Classe",
        exact=False
    ).wait_for(
        state="visible",
        timeout=10000
    )

    print("Formulário de registro de aula aberto.")

def diagnosticar_tabela_alunos(page):
    print("\n=== DIAGNÓSTICO - ALUNOS ===")

    tabelas = page.locator("table")

    print(
        f"Tabelas encontradas: {tabelas.count()}"
    )

    for i in range(tabelas.count()):
        tabela = tabelas.nth(i)

        print(
            f"\n--- TABELA {i} ---"
        )

        try:
            print(
                tabela.inner_text()
            )
        except Exception as erro:
            print(
                f"Não foi possível ler: {erro}"
            )

    print("\n=== FIM DO DIAGNÓSTICO ===")
# def diagnosticar_turmas(page):
#     print("\n===== DIAGNÓSTICO DA PÁGINA DE TURMAS =====")

#     print("URL atual:")
#     print(page.url)

#     print("\nTítulo:")
#     print(page.title())

#     print("\nFrames:")
#     for frame in page.frames:
#         print(" -", frame.url)


#     print("\nQuantidade de #tbDados:")
#     print(
#         page.locator("#tbDados").count()
#     )

#     print("\nQuantidade de tbody:")
#     print(
#         page.locator("#tbDados tbody").count()
#     )

#     print("\nQuantidade de tr:")
#     print(
#         page.locator("#tbDados tbody tr").count()
#     )

#     print("\nQuantidade de td:")
#     print(
#         page.locator("#tbDados tbody td").count()
#     )

#     print("\nQuantidade de ícones editar:")
#     print(
#         page.locator("#tbDados i.edit").count()
#     )

#     print("\n==========================================")

def main():
    try:
        nif, senha = obter_credenciais()

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=False,
                slow_mo=300
            )

            context = browser.new_context()

            page = context.new_page()

            sucesso_login = realizar_login(
                page=page,
                nif=nif,
                senha=senha,
                max_tentativas=3
            )

            # Muito importante:
            # só continua se o login tiver funcionado
            if not sucesso_login:
                print(
                    "\nEncerrando automação."
                )
                browser.close()
                return

            selecionar_unidade_e_abrir_turmas(
                page
            )

            # diagnosticar_turmas(page)

            turma = selecionar_turma(
                page
            )

            abrir_turma_para_edicao(
                page,
                turma
            )

            abrir_registro_aula(
                page
            )

            diagnosticar_tabela_alunos(
                page
            )

            page.pause()

            browser.close()

    except KeyboardInterrupt:
        print(
            "\n\nAutomação cancelada pelo usuário."
        )

    except Exception as erro:
        print(
            f"\nErro: {erro}"
        )


if __name__ == "__main__":
    main()