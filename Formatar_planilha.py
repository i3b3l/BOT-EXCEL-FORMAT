# ============================================================
# FEITO COM <3 POR IZABEL GALVAO :3
# ATUALIZACOES NO MEU GITHUB: https://github.com/i3b3l
# ============================================================

import os
import sys
import re
import unicodedata
from datetime import datetime, date
from decimal import Decimal, InvalidOperation

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


# ============================================================
# CONFIGURAÃ‡ÃƒO
# ============================================================

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
arquivo_entrada = os.path.join(BASE_DIR, 'Pagamentos.xlsx')
arquivo_saida = os.path.join(BASE_DIR, 'Pagamentos_Tratado.xlsx')


# ============================================================
# FUNÃ‡Ã•ES DE LIMPEZA
# ============================================================

def limpar_cpf(val):
    if pd.isna(val):
        return ''

    digits = re.sub(r'\D', '', str(val))
    return digits.zfill(11) if digits else ''


def limpar_valor(val):
    """
    Padroniza o valor com duas casas decimais e vÃ­rgula.

    Exemplos:
    R$ 222,50 -> 222,50
    222.5     -> 222,50
    1.234,56  -> 1234,56
    """
    if pd.isna(val):
        return ''

    texto = str(val).replace('R$', '').strip()

    if not texto:
        return ''

    texto = texto.replace(' ', '')

    if '.' in texto and ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    elif ',' in texto:
        texto = texto.replace(',', '.')

    try:
        numero = Decimal(texto)
        return f'{numero:.2f}'.replace('.', ',')
    except (InvalidOperation, ValueError):
        return str(val).replace('R$', '').strip()


def normalizar_texto(texto):
    """
    Remove acentos, espaÃ§os e sÃ­mbolos para facilitar comparaÃ§Ãµes.

    Exemplos:
    'PROCESSO NÂº' -> 'PROCESSON'
    'GESTÃƒO'      -> 'GESTAO'
    """
    texto = '' if texto is None else str(texto)
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    return re.sub(r'[^A-Z0-9]', '', texto.upper())


def limpar_banco(val):
    """
    Converte o conteÃºdo da coluna BANCO para o cÃ³digo bancÃ¡rio de 3 dÃ­gitos.

    Reconhece cÃ³digos, nomes e apelidos comuns.

    Exemplos:
    Brasil / Banco do Brasil -> 001
    Caixa EconÃ´mica          -> 104
    Nu / Nubank              -> 260
    Sicoob                   -> 756
    C6 / Banco C6            -> 336
    """
    if pd.isna(val):
        return ''

    texto_original = str(val).strip()

    if not texto_original:
        return ''

    texto = normalizar_texto(texto_original)

    # IMPORTANTE:
    # A verificaÃ§Ã£o pelo nome vem antes da extraÃ§Ã£o de nÃºmeros.
    # Isso evita, por exemplo, interpretar "Banco C6" como cÃ³digo "006".
    bancos_por_nome = {
        # Banco do Brasil
        'BANCODOBRASIL': '001',
        'BB': '001',
        'BRASIL': '001',

        # Caixa EconÃ´mica Federal
        'CAIXAECONOMICAFEDERAL': '104',
        'CAIXAECONOMICA': '104',
        'CAIXA': '104',

        # Nubank / Nu Pagamentos
        'NUPAGAMENTOS': '260',
        'NUBANK': '260',
        'NU260': '260',

        # Sicoob
        'SICOOB': '756',
        'SICOOB756': '756',

        # C6 Bank
        'BANCOC6': '336',
        'C6BANK': '336',

        # Outros bancos comuns
        'BRADESCO': '237',
        'ITAU': '341',
        'ITAUNIBANCO': '341',
        'SANTANDER': '033',
        'BANCOINTER': '077',
        'INTER': '077',
        'PICPAY': '380',
        'MERCADOPAGO': '323',
    }

    # Primeiro tenta reconhecer pelo nome.
    for nome, codigo in bancos_por_nome.items():
        # Para nomes muito curtos, exige correspondÃªncia exata.
        if len(nome) <= 3:
            if texto == nome:
                return codigo
        elif nome in texto:
            return codigo

    # Caso especÃ­fico muito comum: "Nu 260".
    if texto.startswith('NU') and '260' in texto:
        return '260'

    # Se jÃ¡ houver um cÃ³digo explÃ­cito de 3 dÃ­gitos, preserva.
    match_3 = re.search(r'(?<!\d)(\d{3})(?!\d)', texto_original)

    if match_3:
        return match_3.group(1)

    # Ãšltimo recurso: usa qualquer sequÃªncia numÃ©rica encontrada.
    match = re.search(r'\d+', texto_original)

    if match:
        return match.group(0).zfill(3)

    # Se nÃ£o reconhecer, mantÃ©m o conteÃºdo original.
    # Assim nenhum dado Ã© apagado silenciosamente.
    return texto_original


def limpar_data(val):
    """Converte datas seriais do Excel ou datas textuais para dd/mm/aaaa."""
    if pd.isna(val):
        return ''
    texto = str(val).strip()
    if not texto:
        return ''

    # Remove o horario quando a data chega como texto com horario.
    # Exemplos: 2026-09-25 00:00:00 e 2026-09-25T00:00:00.
    if ' ' in texto or 'T' in texto:
        texto = texto.split('T', 1)[0].split(' ', 1)[0]

    if re.fullmatch(r'\d+(?:\.\d+)?', texto):
        try:
            serial = float(texto)
            if 20000 <= serial <= 80000:
                convertido = from_excel(serial)
                if isinstance(convertido, (datetime, date)):
                    return convertido.strftime('%d/%m/%Y')
        except (ValueError, OverflowError, TypeError):
            pass
    for formato in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%Y/%m/%d', '%d-%m-%Y', '%d-%m-%y', '%m/%d/%Y'):
        try:
            return datetime.strptime(texto, formato).strftime('%d/%m/%Y')
        except ValueError:
            continue
    return texto


def limpar_alfanumerico(val):
    if pd.isna(val):
        return ''

    return re.sub(r'[^a-zA-Z0-9]', '', str(val))


# ============================================================
# LOCALIZAÃ‡ÃƒO DE COLUNAS
# ============================================================

def encontrar_coluna_por_prefixo(df, prefixo):
    """
    Encontra uma coluna pelo inÃ­cio do cabeÃ§alho.

    Exemplo:
    PROCESSO
    PROCESSO NÂº
    PROCESSO NÂ°
    PROCESSO N.Âº
    PROCESSO NUMERO

    Todos sÃ£o reconhecidos pelo prefixo PROCESSO.
    """
    prefixo_normalizado = normalizar_texto(prefixo)

    for coluna in df.columns:
        nome_normalizado = normalizar_texto(coluna)

        if nome_normalizado.startswith(prefixo_normalizado):
            return coluna

    return None


def encontrar_coluna(df, *nomes):
    """
    Encontra uma coluna mesmo que haja diferenÃ§a de acento,
    espaÃ§os ou sÃ­mbolos no cabeÃ§alho.
    """
    cabecalhos = {
        normalizar_texto(coluna): coluna
        for coluna in df.columns
    }

    for nome in nomes:
        chave = normalizar_texto(nome)

        if chave in cabecalhos:
            return cabecalhos[chave]

    return None


# ============================================================
# PREENCHIMENTO DOS BLOCOS DA ABA ENTRADA
# ============================================================

def preencher_colunas_repetidas(df):
    """
    Preenche para baixo o Ãºltimo valor vÃ¡lido nas colunas:

    - UNIDADE / UNIDADE GESTORA
    - GESTÃƒO
    - PROCESSO
    - EMPENHO
    - DATA
    - OBSERVAÃ‡ÃƒO

    Quando aparece um novo valor no meio da planilha, o novo valor passa
    a valer para as prÃ³ximas linhas.

    Linhas totalmente vazias no fim da planilha NÃƒO sÃ£o preenchidas.
    """
    estatisticas = {}

    # Marca somente as linhas que realmente possuem algum dado.
    linhas_com_dados = (
        df.replace(r'^\s*$', pd.NA, regex=True)
        .notna()
        .any(axis=1)
    )

    colunas_para_preencher = [
        encontrar_coluna(
            df,
            'UNIDADE',
            'UNIDADE GEST',
            'UNIDADE GESTORA'
        ),
        encontrar_coluna(
            df,
            'GESTAO',
            'GESTÃƒO'
        ),
        encontrar_coluna_por_prefixo(
            df,
            'PROCESSO'
        ),
        encontrar_coluna(
            df,
            'EMPENHO'
        ),
        encontrar_coluna(
            df,
            'DATA'
        ),
        encontrar_coluna(
            df,
            'OBSERVACAO',
            'OBSERVAÃ‡ÃƒO'
        ),
    ]

    for coluna_real in colunas_para_preencher:
        if coluna_real is None:
            continue

        serie = df[coluna_real].replace(
            r'^\s*$',
            pd.NA,
            regex=True
        )

        vazios_antes = serie.isna() & linhas_com_dados

        preenchida = serie.ffill()

        df.loc[
            linhas_com_dados,
            coluna_real
        ] = preenchida[linhas_com_dados]

        # Conta somente cÃ©lulas que estavam vazias e receberam valor.
        preenchidos = (
            vazios_antes
            & preenchida.notna()
        ).sum()

        estatisticas[str(coluna_real)] = int(preenchidos)

    return df, estatisticas


# ============================================================
# FORMATAÃ‡ÃƒO DA ABA ENTRADA
# ============================================================

def formatar_entrada(df):
    """
    Executa todos os tratamentos da aba ENTRADA.
    """
    df, preenchimentos = preencher_colunas_repetidas(df)

    coluna_cpf = encontrar_coluna(
        df,
        'CPF'
    )

    coluna_valor = encontrar_coluna(
        df,
        'VALOR'
    )

    coluna_banco = encontrar_coluna(
        df,
        'BANCO'
    )

    coluna_agencia = encontrar_coluna(
        df,
        'AGENCIA',
        'AGÃŠNCIA'
    )

    coluna_conta = encontrar_coluna(
        df,
        'CONTACORRENTE',
        'CONTA CORRENTE',
        'CONTA C',
        'CONTA'
    )

    # Corrige datas antes de converter todos os campos para texto.
    for nome_data in ('DATA', 'DATA DO PAGAMENTO', 'DATA PAGAMENTO'):
        coluna_data = encontrar_coluna(df, nome_data)
        if coluna_data is not None:
            df[coluna_data] = df[coluna_data].apply(limpar_data)

    estatisticas = {
        'linhas': len(df),
        'preenchimentos': preenchimentos,
        'cpf': 0,
        'valor': 0,
        'banco': 0,
        'agencia': 0,
        'conta': 0,
        'bancos_nao_reconhecidos': [],
    }

    if coluna_cpf is not None:
        estatisticas['cpf'] = int(
            df[coluna_cpf].fillna('').astype(str).str.strip().ne('').sum()
        )
        df[coluna_cpf] = df[coluna_cpf].apply(limpar_cpf)

    if coluna_valor is not None:
        estatisticas['valor'] = int(
            df[coluna_valor].fillna('').astype(str).str.strip().ne('').sum()
        )
        df[coluna_valor] = df[coluna_valor].apply(limpar_valor)

    if coluna_banco is not None:
        bancos_originais = df[coluna_banco].copy()

        df[coluna_banco] = df[coluna_banco].apply(limpar_banco)

        estatisticas['banco'] = int(
            df[coluna_banco].fillna('').astype(str).str.strip().ne('').sum()
        )

        # Lista valores de banco que continuaram como nome,
        # indicando que nÃ£o foram reconhecidos como cÃ³digo.
        nao_reconhecidos = []

        for original, tratado in zip(
            bancos_originais.tolist(),
            df[coluna_banco].tolist()
        ):
            if pd.isna(original):
                continue

            original_txt = str(original).strip()
            tratado_txt = str(tratado).strip()

            if not original_txt:
                continue

            if not re.fullmatch(r'\d{3}', tratado_txt):
                nao_reconhecidos.append(original_txt)

        estatisticas['bancos_nao_reconhecidos'] = sorted(
            set(nao_reconhecidos)
        )

    if coluna_agencia is not None:
        estatisticas['agencia'] = int(
            df[coluna_agencia]
            .fillna('')
            .astype(str)
            .str.strip()
            .ne('')
            .sum()
        )

        df[coluna_agencia] = df[coluna_agencia].apply(
            limpar_alfanumerico
        )

    if coluna_conta is not None:
        estatisticas['conta'] = int(
            df[coluna_conta]
            .fillna('')
            .astype(str)
            .str.strip()
            .ne('')
            .sum()
        )

        df[coluna_conta] = df[coluna_conta].apply(
            limpar_alfanumerico
        )

    # GARANTIA EXTRA:
    # Tudo na aba ENTRADA serÃ¡ convertido para string antes
    # de ser gravado no Excel.
    for coluna in df.columns:
        df[coluna] = df[coluna].apply(
            lambda x: '' if pd.isna(x) else str(x).strip()
        )

    return df, estatisticas


# ============================================================
# ABAS QUE DEVEM FICAR SOMENTE COM O CABEÃ‡ALHO
# ============================================================

ABAS_PARA_LIMPAR = {
    'DESPESACERTIFICADA',
    'DESPESASCERTIFICADAS',
    'NOTADELIQUIDACAO',
    'NOTASDELIQUIDACAO',
    'PREPARACAODEPAGAMENTO',
    'PREPARACOESDEPAGAMENTO',
    'ORDEMBANCARIA',
    'ORDENSBANCARIAS',
    'SAIDA',
}


# ============================================================
# SAÃDA VISUAL DO TERMINAL
# ============================================================

def linha_terminal():
    print('=' * 62)


def imprimir_inicio():
    print()
    linha_terminal()
    print(' BOT PAG - FORMATADOR DE PLANILHA')
    linha_terminal()
    print(f' Arquivo de entrada : {arquivo_entrada}')
    print(f' Arquivo de saÃ­da   : {arquivo_saida}')
    print('-' * 62)


def imprimir_resumo(estatisticas_entrada, abas_limpas):
    print()
    linha_terminal()
    print(' PROCESSAMENTO CONCLUÃDO COM SUCESSO')
    linha_terminal()

    if estatisticas_entrada:
        print(f" Linhas processadas na ENTRADA: {estatisticas_entrada['linhas']}")
        print()

        print(' Preenchimento automÃ¡tico por blocos:')

        preenchimentos = estatisticas_entrada['preenchimentos']

        if preenchimentos:
            for coluna, quantidade in preenchimentos.items():
                print(f'   - {coluna}: {quantidade} cÃ©lula(s) preenchida(s)')
        else:
            print('   - Nenhuma cÃ©lula precisou ser preenchida.')

        print()
        print(' Tratamentos realizados:')
        print(f"   - CPF: {estatisticas_entrada['cpf']} registro(s)")
        print(f"   - VALOR: {estatisticas_entrada['valor']} registro(s)")
        print(f"   - BANCO: {estatisticas_entrada['banco']} registro(s)")
        print(f"   - AGÃŠNCIA: {estatisticas_entrada['agencia']} registro(s)")
        print(f"   - CONTA: {estatisticas_entrada['conta']} registro(s)")
        print('   - Todos os dados da ENTRADA gravados como TEXTO.')

        bancos_nao_reconhecidos = (
            estatisticas_entrada['bancos_nao_reconhecidos']
        )

        if bancos_nao_reconhecidos:
            print()
            print(' ATENÃ‡ÃƒO - banco(s) nÃ£o reconhecido(s):')

            for banco in bancos_nao_reconhecidos:
                print(f'   - {banco}')

            print(
                '   O nome original foi mantido para evitar perda de dados.'
            )

    if abas_limpas:
        print()
        print(' Abas mantidas somente com o cabeÃ§alho:')

        for nome in abas_limpas:
            print(f'   - {nome}')

    print()
    print(f' Arquivo gerado: {arquivo_saida}')
    linha_terminal()
    print()


# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================

def main():
    imprimir_inicio()

    if not os.path.exists(arquivo_entrada):
        linha_terminal()
        print(' ERRO: arquivo de entrada nÃ£o encontrado.')
        print(f' Coloque "{arquivo_entrada}" na mesma pasta do programa.')
        linha_terminal()
        return

    estatisticas_entrada = None
    abas_limpas = []

    try:
        excel_file = pd.ExcelFile(arquivo_entrada)

        with pd.ExcelWriter(
            arquivo_saida,
            engine='openpyxl'
        ) as writer:

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(
                    arquivo_entrada,
                    sheet_name=sheet_name,
                    dtype=str
                )

                nome_aba = normalizar_texto(sheet_name)

                if nome_aba == 'ENTRADA':
                    df, estatisticas_entrada = formatar_entrada(df)

                elif nome_aba in ABAS_PARA_LIMPAR:
                    colunas_validas = [
                        coluna
                        for coluna in df.columns
                        if not str(coluna).startswith('Unnamed:')
                    ]

                    df = df[colunas_validas]

                    # MantÃ©m somente a primeira linha:
                    # os cabeÃ§alhos das colunas.
                    df = df.iloc[0:0].copy()

                    abas_limpas.append(sheet_name)

                # Abas nÃ£o previstas sÃ£o copiadas sem alteraÃ§Ã£o.
                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

        # ====================================================
        # SEGUNDA CAMADA DE GARANTIA
        # ====================================================
        #
        # ForÃ§a todas as cÃ©lulas preenchidas da aba ENTRADA
        # para o tipo TEXTO do Excel.
        #
        wb = load_workbook(arquivo_saida)

        for ws in wb.worksheets:
            nome_aba = normalizar_texto(ws.title)

            if nome_aba == 'ENTRADA':
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if (
                            cell.value is not None
                            and str(cell.value).strip() != ''
                        ):
                            cell.value = str(cell.value)
                            cell.number_format = '@'

        wb.save(arquivo_saida)

        imprimir_resumo(
            estatisticas_entrada,
            abas_limpas
        )

    except PermissionError:
        print()
        linha_terminal()
        print(' ERRO DE PERMISSÃƒO')
        linha_terminal()
        print(
            f' Feche o arquivo "{arquivo_saida}" no Excel e tente novamente.'
        )
        print()

    except ImportError as erro:
        print()
        linha_terminal()
        print(' ERRO AO CARREGAR UMA DEPENDÃŠNCIA')
        linha_terminal()
        print(f' {erro}')
        print()
        print(
            ' O cÃ³digo nÃ£o chegou a processar a planilha.'
        )
        print()

    except Exception as erro:
        print()
        linha_terminal()
        print(' ERRO DURANTE O PROCESSAMENTO')
        linha_terminal()
        print(f' Tipo: {type(erro).__name__}')
        print(f' Detalhes: {erro}')
        print()


if __name__ == '__main__':
    main()

