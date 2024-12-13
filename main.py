import json
import math
import os
import time
from datetime import datetime, timedelta

import openpyxl
from openpyxl.utils import get_column_letter

from login_canaime import Login

url_pesquisa = (
    'https://canaime.com.br/sgp2rr/areas/unidades/pesquisa_resultadoVULGO.php?'
    'busca1=nome&busca2=&busca3=SAIDA&Submit2=Pesquisar'
)
url_certidao = 'https://canaime.com.br/sgp2rr/areas/impressoes/UND_CertidaoCarceraria.php?id_cad_preso='


def lista_ids_saida(nome_arquivo: str = "lista_ids_saida.json") -> list:
    """
    Gera uma lista de IDs de presos a partir da página de pesquisa.

    Args:
        nome_arquivo (str): Nome do arquivo JSON para salvar os resultados.

    Returns:
        list: Lista contendo o objeto 'page' e a lista de presos com seus IDs e nomes.
    """
    try:
        page = Login(test=False)
        page.goto(url_pesquisa)
    except Exception as e:
        print("Erro ao acessar a página de pesquisa:", e)
        return []

    try:
        qtd_presos_replace = ' REEDUCANDO(S) CADASTRADO(S)'
        qtd_presos_text = page.locator('.tituloVermelho10').text_content()
        qtd_presos = int(qtd_presos_text.replace(qtd_presos_replace, '').strip())
    except Exception as e:
        print("Erro ao extrair a quantidade de presos:", e)
        return []

    total_paginas = math.ceil(qtd_presos / 10)
    resultados = []

    pagina = 0
    tempos = []  # Lista para armazenar o tempo de cada iteração

    while pagina < total_paginas:
        inicio = time.time()  # Marca o início da iteração
        paginas_restantes = total_paginas - (pagina + 1)
        print(f"Acessando página {pagina + 1} de {total_paginas} páginas, faltam {paginas_restantes} páginas...")

        url_inicial = (
            f'https://canaime.com.br/sgp2rr/areas/unidades/pesquisa_resultadoVULGO.php?pageNum_rsPreso={str(pagina)}'
            f'&totalRows_rsPreso={str(qtd_presos)}&busca1=nome&busca2=&busca3=SAIDA&Submit2=Pesquisar'
        )

        try:
            page.goto(url_inicial)
            elementos = page.locator('a.tituloAzul')
            total_elementos = elementos.count()
        except Exception as e:
            print("Erro ao acessar a página ou localizar elementos:", e)
            break

        for i in range(total_elementos):
            try:
                elemento = elementos.nth(i)
                href = elemento.get_attribute('href')
                if href and "id_cad_preso=" in href:
                    id_cad_preso = href.split('=')[-1]  # Pega os números após "="
                else:
                    id_cad_preso = None

                nome = elemento.text_content().strip()
                resultados.append({'id': id_cad_preso, 'nome': nome})
            except Exception as e:
                print(f"Erro ao processar elemento {i} da página {pagina + 1}:", e)

        pagina += 1
        fim = time.time()
        tempos.append(fim - inicio)  # Adiciona o tempo gasto para a iteração à lista
        if len(tempos) > 0:
            tempo_medio = sum(tempos) / len(tempos)
            tempo_estimado = tempo_medio * paginas_restantes

            # Converte o tempo restante para horas, minutos e segundos
            estimativa_timedelta = timedelta(seconds=tempo_estimado)
            horas, resto = divmod(estimativa_timedelta.seconds, 3600)
            minutos, segundos = divmod(resto, 60)

            print(f"Tempo estimado restante: {horas} horas, {minutos} minutos e {segundos} segundos.")
            print()

    try:
        with open(nome_arquivo, 'w', encoding='utf-8') as f:
            json.dump(resultados, f, ensure_ascii=False, indent=4)
        print(f"Lista de IDs salva no arquivo {nome_arquivo}")
    except Exception as e:
        print("Erro ao salvar o arquivo JSON:", e)

    return [page, resultados]


def busca_dados(lista_arquivo: str = "lista_ids_saida.json") -> list:
    """
    Verifica se a lista de IDs existe. Caso contrário, gera a lista com lista_ids_saida.

    Args:
        lista_arquivo (str): Nome do arquivo JSON contendo a lista de IDs.

    Returns:
        list: Lista contendo o objeto 'page' e a lista de presos.
    """
    if not os.path.exists(lista_arquivo):
        print(f"Arquivo {lista_arquivo} não encontrado. Gerando lista de IDs...")
        lista_de_ids = lista_ids_saida(lista_arquivo)
        return lista_de_ids
    else:
        print(f"Arquivo {lista_arquivo} encontrado. Carregando lista de IDs...")
        try:
            with open(lista_arquivo, 'r', encoding='utf-8') as f:
                lista_de_ids = json.load(f)
            page = Login(test=False)
            return [page, lista_de_ids]
        except Exception as e:
            print("Erro ao carregar o arquivo JSON:", e)
            # Em caso de erro, tenta gerar a lista novamente
            return lista_ids_saida(lista_arquivo)


def busca_datas(lista: list) -> list:
    """
    Acessa as páginas de certidão para cada preso da lista, busca a última data e unidade,
    e verifica se a unidade é 'PAMC' no ano de 2024.

    Args:
        lista (list): Lista contendo o objeto 'page' e a lista de presos.

    Returns:
        list: Lista de dicionários com informações dos presos filtrados.
    """
    try:
        page, lista = lista
    except ValueError:
        print("Lista de IDs inválida. Certifique-se de que a função busca_dados retornou a lista corretamente.")
        return []

    lista_presos_saida = []
    tempos = []  # Lista para armazenar o tempo de cada iteração

    for index, item in enumerate(lista):
        inicio = time.time()  # Marca o início da iteração
        total_items = len(lista)
        items_restantes = total_items - (index + 1)
        print(f"Acessando preso {index + 1} de {total_items}, faltam {items_restantes} presos...")

        try:
            # Acessa a URL específica do preso
            page.goto(url_certidao + item['id'])

            lista_unit = page.locator('table+ table td.titulobk:nth-child(1)').all_text_contents()
            lista_datas = page.locator('table+ table .titulobk+ .titulobk:nth-child(2)').all_text_contents()

            # Itera sobre as unidades e datas
            for index_unit in reversed(range(len(lista_unit))):
                if lista_unit[index_unit] != "SAIDA":
                    ultima_unit = lista_unit[index_unit].strip()
                    ultima_data = lista_datas[index_unit].strip()
                    # Verifica se atende aos critérios
                    data_convertida = datetime.strptime(ultima_data, "%d/%m/%Y")
                    if data_convertida.year == 2024:
                        lista_presos_saida.append({
                            'Código': item['id'],
                            'Preso': item['nome'],
                            'Unidade': ultima_unit,
                            'Data': ultima_data
                        })
                    break
        except Exception as e:
            print(f"Erro ao processar dados do preso {item.get('nome', '')} (ID: {item.get('id', '')}):", e)

        fim = time.time()  # Marca o fim da iteração
        tempos.append(fim - inicio)  # Adiciona o tempo da iteração

        if len(tempos) > 0:
            tempo_medio = sum(tempos) / len(tempos)
            tempo_estimado = tempo_medio * items_restantes

            # Converte o tempo restante para horas, minutos e segundos
            estimativa_timedelta = timedelta(seconds=tempo_estimado)
            horas, resto = divmod(estimativa_timedelta.seconds, 3600)
            minutos, segundos = divmod(resto, 60)

            print(f"Tempo estimado restante: {horas} horas, {minutos} minutos e {segundos} segundos.")
            print()

    return lista_presos_saida


def salvar_excel(lista_presos_saida: list, nome_arquivo: str) -> None:
    """
    Salva a lista de presos com suas informações em um arquivo Excel.

    Args:
        lista_presos_saida (list): Lista de dicionários contendo as informações dos presos.
        nome_arquivo (str): Nome do arquivo Excel a ser salvo.
    """
    try:
        # Cria um novo workbook e adiciona uma planilha
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Presos Saída"

        # Define os cabeçalhos
        cabecalhos = ["Código", "Preso", "Unidade", "Data"]
        ws.append(cabecalhos)

        # Adiciona os dados
        for preso in lista_presos_saida:
            ws.append([
                preso.get("Código", ""),
                preso.get("Preso", ""),
                preso.get("Unidade", ""),
                preso.get("Data", "")
            ])

        # Ajusta a largura das colunas
        for col_num, col_title in enumerate(cabecalhos, 1):
            col_letter = get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = max(len(col_title), 20)

        # Salva o arquivo
        wb.save(nome_arquivo)
        print(f"Arquivo Excel salvo como {nome_arquivo}")
    except Exception as e:
        print("Erro ao salvar o arquivo Excel:", e)


if __name__ == '__main__':
    dados_ids = busca_dados()
    if dados_ids and len(dados_ids) == 2:
        dados_presos = busca_datas(dados_ids)
        if dados_presos:
            salvar_excel(dados_presos, "presos_saida.xlsx")
    else:
        print("Não foi possível obter a lista de IDs para prosseguir.")
