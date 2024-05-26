from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import win32com.client as win32
import pandas as pd

navegador = webdriver.Chrome()

tabela_produtos = pd.read_excel(
    r'C:\Users\natha\Projeto 2 - Automação Web - Aplicação de Mercado de Trabalho\buscas.xlsx')

def verificar_tem_termos_banidos(lista_termos_banidos, nome_texto):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome_texto:
            tem_termos_banidos = True
    return tem_termos_banidos

def verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome_texto):
    tem_todos_os_termos = True
    for nomes_produto in lista_termos_nome_produto:
        if nomes_produto not in nome_texto:
            tem_todos_os_termos = False
    return tem_todos_os_termos

def busca_google_shopping(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")

    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    navegador.get('https://www.google.com/')
    navegador.find_element(By.NAME, 'q').send_keys(produto, Keys.ENTER)

    elementos = navegador.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in elementos:
        if 'Shopping' in item.text:
            item.click()
            break

    time.sleep(3)
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'sh-dlr__list-result')

    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text.lower()
        termos_banidos_presente = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        tem_todos_os_termos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)

        if not termos_banidos_presente and tem_todos_os_termos:
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'T14wmb').text
                preco_texto = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco_texto = float(preco_texto)
                if preco_minimo <= preco_texto <= preco_maximo:
                    link = resultado.find_element(By.CLASS_NAME, 'aULzUe').get_attribute('href')
                    lista_ofertas.append((nome, preco_texto, link))
            except Exception as e:
                print(f"Erro ao processar resultado: {e}")
                continue

    return lista_ofertas

def busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    navegador.get('https://www.buscape.com.br')
    navegador.find_element(By.XPATH,
                           '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(
        produto)
    navegador.find_element(By.CLASS_NAME, 'AutoCompleteStyle_submitButton__GkxPO').click()

    time.sleep(3)  # Aguarde o carregamento da página
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'Paper_Paper__HIHv0')

    for resultado in lista_resultados:
        nome = resultado.find_element(By.TAG_NAME, 'h2').text.lower()
        termos_banidos_presente = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        tem_todos_os_termos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)

        if not termos_banidos_presente and tem_todos_os_termos:
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__Zxam2').text
                preco_texto = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco_texto = float(preco_texto)
                if preco_minimo <= preco_texto <= preco_maximo:
                    link = resultado.find_element(By.CLASS_NAME, 'SearchCard_ProductCard_Inner__7JhKb').get_attribute(
                        'href')
                    lista_ofertas.append((nome, preco_texto, link))
            except Exception as e:
                print(f"Erro ao processar resultado: {e}")
                continue

    return lista_ofertas

tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(navegador, produto, termos_banidos, preco_minimo,
                                                          preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])

    lista_ofertas_buscape = busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])

tabela_ofertas.to_excel('Ofertas.xlsx', index=False)

if len(tabela_ofertas) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'nathan.castro2022@outlook.com'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f'''
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada:</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Att.,</p>
    '''
    mail.Send()

navegador.quit()










