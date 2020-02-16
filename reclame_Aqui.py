from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
from openpyxl import Workbook
driver = webdriver.Firefox()
arquivo_excel = Workbook()
planilha1 = arquivo_excel.active
planilha1.title = "Reclame Aqui"
titulo = [
('ID', 'STATUS', 'TITULO', 'RECLAMAÇÃO', 'CIDADE', 'DATA', 'RESPOSTA', 'DATA RESPOSTA', 'REPLICA', 'DATA REPLICA')
]
for colunas in titulo:
        planilha1.append(colunas)
for j in range(1,11,1):
    driver.get('https://www.reclameaqui.com.br/empresa/valecard/lista-reclamacoes/?pagina='+str(j)+'')
    print(j)
    driver.stop_client()

    for i in range(1,11,1):
        
        click = driver.find_element_by_xpath('/html/body/ui-view/div/div[4]/div[2]/div[2]/div[2]/div/div[3]/div[3]/div[2]/div/ul[1]/li['+str(i)+']/a/div/p')
        status = driver.find_element_by_xpath('/html/body/ui-view/div/div[4]/div[2]/div[2]/div[2]/div/div[3]/div[3]/div[2]/div/ul[1]/li['+str(i)+']/p[2]/company-complain-status/div').text
        click.click()
        time.sleep(15)

        id_solicitaçao = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[1]/div[2]/div[1]/ul[1]/li[2]/b').text
        titulo = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[1]/div[2]/div[1]/h1').text
        reclamamacao = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[2]/p').text
        cidade = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[1]/div[2]/div[1]/ul[1]/li[1]').text
        data = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[1]/div[2]/div[1]/ul[1]/li[3]').text
        if status == 'Não respondida':
            resposta = ''
            dt_res = ''
            replica = ''
            dt_replica = ''
        elif status == 'Em réplica':
            resposta = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[3]/div[3]/p').text
            dt_res = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[3]/div[3]/p').text
            replica = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[4]/div[3]/p').text
            dt_replica = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[4]/div[1]/p[2]').text
        else:
            resposta = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[3]/div[3]/p').text
            dt_res = driver.find_element_by_xpath('/html/body/ui-view/div[3]/div/div[1]/div[2]/div/div[3]/div[3]/p').text
            replica = ''
            dt_replica = ''

        dados = [
            (id_solicitaçao, status, titulo, reclamamacao, cidade, data, resposta, dt_res, replica,dt_replica )
        ]
        for linha in dados:
            planilha1.append(linha)
        arquivo_excel.save("Reclame2.xlsx")
        driver.back()
        print(i)
        time.sleep(15)
arquivo_excel.save("ReclameTotal.xlsx")