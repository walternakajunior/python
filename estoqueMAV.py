from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd

#Carrega ChromeDriver
driver = webdriver.Chrome(r"C:\chromedriver.exe")
#URL do MAV
urlMAV = "http://"
#Abre página
driver.get(urlMAV)

#Usuário e senha
usuarioLogin = driver.find_element_by_xpath('//*[@id="USERNAME"]')
usuarioLogin.send_keys('XX-XX')
senhaLogin = driver.find_element_by_xpath('//*[@id="PASSWORD"]')
senhaLogin.send_keys('XXXX')
senhaLogin.send_keys(Keys.RETURN)

#Seleciona sistema MAV
sistema = driver.find_element_by_xpath('//*[@id="usrSysList"]/option[1]')
actionChains = ActionChains(driver)
actionChains.double_click(sistema).perform()

#Altera janela selecionada para a janela aberta
janela1 = driver.window_handles[1]
driver.switch_to.window(janela1)

#Insere sistema COSAL
cosal = driver.find_element_by_xpath('//*[@id="PROX_TELA"]')
cosal.send_keys('COSAL')
ir = driver.find_element_by_xpath('//*[@id="button1"]')
ir.click()

#Cria um loop enquanto não for selecionado uma fonte válida
fonte = ''
while fonte != '1' and fonte != '2':

	#Pergunta qual fonte de dados usar
	print('Digite o número correspondente a fonte de dados você deseja usar?')
	print('1 - Lista')
	print('2 - Planilha')
	fonte = input()

	#Se escolhido lista
	if fonte == '1':

		#Lista de fórmulas a pesquisar
		formulas = ['5C90006007', '5C90004004']

	#Se escolhido planilha
	elif fonte == '2':

		#Carregando informações da planilha
		planilha = pd.read_excel(r"C:\buscaEstoque.xlsx", engine='openpyxl')
		formulas = planilha['formulas']

	#Mensagem de erro quando escolhido outro valor
	else:

		print('Você escolheu uma opção inválida, selecione apenas entre as opções apresentadas!')

#Criando dataframe
saldoFormulas = pd.DataFrame()

for formula in formulas:

    #Altera janela selecionada para a janela aberta
    janela1 = driver.window_handles[1]
    driver.switch_to.window(janela1)
    #Altera para iframe da janela
    frame = driver.find_element_by_xpath('/html/body/form/iframe')
    driver.switch_to.frame(frame)
    #Insere código da fórmula
    input = driver.find_element_by_xpath('//*[@id="MATERIAL"]')
    input.clear()
    input.send_keys(formula)
    input.send_keys(Keys.ENTER)
    #Pega valores 
    saldoFisico = driver.find_element_by_xpath('//*[@id="QUA_SALFIS"]').get_attribute('value')
    entradasMes = driver.find_element_by_xpath('//*[@id="QUA_ENTACU"]').get_attribute('value')
    saidasMes = driver.find_element_by_xpath('//*[@id="QUA_SAIDAS"]').get_attribute('value')
    saldoFinanceiro = driver.find_element_by_xpath('//*[@id="VL_SALFIN"]').get_attribute('value')

    #Salvando registro
    saldoFormulas = saldoFormulas.append({'Fórmula': formula, 'Saldo Físico': saldoFisico, 'Saldo Financeiro': saldoFinanceiro, 'Entradas no Mês': entradasMes, 'Saídas no Mês': saidasMes}, ignore_index=True)

saldoFormulas.to_excel(r"C:\saldoFormulas.xlsx")
print("Finalizado com sucesso!")