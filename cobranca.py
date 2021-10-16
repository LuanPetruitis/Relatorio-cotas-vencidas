# -*- encoding: utf-8 -*-

import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
import json
import time
from pymongo import MongoClient
import datetime
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import locale
from openpyxl import Workbook



class Cobranca():
	def __init__(self):
		self.url = "URL da administradora"

		# Configuração para abrir o navegador
		option = Options()
		option.headless = True
		self.driver = webdriver.Firefox(options=option)
		self.driver.get(self.url)
		self.todas_mensalidades = []
		self.creditos_vencidos = []
		self.teste = {}
		self.apartamentos_credito = {}
		self.historico= {}
		time.sleep(0.5)


	def login(self):
		self.driver.find_element_by_xpath("//a[@class='setCookies btn btn-primary']").click()
		time.sleep(2)
		login = self.driver.find_element_by_id('UsuarioDcEmail')
		login.send_keys('usuário')
		senha = self.driver.find_element_by_id('UsuarioDcSenha')
		senha.send_keys('senha')
		print("Loguei")
		senha.submit()
		time.sleep(2)
		self.driver.find_element_by_xpath("//div[@class='row']").click()
		# self.driver.find_element_by_xpath("//img[@class='media-object']").click()
		time.sleep(1.5)
		self.driver.find_element_by_xpath("//div[@class='cookieButton']").click()
		print("Passei do cookie")


	def pesquisa(self):
		time.sleep(2)
		self.driver.get('')
		time.sleep(2)
		self.driver.find_element_by_xpath("//a[@title='Filtrar']").click()
		time.sleep(2)
		# filtro = Select(self.driver.find_element_by_xpath("//select[@id='LancamentoFilterStLancamento']"))
		# filtro.select_by_visible_text('Pendentes')
		self.driver.find_element_by_xpath("//div[@id='s2id_LancamentoFilterStLancamento']//span[@class='select2-chosen']").click()
		time.sleep(2)
		self.driver.find_element_by_xpath("//li[@class='select2-results-dept-0 select2-result select2-result-selectable']").click()
		
		time.sleep(2)
		self.driver.find_element_by_xpath("//button[@class='btn btn-sm btn-primary']").click()

		time.sleep(2)
		self.driver.find_element_by_xpath("//div[@id='s2id_autogen15']//span[@class='select2-chosen']").click()
		time.sleep(2)
		self.driver.find_element_by_xpath("//ul[@class='select2-results']//div[contains(text(), '100')]").click()
		time.sleep(10)

		numero_paginas = self.driver.find_element_by_xpath("//div[@id='tblLancamentos_paginate']")
		html_content = numero_paginas.get_attribute('outerHTML')
		time.sleep(3)
		
		soup = BeautifulSoup(html_content, 'html.parser')

		# encontrar o ultimo número
		ultima_pagina = int(soup.find_all('a')[-1].getText())
		# ultima_pagina = 1
		count = 1

		if ultima_pagina > 1:
			while (count <= ultima_pagina):
				time.sleep(2)
				print(f"Pagina {count}")
				self.driver.find_element_by_xpath(f"//a[contains(text(), {count})]").click()
				time.sleep(15)
				self.pegar_dados()
				count = count + 1
		else:
			self.pegar_dados()

		# print(json.dumps(self.todas_mensalidades, indent=4, sort_keys=True))
	

	def pegar_dados(self):
			time.sleep(2)

			element = self.driver.find_element_by_xpath("//table[@id='tblLancamentos']")

			# Traz todo o html da parte que navegamos
			html_content = element.get_attribute('outerHTML')

			# Tratar os dados
			soup = BeautifulSoup(html_content, 'html.parser')
			table = soup.find(name='table')

			# Pandas trata o html
			# Traz um array e coloca o head para trazer todos registros
			df_full = pd.read_html(str(table))[0]

			# Para ver o nome das colunas para extrair os dados
			print(df_full.columns)

			# Organiza os dados para colocar em um dicionário
			df = df_full[['Situação', 'Competência', 'Vencimento', 'Pagamento', 'Crédito', 'Bloco', 'Unidade', 'Item descrição', 'R$ Valor', 'R$ Pago']]

			df.columns = ['situacão', 'competencia', 'vencimento', 'pagamento', 'credito', 'bloco', 'unidade', 'itemDescricao', 'valor', 'pago']

			# Monta o dicionário
			self.teste['contas'] =  df.to_dict('records')
			self.teste['contas'] = json.dumps(self.teste['contas'], indent=4, sort_keys=True).replace('NaN', '"NaN"')  
			# print()
			self.teste['contas'] = json.loads(self.teste['contas'])
			
			for conta in self.teste['contas']:
				self.todas_mensalidades.append(conta)


	def tratar_dados(self):
		arquivo = open('analise/todo_site.txt', 'w')
		print(json.dumps(self.todas_mensalidades, indent=4, sort_keys=True), file=arquivo)
		arquivo.close()

		for credito in self.todas_mensalidades:
			if credito['vencimento'] != "NaN" and credito['unidade'] !=  "NaN" and credito['bloco'] !=  "NaN":
				data_vencimento = datetime.datetime.strptime(credito['vencimento'], '%d/%m/%Y').date()
				data_hoje = datetime.datetime.now().date()
				
				if "ADMINISTRADORA" not in str(credito['unidade']): 
					if "Acordo" in credito['itemDescricao'] or "ACORDO" in credito['itemDescricao']:
						if credito['situac\u00e3o'].lower() != "pago":
							self.creditos_vencidos.append(credito)	
					else:
						if data_vencimento < data_hoje:
							if str(credito['unidade']) != "NaN":
								if credito['situac\u00e3o'].lower() != "pago":
									self.creditos_vencidos.append(credito)


		self.creditos_vencidos = sorted(self.creditos_vencidos, key=lambda k: (str(k['bloco']), str(k['unidade'])))

		arquivo = open('analise/filtrado_ordenado.txt', 'w')
		print(json.dumps(self.creditos_vencidos, indent=4, sort_keys=True), file=arquivo)
		arquivo.close()

		for credito in self.creditos_vencidos:
			if str(int(credito['unidade']))+credito['bloco'] not in self.apartamentos_credito:
				self.apartamentos_credito[str(int(credito['unidade']))+credito['bloco']] = []

			self.apartamentos_credito[str(int(credito['unidade']))+credito['bloco']].append(credito)
				
		print(json.dumps(self.apartamentos_credito, indent=4, sort_keys=True))
		
		arquivo = open('analise/filtrado_separado.txt', 'w')
		print(json.dumps(self.apartamentos_credito, indent=4, sort_keys=True), file=arquivo)
		arquivo.close()

	def get_html(self):
		date_today = datetime.datetime.now()
		data = date_today.date()
		dateFormated = date_today.strftime('%d/%m/%Y')
		locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

		book = Workbook()
		sheet = book.active
		nome_xlsx = 'analise/apartamentos.xlsx'

		sheet['A1'] = 'Unidade'
		sheet['B1'] = 'Total devido'
		sheet['C1'] = 'Total Acordo'
		sheet['D1'] = 'Total'
		cont = 2

		html = """
			<!DOCTYPE html>
			<html lang="en">
			<head>
			<style>
			*{
				box-sizing: border-box;
				-webkit-box-sizing: border-box;
				-moz-box-sizing: border-box;
			}
			body{
				font-family: Helvetica;
				-webkit-font-smoothing: antialiased;
				background: gray;
				min-height: 100vh;
			}
			h2{
				text-align: center;
				font-size: 18px;
				letter-spacing: 1px;
				color: black;
				padding: 30px 0;
			}
			/* Table Styles */
			.table-wrapper{
				margin: 10px 70px 70px;
				box-shadow: 0px 35px 50px rgba( 0, 0, 0, 0.2 );
			}
			.fl-table {
				border-radius: 5px;
				font-size: 12px;
				font-weight: normal;
				border: none;
				border-collapse: collapse;
				width: 100%;
				max-width: 100%;
				white-space: nowrap;
				background-color: white;
			}
			.fl-table td, .fl-table th {
				text-align: center;
				padding: 8px;
			}
			.fl-table td {
				border-right: 1px solid #f8f8f8;
				font-size: 12px;
			}
			.fl-table thead th {
				color: #ffffff;
				background: #72242b;
			}
			.fl-table thead th:nth-child(odd) {
				color: #ffffff;
				background: #72242b;
			}
			.fl-table tr:nth-child(even) {
				background: #F8F8F8;
			}
			/* Responsive */
			@media (max-width: 767px) {
				.fl-table {
					display: block;
					width: 100%;
				}
				.table-wrapper:before{
					content: "Scroll horizontally >";
					display: block;
					text-align: right;
					font-size: 11px;
					color: white;
				}
				.fl-table thead, .fl-table tbody, .fl-table thead th {
					display: block;
				}
				.fl-table thead th:last-child{
					border-bottom: none;
				}
				.fl-table thead {
					float: left;
				}
				.fl-table tbody {
					width: auto;
					position: relative;
					overflow-x: auto;
				}
				.fl-table td, .fl-table th {
					padding: 20px .625em .625em .625em;
					height: 60px;
					vertical-align: middle;
					box-sizing: border-box;
					overflow-x: hidden;
					overflow-y: auto;
					width: 120px;
					font-size: 13px;
					text-overflow: ellipsis;
				}
				.fl-table thead th {
					text-align: left;
					border-bottom: 1px solid #f7f7f9;
				}
				.fl-table tbody tr {
					display: table-cell;
				}
				.fl-table tbody tr:nth-child(odd) {
					background: none;
				}
				.fl-table tr:nth-child(even) {
					background: transparent;
				}
				.fl-table tr td:nth-child(odd) {
					background:  #72242b;
					border-right: 1px solid #E6E4E4;
				}
				.fl-table tr td:nth-child(even) {
					border-right: 1px solid #E6E4E4;
				}
				.fl-table tbody td {
					display: block;
					text-align: center;
				}
			}
			</style>
			</head>
			<body>
			<h2>Apartamentos com crédito a pagar - """ + dateFormated + """ </h2>
			<div class="table-wrapper">
				<table class="fl-table">
					<thead>
					<tr>
						<th>Unidade</th>
						<th>Total devido</th>
						<th>Total Acordo</th>
						<th>Total</th>
					</tr>
					</thead>
					<tbody>
		"""
		# print(json.dumps(self.apartamentos_credito, indent=4, sort_keys=True))

		total = 0
		total_debitos = 0
		total_acordos = 0
		for apartamento in self.apartamentos_credito.keys():
			total_apto_debito = 0
			total_apto_acordo = 0
			for despesa_apto in self.apartamentos_credito[apartamento]: 
				if "Acordo" in despesa_apto['itemDescricao'] or "ACORDO" in despesa_apto['itemDescricao']:
					valor_formatado = float(despesa_apto['valor'].replace("R$ ", "").replace(",", "").replace(".", "")) / 100
					total_apto_acordo = total_apto_acordo + valor_formatado
				else:
					valor_formatado = float(despesa_apto['valor'].replace("R$ ", "").replace(",", "")) / 100
					total_apto_debito = total_apto_debito + valor_formatado

			# print(f"{int(despesa_apto['unidade'])} - {despesa_apto['bloco']} = Débitos {total_apto_debito} || Acordor {total_apto_acordo}")
			total_debitos = total_debitos + total_apto_debito
			total_acordos = total_acordos + total_apto_acordo
			total = total + total_apto_acordo + total_apto_debito

			self.historico[apartamento] = {'total_debitos': total_debitos, 'total_acordos': total_acordos, 'total': total}
		
		# for apartamento in self.apartamentos_credito.keys():
		# 	total_apto = 0
		# 	for despesa_apto in self.apartamentos_credito[apartamento]: 
		# 		valor_formatado = float(despesa_apto['valor'].replace("R$ ", "").replace(",", "")) / 100
		# 		total_apto = total_apto + valor_formatado

			# print(f"{int(despesa_apto['unidade'])} - {despesa_apto['bloco']} = {total_apto}")     
			total_do_apto = total_apto_debito + total_apto_acordo
			total_apto_debito = locale.currency(total_apto_debito, grouping=True, symbol=None)
			total_apto_acordo = locale.currency(total_apto_acordo, grouping=True, symbol=None)
			total_do_apto = locale.currency(total_do_apto, grouping=True, symbol=None)
			html += """
					<tr>
						<td>%s - %s</td>
						<td>R$ %s</td>
						<td>R$ %s</td>
						<td>R$ %s</td>
					</tr>
			""" %(int(despesa_apto['unidade']), despesa_apto['bloco'], total_apto_debito, total_apto_acordo, total_do_apto)

			sheet['A'+str(cont)] = str(int(despesa_apto['unidade'])) + " - " + despesa_apto['bloco']
			sheet['B'+str(cont)] = total_apto_debito
			sheet['C'+str(cont)] = total_apto_acordo
			sheet['D'+str(cont)] = total_do_apto
			cont = cont + 1


		total_debitos = locale.currency(total_debitos, grouping=True, symbol=None)
		total_acordos = locale.currency(total_acordos, grouping=True, symbol=None)
		total = locale.currency(total, grouping=True, symbol=None)

		html += """ 	
					<tr>
						<td>Total</td>
						<td>R$ %s</td>
						<td>R$ %s</td>
						<td>R$ %s</td>
					</tr>
					<tbody>
				</table>
			</div>
			</body>
			</html>
		""" %(total_debitos, total_acordos, total)


		sheet['A'+str(cont)] = "Total"
		sheet['B'+str(cont)] = total_debitos
		sheet['C'+str(cont)] = total_acordos
		sheet['D'+str(cont)] = total

		arquivo = open(f'analise/historico{data}.txt', 'w')
		print(json.dumps(self.historico, indent=4, sort_keys=True), file=arquivo)
		arquivo.close()
		
		# SALVANDO PLANILHA
		book.save(nome_xlsx)

		arquivo = open('analise/index.html', 'w')
		print(html, file=arquivo)
		arquivo.close()

		return html

	def verifica_se_esta_igual_outro_dia(self):
		date_today = datetime.datetime.now()
		data = date_today.date()
		arquivo = open('analise/historico{data}.txt', 'r')
		dia_hoje = eval(arquivo.read())
		arquivo.close()

		data = data - datetime.timedelta(day=1)
		arquivo = open('analise/historico{data}.txt', 'r')
		dia_anterior = eval(arquivo.read())
		arquivo.close()

		if dia_anterior != dia_hoje:
			print("Hoje está com valores diferentes de ontem")
			return True

		return False

	def envia_email(self):
		print("Enviando email.")

		# is_diference = verifica_se_esta_igual_outro_dia()

		# if is_diference:
		sender_email = "luanpetruitis@gmail.com"
		receiver_email = "luanpetruitis@gmail.com"
		password = "futebol2011"
		# password = input("Type your password and press enter:")

		message = MIMEMultipart("alternative")
		date_today = datetime.datetime.now()
		dateFormated = date_today.strftime('%d/%m/%Y')
		message["Subject"] = "Apartamentos com crédito a pagar - %s" %(dateFormated)
		message["From"] = sender_email
		message["To"] = receiver_email

		html = self.get_html()

		text = MIMEText(html, "html")

		message.attach(text)
		# else:
		# 	sender_email = "luanpetruitis@gmail.com"
		# 	receiver_email = "luanpetruitis@gmail.com"
		# 	password = "futebol2011"
		# 	# password = input("Type your password and press enter:")

		# 	message = MIMEMultipart("alternative")
		# 	date_today = datetime.datetime.now()
		# 	dateFormated = date_today.strftime('%d/%m/%Y')
		# 	message["Subject"] = "Apartamentos com crédito a pagar - %s" %(dateFormated)
		# 	message["From"] = sender_email
		# 	message["To"] = receiver_email

		# 	html = self.get_html()

		# 	text = MIMEText(html, "html")

		# 	message.attach(text)

		# Create secure connection with server and send email
		context = ssl.create_default_context()
		with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
			server.login(sender_email, password)
			server.sendmail(
				sender_email, receiver_email, message.as_string()
			)


	def fecha_navegador(self):
		print("Fechando navegador")
		self.driver.quit() 
