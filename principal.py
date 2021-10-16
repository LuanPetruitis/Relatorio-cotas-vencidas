from cobranca import Cobranca 

try:
    cobranca = Cobranca()
    cobranca.login()
    cobranca.pesquisa()
finally:
    cobranca.fecha_navegador()

cobranca.tratar_dados()
cobranca.envia_email()