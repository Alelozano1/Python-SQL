import PySimpleGUI as sg
import requests
import pandas as pd
import mysql.connector

mydb = mysql.connector.connect(
    host="o2-db-ativos-2-aurora-cluster.cluster-ro-cgqle4cw3o4s.us-east-1.rds.amazonaws.com",
    user="user-alexandre",
    password="Esfera2020",
    database="0_transfer"
)

class TelaPython:
    sg.theme('DarkAmber')
    def __init__(self):
        layout = [
            [sg.Multiline('Coloque o pedido', size=(45, 5),key='idss')],
            [sg.Button('Exportar'),sg.Checkbox('Pesquisar por cupom',key='cupon')]

        ]
        
        self.janela = sg.Window('Evento').layout(layout)


    def iniciar(self):
        while True:
            self.button, self.values = self.janela.Read()
            multiline_text = self.values['idss']
            cupon = self.values['cupon']
            multiline_text = multiline_text.strip()  # Remove espaços extras no início e no final
            string_list = multiline_text.split('\n')
            mycursor = mydb.cursor()

            if cupon == False:
                texto = "e.id_evento IN (" + ",".join(["%s"] * len(string_list)) + ")"
            else:
                texto = "cdi.nr_cupom IN (" + ",".join(["%s"] * len(string_list)) + ")"

            query = f'SELECT ' \
                    f'pe.nr_peito,' \
                    f'e.ds_evento "Evento",' \
                    f'p.id_pedido "Protocolo",' \
                    f'IF(p.fl_local_inscricao = 1, u.id_usuario, ub.id_usuario) "id_usuario",' \
                    f'pe.id_pedido_evento "Id Inscrição",' \
                    f'IF(p.fl_local_inscricao = 1, u.nr_documento, ub.nr_documento) "CPF",' \
                    f'u.ds_nomebalcao "Balcão",' \
                    f'em.nm_modalidade "Percurso",' \
                    f'IF (p.fl_local_inscricao = 1, u.ds_nomecompleto, ub.ds_nome) "Nome",' \
                    f'IF(tc.id_tamanho_camiseta = 2, "BL", tc.ds_tamanho) "Camiseta",' \
                    f'IF(p.fl_local_inscricao = 1, u.ds_email, ub.ds_email) "Email",' \
                    f'cd.id_cupom_desconto "Id Cupom",' \
                    f'cdi.nr_cupom "Cupom",' \
                    f'cd.ds_referencia_externa "Referencia Externa",' \
                    f'cd.en_cupom_classificacao "Tipo Cupom"' \
                    f'FROM ' \
                    f'sa_pedido_evento as pe ' \
                    f'JOIN sa_usuario as u ON pe.id_usuario = u.id_usuario ' \
                    f'JOIN sa_evento as e ON pe.id_evento = e.id_evento ' \
                    f'JOIN sa_pedido as p ON pe.id_pedido = p.id_pedido ' \
                    f'JOIN sa_pedido_status as ps ON p.id_pedido_status = ps.id_pedido_status ' \
                    f'JOIN sa_status_detalhado as sd ON pe.id_status_detalhado = sd.id_status_detalhado ' \
                    f'LEFT JOIN sa_tamanho_camiseta as tc ON pe.id_tamanho_camiseta = tc.id_tamanho_camiseta ' \
                    f'LEFT JOIN sa_evento_modalidade as em ON pe.id_modalidade = em.id_modalidade ' \
                    f'LEFT JOIN sa_cupom_desconto_item as cdi on pe.id_cupom_individual = cdi.id_cupom_desconto_item ' \
                    f'LEFT JOIN sa_cupom_desconto as cd on cdi.id_cupom_desconto = cd.id_cupom_desconto ' \
                    f'LEFT JOIN sa_usuario_balcao as ub on pe.id_usuario_balcao = ub.id_usuario ' \
                    f'WHERE ' \
                    f'p.id_pedido_status = 2 and ' \
                    f'{texto} '\
                    f'GROUP BY pe.id_pedido_evento ' \
                    f'ORDER BY pe.nr_peito'



            mycursor.execute(query, string_list)  # Execute the query with the list of values
            teste = mycursor.fetchall()
            df = pd.DataFrame(teste,columns=["N° Peito","Evento","Protocolo","Id Usuario","Id Inscrição","Documento","Balcão","Modalidade","Nome","Camiseta","E-mail","Id Cupom","Cupom","Referencia Externa","Tipo Cupom"])
            xlsx_path = 'output.xlsx'
            df.to_excel(xlsx_path, index=False)
            sg.popup(f'Arquivo XLS criado: {xlsx_path}')

tela = TelaPython()
tela.iniciar()
