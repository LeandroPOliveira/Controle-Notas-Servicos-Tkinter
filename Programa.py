from tkinter import *
from tkinter import ttk
import tkinter.messagebox
import pyodbc
import pandas as pd
from tkinter import filedialog as fd
import os
from datetime import date
from openpyxl.reader.excel import load_workbook
from fpdf import FPDF
from datetime import datetime
from dateutil.relativedelta import relativedelta


class NotasServicos:

    def __init__(self, janela):
        self.janela = janela
        titulo = ' '
        self.janela.title(160 * titulo + 'Notas fiscais de Serviço')
        self.janela.geometry('1200x680+100+20')
        self.janela.resizable(width=False, height=False)

        #definindo frames
        mainframe = Frame(self.janela, bd=5, width=1200, height=680, relief=RIDGE, bg='RoyalBlue1')
        mainframe.place(x=0, y=0)

        topframe = Frame(mainframe, bd=10, width=1190, height=100, bg='cornflowerblue', relief=RIDGE)
        topframe.place(x=0, y=0)
        labeltit = Label(topframe, text='Controle de Notas Fiscais de Serviço', font=('@Microsoft YaHei', 35, 'bold'),
        bg='cornflowerblue', fg='white')
        labeltit.place(x=150, y=5)

        self.comand_frame = Frame(mainframe, bd=5, width=1190, height=100, bg='white smoke', relief=RIDGE)
        self.comand_frame.place(x=43, y=472)

        leftframe = Frame(mainframe, bd=5, width=1190, height=370, bg='white smoke', relief=RIDGE)
        leftframe.place(x=0, y=101)

        self.bottonframe = Frame(mainframe, bd=5, width=1190, height=140, bg='white smoke', relief=RIDGE)
        self.bottonframe.place(x=0, y=525)

        # inserir botões de comando
        btn1 = Button(self.comand_frame, font=('arial', 14, 'bold'), fg='blue', text='Adicionar', bd=4, pady=1, padx=24,
                         width=10, height=1, command=self.adicionar).grid(row=0, column=0, padx=1)
        btn2 = Button(self.comand_frame, font=('arial', 14, 'bold'), text='Limpar', fg='blue', bd=4, pady=1, padx=24,
                         width=10, height=1, command=self.limpar).grid(row=0, column=5, padx=1)
        btn3 = Button(self.comand_frame, font=('arial', 14, 'bold'), text='Listar', fg='blue', bd=4, pady=1, padx=24,
                      width=10, height=1, command=self.mostrar_dados).grid(row=0, column=4, padx=1)
        btn4 = Button(self.comand_frame, font=('arial', 14, 'bold'), text='Apagar', fg='blue',bd=4, pady=1, padx=24,
                      width=10, height=1, command=self.deletar).grid(row=0, column=2, padx=1)
        btn5 = Button(self.comand_frame, font=('arial', 14, 'bold'), text='Procurar', bd=4, fg='blue', pady=1, padx=24,
                      width=10, height=1, command=self.procurar).grid(row=0, column=1, padx=1)
        btn6 = Button(self.comand_frame, font=('arial', 14, 'bold'), text='Atualizar', bd=4, fg='blue', pady=1, padx=24,
                      width=10, height=1, command=self.atualizar).grid(row=0, column=3, padx=1)

        # checkbox para replicar os dados do último lançamento
        self.lembrar = IntVar()
        self.check = tkinter.Checkbutton(leftframe, text='Repetir Lançamento', onvalue=1, offvalue=0, font=('arial', 12),
        variable=self.lembrar, command=self.lembrar_lancamento)
        self.check.place(x=900, y=300)


        #Criar Menu
        meu_menu = Menu(self.janela)
        self.janela.config(menu=meu_menu)

        #Criar itens do Menu Arquivo
        menu_arquivo = Menu(meu_menu, tearoff=0)
        meu_menu.add_cascade(label='Arquivo', menu=menu_arquivo)
        # menu_arquivo.add_command(label='Novo')
        # menu_arquivo.add_separator()
        menu_arquivo.add_command(label='Sair', command=self.janela.quit)

        # ===========================Menu Relatórios================================================================#
        def relatorios():
            relat = Toplevel()
            relat.title('Relatórios')
            relat.geometry('700x500')

            rel_frame = Frame(relat, bd=5, width=700, height=500, relief=RIDGE)
            rel_frame.place(x=0, y=0)
            Label(rel_frame, text='PERÍODO', font=fonte, bd=5).place(x=80, y=70)
            Label(rel_frame, text='A', font=fonte, bd=5).place(x=320, y=70)

            def gerar_relatorios():
                #Criar planilha para gerar arquivo
                writer = pd.ExcelWriter(self.entr_dir.get() + '\Relatórios.xlsx', engine='xlsxwriter')
                #Conectar ao banco
                lmdb = os.getcwd() + '\Base_notas.accdb;'
                cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                cursor = cnx.cursor()

                #=================== Relatório Imposto de Renda ===============================================#
                if check_ir.get() == 1:
                    cursor.execute('select cnpj, fornecedor, sum(irrf) from notas_fiscais '
                    'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) <= '
                                   'DateValue(?) group by cnpj, fornecedor', dt_1.get(), dt_2.get())
                    resultado = cursor.fetchall()
                    lista = [[],[],[]]
                    for i in resultado:
                        for l in range(3):
                            lista[l].append(i[l])
                    tabela = pd.DataFrame(lista).transpose()
                    tabela.columns = ['CNPJ', 'Fornecedor', 'IRRF']
                    tabela.to_excel(writer, sheet_name='Irrf', index=False)

                # =========================== Relatório de Contribuições =========================================#
                if check_crf.get() == 1:
                    cursor.execute('select * from notas_fiscais where data_vencimento <> 0 (select data_vencimento, '
                                   'cnpj, fornecedor, sum(crf) from notas_fiscais '
                    'where DateValue(data_vencimento) >= DateValue(?) and DateValue(data_vencimento) <= DateValue(?) '
                    'group by data_vencimento, cnpj, fornecedor order by fornecedor, data_vencimento)',
                    (dt_1.get(), dt_2.get()))
                    resultado = cursor.fetchall()
                    lista2 = [[], [], [], []]
                    for i in resultado:
                        for l in range(4):
                            lista2[l].append(i[l])
                    tabela2 = pd.DataFrame(lista2).transpose()
                    tabela2.columns = ['Data_Vencimento', 'CNPJ', 'Fornecedor', 'CRF']
                    tabela2.to_excel(writer, sheet_name='Crf', index=False)

                #===========================Relatório ISS ==========================================================#
                if check_iss.get() == 1:
                    lmdb = os.getcwd() + '\\base_notas.accdb;'
                    cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                    cursor = cnx.cursor()

                    cursor.execute('select distinct cidade from notas_fiscais where DateValue(data_analise) >= '
                                   'DateValue(?) and DateValue(data_analise) <= DateValue(?)', (dt_1.get(), dt_2.get()))
                    lista = []
                    for row in cursor:
                        if row[0] != '':
                            lista.append(row[0])

                    for i in lista:
                        cursor.execute('select NF, fornecedor, iss from notas_fiscais '
                                       'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) <= DateValue(?)'
                                       'and cidade = ? order by cidade, cnpj', (dt_1.get(), dt_2.get(), i))


                        vencimentos = pd.read_excel('G:\GECOT\FISCAL - Retenções\\Programa Planilha de retenção.xlsx',
                                                    sheet_name='Relatório ISS', usecols=[9, 10], skiprows=10, dtype=str)

                        for index, row in vencimentos.iterrows():
                            if row['MUNICÍPIOS'] == i.upper():
                                dia = vencimentos.loc[index, 'DIA']
                                data = datetime.strptime(dt_2.get(), '%d/%m/%Y')
                                data = data + relativedelta(months=1)
                                data = data.strftime('%m/%Y')
                                data_venc = dia + '/' + data

                        pdf = FPDF(orientation='P', unit='mm', format='A4')
                        pdf.add_page()
                        pdf_w = 210
                        pdf_h = 297
                        pdf.set_font('Arial', 'B', 10)
                        pdf.image('G:\GECOT\FISCAL - Retenções\logo.png', x=10.0, y=10.0,
                                  h=50.0, w=100.0)
                        pdf.set_xy(10.0, 70.0)
                        pdf.multi_cell(w=125, h=5, txt='ISSQN Município de ' + i)
                        pdf.multi_cell(w=125, h=5, txt='A/C: Contabilidade - Contas a Pagar.')
                        pdf.set_xy(10.0, pdf.get_y() + 5)
                        pdf.multi_cell(w=150, h=5,
                                       txt='Planilha contendo valores a recolher referente ao mês ' + dt_2.get()[3:])
                        pdf.multi_cell(w=125, h=5, txt='Valor a recolher através de BOLETO ANEXO - Contas a Pagar.')
                        pdf.multi_cell(w=125, h=5, txt='Vencimento: ' + data_venc)
                        pdf.set_xy(10.0, pdf.get_y() + 15)
                        pdf.multi_cell(w=30, h=5, txt='Nota Fiscal', border=1, align='C')
                        pdf.set_xy(40.0, pdf.get_y() - 5)
                        pdf.multi_cell(w=80, h=5, txt='Fornecedor', border=1, align='C')
                        pdf.set_xy(120.0, pdf.get_y() - 5)
                        pdf.multi_cell(w=40, h=5, txt='ISS a recolher', border=1, align='C')
                        pdf.set_font('')
                        resultado = cursor.fetchall()
                        soma = []
                        cont = 0
                        for lin in resultado:
                            lin[2] = float(lin[2])
                            soma.append(lin[2])
                            lin[2] = str(lin[2]).replace('.', ',')
                            pdf.multi_cell(w=30, h=5, txt=str(lin[0]), border=1, align='C')
                            pdf.set_xy(40.0, pdf.get_y() - 5)
                            pdf.multi_cell(w=80, h=5, txt=str(lin[1][:32]), border=1, align='C')
                            pdf.set_xy(120.0, pdf.get_y() - 5)
                            pdf.multi_cell(w=40, h=5, txt=str(lin[2]), border=1, align='C')
                            cont += 1
                        for l in range(20 - cont):
                            pdf.multi_cell(w=30, h=5, txt='', border=1)
                            pdf.set_xy(40.0, pdf.get_y() - 5)
                            pdf.multi_cell(w=80, h=5, txt='', border=1)
                            pdf.set_xy(120.0, pdf.get_y() - 5)
                            pdf.multi_cell(w=40, h=5, txt='', border=1)
                        pdf.set_xy(40.0, pdf.get_y())
                        pdf.multi_cell(w=80, h=5, txt='Valor total a recolher', border=1, align='C')
                        pdf.set_xy(120.0, pdf.get_y() - 5)
                        pdf.set_font('Arial', 'B', 10)
                        pdf.multi_cell(w=40, h=5, txt=str(round(sum(soma), 2)), border=1, align='C')
                        pdf.set_xy(10.0, pdf.get_y() + 30)
                        pdf.line(10, pdf.get_y(), 60, pdf.get_y())
                        pdf.multi_cell(w=100, h=5, txt='Pedro Henrique Carrilho')
                        pdf.multi_cell(w=100, h=5, txt='Contador Junior')
                        pdf.multi_cell(w=40, h=5, txt='GECOT')
                        data = dt_2.get()[3:].replace('/', '-')
                        pdf.output(i + ' ' + data + '.pdf', 'F')

                #===========================Relatório INSS==========================================================#
                if check_inss.get() == 1:
                    cursor.execute('select data, NF, cnpj, fornecedor, valor_bruto, inss from notas_fiscais '
                    'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) '
                    '<= DateValue(?)',
                    (dt_1.get(), dt_2.get()))
                    resultado = cursor.fetchall()
                    lista4 = [[], [], [], [], [], []]
                    for i in resultado:
                        for l in range(6):
                            lista4[l].append(i[l])
                    tabela4 = pd.DataFrame(lista4).transpose()
                    tabela4.columns = ['Data Nota Fiscal', 'Nº NF', 'CNPJ', 'Fornecedor', 'Valor Bruto', 'INSS']
                    tabela4.to_excel(writer, sheet_name='INSS', index=False)
                else:
                    pass

                writer.save()

            # intervalo de datas dos relatórios
            dt_1 = Entry(rel_frame, bd=5, font=fonte, width=10)
            dt_1.place(x=200, y=70)
            # dt_1.insert(0, '01/10/2021')
            dt_2 = Entry(rel_frame, bd=5, font=fonte, width=10)
            dt_2.place(x=380, y=70)
            # dt_2.insert(0, '31/10/2021')

            btn_rel = Button(rel_frame, font=('arial', 14, 'bold'), text='Gerar Relatórios', bd=4, pady=1, padx=24,
                             width=12, height=1, command=gerar_relatorios).place(x=230, y=400)

            # checks do menu relatórios
            check_ir = IntVar()
            check_crf = IntVar()
            check_iss = IntVar()
            check_inss = IntVar()

            # seleção dos relatórios a serem gerados
            btncheck_ir = tkinter.Checkbutton(rel_frame, text='Imposto de Renda', onvalue=1, offvalue=0,
            font=('arial', 12), variable=check_ir)
            btncheck_ir.place(x=250, y=160)
            btncheck_crf = tkinter.Checkbutton(rel_frame, text='Contribuições', onvalue=1, offvalue=0,
            font=('arial', 12), variable=check_crf)
            btncheck_crf.place(x=250, y=200)
            btncheck_iss = tkinter.Checkbutton(rel_frame, text='Imposto sobre serviços', onvalue=1, offvalue=0,
            font=('arial', 12), variable=check_iss) #state='disabled')
            btncheck_iss.place(x=250, y=240)
            btncheck_inss = tkinter.Checkbutton(rel_frame, text='Imposto s/ Seguridade Social', onvalue=1, offvalue=0,
            font=('arial', 12), variable=check_inss)
            btncheck_inss.place(x=250, y=280)

            # selecionar diretório
            def sel_diretorio():
                diretorio = fd.askdirectory(title='Abrir diretório')
                self.entr_dir.delete(0, END)
                self.entr_dir.insert(0, diretorio)
                relat.lift()

            self.entr_dir = Entry(rel_frame, width=28, font=12, bd=2)
            self.entr_dir.place(x=150, y=350)
            self.entr_dir.insert(0, os.getcwd())
            btn_dir = Button(rel_frame, text='Selecionar Diretório', command=sel_diretorio, bd=4)
            btn_dir.place(x=420, y=345)

        def exp_banco():
            # exportar banco completo para consultas e geração de guias de recolhimento
            book = load_workbook('G:\GECOT\FISCAL - Retenções\Programa Planilha de retenção.xlsx')
            writer = pd.ExcelWriter('G:\GECOT\FISCAL - Retenções\Programa Planilha de retenção.xlsx', engine='openpyxl')
            writer.book = book

            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

            # Conectar ao banco
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('select * from notas_fiscais')
            resultado = cursor.fetchall()
            lista = [[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []]
            for i in resultado:
                for l in range(20):
                    lista[l].append(i[l])
            tabela = pd.DataFrame(lista).transpose()
            tabela.columns = ['ID', 'data_analise', 'data', 'data_vencimento', 'NF', 'CNPJ', 'Fornecedor', 'cidade',
                              'simples_nacional', 'codigo_servico', 'valor_bruto', 'aliq_irrf', 'irrf', 'aliq_crf',
                              'crf',
                              'aliq_inss', 'inss', 'aliq_iss', 'iss', 'valor_liquido']

            cols = ['valor_bruto', 'aliq_irrf', 'irrf', 'aliq_crf', 'crf',
                    'aliq_inss', 'inss', 'aliq_iss', 'iss', 'valor_liquido']

            tabela[cols] = tabela[cols].apply(pd.to_numeric, errors='coerce')

            frame = pd.DataFrame(tabela)
            frame.to_excel(writer, sheet_name='Geral', index=False)

            writer.save()
            tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Banco exportado com Sucesso!')

        menu_relatorio = Menu(meu_menu, tearoff=0)
        meu_menu.add_cascade(label='Relatórios', menu=menu_relatorio)
        menu_relatorio.add_command(label='Gerar Relatórios', command=relatorios)
        menu_relatorio.add_command(label='Exportar Banco de dados', command=exp_banco)

        #======================================MENU CONSULTAS=========================================================#
        def consultas():
            consult = Toplevel()
            consult.title('Consultas')
            consult.geometry('1300x600')

            # Adicionar estilo, tema e cores ao treeview
            estilo = ttk.Style()
            estilo.theme_use('default')
            estilo.configure('Treeview', background='#D3D3D3', foreground='black', rowheight=25,
                             fieldbackground='#D3D3D3')
            estilo.map('Treeview', background=[('selected', '#347083')])

            # Treeview frame
            tree_frame = Frame(consult)
            tree_frame.pack(pady=10)
            # Barra rolagem
            tree_scroll = Scrollbar(tree_frame)
            tree_scroll.pack(side=RIGHT, fill=Y)
            # Criar Treeview
            nf_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode='extended')
            nf_tree.pack()
            # Configurar Barra Rolagem
            tree_scroll.config(command=nf_tree.yview)
            # Definir colunas
            colunas2 = ['ID', 'DT_Análise', 'DT_NF', 'DT_Venc', 'NF', 'CNPJ', 'Fornecedor', 'Município', 'Simples',
                       'Cod. Ser.', 'Val Bruto', 'Aliq_IR', 'IRRF', 'Aliq_CRF', 'CRF', 'Aliq_INSS', 'INSS',
                       'Aliq_ISS', 'ISS', 'Val Líq']
            nf_tree['columns'] = colunas2
            # formatar colunas
            nf_tree.column('#0', width=0, stretch=NO)
            for coluna2 in colunas2:
                nf_tree.column(coluna2, width=50)
            nf_tree.column('DT_Análise', width=70)
            nf_tree.column('DT_NF', width=70)
            nf_tree.column('DT_Venc', width=70)
            nf_tree.column('ID', width=30)
            nf_tree.column('CNPJ', width=110)
            nf_tree.column('Fornecedor', width=110)
            nf_tree.column('Município', width=70)
            nf_tree.column('Val Bruto', width=70)
            nf_tree.column('Val Líq', width=70)
            # formatar títulos
            nf_tree.heading('#0', text='', anchor=W)
            for h in colunas2:
                nf_tree.heading(h, text=h)

            #inserir dados do banco no treeview
            def inserir_tree(resultado):
                nf_tree.delete(*nf_tree.get_children())
                contagem = 0
                for row in resultado: # loop para inserir cores diferentes nas linhas
                    if contagem % 2 == 0:
                        nf_tree.insert(parent='', index='end', text='', iid=contagem,
                                       values=(row[0], row[1], row[2], row[3],
                                               row[4], row[5], row[6], row[7], row[8], row[9], format_num(row[10]),
                                               format_num(row[11]), format_num(row[12]),
                                               format_num(row[13]), format_num(row[14]), format_num(row[15]),
                                               format_num(row[16]), format_num(row[17]),
                                               format_num(row[18]), format_num(row[19])), tags=('evenrow',))
                    else:
                        nf_tree.insert(parent='', index='end', text='', iid=contagem,
                                       values=(row[0], row[1], row[2], row[3],
                                               row[4], row[5], row[6], row[7], row[8], row[9], format_num(row[10]),
                                               format_num(row[11]), format_num(row[12]),
                                               format_num(row[13]), format_num(row[14]), format_num(row[15]),
                                               format_num(row[16]),
                                               format_num(row[17]), format_num(row[18]), format_num(row[19])),
                                       tags=('oddrow',))
                    contagem += 1

            # conectar banco de dados
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('select * from notas_fiscais order by ID desc')
            resultado1 = cursor.fetchall()
            cnx.commit()
            cnx.close()

            #formatar números para exibição no padrão BR
            def format_num(num):
                num = f'{num:0,.2f}'
                maketrans = num.maketrans
                num = num.translate(maketrans(',.', '.,'))
                return num

            # Selecionar a data de vencimento para modificação
            def NotasInfo2(ev):
                fn_id.delete(0, END)
                verinfo2 = nf_tree.focus()
                dados2 = nf_tree.item(verinfo2)
                row = dados2['values']
                # entr_atual.delete(0, END)
                entr_atual.insert(0, row[3])
                fn_id.insert(0, row[0])

            # adicionar a tela
            nf_tree.tag_configure('oddrow', background='white')
            nf_tree.tag_configure('evenrow', background='lightblue')
            inserir_tree(resultado1)
            nf_tree.bind('<ButtonRelease-1>', NotasInfo2)

            # filtro de datas
            def filtrar_data(ev):
                if dt_inicio.get() != '' and dt_fim.get() != '':
                    lmdb = os.getcwd() + '\Base_notas.accdb;'
                    cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                    cursor = cnx.cursor()
                    cursor.execute(
                        'select * from notas_fiscais where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) <= '
                        'DateValue(?) order by id desc', dt_inicio.get(), dt_fim.get())
                    resultado2 = cursor.fetchall()
                    inserir_tree(resultado2)

            # Adicionar Botões e entradas
            data_frame = LabelFrame(consult, text='Filtros')
            data_frame.pack(fill='x', expand='yes', padx=20)

            label_inic = Label(data_frame, text='Data Início')
            label_inic.grid(row=0, column=0, padx=10, pady=10)
            dt_inicio = Entry(data_frame)
            dt_inicio.grid(row=0, column=1, padx=10, pady=10)

            label_fim = Label(data_frame, text='Data Fim')
            label_fim.grid(row=0, column=2, padx=10, pady=10)
            dt_fim = Entry(data_frame)
            dt_fim.grid(row=0, column=3, padx=10, pady=10)
            dt_fim.bind('<FocusOut>', filtrar_data)


            #função para atualizar ou inserir o vencimento da nota
            def atualizar_venc():
                lmdb = os.getcwd() + '\Base_notas.accdb;'
                cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                cursor = cnx.cursor()
                cursor.execute('update notas_fiscais set DATA_VENCIMENTO=? where ID=?', (entr_atual.get(), fn_id.get()))
                tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro alterado com sucesso!')
                cursor.execute('select * from notas_fiscais order by ID desc')
                resultado2 = cursor.fetchall()
                inserir_tree(resultado2)
                cnx.commit()
                cnx.close()
                entr_atual.delete(0, END)
                consult.lift()

            label_atual = Label(data_frame, text='Atualizar Vencimento')
            label_atual.grid(row=0, column=4, padx=10, pady=10)
            entr_atual = Entry(data_frame)
            entr_atual.grid(row=0, column=5, padx=10, pady=10)
            btn_atual = Button(data_frame, text='Gravar', command=atualizar_venc)
            btn_atual.grid(row=0, column=6, padx=10, pady=10)
            fn_id = Entry(data_frame) # Entrada simbólica para guarda a ID que servirá de índice para atualização vencim


        menu_consulta = Menu(meu_menu, tearoff=0)
        meu_menu.add_cascade(label='Consultas', menu=menu_consulta)
        menu_consulta.add_command(label='Consultar Banco', command=consultas)

        #================================== MENU CADASTRO FORNECEDORES ===============================================#
        def tela_cadastro():
            cadastro = Toplevel()
            cadastro.title('Cadastro de Prestadores')
            cadastro.geometry('700x500')
            # labels
            cad_frame = Frame(cadastro, bd=5, width=700, height=500, relief=RIDGE)
            cad_frame.place(x=0, y=0)
            Label(cad_frame, text='CNPJ', font=fonte, bd=5).place(x=80, y=70)
            Label(cad_frame, text='NOME', font=fonte, bd=5).place(x=80, y=120)
            Label(cad_frame, text='MUNICÍPIO', font=fonte, bd=5).place(x=80, y=170)
            Label(cad_frame, text='OPT. SIMPLES', font=fonte, bd=5).place(x=80, y=220)

            def mascara_cad(ev): # função para formatar CNPJ
                mask = cad_cnpj.get()
                if '/' not in mask and len(mask) >= 14:
                    mask_cnpj = f'{mask[:2]}.{mask[2:5]}.{mask[5:8]}/{mask[8:12]}-{mask[12:14]}'
                    cad_cnpj.delete(0, END)
                    cad_cnpj.insert(0, mask_cnpj)

            def pesquisar_fornecedor():
                try:
                    lmdb = os.getcwd() + '\Base_notas.accdb;'
                    cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                    cursor = cnx.cursor()
                    cursor.execute('SELECT * FROM cadastro WHERE CNPJ=?', (cad_cnpj.get(),))
                    row = cursor.fetchone()
                    cad_cnpj.delete(0, END)
                    cad_nome.delete(0, END)
                    cad_mun.delete(0, END)
                    cad_simples.delete(0, END)
                    cad_cnpj.insert(0, row[0])
                    cad_nome.insert(0, row[1])
                    cad_mun.insert(0, row[2])
                    cad_simples.insert(0, row[3])
                    cnx.commit()
                    cnx.close()
                except:
                    tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro não encontrado!')
                    cadastro.lift()


            # Botões
            cad_cnpj = Entry(cad_frame, width=30, bd=5, font=fonte)
            cad_cnpj.place(x=240, y=70)
            cad_cnpj.bind('<FocusOut>', mascara_cad)
            cad_nome = Entry(cad_frame, width=30, bd=5, font=fonte)
            cad_nome.place(x=240, y=120)
            cad_mun = Entry(cad_frame, width=30, bd=5, font=fonte)
            cad_mun.place(x=240, y=170)
            cad_simples = ttk.Combobox(cad_frame, font=('@Microsoft YaHei', 11, 'bold'), width=30)
            cad_simples['values'] = ('Não', 'Sim')
            cad_simples.current(0)
            cad_simples.place(x=240, y=220)

            def cadastrar_prestador():
                if cad_cnpj.get() == '':
                    tkinter.messagebox.showerror('Notas fiscais de Serviço', 'Coloque todas as informações')
                else:
                    try:
                        lmdb = os.getcwd() + '\Base_notas.accdb;'
                        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                        cursor = cnx.cursor()
                        cursor.execute('INSERT INTO cadastro values (?, ?, ?, ?)', (cad_cnpj.get(), cad_nome.get(),
                                                        cad_mun.get(), cad_simples.get()))
                        cnx.commit()
                        tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro incluído com sucesso!')
                        cnx.close()

                    except:
                        tkinter.messagebox.showerror('Notas Fiscais de Serviço', 'Erro! CNPJ já cadastrado!')
                        cadastro.lift()


            def atualizar_cadastro():
                lmdb = os.getcwd() + '\Base_notas.accdb;'
                self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                cursor = self.cnx.cursor()
                cursor.execute('UPDATE cadastro SET NOME=?, MUNICÍPIO=?, OPTANTE_SIMPLES=? WHERE CNPJ=?',
                               (cad_nome.get(),
                cad_mun.get(),
                cad_simples.get(),
                cad_cnpj.get()))
                self.cnx.commit()
                self.cnx.close()
                tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro incluído com sucesso!')

            btn_pesq = Button(cad_frame, font=('arial', 14, 'bold'), text='Pesquisar', bd=4, pady=1, padx=24,
                             width=7, height=1, command=pesquisar_fornecedor).place(x=70, y=300)

            btn_cad = Button(cad_frame, font=('arial', 14, 'bold'), text='Cadastrar', bd=4, pady=1, padx=24,
                             width=7, height=1, command=cadastrar_prestador).place(x=230, y=300)

            btn_atual = Button(cad_frame, font=('arial', 14, 'bold'), text='Atualizar', bd=4, pady=1, padx=24,
                               width=7, height=1, command=atualizar_cadastro).place(x=390,y=300)

        menu_cadastro = Menu(meu_menu, tearoff=0)
        meu_menu.add_cascade(label='Cadastro', menu=menu_cadastro)
        menu_cadastro.add_command(label='Prestadores', command=tela_cadastro)

        #==================================MENU CONSULTA SERVIÇOS=====================================================#
        def tela_servicos():
            servicos = Toplevel()
            servicos.title('Cadastro de Prestadores')
            servicos.geometry('700x500')

            serv_frame = Frame(servicos, bd=5, width=700, height=500, relief=RIDGE)
            serv_frame.place(x=0, y=0)
            Label(serv_frame, text='CODIGO SERVIÇO', font=fonte, bd=5).place(x=80, y=70)
            Label(serv_frame, text='DESCRIÇÃO', font=fonte, bd=5).place(x=80, y=110)
            Label(serv_frame, text='IRRF', font=fonte, bd=5).place(x=80, y=150)
            Label(serv_frame, text='CRF', font=fonte, bd=5).place(x=80, y=190)
            Label(serv_frame, text='INSS', font=fonte, bd=5).place(x=80, y=230)
            Label(serv_frame, text='ISS', font=fonte, bd=5).place(x=80, y=270)

            def pesquisar_servico():
                if cad_ser.get() == '':
                    tkinter.messagebox.showerror('Notas fiscais de Serviço', 'Serviço Inválido!')
                else:
                    try:
                        lmdb = os.getcwd() + '\Base_notas.accdb;'
                        self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                        cursor = self.cnx.cursor()
                        cursor.execute('SELECT * FROM tabela_iss WHERE servico=?', (cad_ser.get(),))
                        row = cursor.fetchone()
                        cad_ser.delete(0, END)
                        cad_desc.delete(0, END)
                        cad_irrf.delete(0, END)
                        cad_crf.delete(0, END)
                        cad_inss.delete(0, END)
                        cad_iss.delete(0, END)
                        cad_ser.insert(0, row[0])
                        cad_desc.insert(0, row[1])
                        cad_irrf.insert(0, row[2])
                        cad_crf.insert(0, row[3])
                        cad_inss.insert(0, row[4])
                        cad_iss.insert(0, row[5])
                        self.cnx.commit()
                    except:
                        tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Serviço não encontrado!')
                        servicos.lift()
                    self.cnx.close()

            cad_ser = Entry(serv_frame, width=20, bd=5, font=fonte)
            cad_ser.place(x=240, y=70)
            cad_desc = Entry(serv_frame, width=20, bd=5, font=fonte)
            cad_desc.place(x=240, y=110)
            cad_irrf = Entry(serv_frame, width=20, bd=5, font=fonte)
            cad_irrf.place(x=240, y=150)
            cad_crf = Entry(serv_frame, width=20, bd=5, font=fonte)
            cad_crf.place(x=240, y=190)
            cad_inss = Entry(serv_frame, width=20, bd=5, font=fonte)
            cad_inss.place(x=240, y=230)
            cad_iss = Entry(serv_frame, width=20, bd=5, font=fonte)
            cad_iss.place(x=240, y=270)

            btn_serv = Button(serv_frame, font=('arial', 14, 'bold'), text='Pesquisar', bd=4, pady=1, padx=24,
                               width=8, height=1, command=pesquisar_servico).place(x=250, y=350)

        menu_cadastro.add_command(label='Consulta Serviços', command=tela_servicos)


        # ================================== PROGRAMA PRINCIPAL ======================================================#

        # definindo labels
        labels = {'data_ana': 'Data Análise', 'data_nota': 'Data Nota', 'venc_nota': 'Data Vencimento',
                  'num_nota': 'Número da Nota', 'cnpj': 'CNPJ', 'forn': 'Fornecedor', 'mun_iss': 'Município ISS',
                  'simples': 'Simples Nacional', 'cod': 'Código Serviço', 'v_bruto': 'Valor Bruto',
                  'aliq_irrf': 'Alíq. IRRF', 'irrf': 'IRRF', 'aliq_crf': 'Aliq. CRF', 'crf': 'CRF',
                  'aliq_inss':'Aliq. INSS', 'inss': 'INSS', 'aliq_iss': 'Aliq. ISS', 'iss': 'ISS',
                            'v_liqui': 'Valor Líquido'}


        fonte = ('@Microsoft YaHei',11, 'bold') # Fonte Padrão

        #Definindo labels com caracteristicas similares
        px=25
        py=15
        cont = 0
        # variaveis = []
        for i, v in labels.items():
            if cont % 3 == 0 and px == 1075:
                py += 35
                px -= 1050
            self.ldata_nota = Label(leftframe, font=fonte, text=v, bd=7).place(x=px, y=py)
            px += 350
            cont += 1
            if cont == 10:
                break

        # labels restantes
        Label(leftframe, font=fonte, text=labels['aliq_irrf'], bd=7).place(x=25, y=180)
        Label(leftframe, font=fonte, text=labels['irrf'], bd=7).place(x=190, y=180)
        Label(leftframe, font=fonte, text=labels['aliq_crf'], bd=7).place(x=25, y=220)
        Label(leftframe, font=fonte, text=labels['crf'], bd=7).place(x=190, y=220)
        Label(leftframe, font=fonte, text=labels['aliq_inss'], bd=7).place(x=25, y=260)
        Label(leftframe, font=fonte, text=labels['inss'], bd=7).place(x=190, y=260)
        Label(leftframe, font=fonte, text=labels['aliq_iss'], bd=7).place(x=25, y=300)
        Label(leftframe, font=fonte, text=labels['iss'], bd=7).place(x=190, y=300)
        Label(leftframe, font=fonte, text=labels['v_liqui'], bd=7).place(x=450, y=300)

        # Funções para auxiliar e formatar o preenchimento dos dados
        def mascara(ev):
            mask = self.cnpj.get()
            if mask != '' and '/' not in mask and len(mask) >= 14:
                mask_cnpj = f'{mask[:2]}.{mask[2:5]}.{mask[5:8]}/{mask[8:12]}-{mask[12:14]}'
                self.cnpj.delete(0, END)
                self.cnpj.insert(0, mask_cnpj)
            else:
                pass

        # def ignore(ev):
        #     if self.cnpj.get() != '':
        #         return "break"

        def busca_cadastro(ev):
            # self.forn.bind('<FocusIn>', ignore)
            if self.cnpj.get() != '':
                lmdb = os.getcwd() + '\Base_notas.accdb;'
                self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                cursor = self.cnx.cursor()
                cursor.execute('select nome from cadastro where cnpj = ?', (self.cnpj.get(),))
                busca_nome = cursor.fetchone()
                # if busca_nome == None:
                #     tkinter.messagebox.showinfo('Notas fiscais de Serviço', 'Fornecedor não cadastrado!')
                    # resp = tkinter.messagebox.askquestion('Notas Fiscais de Serviço', 'Fornecedor não cadastrado. Deseja cadastrar?')
                    # if resp == 'yes':
                    #     tela_cadastro()
                    # elif resp == 'no':
                    #     pass
                # else:
                self.forn.delete(0, END)
                self.forn.insert(0, busca_nome[0])

                cursor.execute('select optante_simples from cadastro where cnpj = ?', (self.cnpj.get(),))
                busca_simples = cursor.fetchone()
                self.simples.delete(0, END)
                self.simples.insert(0, busca_simples[0])


        def busca_servico(ev):
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = self.cnx.cursor()
            if self.simples.get() in 'nãoNÃOnaoNAONão':
                if self.cod.get() != '':
                    lista = {'irrf': self.aliq_ir, 'crf': self.aliq_crf, 'inss': self.aliq_inss, 'iss': self.aliq_iss}
                    for imp, aliq in lista.items():
                        cursor.execute(f'select {imp} from tabela_iss where servico = ?', (self.cod.get(),))
                        busca = cursor.fetchone()
                        aliq.delete(0, END)
                        aliq.insert(0, str(round(busca[0],2)).replace('.',','))
            cursor.execute(f'select descricao from tabela_iss where servico = ?', (self.cod.get(),))
            busca2 = cursor.fetchone()
            self.descr_serv.delete(1.0, 'end')
            self.descr_serv.insert('end', busca2[0])



        def calcula_irrf(ev):
            if self.aliq_ir.get() != '':
                tupla_ir = (self.aliq_ir.get().replace(',', '.'), self.v_bruto.get().replace(',', '.'))
                self.irrf.delete(0, END)
                self.irrf.insert(0, str(round(float(tupla_ir[1]) * (float(tupla_ir[0]) / 100), 2)).replace('.',','))
            else:
                self.irrf.delete(0, END)
                self.irrf.insert(0, 0)
                self.aliq_ir.insert(0, 0)

        def calcula_crf(ev):
            if self.aliq_crf.get() != '':
                tupla_crf = (self.aliq_crf.get().replace(',', '.'), self.v_bruto.get().replace(',', '.'))
                self.crf.delete(0, END)
                self.crf.insert(0, str(round(float(tupla_crf[1]) * (float(tupla_crf[0]) / 100), 2)).replace('.',','))
            else:
                self.crf.delete(0, END)
                self.crf.insert(0, 0)
                self.aliq_crf.insert(0, 0)

        def calcula_inss(ev):
            if self.aliq_inss.get() != '':
                tupla_inss = (self.aliq_inss.get().replace(',', '.'), self.v_bruto.get().replace(',', '.'))
                self.inss.delete(0, END)
                self.inss.insert(0, str(round(float(tupla_inss[1]) * (float(tupla_inss[0]) / 100), 2)).replace('.',','))
            else:
                self.inss.delete(0, END)
                self.inss.insert(0, 0)
                self.aliq_inss.insert(0, 0)

        def calcula_iss(ev):
            if self.aliq_iss.get() != '':
                tupla_iss = (self.aliq_iss.get().replace(',', '.'), self.v_bruto.get().replace(',', '.'))
                self.iss.delete(0, END)
                self.iss.insert(0, str(round(float(tupla_iss[1]) * (float(tupla_iss[0]) / 100), 2)).replace('.',','))
            else:
                self.iss.delete(0, END)
                self.iss.insert(0, 0)
                self.aliq_iss.insert(0, 0)

        def valor_liq(ev):
            self.v_liq.delete(0, END)
            self.v_liq.insert(0, str(round(float(self.v_bruto.get().replace(',','.')) -
            (sum([float(self.irrf.get().replace(',','.')), float(self.crf.get().replace(',','.')),
            float(self.inss.get().replace(',','.')), float(self.iss.get().replace(',','.'))])), 2)).replace('.',','))

        def data_dia(ev):
            if self.data_nota.get() == '':
                self.data_ana.delete(0, END)
                self.data_ana.insert(0, date.today().strftime('%d/%m/%Y'))
            else:
                pass

        # Definindo entradas
        self.id = Entry(leftframe)
        self.data_ana = Entry(leftframe, width=15, font=fonte, bd=4)
        self.data_ana.place(x=180, y=18)
        self.data_ana.bind('<Motion>', data_dia)
        self.data_nota = Entry(leftframe, width=15, font=fonte, bd=4)
        self.data_nota.place(x=530, y=18)
        self.data_venc = Entry(leftframe, width=15, font=fonte, bd=4)
        self.data_venc.place(x=880, y=18)
        # self.data_venc.bind('<FocusOut>', converter_datas)
        self.num_nota = Entry(leftframe, width=15, font=fonte, bd=4)
        self.num_nota.place(x=180, y=53)
        self.cnpj = Entry(leftframe, width=20, font=fonte, bd=4)
        self.cnpj.place(x=530, y=53)
        self.cnpj.bind("<FocusOut>", mascara)
        self.forn = Entry(leftframe, width=30, font=fonte, bd=4)
        self.forn.place(x=880, y=53)
        self.forn.bind("<FocusIn>", busca_cadastro)
        self.mun_iss = Entry(leftframe, width=15, font=fonte, bd=4)
        self.mun_iss.place(x=180, y=88)
        self.simples = Entry(leftframe, width=15, font=fonte, bd=4)
        self.simples.place(x=530, y=88)
        self.cod = Entry(leftframe, width=14, font=fonte, bd=4)
        self.cod.place(x=880, y=88)
        self.cod.bind('<FocusOut>', busca_servico)
        self.v_bruto = Entry(leftframe, width=18, font=fonte, bd=4)
        self.v_bruto.place(x=180, y=122)
        self.aliq_ir = Entry(leftframe, width=5, font=fonte, bd=4)
        self.aliq_ir.place(x=115, y=180)
        self.irrf = Entry(leftframe, width=15, font=fonte, bd=4)
        self.irrf.place(x=255, y=180)
        self.irrf.bind('<FocusIn>', calcula_irrf)
        self.aliq_crf = Entry(leftframe, width=5, font=fonte, bd=4)
        self.aliq_crf.place(x=115, y=220)
        self.crf = Entry(leftframe, width=15, font=fonte, bd=4)
        self.crf.place(x=255, y=220)
        self.crf.bind('<FocusIn>', calcula_crf)
        self.aliq_inss = Entry(leftframe, width=5, font=fonte, bd=4)
        self.aliq_inss.place(x=115, y=260)
        self.inss = Entry(leftframe, width=15, font=fonte, bd=4)
        self.inss.place(x=255, y=260)
        self.inss.bind('<FocusIn>', calcula_inss)
        self.aliq_iss = Entry(leftframe, width=5, font=fonte, bd=4)
        self.aliq_iss.place(x=115, y=300)
        self.iss = Entry(leftframe, width=15, font=fonte, bd=4)
        self.iss.place(x=255, y=300)
        self.iss.bind('<FocusIn>', calcula_iss)
        self.v_liq = Entry(leftframe, width=15, font=fonte, bd=4)
        self.v_liq.place(x=580, y=300)
        self.v_liq.bind('<FocusIn>', valor_liq)
        self.descr_serv = Text(leftframe, width=35, height=5, font=('arial'), bg='white smoke', bd=0)
        self.descr_serv.place(x=830, y=130)


        # ================= ver tabela (TREEVIEW) ==================================================================#
        scroll_y = Scrollbar(self.bottonframe, orient=VERTICAL)
        colunas = ['ID', 'DT_ANÁ', 'DT_NF', 'DT_Venc', 'NF', 'CNPJ', 'Fornecedor', 'Município', 'Simples Nacional',
                   'Cod. Serviço', 'Val Bruto', 'Aliq_IR', 'IRRF', 'Aliq_CRF', 'CRF', 'Aliq_INSS', 'INSS',
                   'Aliq_ISS', 'ISS', 'Val Líq']
        self.lista_notas = ttk.Treeview(self.bottonframe, height=5, columns=colunas,
                                        yscrollcommand=scroll_y.set)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_y.config(command=self.lista_notas.yview)
        for i in colunas:
            self.lista_notas.heading(i, text=i[:8])

        self.lista_notas['show'] = 'headings'

        for coluna in colunas:
            self.lista_notas.column(coluna, width=50)
        self.lista_notas.column('DT_ANÁ', width=70)
        self.lista_notas.column('DT_NF', width=70)
        self.lista_notas.column('DT_Venc', width=70)
        self.lista_notas.column('ID', width=30)
        self.lista_notas.column('CNPJ', width=110)
        self.lista_notas.column('Fornecedor', width=110)

        self.lista_notas.pack(fill=BOTH, expand=1)
        self.lista_notas.bind('<ButtonRelease-1>', self.NotasInfo)
        self.mostrar_dados()

    # Funções de comando para integração do programa principal com banco de dados
    def isaida(self):
        isaida = tkinter.messagebox.askyesno('Notas Fiscais de Serviço', 'Confirma a saída?')
        if isaida > 0:
            self.janela.destroy()
            return

    def adicionar(self):
        if self.cnpj.get() == '':
            tkinter.messagebox.showerror('Notas fiscais de Serviço', 'Coloque todas as informações')
        else:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('INSERT INTO notas_fiscais (data_analise, data, data_vencimento, NF,	CNPJ, Fornecedor, cidade,'
                           'simples_nacional, codigo_servico, valor_bruto, aliq_irrf, irrf,	aliq_crf, crf, aliq_inss, '
                           'inss,	aliq_iss, iss, valor_liquido) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,'
                           ' ?, ?, ?, ?)', (self.data_ana.get(),
            self.data_nota.get(),
            self.data_venc.get(),
            self.num_nota.get(),
            self.cnpj.get(),
            self.forn.get(),
            self.mun_iss.get(),
            self.simples.get(),
            self.cod.get(),
            self.v_bruto.get(),
            self.aliq_ir.get(),
            self.irrf.get(),
            self.aliq_crf.get(),
            self.crf.get(),
            self.aliq_inss.get(),
            self.inss.get(),
            self.aliq_iss.get(),
            self.iss.get(),
            self.v_liq.get()))
            self.lembrar.set(0)
            self.limpar()
            cnx.commit()
            cnx.close()
            self.mostrar_dados()
            tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro incluído com sucesso!')


    def limpar(self):
        self.data_ana.delete(0, END),
        self.data_nota.delete(0, END),
        self.data_venc.delete(0, END),
        self.num_nota.delete(0, END),
        self.cnpj.delete(0, END),
        self.forn.delete(0, END),
        self.mun_iss.delete(0, END),
        self.simples.delete(0, END),
        self.cod.delete(0, END),
        self.v_bruto.delete(0, END),
        self.aliq_ir.delete(0, END),
        self.irrf.delete(0, END),
        self.aliq_crf.delete(0, END),
        self.crf.delete(0, END),
        self.aliq_inss.delete(0, END),
        self.inss.delete(0, END),
        self.aliq_iss.delete(0, END),
        self.iss.delete(0, END),
        self.v_liq.delete(0, END)
        self.descr_serv.delete('1.0', END)

    def mostrar_dados(self):
        self.limpar()
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('select * from notas_fiscais order by ID desc')
        resultado = cursor.fetchall()
        if len(resultado) != 0:
            self.lista_notas.delete(*self.lista_notas.get_children())
            for row in resultado:
                self.lista_notas.insert(parent='', index='end', text='', values=(row[0], row[1], row[2], row[3],
                row[4], row[5], row[6], row[7], row[8], row[9], str(round(row[10],2)).replace('.',','),
                str(round(row[11],2)).replace('.',','), str(round(row[12],2)).replace('.',','),
                str(round(row[13],2)).replace('.',','), str(round(row[14],2)).replace('.',','),
                str(round(row[15],2)).replace('.',','), str(round(row[16],2)).replace('.',','),
                str(round(row[17],2)).replace('.',','), str(round(row[18],2)).replace('.',','),
                str(round(row[19],2)).replace('.',',')))
        cnx.commit()
        cnx.close()

    def NotasInfo(self, ev):
        self.limpar()
        verinfo = self.lista_notas.focus()
        dados = self.lista_notas.item(verinfo)
        row = dados['values']
        self.id.delete(0, END)
        self.id.insert(0, row[0])
        self.data_ana.insert(0, row[1])
        self.data_nota.insert(0, row[2])
        self.data_venc.insert(0, row[3])
        self.num_nota.insert(0, row[4])
        self.cnpj.insert(0, row[5])
        self.forn.insert(0, row[6])
        self.mun_iss.insert(0, row[7])
        self.simples.insert(0, row[8])
        self.cod.insert(0, row[9])
        self.v_bruto.insert(0, row[10])
        self.aliq_ir.insert(0, row[11])
        self.irrf.insert(0, row[12])
        self.aliq_crf.insert(0, row[13])
        self.crf.insert(0, row[14])
        self.aliq_inss.insert(0, row[15])
        self.inss.insert(0, row[16])
        self.aliq_iss.insert(0, row[17])
        self.iss.insert(0, row[18])
        self.v_liq.insert(0, row[19])

    def deletar(self):
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('DELETE FROM notas_fiscais WHERE ID=?', (self.id.get(),))
        cnx.commit()
        cnx.close()
        tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro apagado com sucesso!')
        self.mostrar_dados()
        self.limpar()

    def procurar(self):
        try:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('select * FROM notas_fiscais WHERE NF=?', (self.num_nota.get(),))
            row = cursor.fetchone()
            self.num_nota.delete(0, END)
            self.id.insert(0, row[0])
            self.data_ana.insert(0, row[1])
            self.data_nota.insert(0, row[2])
            self.data_venc.insert(0, row[3])
            self.num_nota.insert(0, row[4])
            self.cnpj.insert(0, row[5])
            self.forn.insert(0, row[6])
            self.mun_iss.insert(0, row[7])
            self.simples.insert(0, row[8])
            self.cod.insert(0, row[9])
            self.v_bruto.insert(0, str(round(row[10],2)).replace('.',','))
            self.aliq_ir.insert(0, str(round(row[11],2)).replace('.',','))
            self.irrf.insert(0, str(round(row[12],2)).replace('.',','))
            self.aliq_crf.insert(0, str(round(row[13],2)).replace('.',','))
            self.crf.insert(0, str(round(row[14],2)).replace('.',','))
            self.aliq_inss.insert(0, str(round(row[15],2)).replace('.',','))
            self.inss.insert(0, str(round(row[16],2)).replace('.',','))
            self.aliq_iss.insert(0, str(round(row[17],2)).replace('.',','))
            self.iss.insert(0, str(round(row[18],2)).replace('.',','))
            self.v_liq.insert(0, str(round(row[19],2)).replace('.',','))
            cnx.commit()
            cnx.close()
        except:
            tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro não encontrado!')
            self.limpar()


    def atualizar(self):
        try:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('update notas_fiscais set DATA_ANALISE=?, DATA=?, DATA_VENCIMENTO=?, NF=?, CNPJ=?, FORNECEDOR=?, '
                           'CIDADE=?, SIMPLES_NACIONAL=?, CODIGO_SERVICO=?, VALOR_BRUTO=?, ALIQ_IRRF=?, IRRF=?, ALIQ_CRF=?, '
                           'crf=?, ALIQ_INSS=?, INSS=?, ALIQ_ISS=?, ISS=?, VALOR_LIQUIDO=? where ID=?',(self.data_ana.get(),
            self.data_nota.get(),
            self.data_venc.get(),
            self.num_nota.get(),
            self.cnpj.get(),
            self.forn.get(),
            self.mun_iss.get(),
            self.simples.get(),
            self.cod.get(),
            self.v_bruto.get(),
            self.aliq_ir.get(),
            self.irrf.get(),
            self.aliq_crf.get(),
            self.crf.get(),
            self.aliq_inss.get(),
            self.inss.get(),
            self.aliq_iss.get(),
            self.iss.get(),
            self.v_liq.get(),
            self.id.get()))
            cnx.commit()
            cnx.close()
            tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro alterado com sucesso!')
            self.mostrar_dados()
        except:
            tkinter.messagebox.showerror('Notas Fiscais de Serviço', 'Erro!')

    def lembrar_lancamento(self):
        if self.lembrar.get() == 1:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('SELECT TOP 1 data_analise, data, data_vencimento, nf, cnpj, fornecedor, simples_nacional,'
                           'codigo_servico from notas_fiscais order by id desc')
            row = cursor.fetchone()
            self.data_ana.insert(0, row[0])
            self.data_nota.insert(0, row[1]),
            self.data_venc.insert(0, row[2]),
            self.num_nota.insert(0, row[3]+1),
            self.cnpj.insert(0, row[4]),
            self.forn.insert(0, row[5]),
            self.simples.insert(0, row[6]),
            self.cod.insert(0, row[7])
            cnx.commit()
        else:
            self.limpar()



if __name__=='__main__':
    janela = Tk()
    aplicacao = NotasServicos(janela)
    janela.mainloop()