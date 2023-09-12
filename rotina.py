import openpyxl
import PySimpleGUI as sg


workbook = openpyxl.Workbook()
sheet = workbook.active

headers = ['Dia e Horário', 'Assunto', 'Subtemas', 'Atividades de Estudo']
sheet.append(headers)


topics = [
    ("Dia 1", "Aritmética, Gráficos e Tabelas, Análise Combinatória", "10:15 - 12:00", "Operações Fundamentais, Propriedades dos Números, Problemas de Aritmética, Interpretação de Gráficos, Análise de Dados, Conceito de Contagem", "Pratique cálculos de adição, subtração, multiplicação e divisão, Estude números primos, múltiplos e divisores, Resolva exercícios de aplicação de aritmética, Aprenda a ler e interpretar diferentes tipos de gráficos, Identifique tendências e padrões em conjuntos de dados, Entenda como realizar contagem de possibilidades"),
    ("Dia 2", "Sequências, Números Inteiros, Trigonometria", "8:00 - 9:30", "Sequências Numéricas, Progressões Aritméticas e Geométricas, Propriedades dos Números Inteiros, Conceitos Básicos de Trigonometria", "Estude padrões em sequências numéricas, Entenda as características dessas sequências, Explore os números inteiros e suas propriedades, Aprenda sobre seno, cosseno e tangente"),
    ("Dia 2", "Álgebra, Estatística, Polígonos, Funções e Equações", "9:45 - 11:15", "Equações e Inequações, Fatoração e Simplificação, Média Aritmética, Moda e Mediana, Distribuição de Frequência e Histogramas, Classificação de Polígonos, Propriedades dos Polígonos Regulares, Conceitos de Funções, Equações e Inequações de 1° e 2° Graus", "Resolva equações e inequações de primeiro e segundo grau, Aprenda técnicas de fatoração e simplificação de expressões, Entenda medidas de tendência central, Identifique tendências em conjuntos de dados, Aprenda a classificar polígonos, Explore as características dos polígonos regulares, Entenda o que são funções e suas propriedades, Resolva equações e inequações"),
    ("Dia 3", "Mecânica, Eletricidade, Ondulatória, Cinemática, Dinâmica e Estática", "8:00 - 9:30", "Conceitos Básicos de Mecânica, Fundamentos da Eletricidade, Introdução à Ondulatória, Estudo do Movimento (Cinemática), Leis do Movimento (Dinâmica), Estudo de Objetos em Repouso (Estática)", "Estude movimento, força e energia, Explore os conceitos básicos de eletricidade, Aprenda sobre ondas e propagação, Compreenda os conceitos de cinemática, Explore as leis do movimento, Aprenda sobre objetos em repouso"),
    ("Dia 3", "Gravitação, Energia, Impulso e Quantidade de Movimento, Termodinâmica, Óptica, Eletromagnetismo, Radiação", "9:45 - 11:15", "Conceitos de Gravitação, Energia, Impulso e Quantidade de Movimento, Fundamentos da Termodinâmica, Propriedades da Luz e Óptica, Conceitos Básicos de Eletromagnetismo, Noções sobre Radiação Eletromagnética", "Aprenda sobre a influência da gravidade, Compreenda conceitos de energia, Explore impulso e quantidade de movimento, Entenda princípios da termodinâmica, Estude as propriedades da luz e conceitos ópticos, Aprenda sobre campos eletromagnéticos, Entenda os princípios da radiação eletromagnética"),
    
    ("Dia 4", "Concentração das Soluções, Funções Inorgânicas (Ácidos, Bases, Sais)", "8:00 - 9:30", "Conceito de Concentração das Soluções, Tipos de Concentração, Propriedades e Características dos Ácidos, Bases e Sais", "Aprenda a calcular e compreender diferentes tipos de concentração das soluções, Estude as propriedades e características das funções inorgânicas"),
    ("Dia 4", "Reações Orgânicas e Propriedades dos Materiais, Transformações Químicas e Estequiometria", "9:45 - 11:15", "Estudo de Reações Orgânicas, Propriedades dos Materiais Químicos, Conceitos de Transformações Químicas, Cálculos Estequiométricos", "Explore as reações orgânicas e as propriedades dos materiais químicos, Aprenda a realizar cálculos estequiométricos em transformações químicas"),
    ("Dia 5", "Cinética e Equilíbrio Químico, Termoquímica, Ácidos e Bases, Eletroquímica, Orgânica, Polímeros, Biotecnologia", "8:00 - 10:00", "Cinética e Equilíbrio de Reações Químicas, Conceitos de Termoquímica, Propriedades e Características de Ácidos e Bases, Estudo de Reações Eletroquímicas, Fundamentos da Química Orgânica, Introdução aos Polímeros e à Biotecnologia", "Compreenda a cinética e o equilíbrio de reações químicas, Estude termoquímica, ácidos e bases, reações eletroquímicas, química orgânica, polímeros e biotecnologia"),
    ("Dia 5", "Saúde, Doenças e Meio Ambiente, Ecologia, Citologia, Genética, Histologia, Anatomia e Fisiologia, Evolução, Botânica e Zoologia, Ecologia de Populações, Comunidades e Ecossistemas", "10:15 - 12:00", "Relações entre Saúde, Doenças e o Meio Ambiente, Conceitos Básicos de Ecologia, Estudo das Células (Citologia), Conceitos Básicos de Genética, Introdução à Histologia (Tecidos), Exploração da Anatomia e Fisiologia do Corpo Humano, Noções sobre a Teoria da Evolução, Conceitos Básicos de Botânica e Zoologia, Estudo de Ecossistemas, Comunidades e Populações", "Entenda a relação entre saúde, doenças e meio ambiente, Estude conceitos básicos de ecologia, a estrutura celular, genética, histologia, anatomia e fisiologia, evolução, botânica, zoologia, ecossistemas, comunidades e populações")
]



sg.theme('DarkBlue')

layout = [
    [sg.Text('Dia e Horário', size=(12, 0)), sg.Input(key='1', size=(20, 0))],
    [sg.Text('Assunto', size=(12, 0)), sg.Input(size=(20, 0), key='2')],
    [sg.Text('Subtemas(separe por virgula)'), sg.Multiline(size=(40, 5), key='3')],
    [sg.Text('Atividades de Estudo'), sg.Multiline(size=(40, 5), key='4')],
    [sg.Button('Salvar')]
]

window = sg.Window('Cadastro de Estudos', layout)

topics = []

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break
    elif event == 'Salvar':
        day_time = values['1']
        subject = values['2']
        subtopics = values['3']
        activities = values['4']

        topics.append((day_time, subject, subtopics, activities))

        sg.popup('Estudo Cadastrado')
        window['1'].update('')
        window['2'].update('')
        window['3'].update('')
        window['4'].update('')

window.close()
print (topics)

for topic in topics:
    sheet.append(topic)
planilha = input('Qual nome da sua planilha: ') + '.xlsx'
workbook.save(planilha)
print("Planilha criada com sucesso!")
