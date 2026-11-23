"""Conteúdo dos manuais das 4 plataformas novas de 22/11/2026 (Gerencial, Painel do Time B, Central de Cálculos e
Central de Base de Dados). O mesmo texto gera o manual em HTML e a apresentação em PowerPoint (gerar_manuais.py).

Regras: sem travessão, sem senha escrita (o manual fica num site público), sem nome de reclamante, testemunha ou
cliente com dado sensível. Cada seção: (título, resumo de uma linha para o slide, blocos).
Blocos: ("p", texto) parágrafo; ("lista", [itens]) tópicos; ("card", cor, título, texto) destaque;
("passos", [(título, texto)]) passo a passo numerado; ("tabela", [cabeçalho], [[linhas]]).
"""

ACESSO_PADRAO = ("card", "cyan", "Senha de acesso",
                 "O card no painel LAWgico pede a senha de acesso das ferramentas restritas. A senha é passada pela Controladoria "
                 "e não fica escrita neste manual. Depois de acertar uma vez, os outros cards abrem direto enquanto a aba estiver aberta.")

PLATAFORMAS = [
    {
        "id": "Gerencial", "pasta": "Gerencial", "arquivo": "LAWgico_Gerencial_Manual",
        "nome": "Gerencial", "destaque": "Visão Geral ao Vivo",
        "url": "https://peixotoecury.github.io/gerencial/",
        "sub": "Visão gerencial do contencioso trabalhista do escritório, lida direto do LM e atualizada a cada 15 minutos.",
        "secoes": [
            ("Para que serve", "Uma única tela com a carteira, a equipe e a produção do contencioso trabalhista.", [
                ("p", "O Gerencial reúne em uma página a carteira trabalhista ativa do escritório, as entradas e saídas de processos, "
                      "a divisão por grupo, a composição de cada equipe e a produção de cada pessoa (prazos baixados, audiências e agendas)."),
                ("p", "Foi pensado para a coordenação e para os sócios acompanharem a operação sem depender de planilhas: os números "
                      "saem do próprio LM e se atualizam sozinhos durante o expediente."),
                ("lista", ["Quantos processos ativos o escritório tem e como estão distribuídos entre os grupos",
                           "Quantos processos entraram e saíram no mês e no ano",
                           "Quem compõe cada grupo, por cargo, e quantos processos há por advogado",
                           "Quem está com a maior carga e quais pontos pedem atenção"]),
            ]),
            ("Como acessar", "Pelo card no painel LAWgico ou direto pelo endereço, com senha.", [
                ("passos", [("Abra o painel LAWgico", "No painel de ferramentas, procure o card Gerencial · Visão Geral ao Vivo (aba Novidades ou Gestão interna)."),
                            ("Digite a senha do card", "O card abre a plataforma em uma nova aba."),
                            ("Digite a senha da plataforma", "O Gerencial tem uma tela de entrada própria. Depois de entrar, a aba não pede a senha de novo até ser fechada.")]),
                ACESSO_PADRAO,
                ("card", "amber", "Endereço", "peixotoecury.github.io/gerencial. Funciona no computador, no tablet e no celular, de qualquer lugar."),
            ]),
            ("De onde vêm os dados", "LM ao vivo a cada 15 minutos; equipe pelo cadastro do LM e pelo Painel do Time B.", [
                ("lista", ["Carteira, entradas e saídas: view de processos do LM, área trabalhista, todos os status.",
                           "Prazos e agendas por pessoa: views de prazos e de agendas do LM.",
                           "Audiências por pessoa: compilado de audiências da Controladoria.",
                           "Equipe: cadastro de profissionais do LM; os grupos B1, B2 e B3 seguem o Painel do Time B, mantido pela Controladoria.",
                           "Ajustes confirmados pela Controladoria (cargos, grupos e quem saiu) valem por cima do cadastro."]),
                ("card", "cyan", "Atualização", "A cada 15 minutos, das 7h às 20h, em dias úteis. O selo AO VIVO no topo mostra a hora da última leitura."),
            ]),
            ("Seções da página", "Visão geral e 11 seções numeradas, com navegação no topo.", [
                ("tabela", ["Seção", "O que mostra"], [
                    ["Visão geral", "Processos ativos por grupo, pessoas por cargo, entradas e saídas do ano, prazos baixados, audiências e agendas do apoio"],
                    ["01 Entradas e saídas", "Processos que entraram e saíram mês a mês, com opção de tirar as saídas em lote por renúncia"],
                    ["02 Divisão da carteira", "Carteira de cada grupo por cliente e por advogado"],
                    ["03 Grupos e equipe", "Total por equipe (A, B, C e E), um cartão por grupo com processos, pessoas por cargo e processos por advogado, e a tabela do escritório"],
                    ["04 Carga dos advogados", "Carteira e prazos por mês de cada advogado e o índice de carga em relação à média"],
                    ["05 Apoio", "Agendas realizadas e em aberto de assistentes, analistas, estagiários e aprendizes"],
                    ["06 Mapa mensal", "Mapa de calor da produção de cada pessoa mês a mês"],
                    ["07 Por pessoa", "Evolução de uma pessoa: carteira, prazos, audiências e agendas"],
                    ["08 Pontos de atenção", "Alertas automáticos de carga acima da média e de volumetria fora do padrão"],
                    ["09 Por fase", "Advogados por fase processual"],
                    ["10 Clientes", "Entradas e saídas por cliente e grupo, com tendência"],
                    ["11 Saíram", "Quem saiu da equipe, com a última data de movimento"]]),
            ]),
            ("Filtros", "Grupo, função, cliente e pessoa, combináveis entre si.", [
                ("lista", ["Grupo: escolha um ou mais grupos (A1 a E2, Sócios, Sem grupo, Cível).",
                           "Função: advogados, coordenação ou apoio.",
                           "Cliente e Pessoas: listas com busca, com seleção de mais de um item.",
                           "Clique em um grupo, cliente ou pessoa dentro dos gráficos para filtrar a página inteira."]),
                ("card", "cyan", "Dica", "O resumo à direita da barra de filtros mostra quantas pessoas estão no recorte. Para voltar ao todo, use Todos no filtro de grupo."),
            ]),
            ("Cuidados", "Conteúdo interno; números mudam ao longo do dia.", [
                ("lista", ["Os números acompanham o LM: se um processo foi lançado errado no LM, ele aparece errado aqui.",
                           "Quem atua em mais de um grupo aparece em cada um deles; o total do escritório conta cada pessoa uma vez.",
                           "Uso interno do escritório. Não encaminhar prints para fora do P&C."]),
            ]),
        ],
    },
    {
        "id": "PainelTimeB", "pasta": "PainelTimeB", "arquivo": "LAWgico_PainelTimeB_Manual",
        "nome": "Painel do", "destaque": "Time B",
        "url": "https://peixotoecury.github.io/painellawgicoferramentas/PainelTimeB/",
        "sub": "Carga e divisão da carteira dos grupos B1, B2 e B3, volumetria do apoio, pontos de atenção e a transição da gestão.",
        "secoes": [
            ("Para que serve", "Acompanhar a carga do Time B e organizar a equipe.", [
                ("p", "O Painel do Time B mostra como a carteira dos grupos B1, B2 e B3 está dividida entre os advogados, quanto cada "
                      "pessoa produz por mês e onde há sobrecarga. Também reúne a transição da gestão do time: organograma, atividades de "
                      "cada pessoa, relatórios, escala e mudanças na equipe."),
                ("p", "É a fonte oficial da composição do Time B: o Gerencial usa esta lista para os grupos B1, B2 e B3."),
            ]),
            ("Como acessar", "Pelo card no painel LAWgico, com senha.", [
                ("passos", [("Abra o painel LAWgico", "Procure o card Painel do Time B (aba Novidades ou Gestão interna)."),
                            ("Digite a senha do card", "O painel abre em uma nova aba."),
                            ("Digite a senha do painel", "O conteúdo é cifrado e só abre no navegador, com a senha.")]),
                ACESSO_PADRAO,
            ]),
            ("De onde vêm os dados", "LM, Compilado de Baixados, Compilado de Audiências e a planilha de divisão do time.", [
                ("lista", ["Carteira: processos ativos dos grupos B1, B2 e B3 lidos do LM.",
                           "Prazos por advogado: Compilado de Baixados da Controladoria.",
                           "Audiências: Compilado de Audiências.",
                           "Volumetria do apoio: agendas realizadas no LM.",
                           "Transição, atividades, relatórios e escala: planilha Divisão do time B, mantida pela Controladoria."]),
                ("card", "amber", "Atualização", "O painel é regerado pela Controladoria quando chegam os compilados ou a planilha de divisão muda. A data aparece no topo."),
            ]),
            ("Seções do painel", "Visão geral, Transição e 8 seções numeradas.", [
                ("tabela", ["Seção", "O que mostra"], [
                    ["Visão geral", "Processos ativos por grupo, pessoas do time, prazos baixados, audiências e agendas do apoio"],
                    ["00 Transição da gestão", "Organograma, atividades que mudaram de responsável, atividades de todos, relatórios, escala e mudanças da equipe"],
                    ["01 Entradas e saídas", "Processos que entraram e saíram do Grupo B mês a mês"],
                    ["02 Divisão da carteira", "Gráfico de cada grupo por cliente e lista dos advogados de cada grupo"],
                    ["03 Carga dos advogados", "Carteira e prazos por mês de cada advogado e índice de carga em relação à média do time"],
                    ["04 Apoio", "Agendas de assistentes, analistas, estagiários e aprendiz"],
                    ["05 Mapa mensal", "Mapa de calor da produção mês a mês"],
                    ["06 Por pessoa", "Cartão de cada pessoa com três mini gráficos e a média do período"],
                    ["07 Pontos de atenção", "Alertas automáticos de sobrecarga e volumetria"],
                    ["08 Plano de ação", "Ações propostas, com status marcado por quem usa o painel"]]),
            ]),
            ("Filtros e interação", "Barra fixa no topo e cliques que filtram.", [
                ("lista", ["Grupo, função, cliente e busca por nome na barra fixa.",
                           "Clique em um grupo ou cliente para filtrar; clique em uma pessoa para abrir o cartão dela em Por pessoa.",
                           "Colunas das tabelas ordenáveis pelo título.",
                           "Tema claro e escuro no botão do topo."]),
                ("card", "cyan", "Plano de ação", "O status de cada ação fica salvo no navegador de quem marcou. Para combinar o status com a equipe, use a reunião do time."),
            ]),
            ("Cuidados", "Índice de carga é uma referência, não uma avaliação individual.", [
                ("lista", ["O índice compara carteira e prazos com a média de quem tem carteira relevante; coordenação e quem chegou há pouco ficam fora da média.",
                           "Quem não tem sigla no LM ainda aparece com dados parciais até o cadastro ser feito.",
                           "Uso interno do escritório."]),
            ]),
        ],
    },
    {
        "id": "Calculos", "pasta": "Calculos", "arquivo": "LAWgico_Calculos_Manual",
        "nome": "Central de", "destaque": "Cálculos",
        "url": "https://peixotoecury.github.io/calculos/",
        "sub": "Cálculos trabalhistas no padrão do escritório: leitura das decisões, revisão dos parâmetros, cálculo no navegador, impugnação e acordo.",
        "secoes": [
            ("Para que serve", "Fazer, revisar e guardar os cálculos trabalhistas da equipe em um só lugar.", [
                ("p", "A Central de Cálculos lê as peças do processo, ajuda a montar os parâmetros do cálculo a partir das decisões, "
                      "calcula no próprio navegador com as tabelas oficiais do PJe-Calc e gera as saídas: PDF, Excel, discriminação, "
                      "manifestação ou impugnação em Word e simulação de acordo."),
                ("p", "Cada cálculo fica salvo na central, por cliente, com o histórico de quem fez, revisou e aprovou."),
                ("card", "cyan", "Projeto piloto", "O piloto começa pelos clientes Bridgestone e CBC. O motor de cálculo é conferido caso a caso contra o PJe-Calc antes de cada mudança."),
            ]),
            ("Como acessar", "Pelo card no painel LAWgico e login próprio da central.", [
                ("passos", [("Abra o card LAWgico Cálculos P&C", "Digite a senha do card; a central abre em uma nova aba."),
                            ("Entre com o seu e-mail e senha", "Quem ainda não tem acesso clica em Pedir acesso e preenche nome e e-mail."),
                            ("Aguarde a aprovação", "A coordenação aprova o pedido em Pedidos de acesso. Esqueceu a senha? Use Esqueci minha senha.")]),
                ACESSO_PADRAO,
            ]),
            ("Tela inicial", "Cálculos da equipe, últimas movimentações e visão por cliente.", [
                ("lista", ["Cálculos da equipe: lista com situação de cada cálculo.",
                           "Últimas movimentações: o que foi criado, enviado para revisão, aprovado ou entregue.",
                           "Por cliente e por tipo de cálculo: atalhos para a carteira de cálculos.",
                           "Equipe (coordenação): pessoas com acesso e pedidos de acesso para aprovar ou recusar."]),
            ]),
            ("Novo cálculo em 4 etapas", "Peças, leitura das decisões, revisão dos parâmetros e cálculo.", [
                ("passos", [("Peças do processo", "Envie a íntegra ou as peças em PDF. A central separa as peças, identifica o processo e as decisões."),
                            ("Leitura das decisões", "A leitura localiza o que foi deferido em cada decisão e marca o trecho de origem, com a página."),
                            ("Revisão dos parâmetros", "Cada parâmetro aparece com o selo conferido ou a revisar. Quem calcula confere, corrige e aprova item a item."),
                            ("Dados do contrato e cálculo", "Datas do contrato, salário por faixas, jornada e cartão de ponto, férias, faltas, correção e juros, honorários. Depois, Calcular.")]),
                ("card", "amber", "Nada sem conferência", "A leitura automática não substitui a revisão: todo parâmetro passa pela aprovação de quem faz o cálculo."),
            ]),
            ("Resultado e saídas", "Crédito do reclamante, débitos do reclamado, peças e acordo.", [
                ("lista", ["Resultado do cálculo com crédito do reclamante e débitos do reclamado, e as observações do motor.",
                           "Baixar PDF, PDF da discriminação e Excel.",
                           "Manifestação com os nossos cálculos, impugnação ou concordância em Word, com o fundamento de cada item impugnado.",
                           "Simular acordo: valor líquido, honorários e data do pagamento.",
                           "Fluxo: salvar na central, enviar para revisão, aprovar ou devolver para ajuste, marcar como entregue ou arquivar."]),
            ]),
            ("Cuidados", "Sigilo e conferência.", [
                ("lista", ["As peças e os cálculos ficam na central, atrás de login. Não anexar processos de outros escritórios.",
                           "Confira sempre o índice e a taxa de juros da data do cálculo: tabelas novas mudam o resultado.",
                           "Em caso de diferença com o cálculo do perito ou da parte, a tela de impugnação mostra item a item onde está a divergência."]),
            ]),
        ],
    },
    {
        "id": "BaseDeDados", "pasta": "BaseDeDados", "arquivo": "LAWgico_BaseDeDados_Manual",
        "nome": "Central de", "destaque": "Base de Dados",
        "url": "https://basededadospc.web.app/",
        "sub": "Dados, jurimetria e decisões da Justiça do Trabalho por comarca: todas as empresas, peritos, juízes, testemunhas e modelos de peças.",
        "secoes": [
            ("Para que serve", "Saber como a vara julga, não só nos nossos processos.", [
                ("p", "A Central de Base de Dados reúne as publicações e as peças dos processos de todas as empresas nas comarcas "
                      "acompanhadas (Ribeirão Pires e Santo André; São José dos Campos em preparação). Com isso, dá para ver se um "
                      "juiz ou perito é desfavorável só com os nossos clientes ou com todos, e buscar trechos literais para usar nas peças."),
                ("lista", ["Perfil de juízes, peritos e desembargadores, com resultado por empresa",
                           "Testemunhas e depoimentos, com o trecho da ata",
                           "Sinais de alerta e decisões relevantes",
                           "Modelos de peças e trechos prontos, com precedentes oficiais"]),
            ]),
            ("Como acessar", "Pelo card no painel LAWgico; o link pede a senha e mostra o seletor de bases.", [
                ("passos", [("Abra o card Central de Base de Dados", "Digite a senha do card; a central abre em uma nova aba."),
                            ("Digite a senha da central", "A tela de entrada mostra os números da base e, depois da senha, o seletor de comarcas."),
                            ("Escolha a comarca", "Ribeirão Pires ou Santo André. A chave no topo da plataforma troca de comarca sem pedir a senha de novo.")]),
                ACESSO_PADRAO,
            ]),
            ("De onde vêm os dados", "DJEN, peças do PJe e carteira do LM, atualizados todo dia às 7h30.", [
                ("lista", ["Publicações: Diário de Justiça Eletrônico Nacional (DJEN), desde abril de 2024.",
                           "Atas, laudos, sentenças e acórdãos: baixados do PJe e lidos por regras (sem custo por leitura).",
                           "Nossos processos, prazos e casos da CBC e da Bridgestone: LM.",
                           "Precedentes: TST e Banco Nacional de Precedentes do CNJ."]),
                ("card", "amber", "Limite", "Atas, laudos e acórdãos chegam até a data do último download do PJe, que depende de login manual. O mês corrente pode aparecer incompleto e a tela avisa."),
            ]),
            ("Abas principais", "Panorama, perfis, estratégia e ferramentas.", [
                ("tabela", ["Grupo", "Abas"], [
                    ["Panorama", "Visão geral, Movimentação do mês, Comparativo entre comarcas, Decisões, Perfil das ações"],
                    ["Perfis", "Peritos, Juízes, Desembargadores, Testemunhas, Sinais de alerta"],
                    ["Estratégia", "Prazos (nossos), Audiências (calendário e lista), Casos do cliente, Acordos e sugestão"],
                    ["Ferramentas", "Pergunte à base, Gerar painel, Caso novo (contestação), Trechos para peças, Modelos de peças, Processos"]]),
            ]),
            ("Caso novo e modelos", "Da inicial à minuta, no padrão do escritório.", [
                ("passos", [("Envie a inicial", "A aba Caso novo lê a petição no navegador: pedidos e valores, datas, endereço e checagem de texto oculto."),
                            ("Revise os pedidos", "Cada pedido aparece com o trecho de origem."),
                            ("Gere a minuta", "Contestação com preliminares e tópicos de defesa do escritório; trechos de outros casos vêm marcados para adaptar.")]),
                ("card", "cyan", "Precedentes", "Súmulas, OJs e precedentes vêm das bases oficiais do TST e do CNJ. Nunca citar precedente que não esteja na lista."),
            ]),
            ("Cuidados", "Contém nomes de partes e testemunhas.", [
                ("lista", ["A base tem nomes de reclamantes e testemunhas: uso interno, com senha. Não encaminhar telas para fora do P&C.",
                           "Improcedência é calculada só sobre sentenças de mérito (embargos e extinções ficam fora).",
                           "Os números mudam a cada atualização diária; para apresentações, anote a data do dado."]),
            ]),
        ],
    },
]
