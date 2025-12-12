Macro VBA – Consolidação Anual da Produção Nacional (IRP1.1)
Programa Macrorregional de Caracterização de Rendas Petrolíferas (PMCRP)
1. Objetivo do código

Esta macro em VBA, denominada PreencherTabela, foi desenvolvida para automatizar a consolidação anual da produção nacional de petróleo e gás natural, a partir de dados mensais publicados pela Agência Nacional do Petróleo, Gás Natural e Biocombustíveis (ANP).

O código é utilizado como instrumento metodológico de apoio à construção do indicador IRP1.1 – Produção nacional total de petróleo e gás natural, no âmbito do PMCRP.

2. Relação com o indicador IRP1.1

O IRP1.1 corresponde à produção nacional anual consolidada de petróleo e gás natural, expressa em Mboe/d (mil barris de óleo equivalente por dia) ou, quando necessário, convertida para boe (volume anual).

A macro PreencherTabela executa exatamente a etapa de:

agregação dos dados mensais por ano;

separação entre petróleo, gás natural e produção total;

organização padronizada dos resultados anuais.

Dessa forma, o código não cria novos dados, mas estrutura, consolida e organiza informações oficiais da ANP para viabilizar o cálculo e a apresentação do IRP1.1.

3. Fonte dos dados utilizados

Os dados processados pela macro devem ser previamente extraídos do:

Boletim Mensal da Produção de Petróleo e Gás Natural – ANP

sempre a edição de dezembro, que contém o encarte com dados anuais consolidados.

A tabela utilizada como base é:

“Histórico de produção de petróleo e gás natural (Mboe/d)”

4. Estrutura esperada da planilha

Para que a macro funcione corretamente, o arquivo Excel deve conter, no mínimo, as seguintes abas:

Aba “Dados de Produção”

Contém os dados mensais, organizados da seguinte forma:

Coluna	Conteúdo
A	Data (Mês/Ano)
B	Produção de petróleo (Mboe/d)
C	Produção de gás natural (Mboe/d)
D	Produção total (Mboe/d)

A primeira linha deve conter cabeçalhos.

A coluna A deve estar em formato de data reconhecido pelo Excel.

Aba “Gráficos”

Recebe automaticamente a tabela anual consolidada gerada pela macro.

5. O que a macro faz (passo a passo)

De forma resumida, a macro executa as seguintes etapas:

Identifica a última linha preenchida na aba “Dados de Produção”.

Percorre todas as linhas com datas válidas.

Extrai o ano de cada registro mensal.

Agrupa os valores por ano, separando:

petróleo;

gás natural;

produção total.

Armazena os resultados temporariamente em estruturas do tipo Dictionary.

Preenche automaticamente a aba “Gráficos” com:

cabeçalhos dinâmicos de anos;

linhas padronizadas por tipo de produção.

Exibe uma mensagem confirmando até qual ano os dados foram consolidados.

6. Recorte temporal adotado

Ano inicial: 2010

Ano final: determinado automaticamente pelo último ano disponível nos dados.

Esse recorte segue a definição metodológica do PMCRP.

7. Natureza metodológica e limitações

Esta macro possui caráter de processamento interno, sendo aplicada no contexto específico dos arquivos de consolidação do PMCRP.

Embora o código seja disponibilizado por transparência metodológica, a simples execução da macro por terceiros não garante a reprodução exata dos resultados, pois:

depende da estrutura interna dos arquivos do projeto;

pressupõe padronizações previamente adotadas pela equipe técnica;

integra uma cadeia metodológica mais ampla descrita na Nota Técnica.

O código deve ser interpretado como instrumento auxiliar, e não como método isolado.

8. Resultado gerado

O resultado final do processamento é uma tabela anual consolidada, que constitui a base numérica do indicador IRP1.1, posteriormente utilizada:

na construção de gráficos;

em análises comparativas;

como denominador dos indicadores IRP1.2 e IRP1.3.

9. Licença e uso

Código desenvolvido para fins técnicos e institucionais no âmbito do Programa Macrorregional de Caracterização de Rendas Petrolíferas (PMCRP).

Uso permitido para fins acadêmicos, técnicos e informativos, com a devida contextualização metodológica.
