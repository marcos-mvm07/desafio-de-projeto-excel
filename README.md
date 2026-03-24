# desafio-de-projeto-excel
Aprendido ao longo do curso

FÓRMULAS DO EXCEL
= aleatorioentre (x;y) - Traz um número aleatório do intervalo especificado
F2 + CTRL + Enter = Arrasta uma mesma fórmula para todo o intervalo desejado7
CTRL + ALT + V = Colar especial
Digitar F4 = Congela a célula
Dados -> Ferramenta de dados -> Validação de dados -> Lista = Cria uma lista com "valores" que podem ser aceitos na célula
Formatação Condicional -> Nova regra -> Escolhe a célula, valor e diz o que muda quando esta é selecionada
= cont.valores -> Conta a quantidade de células preenchidas, ou seja, a quantidade de valores
= cont.se -> Conta apenas as células que estão com o parâmetro colocado
=VF -> Faz a previsão de um fator futuro com base na taxa de juros, tempo em anos (deve ser multiplicado por 12, já que o valor é em meses) e o valor investido por mês

ctrl + ; = Data do dia atual
Quando valores ficarem em #####, basta selecionar tudo e dar duplo clique para reajustar o tamanho

Dados -> Ferramenta de dados -> Texto para colunas => Caso as informações estejam em uma só linha e queira separar as informações por coluna, basta clicar em "Delimitar"

ctrl + shift + tecla da direita + tecla da esquerda => Seleciona toda a tabela 

ctrl + shift + + => Adiciona linhas

Seleciona uma coluna, vai em edição -> Lupa -> Substituir. Por exemplo, substituir . por ,

Criar tabela dinâmica haverá os campos e é possível selecionar apenas aqueles que se deseja. Ao lado de cada campo, tem uma seta pra baixo que pode configurar o Campo de Valor

Rótulos de linha -> Classificação -> Mais opções de classificação => Classificar de forma crescente ou decrescente

Criar um campo na tabela dinâmica = Análise da tabela dinâmica -> Campos, Itens e Conjunto -> Campo calculado = Inclui a fórmula

Campo calculado pode dar problema pelo cálculo diferente. Duplo clique, vê o racional do número

Seleciona a tabela dinâmica, vai em design e deixa-a mais apresentável. Como cores

ANálise de tabela dinâmica -> Inserir segmentação de dados -> Inserir o filtro => Filtrar de forma visual

Preenchimento relâmpago: Cria uma coluna à direita da qual se quer corrigir, escreve da forma correta, seleciona a nova coluna e ctrl + E

Base normalizadora. Primeiro, Dados -> Ferramenta de dados -> Remover duplicadas. Deixa apenas uma amostra de cada

Base normalizadora = Faz uma tabela com os novos valores, usa o comando PROCV e vai preenchendo os requisitos

Seleciona valores -> Formatação Condicional -> Regras de realce das células => Destaca os valores condicionais

Para alterar em uma célula à parte, seleciona todos os valores, vai em formatação condicional e cria/edita a regra, selecionando a célula em questão

Formatação condicional -> Gerenciar regra -> Nova regra e cria a formatação. Ultima linha deve ser numericamente igual à linha do aplicar-se a

Em validação de dados -> lista => É possível usar uma célula de apoio

Ao selecionar dados, mas se quer apenas uma ocorrência de cada, Dados -> Eliminar duplicadas

=CONT.SE(referência, critérios) -> Possível selecionar de outra planilha, referenciando e colocando o $ pra poder "bloquear"

Alt + = -> Faz o somatório total automaticamente

Cont.SES => Conta a quantidade utilizando mais de 1 critério e referência

Em validação de dados é possível inserir uma mensagem para inserir os valores ou para se os valores estiverem errados

Seleciona as duas ou mais células com filtro -> Classificar -> Classificar personalizado

Formatar células -> Proteção -> Marca oculta = O usuário não vai poder mexer na fórmula

Para uma imagem usada na tabela do excel não ter seu tamanho modificado ao alterar as colunas, vai com botão direito -> Tamanhos e propriedades -> Propriedades -> Não mover ou dimensionar com células 

Ao selecionar uma célula, aparece "F12", se selecioná-la e mudar o nome, esse nome será usado a partir daí. Gerenciador de nomes gerencia os apelidos.

Chave composta 
=C3&"-"&B3

Considerar C3 e B3 como parâmetros dois quais se quer combinar

=procv -> Procura a chave numa coluna e retorna o valor que se quer da referida coluna escolhida. Ex: Na linha do Moderado-PAPEL, seleciona toda a tabela e retorna a porcentagem (40%) que é da coluna 4 (num índice)
