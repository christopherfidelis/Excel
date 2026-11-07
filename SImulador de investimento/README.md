# Simulador de Investimentos --- Invest Mais

## Visão geral

O **Simulador de Investimentos** estima a evolução de aportes mensais ao
longo do tempo, calcula uma estimativa de dividendos mensais e sugere
uma distribuição do aporte entre tipos de Fundos de Investimento
Imobiliário (FIIs), conforme o perfil selecionado.

A versão documentada mantém a estrutura de cálculo e os dados da
planilha recebida. As alterações realizadas nesta etapa foram de
identidade visual: criação/aplicação da logo **Invest Mais** e
reformulação do layout com tons de roxo.

## 1. Quais perguntas a ferramenta responde?

  -----------------------------------------------------------------------
  Pergunta                            Onde aparece a resposta
  ----------------------------------- -----------------------------------
  Qual é o salário informado?         Aba **SIMULADOR**, seção
                                      **CONFIGURAÇÕES**, célula `H10`.

  Quanto a ferramenta sugere investir Seção **CONFIGURAÇÕES**, célula
  por mês com base no salário?        `H12`. A sugestão corresponde a 30%
                                      do salário informado.

  Qual é o rendimento mensal          Seção **CONFIGURAÇÕES**, célula
  considerado para a carteira?        `H11`.

  Quanto investir por mês?            Seção **INVESTIMENTO MENSAL**,
                                      célula `H16`.

  Por quantos anos investir?          Seção **INVESTIMENTO MENSAL**,
                                      célula `H17`.

  Qual taxa mensal será usada na      Seção **INVESTIMENTO MENSAL**,
  projeção?                           célula `H18` (o rótulo da planilha
                                      é "Taxa de Financiamento Mensal?").

  Qual será o patrimônio acumulado ao Resultado **Patrimônio Acumulado**,
  final do período escolhido?         célula `H19`.

  Quanto a carteira poderá gerar de   Resultado **Dividendos Mensais**,
  dividendos por mês?                 célula `H20`.

  Como ficariam o patrimônio e os     Seção **OUTROS CENÁRIOS**, linhas
  dividendos em outros prazos?        24 a 28: prazos de 2, 5, 10, 20 e
                                      30 anos. O patrimônio projetado
                                      aparece na coluna `G` e os
                                      dividendos estimados na coluna `H`.

  Como distribuir o aporte mensal     Seção de distribuição por perfil,
  entre os tipos de FIIs para o       linhas 32 a 42. O perfil é
  perfil escolhido?                   selecionado em `H32`; os
                                      percentuais sugeridos aparecem em
                                      `F36:F41` e os valores
                                      correspondentes em `H36:H41`. O
                                      total distribuído aparece em `H42`.
  -----------------------------------------------------------------------

Os resultados são estimativas matemáticas baseadas nos parâmetros
informados. Não representam garantia de rentabilidade, dividendos ou
valorização dos ativos.

## 2. Como entram o `VF` e o `PROCV` nos cálculos?

### `VF` --- Valor Futuro

A função `VF` (equivalente à função `FV` em instalações do Excel em
inglês) calcula o valor futuro de uma sequência de pagamentos
periódicos, considerando uma taxa por período.

Na planilha, a fórmula do patrimônio acumulado é:

``` excel
=VF(taxa_mensal;qtd_anos*12;aporte)*-1
```

Na versão em inglês do Excel, a mesma função é escrita como:

``` excel
=FV(taxa_mensal,qtd_anos*12,aporte)*-1
```

O significado dos argumentos é:

-   `taxa_mensal`: taxa de rendimento mensal usada na projeção (célula
    `H18`);
-   `qtd_anos*12`: quantidade de aportes mensais durante todo o período;
-   `aporte`: valor investido mensalmente (célula `H16`);
-   `*-1`: ajuste de sinal para apresentar o patrimônio acumulado como
    valor positivo, considerando a convenção de sinais financeiros do
    Excel.

Como o argumento de valor presente foi omitido e o tipo de pagamento
também, o Excel considera valor presente igual a zero e aportes no fim
de cada período. Portanto, a projeção considera aportes mensais
regulares, sem um investimento inicial adicional.

A mesma lógica é aplicada à tabela **OUTROS CENÁRIOS**, substituindo a
quantidade de anos pelo prazo indicado em cada linha.

### `PROCV` --- busca do percentual por perfil

A função `PROCV` procura uma chave na primeira coluna de uma tabela e
retorna o conteúdo de outra coluna da mesma linha.

Na planilha, a fórmula usada para encontrar o percentual de cada tipo de
FII é:

``` excel
=PROCV($H$32&"-"&$B36;AUX!A:D;4;0)
```

O funcionamento é:

1.  `H32` contém o perfil selecionado, como `Agressivo`, `Moderado` ou
    `Conservador`.
2.  `B36` contém o tipo de FII da linha, como `Papel` ou `Tijolo`.
3.  A concatenação cria uma chave, por exemplo `Agressivo-Papel`.
4.  O `PROCV` procura essa chave na primeira coluna da tabela da aba
    `AUX`.
5.  O argumento `4` pede o valor da quarta coluna da tabela, que contém
    o percentual sugerido.
6.  O argumento `0` exige correspondência exata.

Depois da busca, a planilha calcula o valor mensal destinado a cada
categoria multiplicando o aporte pelo percentual encontrado:

``` excel
=$H$33*F36
```

Assim, ao trocar o perfil em `H32`, os percentuais e os valores
sugeridos para cada tipo de FII são atualizados automaticamente.

## 3. Intervalos nomeados

A pasta de trabalho contém os seguintes intervalos nomeados, todos
vinculados a células da aba **SIMULADOR**:

  -------------------------------------------------------------------------
  Nome do intervalo         Referência              Significado
  ------------------------- ----------------------- -----------------------
  `salario`                 `SIMULADOR!$H$10`       Salário informado.

  `rendimento_carteira`     `SIMULADOR!$H$11`       Rendimento mensal usado
                                                    para estimar os
                                                    dividendos.

  `sugestao_investimento`   `SIMULADOR!$H$12`       Sugestão de aporte
                                                    calculada a partir do
                                                    salário.

  `aporte`                  `SIMULADOR!$H$16`       Valor do aporte mensal.

  `qtd_anos`                `SIMULADOR!$H$17`       Prazo do investimento
                                                    em anos.

  `taxa_mensal`             `SIMULADOR!$H$18`       Taxa mensal usada na
                                                    função de valor futuro.

  `pratimonio`              `SIMULADOR!$H$19`       Patrimônio acumulado
                                                    projetado.

  `dividendo`               `SIMULADOR!$H$20`       Estimativa de
                                                    dividendos mensais.
  -------------------------------------------------------------------------

**Observação:** o nome `pratimonio` está escrito dessa forma na pasta de
trabalho e foi mantido para não quebrar referências existentes. Os nomes
`pratimonio` e `dividendo` são os nomes efetivamente cadastrados no
arquivo.

Esses intervalos tornam as fórmulas mais legíveis, pois permitem
referenciar uma célula pelo seu nome em vez de usar apenas seu endereço.

## 4. Percentuais de cada perfil e origem dos dados

Os percentuais ficam cadastrados na aba auxiliar **AUX**, na coluna `D`
(**Percentual Sugerido**). A tabela usa uma linha para cada combinação
de perfil e tipo de FII.

  Tipo de FII         Conservador   Moderado   Agressivo
  ----------------- ------------- ---------- -----------
  Papel                       30%        32%         50%
  Tijolo                      50%        35%         10%
  Híbridos                    10%         8%          5%
  FOF's                       10%         5%          5%
  Desenvolvimento              0%        10%         20%
  Hotelaria                    0%        10%         10%
  **Total**              **100%**   **100%**    **100%**

### De onde vieram?

Esses percentuais foram encontrados na tabela de apoio da aba **AUX** da
planilha original recebida e foram preservados na versão com o novo
layout. A distribuição não foi recalculada nem alterada durante a
reformulação visual.

**Limitação de rastreabilidade:** o arquivo não apresenta uma fonte
externa, estudo, metodologia ou referência bibliográfica que explique a
escolha desses percentuais. Portanto, é correto dizer que eles vieram da
tabela de parâmetros já existente na ferramenta recebida; não é possível
atribuir sua autoria ou validação a uma fonte específica apenas com as
informações registradas na planilha.

Os percentuais são parâmetros de sugestão da ferramenta, não uma
recomendação individualizada de investimento.