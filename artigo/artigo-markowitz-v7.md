# Da planilha ao backtest estático: processo, modelo e o que quatro anos de dados mostram sobre otimização de portfólios de dividendos

**Status:** Rascunho v7 | Junho 2026

> Não constitui recomendação de investimento nem análise de valores mobiliários. O autor não é analista credenciado pela CVM. Rentabilidade passada não garante resultados futuros. As análises utilizam exclusivamente dados históricos retroativos.

---

A pergunta parece simples: um otimizador matemático calibrado com dados históricos consegue superar, fora da amostra, os pesos recomendados por analistas profissionais? A maioria das análises que tentam responder essa pergunta começa com um problema de design que invalida a conclusão antes mesmo de chegar aos números. Usam a composição atual da carteira aplicada retroativamente, como se o analista de 2022 tivesse recomendado hoje o que recomendava então. Ou calibram o otimizador sobre o mesmo período que serve de teste. Os dois erros produzem o mesmo efeito: um resultado que não se sustenta metodologicamente, independente de qual lado vença.

O experimento descrito aqui tentou evitar esses problemas. As carteiras utilizadas são as publicadas em junho de 2022, verificadas em PDFs de arquivo, com os tickers e pesos daquela edição específica. O otimizador foi calibrado exclusivamente com dados de junho de 2019 a maio de 2022 — o que estava disponível antes da decisão de alocação. O período de teste cobre junho de 2022 a junho de 2026, e nenhum dado posterior a maio de 2022 entrou no processo de calibração. É o que a literatura chama de backtest estático com separação IS/OOS: in-sample para calibrar, out-of-sample para avaliar.

---

## O processo: do Excel ao Python com suporte de agente de IA

A versão original dessa análise rodava em Excel com o Solver. A formulação do problema era idêntica à que está no código Python hoje: maximizar o índice de Sharpe de uma carteira sujeito a bounds α = 0,5 e β = 2,0 sobre os pesos recomendados pelo analista, sem venda a descoberto, com soma dos pesos igual a 1. O Solver converge adequadamente para esse problema quando há até 12 ativos. O problema não era a qualidade da solução — era o que ficava de fora.

Qualquer mudança de ticker ou janela de dados exigia retrabalho manual nas fórmulas. A separação IS/OOS com histórico de seis anos, com auditabilidade completa de cada decisão de design, não cabe dentro de uma planilha orientada à célula sem perda de rastreabilidade. E o bootstrap histórico com milhares de trajetórias é simplesmente incompatível com o modelo de execução de uma planilha.

A migração para Python com suporte do Claude (Anthropic) mudou o custo de implementação das decisões metodológicas, não a lógica. O que levaria dias passou a levar horas. Cada decisão de design — qual período usar para calibração, como tratar tickers descontinuados, como verificar ausência de look-ahead bias no código — passou por revisão iterativa entre analista e agente. O código está disponível no repositório GitHub ao final. As decisões permaneceram do lado humano. A execução passou a ser imediata.

---

## O que o otimizador faz — e o que não é

O código resolve o seguinte problema de otimização:

maximizar `[w' μ − rf] / sqrt(w' Σ w)` sujeito a `Σw = 1` e `α × w_rec ≤ w ≤ β × w_rec`

onde μ é o vetor de retornos médios anualizados estimados por média histórica de log-retornos, Σ é a matriz de covariância amostral anualizada e w_rec são os pesos do analista. Ambos calculados exclusivamente sobre o período de calibração (IS). O método de resolução é SLSQP via scipy, com α = 0,5 e β = 2,0.

Isso não é a fronteira eficiente de Markowitz. A fronteira eficiente mapeia todas as combinações ótimas de risco-retorno sem restrição sobre o espaço de pesos. O que o código encontra é o portfólio de máximo Sharpe dentro de uma vizinhança definida pelos pesos do analista — com limite inferior de 50% e superior de 200% de cada posição original. Os bounds transformam o analista em prior implícito. A solução encontrada pode estar distante da fronteira eficiente irrestrita. A denominação correta é otimização por máximo Sharpe com restrições relativas ao analista, não portfólio de Markowitz.

A distinção importa para interpretar os resultados. Michaud (1989) demonstrou que o MVO irrestrito com estimadores de amostra amplifica erros de estimativa: o otimizador concentra posição nos ativos com maior retorno histórico amostral sem mecanismo para separar retorno genuíno de ruído. Os bounds atenuam essa instabilidade ao forçar o vetor de pesos a permanecer próximo do benchmark analítico. Não eliminam o problema — apenas reduzem o espaço em que o overfitting pode ocorrer.

---

## As carteiras e os dados

Duas casas forneceram carteiras verificáveis em junho de 2022 a partir de PDFs de arquivo: Casa A (Carteira Top Dividendos) e Casa B (Carteira Recomendada de Dividendos). As composições foram lidas diretamente dos documentos originais, sem interpolação ou reconstrução retrospectiva.

Casa A — Portfólio Concentrado, 1 de junho de 2022: BBAS3 (20%), CPLE6 (20%), EGIE3 (20%), PETR4 (20%), GRND3 (10%) e TIMS3 (10%). Casa B — Portfólio Equiponderado, junho de 2022: dez ativos com peso igual de 10% — ALUP11, BBAS3, BBDC4, CYRE3, ENGI11, ITUB4, SBSP3, TRPL4, VALE3 e VIVT3.

Dois tickers se mostraram indisponíveis no Yahoo Finance para o período completo. CPLE6 foi descontinuada após a privatização da Copel em agosto de 2023, quando as ações preferenciais foram convertidas para ordinárias (CPLE3) na proporção 1:1. A série foi substituída por CPLE3, que representa a continuidade econômica direta do investimento. TRPL4 (ISA CTEEP PN) não retornou dados em nenhuma variação de ticker testada — TRPL3 também indisponível. A carteira Casa B foi ajustada para nove ativos com pesos redistribuídos proporcionalmente (11,1% cada). Essas são as condições efetivas da análise, declaradas como limitação de dados, não corrigidas por interpolação ou proxy arbitrário.

---

## Resultados: dois portfólios, dois resultados

**Tabela 1. Backtest estático IS/OOS**
*IS: jun/2019–mai/2022 (743 pregões) | OOS: jun/2022–jun/2026 (1.016 pregões)*
*CDI IS: 5,0% a.a. | CDI OOS: 13,1% a.a.*

| | IS Analista | IS Otim. | OOS Analista | OOS Otim. |
|---|---|---|---|---|
| **Casa A — Portfólio Concentrado** | | | | |
| Retorno anual | 12,2% | 17,8% | 19,7% | **25,3%** |
| Volatilidade | 29,5% | 32,5% | 16,9% | 18,7% |
| Sharpe | 0,244 | 0,395 | 0,392 | **0,654** |
| Max Drawdown | -43,5% | -47,2% | -18,0% | -18,3% |
| **Casa B** | | | | |
| Retorno anual | 6,1% | 12,9% | **17,3%** | 14,4% |
| Volatilidade | 29,6% | 26,0% | 18,9% | 16,7% |
| Sharpe | 0,037 | 0,306 | **0,221** | 0,080 |
| Max Drawdown | -42,4% | -34,0% | -19,5% | -17,1% |

*Benchmarks OOS: BOVA11 +12,1% a.a. (Sharpe -0,056) | DIVO11 +14,3% a.a. (Sharpe 0,080) | CDI +13,1% a.a.*

---

## Casa A: o otimizador venceu — e por quê isso exige cautela

O otimizador elevou CPLE6/CPLE3 de 20% para 40% e PETR4 de 20% para 30%, cortando pela metade BBAS3, EGIE3, GRND3 e TIMS3. No OOS, entregou +25,3% a.a. contra +19,7% do analista — delta de 5,6 pontos percentuais anualizados, Sharpe 0,654 contra 0,392.

O ganho é real nos dados. A questão é o mecanismo. CPLE3 subiu expressivamente no período, incluindo o prêmio de privatização da Copel em 2023 e a subsequente expansão dos múltiplos como empresa privada. PETR4 distribuiu dividendos extraordinários relevantes ao longo do período. O otimizador, calibrado com retornos de 2019 a 2022, identificou nesses dois ativos uma combinação com relação risco-retorno favorável naquele histórico — e essa combinação se sustentou no período seguinte.

O problema de atribuição é o seguinte: o otimizador não sabia que a Copel seria privatizada. Não sabia que PETR4 distribuiria dividendos extraordinários. O que sabia era que, no histórico de calibração, esses ativos apresentavam perfil de retorno ajustado ao risco superior aos demais, dadas as covariâncias estimadas. O modelo acertou a direção, mas o mecanismo pelo qual acertou não é replicável como regra: dependeu de eventos corporativos específicos que os dados históricos capturavam como tendência, não como previsão. Isso não invalida o resultado, mas impede que ele seja generalizado como prova de superioridade estrutural do otimizador.

---

## Casa B: o analista venceu — o mecanismo de Michaud em ação

O otimizador elevou ALUP11, VALE3 e VIVT3 para 22,2% cada, concentrando 67% do portfólio em três ativos. Os seis restantes caíram para 5,6% cada. No IS, o Sharpe subiu de 0,037 (analista) para 0,306 (otimizado). No OOS, inverteu: o analista entregou 0,221 e o otimizado caiu para 0,080 — praticamente idêntico ao DIVO11 passivo (0,080).

Esse é o mecanismo que Michaud descreveu: a concentração em ALUP11, VALE3 e VIVT3 era ótima dentro da amostra porque esses três ativos apresentavam, na janela 2019-2022, a combinação mais favorável de retorno esperado e baixa covariância com o restante do portfólio. Parte desse resultado era estrutural. Parte era ruído amostral. O otimizador não distingue. Quando o ruído se dissipou no OOS, o portfólio concentrado não sustentou a vantagem — e o analista, com distribuição mais uniforme, provou ser mais robusto.

A diferença entre os dois casos não é aleatória. Casa A tinha seis ativos com pesos mais assimétricos na recomendação original, o que deixava menos espaço para o otimizador criar concentração adicional dentro dos bounds. Casa B tinha dez ativos equiponderados, o que ampliou o grau de liberdade do otimizador para realocar — e, com isso, amplificou o potencial de overfitting.

---

## O CDI como contexto estrutural

CDI acumulado no período OOS: 64,7% em quatro anos, ou 13,1% ao ano. Com esse patamar de taxa livre de risco, o Sharpe de qualquer portfólio de renda variável parte de uma exigência alta: para Sharpe de 0,40 com volatilidade de 18%, o retorno anual precisa ser de pelo menos 20,3%. Casa B — Analista entregou 17,3% — acima do CDI, mas com Sharpe 0,221. Casa B — Otimizado entregou 14,4% — praticamente idêntico ao CDI em termos absolutos, com Sharpe 0,080.

O patamar de juros brasileiro não é detalhe de contexto. Ele define o denominador financeiro e psicológico contra o qual qualquer portfólio de risco variável será medido pelo investidor. Em um ambiente com CDI a 13%, uma carteira de dividendos precisa entregar retorno real expressivo para justificar o risco. No período analisado, isso foi possível — mas não de forma universal entre as casas testadas.

---

## Bootstrap histórico prospectivo

A simulação por bootstrap histórico reamostra com reposição os retornos diários do período OOS em 5.000 trajetórias de 252 pregões. O objetivo não é projetar performance futura: é estimar a distribuição de resultados possíveis caso o regime de mercado do período observado se repita. Sem premissa de normalidade — a distribuição empírica dos retornos, com suas caudas e assimetrias, é preservada integralmente.

**Tabela 2. Bootstrap histórico | 5.000 trajetórias | Horizonte: 252 pregões**

| | Retorno p5 | Retorno med. | Retorno p95 | Sharpe med. | MDD med. |
|---|---|---|---|---|---|
| Casa A — Analista | -10,0% | +20,3% | +58,3% | 0,424 | -11,9% |
| Casa A — Otimizado | -8,3% | +25,5% | +71,2% | 0,657 | -12,4% |
| Casa B — Analista | -14,5% | +17,9% | +59,3% | 0,253 | -14,1% |
| Casa B — Otimizado | -13,4% | +14,9% | +49,7% | 0,109 | -12,7% |

A assimetria entre Casa A e Casa B no bootstrap reproduz o padrão do backtest estático. A cauda esquerda do Casa A — Otimizado é mais favorável que a do Casa B Analista. A mediana do Casa B Otimizado (14,9%) fica abaixo do CDI (13,1%) por margem estreita — e levemente abaixo do DIVO11. Para o portfólio Casa B, o benchmark passivo de dividendos teria produzido resultado equivalente ao otimizador com menor risco de concentração.

---

## Delimitações

O experimento responde a uma pergunta específica: dados os portfólios exatos de junho de 2022 e dados de calibração disponíveis até maio de 2022, o otimizador por máximo Sharpe com bounds α = 0,5 / β = 2,0 produziu pesos mais eficientes nos quatro anos seguintes? Para Casa A, a resposta é sim, com ressalva de atribuição a eventos corporativos específicos. Para Casa B, a resposta é não, com o mecanismo de overfitting de Michaud como explicação mais provável.

O experimento não responde se o resultado se manteria com outras casas, outros períodos ou outros valores de bounds. Não responde o que ocorreria com a fronteira eficiente irrestrita — que provavelmente teria produzido overfitting mais severo, sem os bounds como amortecedor. Não elimina a possibilidade de que o bom resultado da Casa A seja, em parte, sorte de calibração sobre um período que aconteceu de prefigurar os vencedores do OOS.

CPLE6 foi substituída por CPLE3 como proxy de continuidade econômica, a uma razão de conversão 1:1 confirmada pela privatização. TRPL4 foi removida por indisponibilidade de dados em qualquer variação de ticker. A carteira Casa B analisada tem nove ativos, não dez, com pesos redistribuídos proporcionalmente. Essas são as condições efetivas da análise.

A limitação metodológica mais relevante é que estimadores de média histórica para retornos esperados são ruidosos por construção. O erro padrão da estimativa de retorno médio anualizado de um ativo com volatilidade de 30% ao ano, com 743 pregões de calibração, ainda é da ordem de 11 pontos percentuais. O otimizador não sabe quanto desse retorno histórico era sinal e quanto era ruído. O Ledoit-Wolf shrinkage e o Black-Litterman são as alternativas metodológicas estabelecidas para reduzir esse problema. Não foram implementados aqui. É o próximo passo natural.

---

## Conclusão

A pergunta inicial — o otimizador bate o analista fora da amostra? — produz respostas opostas nas duas carteiras testadas. Casa A: sim, com 5,6 pontos percentuais a mais ao ano e Sharpe 0,654 contra 0,392. Casa B: não, com o analista entregando Sharpe 0,221 contra 0,080 do otimizado, e o DIVO11 passivo empatando com o portfólio "ótimo" nessa mesma métrica.

O resultado assimétrico é mais informativo do que um resultado uniforme seria. Mostra que a utilidade do otimizador depende de quanto do retorno histórico de calibração representa padrão persistente — e com N=1 realização do período OOS, não é possível separar sinal de ruído na atribuição. O que os bounds garantem é um teto de concentração; o que determina o sinal do resultado é o que acontece com os ativos superalocados, não a geometria inicial dos pesos. Em um caso, o otimizador acrescentou valor real. No outro, destruiu o prêmio de diversificação que a equiponderação do analista oferecia.

O CDI a 13,1% ao ano no período testado é o dado mais subestimado da análise. Ele estabelece uma barreira de entrada alta para qualquer argumento de que renda variável de dividendos cria valor consistente ajustado ao risco no Brasil recente. Os portfólios analisados superaram essa barreira — mas por margens que, em dois dos quatro casos, não justificam a complexidade adicional de um otimizador.

---

*Nota metodológica: dados de preços ajustados coletados via Yahoo Finance. Período total: jun/2019–jun/2026. Janela IS (calibração): jun/2019–mai/2022 (743 pregões). Janela OOS (teste): jun/2022–jun/2026 (1.016 pregões). CDI extraído via BCB SGS série 11. Tickers verificados em PDFs originais das casas analisadas. CPLE6 substituída por CPLE3 (mesma empresa, conversão 1:1 na privatização ago/2023). TRPL4 removida por indisponibilidade de dados no Yahoo Finance — carteira Casa B ajustada para 9 ativos com pesos proporcionais. Otimização: SLSQP via scipy, maximização do índice de Sharpe, bounds α = 0,5 e β = 2,0 relativos aos pesos recomendados. Estimadores: média e covariância amostrais, sem shrinkage. Bootstrap histórico MC: 5.000 trajetórias, reamostragem iid de retornos diários do período OOS, sem premissa de normalidade. Análise desenvolvida com suporte do Claude (Anthropic).*

*Referências: Markowitz, H. (1952). Portfolio Selection. Journal of Finance, 7(1), 77-91. Michaud, R. (1989). The Markowitz Optimization Enigma: Is Optimized Optimal? Financial Analysts Journal, 45(1), 31-42. DeMiguel, V., Garlappi, L. e Uppal, R. (2009). Optimal Versus Naive Diversification. Review of Financial Studies, 22(5), 1915-1953.*

---

## Apêndice: prompt de ponto de entrada para análise da sua carteira

O prompt abaixo instrui o Claude (modo Cowork ou Claude.ai) a replicar o processo descrito neste artigo para a carteira informada pelo usuário.

---

```
Você irá analisar minha carteira de investimentos usando o processo descrito no artigo
"Da planilha ao backtest estático".

## Minha carteira

Ticker | Peso
-------|------
[ex: PETR4 | 20%]
[ex: ITUB4 | 15%]
[adicione quantas linhas precisar]

## O que quero que a análise produza

1. Desempenho histórico (últimos 3 anos disponíveis): retorno anualizado, volatilidade,
   índice de Sharpe, Maximum Drawdown. Taxa livre de risco: CDI via BCB SGS série 11.

2. Comparação com benchmarks: DIVO11 e IBOV no mesmo período.

3. Otimização por máximo Sharpe com restrições relativas: bounds α = 0,5 e β = 2,0
   em relação aos pesos atuais. Exibir pesos antes e depois, e Sharpe antes e depois.
   Nomear corretamente: não é "carteira de Markowitz" nem "fronteira eficiente" —
   é otimização por máximo Sharpe com bounds sobre os pesos do analista.

4. Bootstrap histórico não-paramétrico: 2.000 trajetórias de 252 pregões,
   reamostragem iid dos retornos diários históricos. Resultados: percentis
   p5, mediana e p95 de retorno acumulado e índice de Sharpe.

## Premissas

- Taxa livre de risco: CDI do período via API do BCB (SGS série 11)
- Dados: Yahoo Finance, preços ajustados por proventos
- Para tickers sem dados completos: declarar explicitamente, não interpolar

## Formato de saída

Planilha Excel com uma aba por módulo. No chat: três tabelas com os números
e um parágrafo de interpretação para cada uma. Declarar qualquer limitação
de dados explicitamente antes de apresentar os resultados.
```

---

*Este prompt não produz recomendação de investimento. Os resultados são análise retrospectiva de dados históricos públicos.*
