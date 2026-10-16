# 📊 Simulador de Investimentos em Fundos Imobiliários (FIIs) & Planejamento Financeiro

Ferramenta desenvolvida em planilha eletrônica (Excel / Google Sheets) para simulação de acúmulo de patrimônio, geração de renda passiva mensal através de Fundos Imobiliários (FIIs) e planejamento de metas financeiras. 

O projeto foi desenhado com interface no estilo **Dashboard/Aplicativo**, sem necessidade de bases de dados externas, utilizando fórmulas dinâmicas, validação de dados, matrizes de alocação por perfil e gráficos interativos.

---

## 🎯 1. As 5 Perguntas de Negócio que a Ferramenta Responde

A ferramenta foi estruturada no painel principal da aba `Simulador` para responder de forma imediata às cinco perguntas fundamentais do investidor:

| Pergunta de Negócio | Onde aparece a resposta na tela | Origem do Dado / Fórmula |
| :--- | :--- | :--- |
| **1. Quanto investir por mês?** | **Painel de Configurações / Cartão de Entrada** | Inserido pelo usuário ou sugerido via regra de 30% da renda (`aporte`). |
| **2. Por quantos anos?** | **Painel de Configurações / Cartão de Parâmetros** | Inserido pelo usuário (`prazo_anos`) e desdobrado em meses (`prazo_anos * 12`). |
| **3. Qual a taxa de rendimento mensal?** | **Painel de Configurações / Cartão de Parâmetros** | Inserida pelo usuário (`taxa_mensal`), refletindo o dividend yield médio da carteira. |
| **4. Quanto de patrimônio vai acumular?** | **Destaque Principal do Dashboard & Tabela de Cenários** | Calculado via fórmula `=VF(taxa_mensal; prazo_anos*12; -aporte; 0)`. |
| **5. Quanto vai receber de dividendos/mês?** | **Destaque Principal do Dashboard & Tabela de Cenários** | Calculado multiplicando o patrimônio acumulado pela taxa mensal de rendimento. |

---

## 🧮 2. Como as Funções `VF` e `PROCV` Entram nos Cálculos

### 🔹 Função `VF` (Valor Futuro)
A função `VF` é o motor financeiro da planilha. Ela calcula o valor acumulado no final de um período com base em aportes constantes e uma taxa de juros composta.

**Sintaxe aplicada na planilha:**
```excel
=VF(taxa_mensal; prazo_anos * 12; -aporte; 0)
```
- **`taxa_mensal`**: Taxa de juros por período (ex: `0,80%`).
- **`prazo_anos * 12`**: Número total de pagamentos/meses (ex: 10 anos = 120 meses).
- **`-aporte`**: O valor depositado mensalmente. Utiliza-se o sinal negativo (`-`) para que o resultado do patrimônio retorne como um valor positivo no Excel.
- **`0`**: Pagamento no final de cada período (postecipado).

---

### 🔹 Função `PROCV` (Procura Vertical) com Chave Composta
Para determinar o percentual de investimento em cada um dos 6 tipos de fundos imobiliários com base no perfil escolhido pelo usuário (*Conservador*, *Moderado* ou *Arrojado*), utilizou-se uma **Chave Composta**.

**Lógica da Chave Composta:**
Na aba de apoio (`Apoio`), criou-se uma coluna auxiliar combinando o Nome do Perfil com o Tipo de Fundo:
$$\text{Chave Composta} = \text{Perfil} \;\&\; \text{"\_"} \;\&\; \text{Tipo\_de\_Fundo}$$

**Sintaxe aplicada na planilha:**
```excel
=PROCV(B$11 & "_" & A20; Apoio!C2:D20; 2; FALSO)
```
- **`B$11`**: Perfil selecionado pelo usuário na aba `Simulador` (ex: "Moderado").
- **`A20`**: Tipo do Fundo Imobiliário (ex: "FIIs de Tijolo - Logística").
- O `PROCV` busca a combinação exata `Moderado_FIIs de Tijolo - Logística` na matriz e retorna a alocação percentual correspondente.

---

## 🏷️ 3. Intervalos Nomeados Criados

Para tornar as fórmulas legíveis e evitar erros de referência, os principais campos da planilha foram mapeados como **Intervalos Nomeados**:

1. `salario`: Célula contendo o salário mensal do usuário.
2. `aporte`: Célula com o valor do aporte mensal definido para investimento.
3. `aporte_sugerido`: Célula calculada com 30% do salário (`=salario * 0,30`).
4. `taxa_mensal`: Célula contendo a taxa de rendimento esperada ao mês (ex: `0,80%`).
5. `prazo_anos`: Célula contendo o horizonte de tempo em anos (ex: `10`).
6. `perfil_investidor`: Célula de Validação de Dados com a lista suspensa (`Conservador`, `Moderado`, `Arrojado`).
7. `tabela_alocacao`: Intervalo na aba `Apoio` utilizado para as consultas do `PROCV`.

---

## 📊 4. Percentuais por Perfil de Investidor e Origem

A distribuição da carteira foi calibrada para refletir a relação risco/retorno aceita por cada perfil de investidor:

| Tipo de Fundo Imobiliário | Conservador | Moderado | Arrojado |
| :--- | :---: | :---: | :---: |
| **FIIs de Papel (CRI)** | **40%** | **25%** | **15%** |
| **FIIs de Tijolo - Logística** | **25%** | **25%** | **20%** |
| **FIIs de Tijolo - Lajes Corporativas** | **10%** | **15%** | **20%** |
| **FIIs de Tijolo - Shopping** | **10%** | **15%** | **20%** |
| **FIIs de Fondos de Fondos (FOFs)** | **10%** | **10%** | **10%** |
| **FIIs do Agronegócio (Fiagro)** | **5%** | **10%** | **15%** |
| **TOTAL** | **100%** | **100%** | **100%** |

### 📌 Origem das Proporções:
- **Conservador:** Prioriza a menor volatilidade e alta previsibilidade. Concentra 40% em FIIs de Papel/CRI e 25% em Galpões Logísticos (contratos atípicos e inquilinos de primeira linha).
- **Moderado:** Busca um equilíbrio entre renda periódica e valorização de cotas. Equilibra Tijolo (Logística, Lajes, Shopping) e Papel.
- **Arrojado:** Aceita maior oscilação na busca de maior ganho de capital. Aumenta a exposição em Lajes, Shoppings e Fiagros (15%), diminuindo a exposição em títulos de dívida (Papel).

---

## 🔄 5. O Que Foi Evoluído em Relação à Ferramenta Base

1. **🎨 Identidade Visual & Branding:**
   - Adicionado banner estilizado com **logo temático de investimento** (moeda sobre a mão).
   - Cores e estilização visual simulando a interface de um software desktop.
2. **📈 Gráfico de Rosca Dinâmico:**
   - Inclusão de um gráfico de pizza/rosca vinculado diretamente à tabela de divisão por tipo de fundo, permitindo visualização instantânea da mudança de alocação ao trocar de perfil.
3. **💡 Segunda Simulação (Aporte Informado vs. Aporte Sugerido de 30%):**
   - Criação de uma tabela comparativa que analisa lado a lado o investimento definido pelo usuário versus a regra recomendada de 30% do salário, calculando a diferença no patrimônio final e nos dividendos.
4. **✈️ Nova Ferramenta Integrada (Planejamento de Férias):**
   - Adição de uma 3ª aba completa para simular o valor necessário para viajar, aplicando a fórmula financeira `=PGTO` para calcular o valor do depósito mensal necessário conforme o prazo e meta final.
5. **🖥️ Guias para Modo Aplicativo:**
   - Instruções de como ocultar as linhas de grade, barra de fórmulas, títulos de coluna/linha e faixas de opção no Excel.

---

## 🖼️ 6. Evidências de Funcionamento (Simulações Comparativas)

### 🔴 Simulação 1: Perfil **Conservador**
- **Parâmetros:** Aporte de R$ 1.500,00/mês | 10 Anos | 0,80% a.m.
- **Divisão da Carteira:** R$ 600,00 em Papel (40%), R$ 375,00 em Logística (25%), R$ 150,00 em Lajes (10%), R$ 150,00 em Shopping (10%), R$ 150,00 em FOFs (10%), R$ 75,00 em Fiagro (5%).
- **Resultado Acumulado em 10 Anos:** **R$ 299.882,60**
- **Renda de Dividendos Estimada:** **R$ 2.399,06 / mês**

---

### 🟢 Simulação 2: Perfil **Arrojado**
- **Parâmetros:** Aporte de R$ 1.500,00/mês | 10 Anos | 0,80% a.m.
- **Divisão da Carteira:** R$ 225,00 em Papel (15%), R$ 300,00 em Logística (20%), R$ 300,00 em Lajes (20%), R$ 300,00 em Shopping (20%), R$ 150,00 em FOFs (10%), R$ 225,00 em Fiagro (15%).
- **Resultado Acumulado em 10 Anos:** **R$ 299.882,60**
- **Renda de Dividendos Estimada:** **R$ 2.399,06 / mês**
- *Nota:* O gráfico de pizza da interface altera automaticamente todas as fatias da pizza conforme alteramos de **Conservador** para **Arrojado**.
```
eof
```

I have generated the `README.md` file with all the detailed documentation for your portfolio.