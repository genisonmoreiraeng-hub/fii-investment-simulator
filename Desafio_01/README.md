# 🏢 Simulador Interativo de Investimentos em FIIs & Planejamento Financeiro

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
![Status](https://img.shields.io/badge/Status-Conclu%C3%ADdo-success?style=for-the-badge)

> Dashboard em Excel/Google Sheets com interface em estilo aplicativo para projeção de patrimônio, acúmulo de dividendos mensais, cenários de juros compostos e alocação por perfil de investidor.

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

* **`taxa_mensal`**: Taxa de juros por período (ex: `0,80%`).
* **`prazo_anos * 12`**: Número total de pagamentos/meses (ex: 10 anos = 120 meses).
* **`-aporte`**: O valor depositado mensalmente. Utiliza-se o sinal negativo (`-`) para que o resultado do patrimônio retorne como um valor positivo no Excel.
* **`0`**: Pagamento no final de cada período (postecipado).

---

### 🔹 Função `PROCV` (Procura Vertical) com Chave Composta

Para determinar o percentual de investimento em cada um dos 6 tipos de fundos imobiliários com base no perfil escolhido pelo usuário (*Conservador*, *Moderado* ou *Arrojado*), utilizou-se uma **Chave Composta**.

**Lógica da Chave Composta:**
Na aba de apoio (`Apoio`), criou-se uma coluna auxiliar combinando o Nome do Perfil com o Tipo de Fundo:

```text
Chave Composta = Perfil & "_" & Tipo_de_Fundo
```

*Exemplo prático de valor gerado:* `"Moderado_FIIs de Tijolo - Logística"`

**Sintaxe aplicada na planilha:**

```excel
=PROCV(B$11 & "_" & A20; Apoio!C2:D20; 2; FALSO)
```

* **`B$11`**: Célula com o Perfil selecionado pelo usuário na aba `Simulador` (ex: `"Moderado"`).
* **`A20`**: Célula com o Tipo do Fundo Imobiliário (ex: `"FIIs de Tijolo - Logística"`).
* **`Apoio!C2:D20`**: Matriz de busca na aba de apoio contendo as chaves compostas na primeira coluna e os percentuais na segunda coluna.
* O `PROCV` busca a combinação exata `Moderado_FIIs de Tijolo - Logística` e retorna a alocação percentual correspondente.

---

## 🏷️ 3. Intervalos Nomeados Criados

Para tornar as fórmulas legíveis e evitar erros de referência ao arrastar células, os principais campos da planilha foram mapeados como **Intervalos Nomeados**:

1. `salario`: Célula contendo o salário mensal informado pelo usuário.
2. `aporte`: Célula com o valor do aporte mensal efetivo definido para investimento.
3. `aporte_sugerido`: Célula calculada com a sugestão de 30% do salário (`=salario * 0,30`).
4. `taxa_mensal`: Célula contendo a taxa de rendimento esperada ao mês (ex: `0,80%`).
5. `prazo_anos`: Célula contendo o horizonte de tempo em anos (ex: `10`).
6. `perfil_investidor`: Célula com a lista suspensa de Validação de Dados (`Conservador`, `Moderado`, `Arrojado`).
7. `tabela_alocacao`: Intervalo na aba `Apoio` utilizado para as consultas dinâmicas do `PROCV`.

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
* **Conservador:** Prioriza menor volatilidade e alta previsibilidade. Concentra 40% em FIIs de Papel/CRI e 25% em Galpões Logísticos (contratos atípicos e inquilinos corporativos).
* **Moderado:** Busca equilíbrio entre renda periódica e valorização de cotas. Equilibra ativamente segmentos de Tijolo (Logística, Lajes, Shopping) com títulos de dívida (Papel).
* **Arrojado:** Aceita maior oscilação em busca de maior valorização patrimonial. Aumenta a exposição em Lajes Corporativas, Shoppings e Fiagros (15%), reduzindo títulos fixos de dívida.

---

## 🔄 5. O Que Foi Evoluído em Relação à Ferramenta Base

1. **🎨 Identidade Visual & Branding:**
   * Adicionado banner estilizado no topo com **logo temático de investimentos** (moeda sobre a mão).
   * Paleta visual profissional simulando a interface de um aplicativo desktop.

2. **📈 Gráfico de Rosca Dinâmico:**
   * Inclusão de um gráfico de rosca vinculado à tabela de divisão de carteira, permitindo visualização gráfica automática ao trocar de perfil no menu suspenso.

3. **💡 Segunda Simulação (Aporte Informado vs. Aporte Sugerido de 30%):**
   * Tabela comparativa que analisa lado a lado o investimento informado versus a sugestão de 30% da renda, exibindo o ganho adicional em patrimônio final e dividendos.

4. **✈️ Nova Ferramenta Integrada (Planejamento de Férias):**
   * Adição de uma 3ª aba completa para simular custos de viagem, aplicando a fórmula financeira `=PGTO` para calcular a meta de reserva mensal necessária.

5. **🖥️ Configuração do Modo Aplicativo:**
   * Orientações para ocultar linhas de grade, barra de fórmulas, títulos de coluna e faixas de opção no Excel.

---

## 🖼️ 6. Evidências de Funcionamento (Simulações Comparativas)

### 🔴 Simulação 1: Perfil **Conservador**
* **Parâmetros:** Salário: R$ 5.000,00 | Aporte: R$ 1.500,00/mês | 10 Anos | 0,80% a.m.
* **Divisão da Carteira:** R$ 600,00 em Papel (40%), R$ 375,00 em Logística (25%), R$ 150,00 em Lajes (10%), R$ 150,00 em Shopping (10%), R$ 150,00 em FOFs (10%), R$ 75,00 em Fiagro (5%).
* **Resultado Acumulado em 10 Anos:** **R$ 299.882,60**
* **Renda de Dividendos Estimada:** **R$ 2.399,06 / mês**

### 🟢 Simulação 2: Perfil **Arrojado**
* **Parâmetros:** Salário: R$ 5.000,00 | Aporte: R$ 1.500,00/mês | 10 Anos | 0,80% a.m.
* **Divisão da Carteira:** R$ 225,00 em Papel (15%), R$ 300,00 em Logística (20%), R$ 300,00 em Lajes (20%), R$ 300,00 em Shopping (20%), R$ 150,00 em FOFs (10%), R$ 225,00 em Fiagro (15%).
* **Resultado Acumulado em 10 Anos:** **R$ 299.882,60**
* **Renda de Dividendos Estimada:** **R$ 2.399,06 / mês**
* *Nota:* O gráfico de rosca atualiza automaticamente a proporção de cada fatia de acordo com o perfil selecionado.