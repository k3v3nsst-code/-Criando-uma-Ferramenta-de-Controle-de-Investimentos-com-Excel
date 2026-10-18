# Simulador de Renda Passiva com FIIs (Excel)

Uma planilha com cara de aplicativo que responde às cinco perguntas de negócio de quem quer viver de dividendos de fundos imobiliários:

1. Quanto investir por mês?
2. Por quantos anos?
3. Qual a taxa de rendimento mensal?
4. Quanto de patrimônio vai acumular?
5. Quanto vai receber de dividendos por mês?

Além disso, ela projeta o patrimônio em cenários de 2 a 30 anos, compara o aporte informado com a sugestão de 30% do salário, calcula quanto tempo leva para atingir uma meta de renda e divide o aporte entre seis tipos de fundo imobiliário conforme o perfil de investidor escolhido numa lista.

> **Aviso:** os percentuais de cada perfil, o salário, a taxa de 0,8% ao mês e a meta de R$ 5.000 são **valores de exemplo**. Não constituem recomendação de investimento. Use os seus.

## Arquivos

| Arquivo | Para que serve |

| `simulador_fiis_final.xlsx` | Mesma planilha com guias e títulos escondidos (navegação por links) |

## Estrutura da pasta de trabalho

| Aba | Conteúdo |
|---|---|
| **Simulador** | A "tela do app": entradas, resultado, cenários, divisão por perfil, meta de renda e gráficos |
| **Configurações** | Salário, rendimento da carteira, % do aporte, meta de renda e sugestão de aporte |
| **Apoio** | Tabela perfil × tipo de fundo, listas e tabela de evolução ano a ano |

### Convenção de cores

| Cor | Significado |
|---|---|
| Azul sobre amarelo claro (`FFF9DB`) | Célula que você edita |
| Cinza (`E4E4E4`) | Célula calculada — não edite |
| Verde | Valor que vem de outra aba |
| Azul-marinho (`1F3864`) | Banner |
| Azul (`2E75B6`) | Faixas de seção |

> As fórmulas abaixo estão em português (Excel em PT-BR, separador `;`). Em Excel em inglês: `VF`=`FV`, `PROCV`=`VLOOKUP`, `SEERRO`=`IFERROR`, `SOMASE`=`SUMIF`, `SE`=`IF`, `SOMA`=`SUM`, `ARRED`=`ROUND`.

---

## Passo a passo

> **Se você está começando:** faça primeiro os passos 1 e 2 (as cinco perguntas funcionando). Cenários e divisão por perfil vêm depois. Uma parte bem feita vale mais que a ferramenta inteira pela metade.

### Passo 1 — Monte a tela

1. Abra uma pasta de trabalho em branco e renomeie a primeira aba para **Simulador**.
2. Em *Exibição*, desmarque **Linhas de Grade**.
3. Mescle `B1:Q2` e escreva o nome da ferramenta (fonte grande, branca, negrito). Mescle `B3:Q3` para um subtítulo. Pinte as duas faixas de azul-marinho. *(O banner é opcional.)*
4. Defina as larguras: coluna A estreita (2), B larga (38), C a F com 22.
5. Monte o bloco **1 · ENTRADAS**:

| Célula | Rótulo | Valor de exemplo |
|---|---|---|
| `B6` / `C6` | Quanto investir por mês (R$) | `1000` |
| `B7` / `C7` | Por quantos anos | `20` |
| `B8` / `C8` | Taxa de rendimento mensal | `0,8%` |
| `B9` / `C9` | Perfil de investidor | `Moderado` |

6. Pinte de **cinza** toda célula que contém fórmula, para sinalizar que ninguém deve editá-la.

### Passo 2 — Calcule o patrimônio e os dividendos

Com a função `VF` (valor futuro), chegue ao patrimônio acumulado e, a partir dele, aos dividendos mensais:

```excel
Patrimônio acumulado:  =VF(C8; C7*12; -C6)
Dividendos por mês:    =patrimônio * C8
Total investido:       =C6*C7*12
Ganho com juros:       =patrimônio - total investido
```

- `C8` é a taxa mensal, `C7*12` converte anos em meses e `-C6` é o aporte (negativo, porque é saída de caixa; assim o resultado vem positivo).
- Os aportes são feitos no **fim de cada mês** (argumento `tipo` omitido = 0).
- Os dividendos mensais são o patrimônio multiplicado pela taxa mensal.

Neste ponto as **cinco perguntas já funcionam**.

### Passo 3 — Projete os cenários

Crie a tabela de cenários com os prazos **2, 5, 10, 20 e 30 anos** (um por linha) e repita o cálculo do `VF` em cada linha.

Use **referências absolutas** (`$`) para as células que não mudam, assim a fórmula não se perde ao ser arrastada:

```excel
=VF($C$8; B21*12; -$C$6)
```

- `B21` (prazo da linha) é relativa: muda ao arrastar.
- `$C$8` e `$C$6` são travadas: continuam apontando para a taxa e o aporte.

### Passo 4 — Crie as configurações e os intervalos nomeados

1. Crie a aba **Configurações** e preencha:

| Célula | Rótulo | Exemplo | Nome |
|---|---|---|---|
| `C6` | Salário mensal líquido (R$) | `6000` | `salario` |
| `C7` | Rendimento mensal da carteira | `0,8%` | `taxa_mensal` |
| `C8` | % do salário destinado ao aporte | `30%` | `pct_aporte` |
| `C9` | Meta de dividendos por mês (R$) | `5000` | `meta_renda` |
| `C10` | **Sugestão de aporte** (calculada) | `=salario*pct_aporte` | `sugestao_aporte` |

2. Dê nome às células: selecione a célula, clique na *Caixa de Nome* (à esquerda da barra de fórmulas), digite o nome e tecle Enter. Ou use *Fórmulas > Gerenciador de Nomes*.
3. No **Simulador**, nomeie também:

| Célula | Nome |
|---|---|
| `C6` | `aporte` |
| `C7` | `prazo_anos` |
| `C9` | `perfil` |
| `C14` | `patrimonio` |
| `C15` | `dividendos` |

4. **Troque os endereços de célula pelos nomes** nas fórmulas. A fórmula do passo 3 fica legível:

```excel
=VF(taxa_mensal; B21*12; -aporte)
```

5. Mostre a sugestão ao lado do aporte (`D6 = sugestao_aporte`) e limite as entradas com **Validação de Dados** (*Dados > Validação de Dados*):

| Célula | Regra |
|---|---|
| `C6` | Decimal ≥ 0 |
| `C7` | Número inteiro entre 1 e 50 |
| `Configurações!C7` | Decimal entre 0% e 10% |

### Passo 5 — Divida o aporte por perfil

**5.1 Crie a aba Apoio** com três blocos:

| Bloco | Intervalo | Nome |
|---|---|---|
| Lista de perfis (Conservador, Moderado, Arrojado) | `G7:G9` | `lista_perfis` |
| Lista de tipos de fundo | `K7:K12` | `lista_tipos` |
| Tabela perfil × tipo × percentual (18 linhas) | `B7:E24` | `tabela_perfis` |

Os seis tipos de fundo: **Papel (CRI)**, **Logística**, **Lajes Corporativas**, **Shoppings**, **Fundo de Fundos (FoF)** e **Híbridos**.

**5.2 Monte a tabela com chave composta.** O `PROCV` só procura na coluna mais à esquerda, então a primeira coluna da tabela (`B`) concatena perfil e tipo:

| B — Chave | C — Perfil | D — Tipo de fundo | E — Percentual |
|---|---|---|---|
| `=C7&"\|"&D7` | `=$G$7` | `=$K$7` | `35%` |

O separador `|` evita ambiguidade entre chaves. As colunas `C` e `D` são fórmulas que apontam para as listas — se você renomear um perfil na lista, a tabela acompanha.

**Percentuais de exemplo** (cada perfil soma 100%):

| Tipo de fundo | Conservador | Moderado | Arrojado |
|---|---:|---:|---:|
| Papel (CRI) | 35% | 25% | 15% |
| Logística | 20% | 20% | 20% |
| Lajes Corporativas | 5% | 15% | 20% |
| Shoppings | 10% | 15% | 20% |
| Fundo de Fundos (FoF) | 25% | 15% | 5% |
| Híbridos | 5% | 10% | 20% |
| **Total** | **100%** | **100%** | **100%** |

**5.3 Confira se cada perfil soma 100%** (`H7:I9`):

```excel
Soma:    =SOMASE($C$7:$C$24; G7; $E$7:$E$24)
Status:  =SE(ARRED(H7;4)=1; "OK"; "Revisar: deve somar 100%")
```

Aplique *Formatação Condicional* para destacar em vermelho quando o status for diferente de "OK".

**5.4 Crie a lista suspensa de perfis** no Simulador: selecione `C9` → *Dados > Validação de Dados* → **Lista** → origem `=lista_perfis`.

**5.5 Traga cada percentual com `PROCV` e a chave composta.** No Simulador, seção **4 · DIVISÃO DO APORTE POR PERFIL**, com os tipos de fundo em `B29:B34` (`=Apoio!$K$7` etc.):

```excel
Percentual (C29):             =SEERRO(PROCV(perfil&"|"&B29; tabela_perfis; 4; FALSO); 0)
Aporte mensal (D29):          =aporte * C29
Patrimônio projetado (E29):   =patrimonio * C29
Dividendos por mês (F29):     =dividendos * C29
```

Some a linha de total em `C35:F35` com `=SOMA(C29:C34)`. O total de percentuais deve ser 100% (use formatação condicional para alertar se não for).

---

## Evoluções incluídas

### Segunda simulação: aporte informado × sugestão de 30% do salário

A seção **2 · RESULTADO** tem duas colunas lado a lado. A coluna "Aporte sugerido" usa o mesmo cálculo, trocando `aporte` por `sugestao_aporte`:

```excel
Patrimônio (informado):  =VF(taxa_mensal; prazo_anos*12; -aporte)
Patrimônio (sugerido):   =VF(taxa_mensal; prazo_anos*12; -sugestao_aporte)
Dividendos (sugerido):   =patrimonio_sugestao * taxa_mensal
```

Nomes adicionais: `patrimonio_sugestao` (`D14`) e `dividendos_sugestao` (`D15`). A tabela de cenários também ganha as colunas "sugerido".

### Meta de renda: quanto tempo para chegar lá

```excel
Patrimônio necessário (C39):  =SE(taxa_mensal>0; meta_renda/taxa_mensal; 0)     → nome: patrimonio_meta
Anos para a meta (C41):       =SEERRO(NPER(taxa_mensal; -aporte; 0; patrimonio_meta)/12; "—")
Anos para a meta (D41):       =SEERRO(NPER(taxa_mensal; -sugestao_aporte; 0; patrimonio_meta)/12; "—")
```

Lógica: para receber uma renda mensal `R` com taxa mensal `i`, é preciso um patrimônio de `R / i`. A função `NPER` diz em quantos meses o aporte mensal chega a esse patrimônio.

### Tabela de apoio de evolução anual

Na aba **Apoio**, o intervalo `M7:Q36` (nome `tabela_evolucao`) tem 30 linhas, uma por ano:

| Coluna | Conteúdo | Fórmula (linha 7) |
|---|---|---|
| M | Ano (1 a 30) | `1`, depois `=M7+1` |
| N | Patrimônio · informado | `=VF(taxa_mensal; M7*12; -aporte)` |
| O | Patrimônio · sugerido | `=VF(taxa_mensal; M7*12; -sugestao_aporte)` |
| P | Dividendos/mês · informado | `=N7*taxa_mensal` |
| Q | Dividendos/mês · sugerido | `=O7*taxa_mensal` |

### Gráficos

- **Linha — evolução do patrimônio:** *Inserir > Gráfico de Linhas*, usando `Apoio!M6:O36` (anos como categorias e as duas colunas de patrimônio como séries).
- **Pizza — divisão do aporte:** *Inserir > Gráfico de Pizza*, usando `Simulador!B29:C34`. Como os percentuais vêm do `PROCV`, o gráfico **muda sozinho** quando você troca o perfil na lista. Ative os rótulos de dados em percentual.

### Versão final (interface limpa)

Para a versão de entrega:

1. *Exibição*: desmarque **Títulos** (cabeçalhos de linha e coluna) em cada aba.
2. *Arquivo > Opções > Avançado*: desmarque **Mostrar guias de planilha**.
3. Como as guias somem, crie **links de navegação** com `=HIPERLINK("#'Configurações'!C6"; "→ Configurações")` e `=HIPERLINK("#Simulador!C6"; "← Voltar ao Simulador")`.
4. **Barra de fórmulas:** é uma configuração do aplicativo Excel, não do arquivo. Desmarque em *Exibição > Barra de Fórmulas* ou use uma macro (`Application.DisplayFormulaBar = False`).

---

## Referência rápida dos intervalos nomeados

| Nome | Local | Descrição |
|---|---|---|
| `salario` | `Configurações!C6` | Salário mensal líquido |
| `taxa_mensal` | `Configurações!C7` | Rendimento mensal da carteira |
| `pct_aporte` | `Configurações!C8` | % do salário destinado ao aporte |
| `meta_renda` | `Configurações!C9` | Meta de dividendos por mês |
| `sugestao_aporte` | `Configurações!C10` | Salário × percentual |
| `aporte` | `Simulador!C6` | Aporte mensal informado |
| `prazo_anos` | `Simulador!C7` | Prazo em anos |
| `perfil` | `Simulador!C9` | Perfil escolhido na lista |
| `patrimonio` | `Simulador!C14` | Patrimônio com o aporte informado |
| `dividendos` | `Simulador!C15` | Dividendos com o aporte informado |
| `patrimonio_sugestao` | `Simulador!D14` | Patrimônio com o aporte sugerido |
| `dividendos_sugestao` | `Simulador!D15` | Dividendos com o aporte sugerido |
| `patrimonio_meta` | `Simulador!C39` | Patrimônio necessário para a meta |
| `tabela_perfis` | `Apoio!B7:E24` | Chave, perfil, tipo e percentual |
| `lista_perfis` | `Apoio!G7:G9` | Lista suspensa de perfis |
| `lista_tipos` | `Apoio!K7:K12` | Tipos de fundo |
| `tabela_evolucao` | `Apoio!M7:Q36` | Evolução ano a ano |

## Premissas e limitações

- Aportes no fim de cada mês, com rendimento **reinvestido** (juros compostos).
- Dividendos mensais = patrimônio acumulado × taxa de rendimento mensal.
- Sem impostos, taxas, inflação ou variação do preço das cotas.
- A taxa de rendimento é constante ao longo de todo o período.
- É uma ferramenta de simulação e aprendizado, **não** uma recomendação de investimento.

## Ideia para a próxima ferramenta

Planejamento do dinheiro das férias, com as mesmas técnicas: quanto custa a viagem, em quantos meses, quanto guardar por mês e onde guardar (`VF`, `NPER`, `PGTO`, intervalos nomeados, tabela de apoio e lista suspensa).
