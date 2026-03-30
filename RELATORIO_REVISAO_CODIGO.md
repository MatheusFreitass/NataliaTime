# Relatório de Revisão de Código — NataliaTime v1.1.0
**Data:** 2026-03-30
**Arquivos revisados:** `analise_natalia_time.py`, `NT_app_ctk.py`, `Migrador NT/migrador.py`
**Nota:** Este é um relatório de observação. Nenhuma alteração foi feita no código.

---

## 1. `analise_natalia_time.py`

### 1.1 Bugs e Problemas Confirmados

#### 🔴 ALTO — Leitura de Excel silenciosamente incompleta (linhas ~344–356)
Células Excel são lidas com `try/except` genérico. Se uma célula obrigatória retornar `None` ou um tipo inesperado, o parâmetro fica como `None` sem aviso explícito ao usuário. Parâmetros críticos como `m_frag`, `DE`, `LL` podem entrar na análise como `None`, causando falha silenciosa ou resultados fisicamente absurdos. O fallback `FIXOS_FALLBACK` cobre alguns casos, mas não todos.

#### 🔴 ALTO — Crash se leitura CSV falha parcialmente (linha ~462, ~490)
Se `ler_dados_csv()` lança `ValueError` após percorrer o arquivo mas antes de chamar `inicializar_grade()`, a variável global `GRADE_BM` permanece `None`. O código subsequente acessa `GRADE_BM[0]` sem verificação, causando `AttributeError` ou `TypeError`.

#### 🟡 MÉDIO — Divisão quase-por-zero na normalização de amplitudes (linhas ~601, 609, 620, 631, 642)
A guarda `1.0 / int_e if int_e > 0 else 1.0` protege contra zero exato, mas não contra valores extremamente próximos de zero (ex: `1e-15`). Isso ocorre quando β é muito grande junto com energia alta, gerando amplitudes numericamente gigantes que contaminam a soma ponderada e produzem ajuste aparentemente válido com parâmetros fisicamente impossíveis.

#### 🟡 MÉDIO — Race condition em multiprocessing com variável global `INTERVALOS_BETA` (linha ~1216)
`busca_duas_fases()` pode modificar `INTERVALOS_BETA` em diferentes ciclos. Como `multiprocessing.Pool` usa `spawn` no Windows (e `fork` no Linux), o estado da variável global pode diferir entre parent e workers dependendo de quando os processos são criados.

#### 🟡 MÉDIO — Extrapolação silenciosa em `np.interp` (linha ~795)
`np.interp(tempos_exp, GRADE_BM, bs)` extrapola silenciosamente com o valor do extremo se `tempos_exp` contém valores fora de `[GRADE_BM[0], GRADE_BM[-1]]`. Com dados mal formatados (offset de tempo errado no Migrador, por exemplo), o espectro teórico fica errado mas o fitting continua.

#### 🟡 MÉDIO — Memory leak potencial no `multiprocessing.Pool` (linhas ~1122–1147)
Se um processo worker morre abruptamente (segfault, OOM) enquanto o pipe de comunicação está aberto, `pool.join()` após `pool.terminate()` pode bloquear indefinidamente. O código não tem timeout no join.

#### 🟢 BAIXO — Plotagem com `sinal_exp` todo negativo ou zero (linha ~1737)
`spos = sinal_exp[sinal_exp > 0]` pode retornar array vazio se os dados experimentais são todos ≤ 0 (dados corrompidos). `np.min(spos)` lançaria `ValueError`. Caso raro mas possível com arquivos CSV malformados.

#### 🟢 BAIXO — PARAR.txt residual pode parar próxima execução (linhas ~1099, ~2374)
Linha 2374 remove o arquivo residual *antes* de iniciar, o que está correto. Porém, se o processo é morto com SIGKILL (não gracefully), o arquivo pode persistir no disco. A próxima execução do *mesmo* experimento na *mesma* pasta iria detectá-lo antes de limpá-lo — mas como a pasta é nova (timestamp no nome), esse cenário não ocorre. Não é um bug ativo.

#### 🟢 BAIXO — Edge case em clustering com 1 candidato válido (linha ~1399)
Se `len(validos_df) == 1`, o clustering hierárquico produz 1 cluster com threshold calculado. Nelder-Mead refina apenas 1 candidato em vez de `NM_TOP_CANDIDATOS`, reduzindo cobertura do espaço de parâmetros. Não falha, mas não é o comportamento esperado.

---

### 1.2 Gargalos de Performance

#### 🟡 MÉDIO — Recálculo desnecessário de vetores base em Nelder-Mead
`gh_exp()` e `gh_gauss()` (chamadas dentro de `calcular_espectro()`) recalculam via `np.einsum` mesmo quando β não mudou entre iterações do NM. Para cada candidato NM com múltiplas avaliações da função objetivo, esses vetores são idênticos. Um cache simples com chave `(beta_exp, beta_g1, beta_g2, beta_g3)` eliminaria ~30% do overhead na fase 2.

#### 🟡 MÉDIO — Geração de múltiplos PNGs sequencial e síncrona (linhas ~2781–2814)
Com `SALVAR_GRAFICOS_TOP = 20`, gera 20 gráficos sequencialmente na thread principal usando matplotlib (que reinicializa figura a cada vez). Isso pode levar 10–30 segundos dependendo da máquina. Solução natural seria usar Pool para paralelizar, mas matplotlib não é thread-safe.

#### 🟢 BAIXO — DataFrames recriados repetidamente (linhas ~2489, ~2572, ~2735)
`df[df["status"]=="VÁLIDO"].sort_values("rmse")` é executado 3 vezes em locais diferentes com o mesmo `historico_completo`. O resultado deveria ser calculado uma vez e reutilizado.

---

## 2. `NT_app_ctk.py`

### 2.1 Bugs e Problemas Confirmados

#### 🟡 MÉDIO — Arquivos temporários não deletados em caso de exceção (linhas ~1151–1152)
O `try: os.unlink(tmp) except: pass` usa bare `except`, silenciando qualquer erro. No Windows, se o processo filho ainda tem o arquivo aberto (não fechou stdout pipe antes do unlink), a deleção falha silenciosamente. Ao longo do tempo, acumula-se arquivos `.py` temporários no diretório de temp do sistema.

#### 🟡 MÉDIO — Deadlock potencial no loop de leitura de stdout (linhas ~1138–1162)
A thread `_runner()` usa `proc.stdout.readline()` bloqueante enquanto a main thread usa `Queue.get_nowait()` com sleep de 300ms. Se o processo filho produz output muito rapidamente (>= buffer do pipe), `readline()` pode bloquear enquanto a queue está cheia, criando pressão de back-pressure. Com output de >~10MB, a UI pode congelar.

#### 🟡 MÉDIO — Geração de script Python por f-string sem escape de paths (linhas ~159–247)
`gerar_config_python()` usa `repr()` para caminhos, o que é correto para paths simples. Porém, se `cfg['excel_path']` contém caracteres que quebram a f-string (ex: `\N{LATIN SMALL LETTER...}` interpretado como escape Unicode), o script gerado pode ter erro de sintaxe que só é percebido no subprocess.

#### 🟡 MÉDIO — Detecção de pasta de resultados por string matching (linhas ~1142–1143)
`"Resultados" in token` é um matching simples de substring no output do subprocess. Se o usuário tem pastas com "Resultados" no caminho pai (ex: `C:\Resultados da Faculdade\NataliaTime\...`), o caminho capturado pode estar errado.

#### 🟢 BAIXO — `_atualizar_lista_seq` destrói e recria todos os widgets (linhas ~993–1004)
A cada adição/remoção na sequência, todos os widgets são destruídos e recriados. Com sequências longas (>20 experimentos), isso gera ~100ms de latência perceptível na UI.

#### 🟢 BAIXO — Valores `NaN` ou `Inf` não validados antes de escrever TOML (linhas ~945–962)
Campos numéricos do formulário são passados direto para o TOML sem verificar se são valores válidos. Se o usuário digita "inf" ou "nan" em um campo, o TOML gerado é sintaticamente inválido e o subprocess falha com mensagem obscura.

---

### 2.2 Gargalos de Performance

#### 🟢 BAIXO — Logo carregada duas vezes em resolução diferente (linhas ~283, ~396)
A imagem `logo_lms.png` é aberta com PIL e convertida para CTkImage duas vezes (34px para o header, 78px para a aba Início). Pode ser carregada uma vez e redimensionada nas duas resoluções.

---

## 3. `Migrador NT/migrador.py`

### 3.1 Bugs e Problemas Confirmados

#### 🟡 MÉDIO — Parsing de dados com tab e vírgula simultâneos (linha ~341)
O parser usa: `linha.split("\t") if "\t" in linha else linha.replace(";",",").split(",")`. Se uma linha contém ambos tab e vírgula (ex: dado copiado de Excel com formatação mista), o split por tab é usado e a vírgula no valor numérico não é processada corretamente — `"0.9,5"` não seria convertível para float.

#### 🟢 BAIXO — Offset de tempo aplicado sem validação de sinal (linha ~441)
`f"{t + t_offset},{s}\n"` aplica o offset sem verificar se o resultado é negativo. Tempo negativo não tem sentido físico no contexto DETOF e pode causar comportamento indefinido na análise.

#### 🟢 BAIXO — `os.startfile()` falha silenciosamente (linhas ~476–479)
O `try/except` silencia falhas ao abrir a pasta no Explorer. Se a pasta foi deletada entre criação e abertura, o usuário não recebe feedback.

#### 🟢 BAIXO — Placeholder não resetado em cola com Middle-Click (linhas ~266–277)
O sistema de placeholder (texto cinza que some ao focar) é controlado por `<FocusIn>/<FocusOut>`. Colar texto via middle-click (Linux/X11) não dispara `<FocusIn>`, então o placeholder pode coexistir com dados reais no widget.

---

## Resumo Executivo

| Severidade | analise_natalia_time.py | NT_app_ctk.py | migrador.py | Total |
|---|---|---|---|---|
| 🔴 ALTO | 2 | 0 | 0 | **2** |
| 🟡 MÉDIO | 3 | 4 | 1 | **8** |
| 🟢 BAIXO | 4 | 2 | 3 | **9** |
| **Total** | **9** | **6** | **4** | **19** |

**Prioridade imediata sugerida (quando o usuário retornar):**
1. Validar parâmetros `None` após leitura de Excel (`analise_natalia_time.py`)
2. Proteger contra `GRADE_BM` não inicializada em modo CSV (`analise_natalia_time.py`)
3. Adicionar timeout ao `pool.join()` (`analise_natalia_time.py`)
4. Validar campos numéricos antes de escrever TOML (`NT_app_ctk.py`)

---
*Relatório gerado automaticamente em revisão estática. Nenhuma alteração foi feita no código.*
