# Saldo · Leitor da sua planilha financeira

Um painel estático feito para o formato da sua **planilha original**. Mantém o Excel/OneDrive como fonte dos dados e acrescenta gráficos, leitura dos ciclos, parcelas e acompanhamento de metas.

## Abrir no computador

1. Extraia **todo** o ZIP em uma pasta.
2. Abra `index.html` no Chrome, Edge, Firefox ou outro navegador moderno.
3. Clique em **Escolher planilha** e selecione a cópia `.xlsx` da sua planilha original.

Não precisa de Node.js, npm, Vue, servidor local, conta ou instalação. As bibliotecas estão incluídas na pasta `vendor`. Funciona sem internet depois de baixar o projeto. Não abra o HTML de dentro do visualizador do ZIP: extraia a pasta primeiro.

Use **Explorar um exemplo** para conhecer a interface com dados fictícios. O exemplo não contém seus dados.

## O que há no painel

- **Visão geral:** saldo dos ciclos fechados, previsão final, sobra do ciclo, gráficos de saldo e de entradas/saídas, grupos de gastos e metas prioritárias.
- **Ciclos e gastos:** os 12 ciclos do ano, consulta aos lançamentos, anotações opcionais de orçamento, parcelas e resumo do pagamento da moto.
- **Metas:** metas originais sem duplicar a lista de prioridades; valor guardado, aporte, prazo, progresso, aporte necessário para o prazo e metas locais. As outras listas ficam disponíveis para inclusão voluntária.
- **Planilhas:** consulta de todas as abas, busca e opção de mostrar as fórmulas como texto.

### Atualizar e fechar os ciclos

Continue lançando os dados no Excel Web. Quando quiser atualizar o painel, baixe a planilha atualizada e use **Trocar planilha**.

O leitor **não sincroniza com o OneDrive**, não escreve no arquivo e não recalcula fórmulas. Usa os resultados salvos pelo Excel. Se faltar um resultado necessário, pede uma nova cópia salva no Excel Web em vez de inventar valores.

No seletor **Ciclos fechados até**, ajuste o marco que separa os ciclos fechados da previsão. O padrão para 2026 é **Out–Nov**, conforme seu fechamento informado. Outros anos começam sem ciclos fechados definidos. “Fechado” é uma preferência do painel, não uma validação de extrato bancário.

O saldo inicial, as receitas e as despesas permanecem conforme os resultados do arquivo. A renda de cada ciclo é obtida por `despesas + sobra`, preservando o valor usado no quadro daquele ciclo. O orçamento local é opcional; diferenças entre realizado e orçamento só aparecem com o ciclo fechado e um orçamento informado. A previsão original dos ciclos passados não é inventada.

### Metas e backup

Os objetivos são lidos da aba Financeiro mais recente que tenha metas preenchidas. Itens repetidos na coluna de prioridades são contados uma única vez e mantêm prioridade alta.

Nome e custo das metas importadas vêm da planilha. Atualize-os no Excel e reabra o arquivo. Você também pode criar metas locais. Valor guardado, aporte mensal, prazo, prioridades, metas ocultas, orçamentos e marco de fechamento ficam no armazenamento deste navegador.

**Exporte um backup** depois de atualizar suas metas. O backup é um JSON com suas anotações, não com a planilha inteira. Pode ser restaurado depois de abrir seu `.xlsx` em outro navegador ou computador. As anotações não sincronizam entre dispositivos. Limpar os dados do navegador, mudar o endereço do site ou renomear o arquivo pode mudar o contexto do armazenamento; use o backup para transferir os dados.

Deixar “guardado” em branco significa **não informado**. Preencher `0` significa que você ainda não reservou dinheiro. O aporte estimado usa meses do calendário, sem rendimentos e sem prometer uma data de compra. Aportes e parcelas não são descontados uma segunda vez dos saldos do arquivo.

## Subir no GitHub Pages

1. Crie um repositório e envie **o conteúdo desta pasta**, mantendo `index.html` na raiz e a pasta `vendor` completa.
2. Em **Settings → Pages**, selecione **Deploy from a branch**.
3. Escolha a branch `main` e a pasta `/ (root)`, e salve.
4. Abra o endereço informado pelo GitHub e escolha seu `.xlsx` pelo botão do painel.

Não há comando de build ou GitHub Actions personalizado para configurar. A documentação oficial é [Criar um site do GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site).

**Publique só o código. Não envie sua planilha ou seus backups ao repositório.** Não há dados reais incluídos no projeto. O `.gitignore` ajuda a evitar arquivos `.xlsx` e backups em commits feitos por Git; uploads manuais no site do GitHub exigem esse mesmo cuidado.

## Privacidade e funcionamento

O arquivo é processado na memória do navegador. O projeto não contém analytics, telemetria, chamadas de API, backend ou envio de dados. A política de conteúdo bloqueia conexões de dados. O acesso normal aos arquivos do site, quando hospedado, continua sujeito à infraestrutura do provedor da hospedagem.

As bibliotecas são locais. Para ver o funcionamento, use as ferramentas de rede do navegador: depois de carregar HTML, CSS e JavaScript, escolher uma planilha não dispara upload ou requisição.

## Arquivos do projeto

| Arquivo | Papel |
| --- | --- |
| `index.html` | Estrutura e entrada do painel |
| `styles.css` | Estilo e responsividade |
| `model.js` | Leitura do formato financeiro, metas e validação do backup |
| `app.js` | Interface, gráficos e armazenamento local |
| `vendor/` | SheetJS, Chart.js e suas licenças |

O leitor reconhece abas chamadas `Financeiro 20XX`, com os 12 blocos de `Sobra Total` no formato da sua planilha. Não é um importador universal de qualquer planilha financeira. O visualizador consulta até 500 linhas e 26 colunas por aba. Mudanças substanciais na estrutura podem exigir adaptar `model.js`.

## Bibliotecas

- [SheetJS CE 0.20.3](https://docs.sheetjs.com/docs/getting-started/installation/standalone/) — Apache-2.0, incluída localmente conforme a orientação oficial de vendoring.
- [Chart.js 4.6.1](https://www.chartjs.org/docs/latest/getting-started/integration.html) — MIT, distribuição UMD incluída localmente.

As licenças completas estão na pasta `vendor`.

## Verificação

A leitura foi conferida com o arquivo original, incluindo saldos de 2026/2027, 14 metas sem duplicatas, parcelas, valores ausentes e backups. Os fluxos de importação, troca de ano, metas, orçamento, backup, restauração e persistência foram testados na interface em uma prévia HTTP local.

O navegador integrado de testes bloqueia endereços `file://`, por isso a abertura por duplo clique não foi validada por essa ferramenta. O projeto usa scripts clássicos e arquivos locais, sem módulos ou downloads externos, para permitir esse modo de uso sem servidor.
