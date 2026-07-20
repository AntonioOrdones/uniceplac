# Painel de Alunos — UNICEPLAC 2026/1

Dashboard estático (HTML + CSS/SASS + JavaScript puro) com estatísticas descritivas dos alunos da UNICEPLAC, construído sobre a identidade visual do manual da marca da instituição e inspirado no padrão de painéis gov.br.

A página lê a planilha `dados.xlsx` diretamente no navegador (via [SheetJS](https://sheetjs.com/)) e monta todos os indicadores, gráficos e tabelas em tempo real — **não há backend, banco de dados nem etapa de build obrigatória**. Basta hospedar os arquivos.

## O que o painel mostra

- **KPIs**: total de alunos, alunos cursando, número de cursos, média e mediana de idade.
- **Gráficos (Chart.js)**: alunos por curso; média de idade por curso (com linha de referência da média institucional); faixas etárias; sexo; situação acadêmica (gráfico polar); estado civil; UF de origem; principais cidades.
- **Radar comparativo**: perfil de um curso escolhido × média institucional em 5 dimensões normalizadas (idade média, % feminino, % cursando, % residentes no DF, % solteiros) — inspirado nos exemplos de radar do amCharts.
- **Mapa de calor** curso × situação acadêmica.
- **Mapa geoespacial (Leaflet)** com a distribuição residencial dos alunos: círculos agregados por **bairro** dentro do DF e por **município** nas demais UFs, com cor e tamanho proporcionais à quantidade (tons quentes = mais alunos). Tem enquadramentos "DF e Entorno" e "Brasil", tooltips com contagens e uma nota automática com o total georreferenciado, registros sem bairro identificável e, se houver, alunos residentes no exterior.
- **Tabela de estatísticas descritivas por curso** (n, média, mediana, moda, desvio-padrão, quartis, mínimo e máximo de idade), ordenável e com exportação para CSV (formato pt-BR: `;` como separador e vírgula decimal).
- **Lista de alunos** com busca livre e paginação, respeitando os filtros globais (curso, situação, sexo e UF).

### Tratamentos de dados feitos no navegador

- Preenchimento do nome do curso nas linhas em branco (a planilha só traz o curso na primeira linha de cada bloco).
- Mesclagem de grafias duplicadas por acentuação (ex.: "EDUCAÇÃO FÍSICA" × "EDUCAÇÃO FISICA") e normalização de cidades ("Brasília" × "BRASILIA" etc.), escolhendo sempre a grafia mais frequente.
- Cálculo de idade a partir da data de nascimento, descartando valores impossíveis.
- Registros de transferência sem RA/nome são mantidos nas estatísticas demográficas e exibidos com "—" na lista.
- Para o mapa, o texto livre do campo `BAIRRO` é classificado em ~65 localidades canônicas do DF e Entorno (RAs, setores do Gama, cidades goianas limítrofes etc.), inclusive corrigindo erros comuns de digitação; fora do DF a agregação é por município, com recuo para a capital da UF (marcado como "posição aproximada") quando o município não é reconhecido. As coordenadas são aproximadas — representam o centro do bairro/município, não endereços.

## Estrutura do projeto

```
painel-alunos-uniceplac/
├── index.html                  # Página única do painel
├── dados.xlsx                  # Planilha de dados (lida via fetch)
├── css/
│   └── style.css               # CSS compilado (usado pela página)
├── scss/
│   └── style.scss              # Fonte SASS (edite aqui e recompile)
├── js/
│   └── app.js                  # Leitura da planilha, estatísticas e gráficos
├── assets/
│   ├── logo-horizontal.png
│   └── logo-vertical.png
├── .nojekyll                   # Evita processamento Jekyll no GitHub Pages
└── README.md
```

Dependências carregadas por CDN (exigem internet ao abrir a página): Chart.js 4.4.3, SheetJS 0.18.5, Leaflet 1.9.4 (com mapas-base CARTO/OpenStreetMap), Font Awesome 6.5.2 e Google Fonts (Fira Sans / Fira Mono).

## Como publicar no GitHub Pages

1. Crie um repositório novo no GitHub (ex.: `painel-alunos-uniceplac`), público.
2. Envie **todo o conteúdo desta pasta** para a raiz do repositório — o `index.html` precisa ficar na raiz. Pelo site do GitHub: *Add file → Upload files*, arraste os arquivos e pastas e confirme o commit. Por linha de comando:

   ```bash
   cd painel-alunos-uniceplac
   git init
   git add .
   git commit -m "Painel de alunos UNICEPLAC 2026/1"
   git branch -M main
   git remote add origin https://github.com/SEU_USUARIO/painel-alunos-uniceplac.git
   git push -u origin main
   ```

3. No repositório, abra **Settings → Pages**. Em *Build and deployment*, escolha **Deploy from a branch**, branch `main`, pasta `/ (root)`, e salve.
4. Aguarde 1–2 minutos. A página ficará disponível em:
   `https://SEU_USUARIO.github.io/painel-alunos-uniceplac/`

O arquivo `.nojekyll` já está incluído para o GitHub servir os arquivos exatamente como estão.

## Como testar localmente

O navegador bloqueia `fetch()` de arquivos abertos direto do disco (`file://`), então **abrir o index.html com duplo clique não carrega os dados** — o próprio painel exibe um aviso explicando isso. Sirva a pasta por um servidor local:

```bash
cd painel-alunos-uniceplac
python -m http.server 8080
```

e acesse <http://localhost:8080>. (Alternativas: `npx serve .` ou a extensão *Live Server* do VS Code.)

## Como atualizar os dados

1. Gere a nova planilha **com o mesmo nome** (`dados.xlsx`) e as mesmas colunas: `NOME_CURSO`, `RA`, `NOME_ALUNO`, `SITUACAO`, `DTNASCIMENTO`, `SEXO`, `ESTADO_CIVIL`, `CIDADE`, `BAIRRO`, `ESTADO`.
2. Substitua o arquivo na raiz do repositório e faça o commit. O painel recalcula tudo sozinho no próximo carregamento.

Para usar outro nome de arquivo, ajuste a constante `ARQUIVO_DADOS` no início de `js/app.js`.

## Como alterar o visual (SASS)

O CSS entregue já está compilado. Para mexer em cores, espaçamentos ou componentes, edite `scss/style.scss` (os tokens da marca estão no topo do arquivo) e recompile:

```bash
npm install -g sass
sass scss/style.scss css/style.css --style=compressed
```

As cores institucionais usadas — verde `#02744F`, verde-claro `#A6D3C1`, laranja `#F07F3D` e roxo `#41276F` — foram extraídas do manual da marca da UNICEPLAC. A tipografia FF Meta do manual foi substituída pela Fira Sans (fonte livre do mesmo designer, Erik Spiekermann).

## Créditos

- Identidade visual: [Manual da marca UNICEPLAC](https://www.uniceplac.edu.br/manual-da-marca/).
- Estrutura e componentes adaptados do padrão de painéis gov.br (Design System do Governo Federal), a partir de um dashboard da Codevasf.
- Gráficos: [Chart.js](https://www.chartjs.org/) · Mapa: [Leaflet](https://leafletjs.com/) com tiles © [OpenStreetMap](https://www.openstreetmap.org/copyright) © [CARTO](https://carto.com/attributions) · Leitura de planilha: [SheetJS CE](https://sheetjs.com/) · Ícones: [Font Awesome](https://fontawesome.com/) · Radar comparativo inspirado nas demos do [amCharts](https://www.amcharts.com/).

Uso educacional. Os dados dos alunos pertencem à UNICEPLAC — avalie a necessidade de anonimização antes de publicar o repositório como público.
