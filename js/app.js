/* ════════════════════════════════════════════════════════════════════
   Painel de Alunos · UNICEPLAC — 2026/1
   ─────────────────────────────────────────────────────────────────────
   Lê `dados_alunos_2026_1.xlsx` (mesma pasta do index.html) com SheetJS,
   calcula estatísticas descritivas e desenha os gráficos com Chart.js.
   Estrutura adaptada do painel Codevasf/gov.br para o padrão UNICEPLAC.

   Colunas esperadas na planilha (aba única):
   NOME_CURSO · RA · NOME_ALUNO · SITUACAO · DTNASCIMENTO · SEXO ·
   ESTADO_CIVIL · CIDADE · BAIRRO · ESTADO
   Observações tratadas aqui:
   – NOME_CURSO vem preenchido só na 1ª linha de cada curso (células
     mescladas) → é propagado para baixo ("forward fill");
   – DTNASCIMENTO pode chegar como data ou como número serial do Excel;
   – nomes de curso/cidade têm grafias duplicadas (com/sem acento,
     caixa alta) → são unificados por chave sem acentos.
   ════════════════════════════════════════════════════════════════════ */
'use strict';

/* ── Configuração ─────────────────────────────────────────────────── */
const ARQUIVO_DADOS = 'dados.xlsx';
const DATA_REFERENCIA = new Date(); // data-base do cálculo de idade

const FAIXAS = [
  { rot: 'até 17', min: -Infinity, max: 17 },
  { rot: '18–20', min: 18, max: 20 },
  { rot: '21–24', min: 21, max: 24 },
  { rot: '25–29', min: 25, max: 29 },
  { rot: '30–34', min: 30, max: 34 },
  { rot: '35–39', min: 35, max: 39 },
  { rot: '40–49', min: 40, max: 49 },
  { rot: '50 ou +', min: 50, max: Infinity }
];

const GRUPOS_SITUACAO = [
  { rot: 'Cursando',      teste: s => s === 'CURSANDO',            cls: 't-cursando' },
  { rot: 'Trancado',      teste: s => s === 'TRANCADO',            cls: 't-trancado' },
  { rot: 'Cancelado',     teste: s => s === 'CANCELADO',           cls: 't-cancelado' },
  { rot: 'Transferência', teste: s => s.startsWith('TRANSFER'),    cls: 't-transf' },
  { rot: 'Pré-matrícula', teste: s => s.startsWith('PRÉ') || s.startsWith('PRE'), cls: 't-outros' },
  { rot: 'Outros',        teste: () => true,                       cls: 't-outros' }
];

/* ── Utilidades ───────────────────────────────────────────────────── */
const $ = sel => document.querySelector(sel);
const nf0 = new Intl.NumberFormat('pt-BR');
const nf1 = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const fmtPct = (n, d) => (d ? nf1.format((n / d) * 100) + '%' : '—');

const semAcento = s => String(s ?? '')
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toUpperCase();

const MINUSCULAS = new Set(['de', 'da', 'do', 'das', 'dos', 'e', 'em', 'a', 'o']);
function titulo(s) {
  return String(s ?? '').toLocaleLowerCase('pt-BR').split(/\s+/).filter(Boolean)
    .map((w, i) => {
      if (i > 0 && MINUSCULAS.has(w)) return w;
      return w.charAt(0).toLocaleUpperCase('pt-BR') + w.slice(1);
    }).join(' ');
}

/** Rótulo curto e legível do curso (para gráficos e tabelas). */
function rotuloCurso(nomeOriginal) {
  let s = String(nomeOriginal).replace(/\s+/g, ' ').trim();
  let tec = false;
  s = s.replace(/^CURSO SUPERIOR (DE )?TECNOLOGIA EM\s+/i, () => { tec = true; return ''; });
  s = s.replace(/^CURSO DE\s+/i, '');
  // preserva o complemento entre parênteses — ex.: (Licenciatura)
  let compl = '';
  s = s.replace(/\s*\(([^)]+)\)\s*$/, (_, c) => { compl = ' (' + titulo(c) + ')'; return ''; });
  let base = titulo(s)
    .replace(/\bDesenvolvimento\b/i, 'Desenv.')
    .replace(/\bInteligência\b/i, 'Intelig.');
  return (tec ? 'Tec. ' : '') + base + compl;
}

/** Converte DTNASCIMENTO (Date do SheetJS ou serial do Excel) em Date. */
function paraData(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && isFinite(v) && v > 59) {
    return new Date(Math.round((v - 25569) * 86400 * 1000)); // época Excel→Unix
  }
  return null;
}

/** Idade em anos completos na data de referência. */
function idadeEm(nasc, ref) {
  if (!nasc) return null;
  let a = ref.getFullYear() - nasc.getFullYear();
  const m = ref.getMonth() - nasc.getMonth();
  if (m < 0 || (m === 0 && ref.getDate() < nasc.getDate())) a--;
  return (a >= 0 && a <= 110) ? a : null;
}

/* ── Estatística descritiva ───────────────────────────────────────── */
function quantil(ordenado, q) {
  if (!ordenado.length) return null;
  const pos = (ordenado.length - 1) * q;
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  return ordenado[lo] + (ordenado[hi] - ordenado[lo]) * (pos - lo);
}

/** Média, mediana, moda, desvio-padrão (amostral), quartis, mín e máx. */
function descreve(valores) {
  const v = valores.filter(x => x != null).slice().sort((a, b) => a - b);
  const n = v.length;
  if (!n) return { n: 0, media: null, mediana: null, moda: null, dp: null, min: null, q1: null, q3: null, max: null };
  const soma = v.reduce((s, x) => s + x, 0);
  const media = soma / n;
  const dp = n > 1 ? Math.sqrt(v.reduce((s, x) => s + (x - media) ** 2, 0) / (n - 1)) : 0;
  const freq = new Map();
  let moda = v[0], fMax = 0;
  for (const x of v) {
    const f = (freq.get(x) || 0) + 1;
    freq.set(x, f);
    if (f > fMax) { fMax = f; moda = x; }
  }
  return {
    n, media, dp, moda,
    mediana: quantil(v, .5), q1: quantil(v, .25), q3: quantil(v, .75),
    min: v[0], max: v[n - 1]
  };
}

/* ── Leitura e preparação da planilha ─────────────────────────────── */
function prepararRegistros(linhas) {
  const regs = [];
  const cursoNome = new Map();   // chave sem acento → contagem por grafia
  const cidadeNome = new Map();
  let cursoAtual = null;

  for (const l of linhas) {
    const bruto = l.NOME_CURSO != null ? String(l.NOME_CURSO).trim() : '';
    if (bruto) cursoAtual = bruto;                       // forward fill
    const temDado = ['RA', 'NOME_ALUNO', 'SITUACAO', 'DTNASCIMENTO', 'SEXO', 'CIDADE']
      .some(c => l[c] != null && String(l[c]).trim() !== '');
    if (!cursoAtual || !temDado) continue;               // ignora linhas vazias

    const nasc = paraData(l.DTNASCIMENTO);
    const cidadeBruta = l.CIDADE != null ? String(l.CIDADE).replace(/\s+/g, ' ').trim() : '';
    const r = {
      cursoKey: semAcento(cursoAtual),
      cursoBruto: cursoAtual,
      ra: l.RA != null ? String(l.RA).trim() : '',
      nome: l.NOME_ALUNO != null ? String(l.NOME_ALUNO).trim() : '',
      situacao: l.SITUACAO != null ? String(l.SITUACAO).trim().toUpperCase() : '',
      idade: idadeEm(nasc, DATA_REFERENCIA),
      sexo: l.SEXO != null ? String(l.SEXO).trim().toUpperCase() : '',
      estadoCivil: l.ESTADO_CIVIL != null ? titulo(String(l.ESTADO_CIVIL)) : '',
      cidadeKey: semAcento(cidadeBruta),
      cidadeBruta,
      bairro: l.BAIRRO != null ? titulo(String(l.BAIRRO)) : '',
      uf: l.ESTADO != null ? String(l.ESTADO).trim().toUpperCase() : ''
    };
    regs.push(r);

    if (!cursoNome.has(r.cursoKey)) cursoNome.set(r.cursoKey, new Map());
    const cn = cursoNome.get(r.cursoKey);
    cn.set(cursoAtual, (cn.get(cursoAtual) || 0) + 1);
    if (cidadeBruta) {
      if (!cidadeNome.has(r.cidadeKey)) cidadeNome.set(r.cidadeKey, new Map());
      const dn = cidadeNome.get(r.cidadeKey);
      dn.set(cidadeBruta, (dn.get(cidadeBruta) || 0) + 1);
    }
  }

  const maisFrequente = m => [...m.entries()].sort((a, b) => b[1] - a[1])[0][0];
  const rotCurso = new Map([...cursoNome].map(([k, m]) => [k, rotuloCurso(maisFrequente(m))]));
  const rotCidade = new Map([...cidadeNome].map(([k, m]) => [k, titulo(maisFrequente(m))]));
  for (const r of regs) {
    r.curso = rotCurso.get(r.cursoKey);
    r.cidade = r.cidadeKey ? rotCidade.get(r.cidadeKey) : '';
    const g = GRUPOS_SITUACAO.find(g => g.teste(r.situacao)) || GRUPOS_SITUACAO.at(-1);
    r.grupoSituacao = g.rot;
    r.grupoCls = g.cls;
  }
  return regs;
}

/* ── Estado global ────────────────────────────────────────────────── */
const estado = {
  todos: [],
  filtro: { curso: '', situacao: '', sexo: '', uf: '' },
  busca: '',
  pagina: 1,
  porPagina: 12,
  ordem: { col: 'media', dir: 'desc' },
  radarCurso: '',
  charts: {},
  mapa: null,
  mapaCamada: null,
  mapaBounds: null
};

const filtrados = () => estado.todos.filter(r =>
  (!estado.filtro.curso || r.cursoKey === estado.filtro.curso) &&
  (!estado.filtro.situacao || r.grupoSituacao === estado.filtro.situacao) &&
  (!estado.filtro.sexo || r.sexo === estado.filtro.sexo) &&
  (!estado.filtro.uf || r.uf === estado.filtro.uf));

function contar(regs, chave) {
  const m = new Map();
  for (const r of regs) {
    const k = typeof chave === 'function' ? chave(r) : r[chave];
    if (k == null || k === '') continue;
    m.set(k, (m.get(k) || 0) + 1);
  }
  return [...m.entries()].sort((a, b) => b[1] - a[1]);
}

/* ── Cores e padrões do Chart.js ──────────────────────────────────── */
let COR = {};
function lerCores() {
  const css = getComputedStyle(document.documentElement);
  const v = n => css.getPropertyValue(n).trim();
  COR = {
    verde: v('--un-green'), verdeEscuro: v('--un-green-dark'), verdeProfundo: v('--un-green-deep'),
    mint: v('--un-mint'), laranja: v('--un-orange'), laranjaEscuro: v('--un-orange-dark'),
    roxo: v('--un-purple'), ink: v('--ink'), muted: v('--muted'), borda: v('--border')
  };
  COR.paleta = [COR.verde, COR.laranja, COR.roxo, '#63A98C', COR.laranjaEscuro, '#7E6BA8', COR.mint, '#0E8A63'];
}

function configurarChartJS() {
  Chart.defaults.font.family = "'Fira Sans','Segoe UI',Arial,sans-serif";
  Chart.defaults.font.size = 12;
  Chart.defaults.color = COR.muted;
  Chart.defaults.borderColor = COR.borda;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = COR.verdeProfundo;
  Chart.defaults.plugins.tooltip.titleFont = { weight: '700' };
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 8;
}

function desenhar(id, cfg) {
  if (estado.charts[id]) estado.charts[id].destroy();
  estado.charts[id] = new Chart(document.getElementById(id), cfg);
}

/* Linha vertical tracejada na média geral (gráfico de idade por curso). */
const pluginLinhaMedia = {
  id: 'linhaMedia',
  afterDatasetsDraw(chart, _args, opts) {
    if (opts.valor == null) return;
    const { ctx, chartArea, scales } = chart;
    const x = scales.x.getPixelForValue(opts.valor);
    if (x < chartArea.left || x > chartArea.right) return;
    ctx.save();
    ctx.strokeStyle = COR.laranja;
    ctx.setLineDash([6, 5]);
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(x, chartArea.top);
    ctx.lineTo(x, chartArea.bottom);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = COR.laranjaEscuro;
    ctx.font = "700 11px 'Fira Sans'";
    ctx.textAlign = 'center';
    ctx.fillText('média geral: ' + nf1.format(opts.valor), x, chartArea.top - 6);
    ctx.restore();
  }
};


/* ── Geolocalização: bairros do DF e municípios ───────────────────── */
/* Coordenadas aproximadas (centro do bairro/RA ou do município), suficientes
   para a visualização espacial. Posições marcadas como "aprox" usam a capital
   da UF quando o município não está no dicionário. */
const normGeo = s => semAcento(s).replace(/[^A-Z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();

const GEO_PONTOS = {
  'Gama': [-16.0205, -48.0646],
  'Setor Central (Gama)': [-16.0170, -48.0615],
  'Setor Leste (Gama)': [-16.0125, -48.0500],
  'Setor Sul (Gama)': [-16.0295, -48.0570],
  'Setor Oeste (Gama)': [-16.0230, -48.0768],
  'Setor Norte (Gama)': [-16.0062, -48.0610],
  'Setor de Indústria (Gama)': [-16.0030, -48.0478],
  'Ponte Alta Norte (Gama)': [-15.9830, -48.1240],
  'Ponte Alta (Gama)': [-16.0420, -48.1330],
  'Santa Maria': [-16.0125, -47.9880],
  'Santa Maria Norte': [-16.0005, -47.9845],
  'Santa Maria Sul': [-16.0265, -47.9910],
  'Setor Meireles (Santa Maria)': [-16.0330, -47.9745],
  'Plano Piloto': [-15.7942, -47.8825],
  'Asa Sul': [-15.8330, -47.9135],
  'Asa Norte': [-15.7630, -47.8770],
  'Sudoeste/Octogonal': [-15.7950, -47.9260],
  'Noroeste': [-15.7400, -47.9100],
  'Vila Planalto': [-15.7880, -47.8560],
  'Granja do Torto': [-15.7100, -47.9300],
  'Cruzeiro': [-15.7905, -47.9370],
  'Lago Sul': [-15.8480, -47.8740],
  'Lago Norte': [-15.7380, -47.8600],
  'Varjão': [-15.7100, -47.8770],
  'Taguatinga': [-15.8386, -48.0553],
  'Taguatinga Norte': [-15.8160, -48.0570],
  'Taguatinga Sul': [-15.8520, -48.0580],
  'Ceilândia': [-15.8195, -48.1073],
  'Ceilândia Norte': [-15.8030, -48.1080],
  'Ceilândia Sul': [-15.8330, -48.1000],
  'Sol Nascente/Pôr do Sol': [-15.8150, -48.1350],
  'Samambaia': [-15.8763, -48.0847],
  'Samambaia Norte': [-15.8680, -48.0820],
  'Samambaia Sul': [-15.8850, -48.0880],
  'Águas Claras': [-15.8340, -48.0260],
  'Arniqueira': [-15.8560, -48.0080],
  'Vicente Pires': [-15.8060, -48.0200],
  'Guará': [-15.8266, -47.9770],
  'Guará I': [-15.8180, -47.9700],
  'Guará II': [-15.8355, -47.9787],
  'SIA': [-15.8020, -47.9540],
  'SCIA/Estrutural': [-15.7840, -48.0000],
  'Núcleo Bandeirante': [-15.8697, -47.9668],
  'Candangolândia': [-15.8547, -47.9530],
  'Park Way': [-15.9006, -47.9631],
  'Riacho Fundo': [-15.8820, -48.0170],
  'Riacho Fundo II': [-15.8990, -48.0435],
  'Recanto das Emas': [-15.9057, -48.0625],
  'Jardim Botânico': [-15.8730, -47.8100],
  'São Sebastião': [-15.9020, -47.7800],
  'Paranoá': [-15.7754, -47.7776],
  'Itapoã': [-15.7550, -47.7680],
  'Sobradinho': [-15.6520, -47.7900],
  'Sobradinho II': [-15.6525, -47.8250],
  'Fercal': [-15.5900, -47.8700],
  'Planaltina': [-15.6200, -47.6550],
  'Brazlândia': [-15.6770, -48.2030],
  'Novo Gama (GO)': [-16.0580, -48.0380],
  'Valparaíso de Goiás (GO)': [-16.0651, -47.9757],
  'Cidade Ocidental (GO)': [-16.1064, -47.9264],
  'Luziânia (GO)': [-16.2525, -47.9502],
  'Águas Lindas de Goiás (GO)': [-15.7617, -48.2816],
  'Santo Antônio do Descoberto (GO)': [-15.9339, -48.2582]
};

const GEO_CIDADES = {
  'VALPARAISO DE GOIAS': ['Valparaíso de Goiás (GO)', -16.0651, -47.9757],
  'NOVO GAMA': ['Novo Gama (GO)', -16.0580, -48.0380],
  'LUZIANIA': ['Luziânia (GO)', -16.2525, -47.9502],
  'CIDADE OCIDENTAL': ['Cidade Ocidental (GO)', -16.1064, -47.9264],
  'AGUAS LINDAS DE GOIAS': ['Águas Lindas de Goiás (GO)', -15.7617, -48.2816],
  'SANTO ANTONIO DO DESCOBERTO': ['Santo Antônio do Descoberto (GO)', -15.9339, -48.2582],
  'GOIANIA': ['Goiânia (GO)', -16.6864, -49.2643],
  'ANAPOLIS': ['Anápolis (GO)', -16.3281, -48.9530],
  'APARECIDA DE GOIANIA': ['Aparecida de Goiânia (GO)', -16.8220, -49.2470],
  'GOIAS': ['Cidade de Goiás (GO)', -15.9337, -50.1400],
  'CRISTALINA': ['Cristalina (GO)', -16.7688, -47.6132],
  'FORMOSA': ['Formosa (GO)', -15.5372, -47.3340],
  'PLANALTINA': ['Planaltina (GO)', -15.4529, -47.6142],
  'ALEXANIA': ['Alexânia (GO)', -16.0837, -48.5076],
  'POSSE': ['Posse (GO)', -14.0930, -46.3690],
  'CALDAS NOVAS': ['Caldas Novas (GO)', -17.7440, -48.6250],
  'CAMPO LIMPO DE GOIAS': ['Campo Limpo de Goiás (GO)', -16.2960, -49.0920],
  'GOIANESIA': ['Goianésia (GO)', -15.3170, -49.1170],
  'CERES': ['Ceres (GO)', -15.3080, -49.6010],
  'COCALZINHO DE GOIAS': ['Cocalzinho de Goiás (GO)', -15.7940, -48.7770],
  'PADRE BERNARDO': ['Padre Bernardo (GO)', -15.1600, -48.2830],
  'LUIS EDUARDO MAGALHAES': ['Luís Eduardo Magalhães (BA)', -12.0920, -45.8000],
  'BARREIRAS': ['Barreiras (BA)', -12.1530, -44.9900],
  'CORRENTINA': ['Correntina (BA)', -13.3430, -44.6370],
  'COCOS': ['Cocos (BA)', -14.1800, -44.5330],
  'UNAI': ['Unaí (MG)', -16.3570, -46.9060],
  'BURITIS': ['Buritis (MG)', -15.6180, -46.4230],
  'ARINOS': ['Arinos (MG)', -15.9170, -46.1060],
  'UBERLANDIA': ['Uberlândia (MG)', -18.9186, -48.2772],
  'JOAO PINHEIRO': ['João Pinheiro (MG)', -17.7400, -46.1720],
  'FORTALEZA': ['Fortaleza (CE)', -3.7319, -38.5267],
  'JUAZEIRO DO NORTE': ['Juazeiro do Norte (CE)', -7.2130, -39.3150],
  'CORRENTE': ['Corrente (PI)', -10.4430, -45.1620],
  'TUCUMA': ['Tucumã (PA)', -6.7480, -51.1620],
  'CONFRESA': ['Confresa (MT)', -10.6440, -51.5700],
  'CAMPO GRANDE': ['Campo Grande (MS)', -20.4697, -54.6201],
  'CAMPINAS': ['Campinas (SP)', -22.9099, -47.0626],
  'FLORIANOPOLIS': ['Florianópolis (SC)', -27.5954, -48.5480]
};

const GEO_CAPITAIS = {
  AC: [-9.9750, -67.8240], AL: [-9.6660, -35.7350], AP: [0.0349, -51.0694], AM: [-3.1190, -60.0217],
  BA: [-12.9714, -38.5014], CE: [-3.7319, -38.5267], DF: [-15.7942, -47.8825], ES: [-20.3155, -40.3128],
  GO: [-16.6864, -49.2643], MA: [-2.5307, -44.3068], MT: [-15.6010, -56.0974], MS: [-20.4697, -54.6201],
  MG: [-19.9167, -43.9345], PA: [-1.4558, -48.4902], PB: [-7.1195, -34.8450], PR: [-25.4284, -49.2733],
  PE: [-8.0476, -34.8770], PI: [-5.0892, -42.8016], RJ: [-22.9068, -43.1729], RN: [-5.7945, -35.2110],
  RS: [-30.0346, -51.2177], RO: [-8.7612, -63.9004], RR: [2.8235, -60.6758], SC: [-27.5954, -48.5480],
  SP: [-23.5505, -46.6333], SE: [-10.9472, -37.0731], TO: [-10.1840, -48.3336]
};

/* Converte o texto livre do campo BAIRRO em um lugar canônico do DF/Entorno. */
function classificaBairro(bruto) {
  const t = ' ' + normGeo(bruto) + ' ';
  const tem = (...ws) => ws.some(w => t.includes(' ' + w + ' ') || t.includes(w));
  const palavra = w => new RegExp('\\b' + w + '\\b').test(t);

  if (tem('MEIRELES')) return 'Setor Meireles (Santa Maria)';
  if (tem('PONTE ALTA NORTE', 'P ALTA NORTE')) return 'Ponte Alta Norte (Gama)';
  if (tem('PONTE ALTA', 'POBTE ALTA', 'POLO JK', 'PONTE DE TERRA') || (tem('P ALTA') && palavra('GAMA'))) return 'Ponte Alta (Gama)';
  if (tem('SANTA MARIA', 'SANTA MARA', 'SENTA MARIA', 'TOTAL VILLE')) {
    if (tem('NORTE')) return 'Santa Maria Norte';
    if (tem('SUL')) return 'Santa Maria Sul';
    return 'Santa Maria';
  }
  if (tem('NOVO GAMA', 'PEDREGAL', 'LAGO AZUL')) return 'Novo Gama (GO)';
  if (tem('VALPARAISO', 'CEU AZUL', 'PARQUE ESPLANADA')) return 'Valparaíso de Goiás (GO)';
  if (tem('OCIDENTAL', 'JARDIM ABC')) return 'Cidade Ocidental (GO)';
  if (tem('LUZIANIA', 'JARDIM INGA')) return 'Luziânia (GO)';
  if (tem('AGUAS LINDAS')) return 'Águas Lindas de Goiás (GO)';
  if (palavra('GAMA') || /\bSETOR (LESTE|OESTE|SUL|NORTE|CENTRAL)\b/.test(t) || tem('SETOR DE INDUSTRIA', 'SETOR INDUSTRIAL')) {
    if (tem('LESTE')) return 'Setor Leste (Gama)';
    if (tem('OESTE')) return 'Setor Oeste (Gama)';
    if (tem('INDUSTRIA')) return 'Setor de Indústria (Gama)';
    if (tem('CENTRAL')) return 'Setor Central (Gama)';
    if (tem('NORTE')) return 'Setor Norte (Gama)';
    if (tem('SUL')) return 'Setor Sul (Gama)';
    return 'Gama';
  }
  if (tem('TAGUATINGA')) {
    if (tem('NORTE')) return 'Taguatinga Norte';
    if (tem('SUL')) return 'Taguatinga Sul';
    return 'Taguatinga';
  }
  if (tem('SOL NASCENTE', 'POR DO SOL')) return 'Sol Nascente/Pôr do Sol';
  if (tem('CEILANDIA') || /\bP (SUL|NORTE)\b/.test(t) || /\bQN[MNOPQR]\b/.test(t)) {
    if (tem('NORTE')) return 'Ceilândia Norte';
    if (tem('SUL')) return 'Ceilândia Sul';
    return 'Ceilândia';
  }
  if (tem('SAMAMBAIA')) {
    if (tem('NORTE')) return 'Samambaia Norte';
    if (tem('SUL')) return 'Samambaia Sul';
    return 'Samambaia';
  }
  if (tem('AGUAS CLARAS', 'ADE AGUAS')) return 'Águas Claras';
  if (tem('ARNIQUEIRA')) return 'Arniqueira';
  if (tem('RECANTO DAS EMAS', 'RECONTO DAS EMAS')) return 'Recanto das Emas';
  if (tem('RIACHO FUNDO')) return (palavra('II') || palavra('2')) ? 'Riacho Fundo II' : 'Riacho Fundo';
  if (tem('GUARA')) {
    if (palavra('II') || palavra('2')) return 'Guará II';
    if (palavra('I') || palavra('1')) return 'Guará I';
    return 'Guará';
  }
  if (tem('BANDEIRANTE')) return 'Núcleo Bandeirante';
  if (tem('CANDANGOLANDIA')) return 'Candangolândia';
  if (tem('PARK WAY', 'PARKWAY')) return 'Park Way';
  if (tem('LAGO SUL')) return 'Lago Sul';
  if (tem('LAGO NORTE')) return 'Lago Norte';
  if (tem('JARDIM BOTANICO', 'MANGUEIRAL')) return 'Jardim Botânico';
  if (tem('SAO SEBASTIAO')) return 'São Sebastião';
  if (tem('SOBRADINHO')) return (palavra('II') || palavra('2')) ? 'Sobradinho II' : 'Sobradinho';
  if (tem('PLANALTINA')) return 'Planaltina';
  if (tem('PARANOA', 'PARONOA')) return 'Paranoá';
  if (tem('ITAPOA')) return 'Itapoã';
  if (tem('VARJAO')) return 'Varjão';
  if (tem('CRUZEIRO')) return 'Cruzeiro';
  if (tem('SUDOESTE', 'OCTOGONAL')) return 'Sudoeste/Octogonal';
  if (tem('VICENTE PIRES')) return 'Vicente Pires';
  if (tem('ESTRUTURAL', 'SCIA')) return 'SCIA/Estrutural';
  if (palavra('SIA')) return 'SIA';
  if (tem('BRAZLANDIA')) return 'Brazlândia';
  if (tem('FERCAL')) return 'Fercal';
  if (tem('ASA SUL') || palavra('SQS') || palavra('SHIS') || palavra('SGAS') || palavra('CLS')) return 'Asa Sul';
  if (tem('ASA NORTE') || palavra('SQN') || palavra('SGAN') || palavra('CLN')) return 'Asa Norte';
  if (tem('VILA PLANALTO')) return 'Vila Planalto';
  if (tem('NOROESTE')) return 'Noroeste';
  if (tem('GRANJA DO TORTO') || palavra('TORTO')) return 'Granja do Torto';
  if (tem('BRASILIA', 'PLANO PILOTO')) return 'Plano Piloto';
  return null;
}

/* Ponto no mapa para um registro: bairro (DF) ou município (demais UFs). */
function pontoDoRegistro(r) {
  if (r.uf === 'DF') {
    const canon = classificaBairro(r.bairro);
    if (canon && GEO_PONTOS[canon]) return { chave: canon, coords: GEO_PONTOS[canon], aprox: false };
    return null;
  }
  const cid = GEO_CIDADES[normGeo(r.cidade)];
  if (cid) return { chave: cid[0], coords: [cid[1], cid[2]], aprox: false };
  const cap = GEO_CAPITAIS[r.uf];
  if (cap) return { chave: 'Outros municípios — ' + r.uf, coords: cap, aprox: true };
  return { exterior: true };
}

const CALOR = ['#FBE4D3', '#F6BC90', '#F08F4E', '#D96524', '#A93E0F'];
const corCalor = f => CALOR[Math.max(0, Math.min(CALOR.length - 1, Math.floor(f * CALOR.length)))];
const VISTA_DF = { centro: [-15.95, -48.00], zoom: 10 };

function vistaMapa(qual) {
  if (!estado.mapa) return;
  $('#btnMapaDF').classList.toggle('active', qual === 'df');
  $('#btnMapaBR').classList.toggle('active', qual === 'br');
  if (qual === 'br' && estado.mapaBounds) estado.mapa.fitBounds(estado.mapaBounds.pad(0.18));
  else estado.mapa.setView(VISTA_DF.centro, VISTA_DF.zoom);
}

function renderMapa(regs) {
  const nota = $('#mapaNota');
  if (typeof L === 'undefined') {
    nota.innerHTML = '<i class="fas fa-triangle-exclamation" aria-hidden="true"></i>A biblioteca de mapas (Leaflet) não pôde ser carregada — verifique a conexão com a internet.';
    return;
  }
  if (!estado.mapa) {
    estado.mapa = L.map('mapaAlunos', { scrollWheelZoom: false });
    L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener">OpenStreetMap</a> &copy; <a href="https://carto.com/attributions" target="_blank" rel="noopener">CARTO</a>',
      subdomains: 'abcd', maxZoom: 19
    }).addTo(estado.mapa);
    estado.mapaCamada = L.layerGroup().addTo(estado.mapa);
    estado.mapa.setView(VISTA_DF.centro, VISTA_DF.zoom);
    $('#btnMapaDF').addEventListener('click', () => vistaMapa('df'));
    $('#btnMapaBR').addEventListener('click', () => vistaMapa('br'));
    setTimeout(() => estado.mapa.invalidateSize(), 250);
  }

  const ag = new Map();
  let semLocal = 0, exterior = 0;
  for (const r of regs) {
    const p = pontoDoRegistro(r);
    if (!p) { semLocal++; continue; }
    if (p.exterior) { exterior++; continue; }
    const atual = ag.get(p.chave) || { coords: p.coords, n: 0, aprox: p.aprox };
    atual.n++;
    ag.set(p.chave, atual);
  }

  estado.mapaCamada.clearLayers();
  const itens = [...ag.entries()].sort((a, b) => b[1].n - a[1].n); // maiores primeiro: círculos pequenos ficam por cima
  const nmax = itens.length ? itens[0][1].n : 0;
  const pontos = [];
  for (const [chave, d] of itens) {
    const f = nmax > 0 ? Math.log(1 + d.n) / Math.log(1 + nmax) : 0;
    pontos.push(d.coords);
    L.circleMarker(d.coords, {
      radius: 6 + 22 * Math.sqrt(nmax ? d.n / nmax : 0),
      color: '#ffffff', weight: 1.5,
      fillColor: corCalor(f), fillOpacity: 0.78
    }).bindTooltip(
      '<b>' + chave + '</b><br>' + nf0.format(d.n) + ' aluno(s) · ' + fmtPct(d.n, regs.length) + ' do filtro' +
      (d.aprox ? '<br><i>posição aproximada (capital da UF)</i>' : ''),
      { className: 'map-tip', direction: 'top', sticky: true }
    ).addTo(estado.mapaCamada);
  }
  estado.mapaBounds = pontos.length ? L.latLngBounds(pontos) : null;

  const georef = regs.length - semLocal - exterior;
  const partes = ['<i class="fas fa-circle-info" aria-hidden="true"></i>' +
    nf0.format(georef) + ' de ' + nf0.format(regs.length) +
    ' aluno(s) do filtro posicionados pelo centro aproximado do bairro (DF) ou do município (demais UFs).'];
  if (semLocal) partes.push(nf0.format(semLocal) + ' registro(s) com bairro não identificável ficam fora do mapa.');
  if (exterior) partes.push(nf0.format(exterior) + ' aluno(s) residem no exterior e não aparecem no mapa.');
  nota.innerHTML = partes.join(' ');
}

/* ── Blocos de renderização ───────────────────────────────────────── */
function renderHero() {
  const t = estado.todos;
  $('#hsAlunos').textContent = nf0.format(t.length);
  $('#hsCursos').textContent = new Set(t.map(r => r.cursoKey)).size;
  const d = descreve(t.map(r => r.idade));
  $('#hsIdade').textContent = d.media != null ? nf1.format(d.media) + ' anos' : '—';
  $('#hsMunicipios').textContent = new Set(t.filter(r => r.cidadeKey).map(r => r.cidadeKey)).size;
  $('#heroData').textContent = DATA_REFERENCIA.toLocaleDateString('pt-BR');
}

function renderKPIs(regs) {
  const n = regs.length;
  $('#fcAtual').textContent = nf0.format(n);
  $('#fcTotal').textContent = nf0.format(estado.todos.length);
  $('#kpiTotal').textContent = nf0.format(n);

  const cursando = regs.filter(r => r.grupoSituacao === 'Cursando').length;
  $('#kpiCursando').textContent = nf0.format(cursando);
  $('#kpiCursandoPct').textContent = fmtPct(cursando, n) + ' do filtro';

  const d = descreve(regs.map(r => r.idade));
  $('#kpiIdade').textContent = d.media != null ? nf1.format(d.media) : '—';
  $('#kpiIdadeMed').textContent = d.mediana != null
    ? 'mediana: ' + nf1.format(d.mediana) + ' · DP: ' + nf1.format(d.dp) : 'mediana: —';

  const fem = regs.filter(r => r.sexo === 'FEMININO').length;
  $('#kpiFem').textContent = fmtPct(fem, n);
  $('#kpiFemN').textContent = nf0.format(fem) + ' alunas';

  const df = regs.filter(r => r.uf === 'DF').length;
  $('#kpiDF').textContent = fmtPct(df, n);
  $('#kpiDFN').textContent = nf0.format(df) + ' alunos no DF';
}

function statsPorCurso(regs) {
  const porCurso = new Map();
  for (const r of regs) {
    if (!porCurso.has(r.cursoKey)) porCurso.set(r.cursoKey, []);
    porCurso.get(r.cursoKey).push(r);
  }
  const linhas = [...porCurso.entries()].map(([key, rs]) => {
    const d = descreve(rs.map(r => r.idade));
    return {
      key, curso: rs[0].curso, total: rs.length, ...d,
      pctFem: rs.length ? rs.filter(r => r.sexo === 'FEMININO').length / rs.length * 100 : null,
      pctCursando: rs.length ? rs.filter(r => r.grupoSituacao === 'Cursando').length / rs.length * 100 : null,
      pctDF: rs.length ? rs.filter(r => r.uf === 'DF').length / rs.length * 100 : null,
      pctSolteiro: rs.length ? rs.filter(r => semAcento(r.estadoCivil) === 'SOLTEIRO').length / rs.length * 100 : null
    };
  });
  return linhas;
}

function renderGraficosCursos(regs, stats) {
  const porQtd = stats.slice().sort((a, b) => b.total - a.total);
  desenhar('cAlunosCurso', {
    type: 'bar',
    data: {
      labels: porQtd.map(s => s.curso),
      datasets: [{
        data: porQtd.map(s => s.total),
        backgroundColor: COR.verde, hoverBackgroundColor: COR.laranja,
        borderRadius: 5, barPercentage: .78
      }]
    },
    options: {
      indexAxis: 'y', maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => nf0.format(c.parsed.x) + ' aluno(s) · ' + fmtPct(c.parsed.x, regs.length) + ' do filtro' } }
      },
      scales: {
        x: { grid: { color: COR.borda }, ticks: { precision: 0 } },
        y: { grid: { display: false }, ticks: { autoSkip: false, font: { size: 11 } } }
      }
    }
  });

  const comIdade = stats.filter(s => s.n > 0).sort((a, b) => b.media - a.media);
  const mediaGeral = descreve(regs.map(r => r.idade)).media;
  desenhar('cIdadeCurso', {
    type: 'bar',
    data: {
      labels: comIdade.map(s => s.curso),
      datasets: [{
        data: comIdade.map(s => +s.media.toFixed(2)),
        backgroundColor: comIdade.map(s => s.media >= (mediaGeral ?? 0) ? COR.roxo : COR.verde),
        hoverBackgroundColor: COR.laranja, borderRadius: 5, barPercentage: .78
      }]
    },
    options: {
      indexAxis: 'y', maintainAspectRatio: false,
      layout: { padding: { top: 30 } },
      plugins: {
        legend: { display: false },
        linhaMedia: { valor: mediaGeral },
        tooltip: {
          callbacks: {
            label: c => 'Idade média: ' + nf1.format(c.parsed.x) + ' anos',
            afterBody: items => {
              const s = comIdade[items[0].dataIndex];
              return [
                'Mediana: ' + nf1.format(s.mediana) + ' · DP: ' + nf1.format(s.dp),
                'Mín–máx: ' + s.min + '–' + s.max + ' · N = ' + nf0.format(s.n)
              ];
            }
          }
        }
      },
      scales: {
        x: { grid: { color: COR.borda }, title: { display: true, text: 'anos' } },
        y: { grid: { display: false }, ticks: { autoSkip: false, font: { size: 11 } } }
      }
    },
    plugins: [pluginLinhaMedia]
  });
}

function renderGraficosPerfil(regs) {
  const idades = regs.map(r => r.idade).filter(i => i != null);
  const porFaixa = FAIXAS.map(f => idades.filter(i => i >= f.min && i <= f.max).length);
  desenhar('cFaixa', {
    type: 'bar',
    data: {
      labels: FAIXAS.map(f => f.rot),
      datasets: [{ data: porFaixa, backgroundColor: COR.verde, hoverBackgroundColor: COR.laranja, borderRadius: 5 }]
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => nf0.format(c.parsed.y) + ' aluno(s) · ' + fmtPct(c.parsed.y, idades.length) } }
      },
      scales: { x: { grid: { display: false } }, y: { grid: { color: COR.borda }, ticks: { precision: 0 } } }
    }
  });

  const sexo = contar(regs, 'sexo');
  desenhar('cSexo', {
    type: 'doughnut',
    data: {
      labels: sexo.map(([k]) => titulo(k)),
      datasets: [{ data: sexo.map(([, v]) => v), backgroundColor: [COR.verde, COR.laranja, COR.roxo], borderWidth: 2, borderColor: '#fff' }]
    },
    options: {
      maintainAspectRatio: false, cutout: '58%',
      plugins: {
        legend: { position: 'bottom' },
        tooltip: { callbacks: { label: c => ' ' + nf0.format(c.parsed) + ' · ' + fmtPct(c.parsed, regs.length) } }
      }
    }
  });

  const ORDEM_GRUPOS = GRUPOS_SITUACAO.map(g => g.rot);
  const sit = contar(regs, 'grupoSituacao')
    .sort((a, b) => ORDEM_GRUPOS.indexOf(a[0]) - ORDEM_GRUPOS.indexOf(b[0]));
  desenhar('cSituacao', {
    type: 'polarArea',
    data: {
      labels: sit.map(([k]) => k),
      datasets: [{
        data: sit.map(([, v]) => v),
        backgroundColor: [COR.verde, COR.laranja, '#8E2F3C', COR.roxo, COR.mint, '#9AA8A1'].map(c => c + 'D0'),
        borderColor: '#fff', borderWidth: 2
      }]
    },
    options: {
      maintainAspectRatio: false,
      scales: { r: { ticks: { display: false }, grid: { color: COR.borda } } },
      plugins: {
        legend: { position: 'bottom' },
        tooltip: { callbacks: { label: c => ' ' + nf0.format(c.parsed.r) + ' · ' + fmtPct(c.parsed.r, regs.length) } }
      }
    }
  });

  const ec = contar(regs, 'estadoCivil').slice(0, 6);
  desenhar('cEstadoCivil', {
    type: 'bar',
    data: {
      labels: ec.map(([k]) => k),
      datasets: [{ data: ec.map(([, v]) => v), backgroundColor: COR.paleta, borderRadius: 5 }]
    },
    options: {
      indexAxis: 'y', maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => nf0.format(c.parsed.x) + ' · ' + fmtPct(c.parsed.x, regs.length) } }
      },
      scales: { x: { grid: { color: COR.borda }, ticks: { precision: 0 } }, y: { grid: { display: false } } }
    }
  });

  const ufs = contar(regs, 'uf');
  const top = ufs.slice(0, 2);
  const resto = ufs.slice(2).reduce((s, [, v]) => s + v, 0);
  const dataUF = resto ? [...top, ['Outras UFs', resto]] : top;
  desenhar('cUF', {
    type: 'doughnut',
    data: {
      labels: dataUF.map(([k]) => k),
      datasets: [{ data: dataUF.map(([, v]) => v), backgroundColor: [COR.verde, COR.laranja, COR.roxo], borderWidth: 2, borderColor: '#fff' }]
    },
    options: {
      maintainAspectRatio: false, cutout: '58%',
      plugins: {
        legend: { position: 'bottom' },
        tooltip: { callbacks: { label: c => ' ' + nf0.format(c.parsed) + ' · ' + fmtPct(c.parsed, regs.length) } }
      }
    }
  });

  const cidades = contar(regs, 'cidade').slice(0, 10);
  desenhar('cCidades', {
    type: 'bar',
    data: {
      labels: cidades.map(([k]) => k),
      datasets: [{ data: cidades.map(([, v]) => v), backgroundColor: COR.verde, hoverBackgroundColor: COR.laranja, borderRadius: 5 }]
    },
    options: {
      indexAxis: 'y', maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => nf0.format(c.parsed.x) + ' · ' + fmtPct(c.parsed.x, regs.length) } }
      },
      scales: {
        x: { grid: { color: COR.borda }, ticks: { precision: 0 } },
        y: { grid: { display: false }, ticks: { font: { size: 11 } } }
      }
    }
  });
}

/* Radar: curso selecionado × média institucional (eixos normalizados). */
function renderRadar(regs, stats) {
  const sel = $('#radarCurso');
  const ordenados = stats.slice().sort((a, b) => a.curso.localeCompare(b.curso, 'pt-BR'));
  sel.innerHTML = ordenados.map(s => `<option value="${s.key}">${s.curso}</option>`).join('');
  if (!ordenados.some(s => s.key === estado.radarCurso)) {
    const maior = stats.slice().sort((a, b) => b.total - a.total)[0];
    estado.radarCurso = maior ? maior.key : '';
  }
  sel.value = estado.radarCurso;

  const alvo = stats.find(s => s.key === estado.radarCurso);
  if (!alvo) {
    if (estado.charts.cRadar) { estado.charts.cRadar.destroy(); delete estado.charts.cRadar; }
    return;
  }

  const medias = stats.filter(s => s.n > 0).map(s => s.media);
  const minM = Math.min(...medias), maxM = Math.max(...medias);
  const normIdade = m => maxM > minM ? ((m - minM) / (maxM - minM)) * 100 : 50;

  const dGeral = descreve(regs.map(r => r.idade));
  const geral = {
    media: dGeral.media,
    pctFem: regs.length ? regs.filter(r => r.sexo === 'FEMININO').length / regs.length * 100 : 0,
    pctCursando: regs.length ? regs.filter(r => r.grupoSituacao === 'Cursando').length / regs.length * 100 : 0,
    pctDF: regs.length ? regs.filter(r => r.uf === 'DF').length / regs.length * 100 : 0,
    pctSolteiro: regs.length ? regs.filter(r => semAcento(r.estadoCivil) === 'SOLTEIRO').length / regs.length * 100 : 0
  };

  const eixos = ['Idade média', '% feminino', '% cursando', '% residentes no DF', '% solteiros(as)'];
  const reais = s => [s.media, s.pctFem, s.pctCursando, s.pctDF, s.pctSolteiro];
  const norm = s => [normIdade(s.media), s.pctFem, s.pctCursando, s.pctDF, s.pctSolteiro].map(x => +(x ?? 0).toFixed(1));
  const fmtReal = (i, v) => i === 0 ? nf1.format(v) + ' anos' : nf1.format(v) + '%';

  desenhar('cRadar', {
    type: 'radar',
    data: {
      labels: eixos,
      datasets: [
        {
          label: alvo.curso, data: norm(alvo), _reais: reais(alvo),
          borderColor: COR.laranja, backgroundColor: COR.laranja + '33',
          pointBackgroundColor: COR.laranja, borderWidth: 2.5
        },
        {
          label: 'Média institucional (filtro)', data: norm(geral), _reais: reais(geral),
          borderColor: COR.verde, backgroundColor: COR.verde + '26',
          pointBackgroundColor: COR.verde, borderWidth: 2, borderDash: [5, 4]
        }
      ]
    },
    options: {
      maintainAspectRatio: false,
      scales: {
        r: {
          min: 0, max: 100, ticks: { stepSize: 25, backdropColor: 'transparent', color: COR.muted },
          grid: { color: COR.borda }, angleLines: { color: COR.borda },
          pointLabels: { font: { size: 12, weight: '600' }, color: COR.ink }
        }
      },
      plugins: {
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: c => ' ' + c.dataset.label + ': ' + fmtReal(c.dataIndex, c.dataset._reais[c.dataIndex])
              + (c.dataIndex === 0 ? ' (escala relativa entre cursos)' : '')
          }
        }
      }
    }
  });
}

function renderHeatmap(regs, stats) {
  const grupos = GRUPOS_SITUACAO.map(g => g.rot);
  const porCurso = stats.slice().sort((a, b) => b.total - a.total);
  const linhas = porCurso.map(s => {
    const rs = regs.filter(r => r.cursoKey === s.key);
    const cont = grupos.map(g => rs.filter(r => r.grupoSituacao === g).length);
    return { curso: s.curso, total: s.total, cont };
  });
  const cor = (n, tot) => {
    if (!n) return 'background:#fff';
    const p = n / tot; // intensidade pelo % dentro do curso
    const alfa = Math.min(.92, .12 + p * .95);
    return `background:rgba(2,116,79,${alfa.toFixed(2)});color:${p > .45 ? '#fff' : 'var(--un-green-deep)'}`;
  };
  $('#heatmapWrap').innerHTML = `
    <table class="heatmap">
      <thead><tr><th style="text-align:left">Curso</th>${grupos.map(g => `<th>${g}</th>`).join('')}<th>Total</th></tr></thead>
      <tbody>${linhas.map(l => `
        <tr><th scope="row">${l.curso}</th>
        ${l.cont.map(n => `<td class="${n ? '' : 'hm-0'}" style="${cor(n, l.total)}">${n ? nf0.format(n) : '·'}</td>`).join('')}
        <td style="font-weight:700">${nf0.format(l.total)}</td></tr>`).join('')}
      </tbody>
    </table>`;
}

/* ── Tabela de estatísticas descritivas ───────────────────────────── */
const COLS_STATS = [
  { id: 'curso', rot: 'Curso', num: false },
  { id: 'n', rot: 'N', num: true, f: v => nf0.format(v) },
  { id: 'media', rot: 'Média', num: true, f: v => nf1.format(v) },
  { id: 'mediana', rot: 'Mediana', num: true, f: v => nf1.format(v) },
  { id: 'moda', rot: 'Moda', num: true, f: v => nf0.format(v) },
  { id: 'dp', rot: 'DP', num: true, f: v => nf1.format(v) },
  { id: 'min', rot: 'Mín', num: true, f: v => nf0.format(v) },
  { id: 'q1', rot: 'Q1', num: true, f: v => nf1.format(v) },
  { id: 'q3', rot: 'Q3', num: true, f: v => nf1.format(v) },
  { id: 'max', rot: 'Máx', num: true, f: v => nf0.format(v) },
  { id: 'pctFem', rot: '% Fem.', num: true, f: v => nf1.format(v) + '%' },
  { id: 'pctCursando', rot: '% Curs.', num: true, f: v => nf1.format(v) + '%' }
];

let statsAtuais = [];
function renderTabelaStats(regs, stats) {
  statsAtuais = stats;
  const { col, dir } = estado.ordem;
  const mult = dir === 'asc' ? 1 : -1;
  const ordenado = stats.slice().sort((a, b) => {
    const va = a[col], vb = b[col];
    if (va == null) return 1;
    if (vb == null) return -1;
    if (typeof va === 'string') return va.localeCompare(vb, 'pt-BR') * mult;
    return (va - vb) * mult;
  });

  $('#tblStatsHead').innerHTML = COLS_STATS.map(c => {
    const ativo = c.id === col;
    const seta = ativo ? (dir === 'asc' ? '▲' : '▼') : '↕';
    return `<th data-col="${c.id}" aria-sort="${ativo ? (dir === 'asc' ? 'ascending' : 'descending') : 'none'}">
      ${c.rot}<span class="sort-ind">${seta}</span></th>`;
  }).join('');

  const celula = (s, c) => {
    const v = s[c.id];
    const txt = v == null ? '—' : (c.f ? c.f(v) : v);
    return `<td class="${c.num ? 'num' : ''}">${txt}</td>`;
  };
  const geral = { curso: 'Geral (filtro atual)', ...descreve(regs.map(r => r.idade)),
    pctFem: regs.length ? regs.filter(r => r.sexo === 'FEMININO').length / regs.length * 100 : null,
    pctCursando: regs.length ? regs.filter(r => r.grupoSituacao === 'Cursando').length / regs.length * 100 : null };

  $('#tblStatsBody').innerHTML =
    ordenado.map(s => `<tr>${COLS_STATS.map(c => celula(s, c)).join('')}</tr>`).join('') +
    `<tr class="total-row">${COLS_STATS.map(c => celula(geral, c)).join('')}</tr>`;

  document.querySelectorAll('#tblStatsHead th').forEach(th => {
    th.addEventListener('click', () => {
      const c = th.dataset.col;
      if (estado.ordem.col === c) estado.ordem.dir = estado.ordem.dir === 'asc' ? 'desc' : 'asc';
      else estado.ordem = { col: c, dir: c === 'curso' ? 'asc' : 'desc' };
      renderTabelaStats(filtrados(), statsAtuais);
    });
  });
}

function exportarCSV() {
  const dec = s => s == null ? '' : String(s).replace('.', ',');
  const cab = ['Curso', 'N', 'Media', 'Mediana', 'Moda', 'Desvio_Padrao', 'Minimo', 'Q1', 'Q3', 'Maximo', 'Pct_Feminino', 'Pct_Cursando'];
  const linhas = statsAtuais.slice().sort((a, b) => b.total - a.total).map(s => [
    '"' + s.curso.replace(/"/g, '""') + '"', s.n,
    dec(s.media?.toFixed(2)), dec(s.mediana?.toFixed(2)), s.moda ?? '',
    dec(s.dp?.toFixed(2)), s.min ?? '', dec(s.q1?.toFixed(2)), dec(s.q3?.toFixed(2)), s.max ?? '',
    dec(s.pctFem?.toFixed(1)), dec(s.pctCursando?.toFixed(1))
  ].join(';'));
  const csv = '\ufeff' + cab.join(';') + '\n' + linhas.join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = 'estatisticas_idade_por_curso_uniceplac.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ── Lista de alunos (busca + paginação) ──────────────────────────── */
function renderTabelaAlunos(regs) {
  const q = semAcento(estado.busca);
  const base = q
    ? regs.filter(r => semAcento(`${r.nome} ${r.ra} ${r.cidade} ${r.bairro}`).includes(q))
    : regs;
  $('#alunosCount').textContent = nf0.format(base.length);

  const paginas = Math.max(1, Math.ceil(base.length / estado.porPagina));
  if (estado.pagina > paginas) estado.pagina = paginas;
  const ini = (estado.pagina - 1) * estado.porPagina;
  const vis = base.slice(ini, ini + estado.porPagina);

  $('#tblAlunosBody').innerHTML = vis.length ? vis.map(r => `
    <tr>
      <td class="num">${r.ra || '—'}</td>
      <td>${r.nome ? titulo(r.nome) : '<span style="color:var(--muted)">— (transferência)</span>'}</td>
      <td>${r.curso}</td>
      <td class="num">${r.idade ?? '—'}</td>
      <td>${r.sexo ? titulo(r.sexo) : '—'}</td>
      <td><span class="tag ${r.grupoCls}">${r.situacao ? titulo(r.situacao) : '—'}</span></td>
      <td>${r.cidade || '—'}${r.uf ? ' / ' + r.uf : ''}</td>
    </tr>`).join('')
    : `<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:26px">Nenhum aluno corresponde à busca e aos filtros atuais.</td></tr>`;

  // Paginação com janela e reticências
  const btn = (rot, pag, extra = '') => `<button type="button" data-pag="${pag}" ${extra}>${rot}</button>`;
  let seq = [];
  const win = 2;
  for (let p = 1; p <= paginas; p++) {
    if (p === 1 || p === paginas || Math.abs(p - estado.pagina) <= win) seq.push(p);
    else if (seq.at(-1) !== '…') seq.push('…');
  }
  $('#alunosPag').innerHTML =
    `<span class="page-info">Página ${estado.pagina} de ${paginas} · ${nf0.format(base.length)} registro(s)</span>` +
    btn('‹', estado.pagina - 1, estado.pagina === 1 ? 'disabled' : '') +
    seq.map(p => p === '…' ? '<span>…</span>' : btn(p, p, p === estado.pagina ? 'class="current" aria-current="page"' : '')).join('') +
    btn('›', estado.pagina + 1, estado.pagina === paginas ? 'disabled' : '');

  document.querySelectorAll('#alunosPag button[data-pag]').forEach(b => {
    b.addEventListener('click', () => {
      estado.pagina = +b.dataset.pag;
      renderTabelaAlunos(filtrados());
      $('#sec-alunos').scrollIntoView({ block: 'start' });
    });
  });
}

/* ── Filtros e navegação ──────────────────────────────────────────── */
function montarFiltros() {
  const t = estado.todos;
  const opt = (v, rot) => `<option value="${v}">${rot}</option>`;

  const cursos = contar(t, 'cursoKey')
    .map(([k]) => ({ k, rot: t.find(r => r.cursoKey === k).curso }))
    .sort((a, b) => a.rot.localeCompare(b.rot, 'pt-BR'));
  $('#fCurso').innerHTML = opt('', 'Todos os cursos') + cursos.map(c => opt(c.k, c.rot)).join('');

  const grupos = GRUPOS_SITUACAO.map(g => g.rot).filter(g => t.some(r => r.grupoSituacao === g));
  $('#fSituacao').innerHTML = opt('', 'Todas') + grupos.map(g => opt(g, g)).join('');

  $('#fSexo').innerHTML = opt('', 'Todos') + contar(t, 'sexo').map(([k]) => opt(k, titulo(k))).join('');
  $('#fUF').innerHTML = opt('', 'Todas') + contar(t, 'uf').map(([k, v]) => opt(k, `${k} (${nf0.format(v)})`)).join('');

  const aoMudar = () => {
    estado.filtro = {
      curso: $('#fCurso').value, situacao: $('#fSituacao').value,
      sexo: $('#fSexo').value, uf: $('#fUF').value
    };
    estado.pagina = 1;
    renderTudo();
  };
  ['#fCurso', '#fSituacao', '#fSexo', '#fUF'].forEach(s => $(s).addEventListener('change', aoMudar));

  $('#btnLimpar').addEventListener('click', () => {
    ['#fCurso', '#fSituacao', '#fSexo', '#fUF'].forEach(s => { $(s).value = ''; });
    $('#buscaAluno').value = '';
    estado.busca = '';
    aoMudar();
  });

  $('#buscaAluno').addEventListener('input', e => {
    estado.busca = e.target.value;
    estado.pagina = 1;
    renderTabelaAlunos(filtrados());
  });

  $('#radarCurso').addEventListener('change', e => {
    estado.radarCurso = e.target.value;
    const regs = filtrados();
    renderRadar(regs, statsPorCurso(regs));
  });

  $('#btnCSV').addEventListener('click', exportarCSV);
  $('#btnRecarregar').addEventListener('click', () => carregar(true));
}

function observarSecoes() {
  const links = [...document.querySelectorAll('.section-nav a')];
  const alvos = links.map(a => document.querySelector(a.getAttribute('href'))).filter(Boolean);
  const obs = new IntersectionObserver(entries => {
    entries.forEach(en => {
      if (!en.isIntersecting) return;
      links.forEach(l => l.classList.toggle('active', l.getAttribute('href') === '#' + en.target.id));
    });
  }, { rootMargin: '-40% 0px -55% 0px' });
  alvos.forEach(a => obs.observe(a));
}

/* ── Ciclo principal ──────────────────────────────────────────────── */
function renderTudo() {
  const regs = filtrados();
  const stats = statsPorCurso(regs);
  renderKPIs(regs);
  renderGraficosCursos(regs, stats);
  renderGraficosPerfil(regs);
  renderRadar(regs, stats);
  renderHeatmap(regs, stats);
  renderMapa(regs);
  renderTabelaStats(regs, stats);
  renderTabelaAlunos(regs);
}

async function carregar(forcar = false) {
  const loader = $('#loader');
  loader.classList.remove('hide');
  $('#errorPanel').classList.remove('show');
  try {
    const url = ARQUIVO_DADOS + (forcar ? '?t=' + Date.now() : '');
    const resp = await fetch(url, { cache: forcar ? 'reload' : 'default' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status + ' ao buscar ' + ARQUIVO_DADOS);
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const linhas = XLSX.utils.sheet_to_json(ws, { defval: null });
    estado.todos = prepararRegistros(linhas);
    if (!estado.todos.length) throw new Error('A planilha foi lida, mas nenhum registro de aluno foi reconhecido.');
    renderHero();
    renderTudo();
  } catch (e) {
    console.error(e);
    $('#errorDetail').textContent = e.message;
    $('#errorPanel').classList.add('show');
  } finally {
    loader.classList.add('hide');
  }
}

function boot() {
  lerCores();
  configurarChartJS();
  $('#anoAtual').textContent = new Date().getFullYear();
  montarFiltros();
  observarSecoes();
  carregar();
}

if (typeof window !== 'undefined' && typeof document !== 'undefined' && typeof Chart !== 'undefined') {
  boot();
}

/* Exporta funções puras para testes em Node (não afeta o navegador). */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { prepararRegistros, descreve, quantil, rotuloCurso, paraData, idadeEm, semAcento, titulo, GRUPOS_SITUACAO, normGeo, classificaBairro, pontoDoRegistro };
}
