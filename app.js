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
const ARQUIVO_DADOS = 'dados_alunos_2026_1.xlsx';
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
  charts: {}
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
  module.exports = { prepararRegistros, descreve, quantil, rotuloCurso, paraData, idadeEm, semAcento, titulo, GRUPOS_SITUACAO };
}
