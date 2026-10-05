const NOMES_ESTADOS = {
  AC: "Acre", AL: "Alagoas", AM: "Amazonas", AP: "Amapá", BA: "Bahia",
  CE: "Ceará", DF: "Distrito Federal", ES: "Espírito Santo", GO: "Goiás",
  MA: "Maranhão", MG: "Minas Gerais", MS: "Mato Grosso do Sul",
  MT: "Mato Grosso", PA: "Pará", PB: "Paraíba", PE: "Pernambuco",
  PI: "Piauí", PR: "Paraná", RJ: "Rio de Janeiro", RN: "Rio Grande do Norte",
  RO: "Rondônia", RR: "Roraima", RS: "Rio Grande do Sul",
  SC: "Santa Catarina", SE: "Sergipe", SP: "São Paulo", TO: "Tocantins",
};

const UFS = Object.keys(NOMES_ESTADOS).sort();

const HUB_CORES = {
  "Vilhena": "#e07a5f",
  "Ji-Paraná": "#81b29a",
  "Porto Velho": "#4ea8de",
  "Cacoal": "#f2cc8f",
};
const HUB_CORES_PR = {
  "Cascavel": "#e07a5f",
  "Curitiba": "#4ea8de",
};
const COR_SEM_CEEM = "#3a4453";

let dadosEstados = {};
let dadosCidadesRO = {};
let hubsRO = [];
let dadosCidadesPR = {};
let hubsPR = [];
let viewAtual = "brasil"; // "brasil" | "rondonia"

/* ---------- elementos ---------- */

const overlay = document.getElementById("overlay");
const modalUf = document.getElementById("modal-uf");
const modalNomeEstado = document.getElementById("modal-nome-estado");
const infoSwatch = document.getElementById("info-swatch");
const infoCeemValor = document.getElementById("info-ceem-valor");

const ufListEl = document.getElementById("uf-list");
const roListEl = document.getElementById("ro-list");
const prListEl = document.getElementById("pr-list");
const hubLegendEl = document.getElementById("hub-legend");
const hubLegendPrEl = document.getElementById("hub-legend-pr");
const statPreenchidos = document.getElementById("stat-preenchidos");
const statPendentes = document.getElementById("stat-pendentes");
const mapHint = document.getElementById("map-hint");
const roHint = document.getElementById("ro-hint");
const prHint = document.getElementById("pr-hint");

const mapTooltip = document.getElementById("map-tooltip");
const citySearchInput = document.getElementById("city-search-input");
const citySearchResults = document.getElementById("city-search-results");

const sideBrasil = document.getElementById("side-brasil");
const sideRondonia = document.getElementById("side-rondonia");
const sideParana = document.getElementById("side-parana");
const viewBrasilEl = document.getElementById("view-brasil");
const viewRondoniaEl = document.getElementById("view-rondonia");
const viewParanaEl = document.getElementById("view-parana");
const mapToggle = document.getElementById("map-toggle");

const ROTULOS_PAPEIS = [
  ["gerente_nome", "Gerente"],
  ["coord_operacional_nome", "Coord. Operacional"],
  ["coord_comercial_nome", "Coord. Comercial"],
];

/* ---------- carregar dados (lê data/ceem.xlsx direto no navegador) ---------- */

const PLANILHA_URL = "data/ceem.xlsx";
const HUBS_RO = ["Vilhena", "Ji-Paraná", "Porto Velho", "Cacoal"];
const HUBS_PR = ["Cascavel", "Curitiba"];

const COLUNAS = [
  "ceem", "cor",
  "gerente_nome", "gerente_telefone", "gerente_email",
  "coord_operacional_nome", "coord_operacional_telefone", "coord_operacional_email",
  "coord_comercial_nome", "coord_comercial_telefone", "coord_comercial_email",
];

function valorCelula(v) {
  return v === undefined || v === null ? "" : String(v).trim();
}

function linhaParaDict(chave, linha) {
  const dado = { chave };
  COLUNAS.forEach((campo, i) => { dado[campo] = valorCelula(linha[i + 1]); });
  return dado;
}

// linhas a partir da 4ª (índice 3), igual ao backend antigo
function lerAba(wb, nome) {
  const ws = wb.Sheets[nome];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true }).slice(3);
}

async function carregarPlanilha() {
  const resp = await fetch(PLANILHA_URL, { cache: "no-cache" });
  if (!resp.ok) throw new Error("Não foi possível carregar " + PLANILHA_URL);
  const wb = XLSX.read(await resp.arrayBuffer(), { type: "array" });

  const estados = {};
  lerAba(wb, "CEEM").forEach((linha) => {
    const chave = valorCelula(linha[0]).toUpperCase();
    if (UFS.includes(chave)) estados[chave] = linhaParaDict(chave, linha);
  });

  const porNome = (aba) => {
    const out = {};
    lerAba(wb, aba).forEach((linha) => {
      const nome = valorCelula(linha[0]);
      if (nome) out[nome] = linhaParaDict(nome, linha);
    });
    return out;
  };

  return { estados, cidadesRO: porNome("Cidades RO"), cidadesPR: porNome("Cidades PR") };
}

async function carregarTudo() {
  try {
    const dados = await carregarPlanilha();
    dadosEstados = dados.estados;
    dadosCidadesRO = dados.cidadesRO;
    dadosCidadesPR = dados.cidadesPR;
  } catch (err) {
    console.error(err);
    alert("Erro ao carregar a planilha data/ceem.xlsx. Se abriu o index.html direto do disco, rode por um servidor (ex.: python -m http.server).");
    return;
  }
  hubsRO = HUBS_RO;
  hubsPR = HUBS_PR;

  aplicarCoresNoMapaBrasil();
  renderizarListaEstados();
  atualizarStats();
  montarMapaRondonia();
  renderizarListaCidadesRO();
  montarMapaParana();
  renderizarListaCidadesPR();
  renderizarHubLegend(hubLegendEl, hubsRO, HUB_CORES);
  renderizarHubLegend(hubLegendPrEl, hubsPR, HUB_CORES_PR);
}

/* ---------- mapa Brasil ---------- */

function aplicarCoresNoMapaBrasil() {
  UFS.forEach((uf) => {
    const el = document.getElementById(uf);
    if (!el) return;
    const dado = dadosEstados[uf];
    if (dado && dado.ceem) {
      el.style.fill = dado.cor || "#4ea8de";
      el.classList.add("is-filled");
    } else {
      el.style.fill = "";
      el.classList.remove("is-filled");
    }
  });
}

function renderizarListaEstados() {
  ufListEl.innerHTML = "";
  UFS.forEach((uf) => {
    const dado = dadosEstados[uf];
    const li = document.createElement("li");
    li.className = "uf-row";

    const swatch = document.createElement("span");
    swatch.className = "uf-row__swatch";
    if (dado && dado.ceem) swatch.style.background = dado.cor || "#4ea8de";

    const ufLabel = document.createElement("span");
    ufLabel.className = "uf-row__uf";
    ufLabel.textContent = uf;

    const info = document.createElement("span");
    info.className = "uf-row__info";

    if (dado && dado.ceem) {
      const ceemLine = document.createElement("span");
      ceemLine.className = "uf-row__ceem";
      ceemLine.textContent = dado.ceem;
      const nomeLine = document.createElement("span");
      nomeLine.className = "uf-row__status";
      nomeLine.textContent = NOMES_ESTADOS[uf];
      info.appendChild(ceemLine);
      info.appendChild(nomeLine);
    } else {
      const empty = document.createElement("span");
      empty.className = "uf-row__empty";
      empty.textContent = NOMES_ESTADOS[uf] + " — não preenchido";
      info.appendChild(empty);
    }

    li.appendChild(swatch);
    li.appendChild(ufLabel);
    li.appendChild(info);
    li.addEventListener("click", () => abrirModalEstado(uf));
    ufListEl.appendChild(li);
  });
}

function atualizarStats() {
  const preenchidos = UFS.filter((uf) => dadosEstados[uf] && dadosEstados[uf].ceem).length;
  statPreenchidos.textContent = preenchidos;
  statPendentes.textContent = UFS.length - preenchidos;
}

/* ---------- mapa Rondônia / Paraná (genérico) ---------- */

function corDoHub(nomeHub, cores) {
  if (!nomeHub || nomeHub === "--") return null;
  return cores[nomeHub] || "#8b95a7";
}

function montarMapaSub(opcoes) {
  const { contornoId, marcadoresId, pontos, poligono, dadosCidades, hubs, cores, abrirModal, mostrarTooltip } = opcoes;
  const contorno = document.getElementById(contornoId);
  contorno.setAttribute("points", poligono);

  const grupo = document.getElementById(marcadoresId);
  grupo.innerHTML = "";

  pontos.forEach((pos) => {
    const dado = dadosCidades[pos.nome];
    const cor = dado ? corDoHub(dado.ceem, cores) : null;

    const circ = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circ.setAttribute("cx", pos.x);
    circ.setAttribute("cy", pos.y);
    circ.setAttribute("r", hubs.includes(pos.nome) ? 5.5 : 3.4);
    circ.setAttribute("class", "ro-marcador" + (hubs.includes(pos.nome) ? " ro-marcador--hub" : ""));
    circ.style.fill = cor || COR_SEM_CEEM;
    circ.dataset.nome = pos.nome;

    circ.addEventListener("click", () => abrirModal(pos.nome));
    circ.addEventListener("mouseenter", (e) => mostrarTooltip(pos.nome, e));
    circ.addEventListener("mousemove", posicionarTooltip);
    circ.addEventListener("mouseleave", esconderTooltip);

    grupo.appendChild(circ);
  });
}

function montarMapaRondonia() {
  montarMapaSub({
    contornoId: "ro-contorno", marcadoresId: "ro-marcadores",
    pontos: CIDADES_RO_POS, poligono: RO_POLYGON_POINTS,
    dadosCidades: dadosCidadesRO, hubs: hubsRO, cores: HUB_CORES,
    abrirModal: abrirModalCidadeRO, mostrarTooltip: mostrarTooltipCidadeRO,
  });
}

function montarMapaParana() {
  montarMapaSub({
    contornoId: "pr-contorno", marcadoresId: "pr-marcadores",
    pontos: CIDADES_PR_POS, poligono: PR_POLYGON_POINTS,
    dadosCidades: dadosCidadesPR, hubs: hubsPR, cores: HUB_CORES_PR,
    abrirModal: abrirModalCidadePR, mostrarTooltip: mostrarTooltipCidadePR,
  });
}

function renderizarListaCidadesSub(opcoes) {
  const { listaEl, pontos, dadosCidades, cores, abrirModal } = opcoes;
  listaEl.innerHTML = "";
  const cidadesOrdenadas = [...pontos].sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));

  cidadesOrdenadas.forEach((pos) => {
    const dado = dadosCidades[pos.nome];
    const li = document.createElement("li");
    li.className = "uf-row";

    const swatch = document.createElement("span");
    swatch.className = "uf-row__swatch";
    const cor = dado ? corDoHub(dado.ceem, cores) : null;
    swatch.style.background = cor || "";

    const info = document.createElement("span");
    info.className = "uf-row__info";

    const nomeLine = document.createElement("span");
    nomeLine.className = "uf-row__ceem";
    nomeLine.textContent = pos.nome;

    const hubLine = document.createElement("span");
    hubLine.className = "uf-row__status";
    hubLine.textContent = dado && dado.ceem ? "CEEM: " + dado.ceem : "Sem CEEM definido";

    info.appendChild(nomeLine);
    info.appendChild(hubLine);
    li.appendChild(swatch);
    li.appendChild(info);
    li.addEventListener("click", () => abrirModal(pos.nome));
    listaEl.appendChild(li);
  });
}

function renderizarListaCidadesRO() {
  renderizarListaCidadesSub({
    listaEl: roListEl, pontos: CIDADES_RO_POS, dadosCidades: dadosCidadesRO,
    cores: HUB_CORES, abrirModal: abrirModalCidadeRO,
  });
}

function renderizarListaCidadesPR() {
  renderizarListaCidadesSub({
    listaEl: prListEl, pontos: CIDADES_PR_POS, dadosCidades: dadosCidadesPR,
    cores: HUB_CORES_PR, abrirModal: abrirModalCidadePR,
  });
}

function renderizarHubLegend(elemento, hubs, cores) {
  elemento.innerHTML = "";
  hubs.forEach((hub) => {
    const li = document.createElement("li");
    li.className = "hub-legend__item";
    const dot = document.createElement("span");
    dot.className = "hub-legend__dot";
    dot.style.background = cores[hub] || "#8b95a7";
    const label = document.createElement("span");
    label.textContent = hub;
    li.appendChild(dot);
    li.appendChild(label);
    elemento.appendChild(li);
  });
  const liVazio = document.createElement("li");
  liVazio.className = "hub-legend__item";
  const dotVazio = document.createElement("span");
  dotVazio.className = "hub-legend__dot";
  dotVazio.style.background = COR_SEM_CEEM;
  const labelVazio = document.createElement("span");
  labelVazio.textContent = "Sem CEEM";
  liVazio.appendChild(dotVazio);
  liVazio.appendChild(labelVazio);
  elemento.appendChild(liVazio);
}

/* ---------- toggle Brasil / Rondônia ---------- */

mapToggle.addEventListener("click", (e) => {
  const btn = e.target.closest(".map-toggle__btn");
  if (!btn) return;
  const view = btn.dataset.view;
  if (view === viewAtual) return;
  viewAtual = view;

  mapToggle.querySelectorAll(".map-toggle__btn").forEach((b) => {
    b.classList.toggle("is-active", b.dataset.view === view);
  });

  viewBrasilEl.style.display = view === "brasil" ? "" : "none";
  viewRondoniaEl.style.display = view === "rondonia" ? "" : "none";
  viewParanaEl.style.display = view === "parana" ? "" : "none";

  mapHint.style.display = view === "brasil" ? "" : "none";
  roHint.style.display = view === "rondonia" ? "" : "none";
  prHint.style.display = view === "parana" ? "" : "none";

  sideBrasil.style.display = view === "brasil" ? "" : "none";
  sideRondonia.style.display = view === "rondonia" ? "" : "none";
  sideParana.style.display = view === "parana" ? "" : "none";
});

/* ---------- tooltip customizado ---------- */

function montarTooltipEstado(uf) {
  const dado = dadosEstados[uf];
  let html = `<div class="map-tooltip__title">${uf} — ${NOMES_ESTADOS[uf]}</div>`;
  if (dado && dado.ceem) {
    html += `<div class="map-tooltip__ceem">CEEM: ${dado.ceem}</div>`;
    ROTULOS_PAPEIS.forEach(([campo, rotulo]) => {
      if (dado[campo]) html += `<div class="map-tooltip__role"><b>${rotulo}:</b> ${dado[campo]}</div>`;
    });
  } else {
    html += `<div class="map-tooltip__empty">Não preenchido</div>`;
  }
  return html;
}

function montarTooltipCidadeRO(nome) {
  const dado = dadosCidadesRO[nome];
  let html = `<div class="map-tooltip__title">${nome} — RO</div>`;
  if (dado && dado.ceem) {
    html += `<div class="map-tooltip__ceem">CEEM: ${dado.ceem}</div>`;
    ROTULOS_PAPEIS.forEach(([campo, rotulo]) => {
      if (dado[campo]) html += `<div class="map-tooltip__role"><b>${rotulo}:</b> ${dado[campo]}</div>`;
    });
  } else {
    html += `<div class="map-tooltip__empty">Sem CEEM definido</div>`;
  }
  return html;
}

function montarTooltipCidadePR(nome) {
  const dado = dadosCidadesPR[nome];
  let html = `<div class="map-tooltip__title">${nome} — PR</div>`;
  if (dado && dado.ceem) {
    html += `<div class="map-tooltip__ceem">CEEM: ${dado.ceem}</div>`;
    ROTULOS_PAPEIS.forEach(([campo, rotulo]) => {
      if (dado[campo]) html += `<div class="map-tooltip__role"><b>${rotulo}:</b> ${dado[campo]}</div>`;
    });
  } else {
    html += `<div class="map-tooltip__empty">Sem CEEM definido</div>`;
  }
  return html;
}

function mostrarTooltip(uf, evento) {
  mapTooltip.innerHTML = montarTooltipEstado(uf);
  mapTooltip.classList.add("is-visible");
  posicionarTooltip(evento);
}

function mostrarTooltipCidadeRO(nome, evento) {
  mapTooltip.innerHTML = montarTooltipCidadeRO(nome);
  mapTooltip.classList.add("is-visible");
  posicionarTooltip(evento);
}

function mostrarTooltipCidadePR(nome, evento) {
  mapTooltip.innerHTML = montarTooltipCidadePR(nome);
  mapTooltip.classList.add("is-visible");
  posicionarTooltip(evento);
}

function posicionarTooltip(evento) {
  const margem = 16;
  let x = evento.clientX + margem;
  let y = evento.clientY + margem;
  const rect = mapTooltip.getBoundingClientRect();
  if (x + rect.width > window.innerWidth - 8) x = evento.clientX - rect.width - margem;
  if (y + rect.height > window.innerHeight - 8) y = evento.clientY - rect.height - margem;
  mapTooltip.style.left = x + "px";
  mapTooltip.style.top = y + "px";
}

function esconderTooltip() {
  mapTooltip.classList.remove("is-visible");
}

/* ---------- modal de informação (somente leitura) ---------- */

function preencherPapel(prefixo, dado, campoBase) {
  const nome = dado ? dado[campoBase + "_nome"] : "";
  const tel = dado ? dado[campoBase + "_telefone"] : "";
  const email = dado ? dado[campoBase + "_email"] : "";

  document.getElementById(`info-${prefixo}-nome`).textContent = nome || "--";
  const telEl = document.getElementById(`info-${prefixo}-tel`);
  const emailEl = document.getElementById(`info-${prefixo}-email`);
  telEl.textContent = tel || "--";
  emailEl.textContent = email || "--";
}

function abrirModalComDados(titulo, subtitulo, dado) {
  modalUf.textContent = titulo;
  modalNomeEstado.textContent = subtitulo;

  const temCeem = dado && dado.ceem;
  infoCeemValor.textContent = temCeem ? dado.ceem : "Não preenchido";
  infoSwatch.style.background = temCeem ? (dado.cor || "#4ea8de") : "transparent";

  preencherPapel("gerente", dado, "gerente");
  preencherPapel("operacional", dado, "coord_operacional");
  preencherPapel("comercial", dado, "coord_comercial");

  overlay.classList.add("is-open");
}

function abrirModalEstado(uf) {
  abrirModalComDados(uf, NOMES_ESTADOS[uf], dadosEstados[uf]);
}

function abrirModalCidadeRO(nome) {
  abrirModalComDados(nome, "Rondônia (RO)", dadosCidadesRO[nome]);
}

function abrirModalCidadePR(nome) {
  abrirModalComDados(nome, "Paraná (PR)", dadosCidadesPR[nome]);
}

function fecharModal() {
  overlay.classList.remove("is-open");
}

document.getElementById("modal-close").addEventListener("click", fecharModal);
overlay.addEventListener("click", (e) => { if (e.target === overlay) fecharModal(); });
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && overlay.classList.contains("is-open")) fecharModal();
});

UFS.forEach((uf) => {
  const el = document.getElementById(uf);
  if (!el) return;
  el.addEventListener("click", () => abrirModalEstado(uf));
  el.addEventListener("mouseenter", (e) => mostrarTooltip(uf, e));
  el.addEventListener("mousemove", posicionarTooltip);
  el.addEventListener("mouseleave", esconderTooltip);
});

/* ---------- busca por cidade (nacional, abre o estado) ---------- */

function normalizar(txt) {
  return txt.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function buscarCidades(termo) {
  const alvo = normalizar(termo);
  const resultados = [];
  for (let i = 0; i < CIDADES_BR.length && resultados.length < 8; i++) {
    const cidade = CIDADES_BR[i];
    if (normalizar(cidade.n).includes(alvo)) resultados.push(cidade);
  }
  return resultados;
}

function renderizarResultadosCidade(termo) {
  citySearchResults.innerHTML = "";
  if (termo.trim().length < 2) {
    citySearchResults.classList.remove("is-open");
    return;
  }
  const resultados = buscarCidades(termo);
  if (resultados.length === 0) {
    const li = document.createElement("li");
    li.className = "city-search__empty";
    li.textContent = "Nenhuma cidade encontrada.";
    citySearchResults.appendChild(li);
  } else {
    resultados.forEach((cidade) => {
      const li = document.createElement("li");
      li.className = "city-search__result";
      const nome = document.createElement("span");
      nome.textContent = cidade.n;
      const uf = document.createElement("span");
      uf.className = "city-search__result-uf";
      uf.textContent = cidade.uf;
      li.appendChild(nome);
      li.appendChild(uf);
      li.addEventListener("click", () => {
        citySearchInput.value = `${cidade.n} — ${cidade.uf}`;
        citySearchResults.classList.remove("is-open");
        abrirModalEstado(cidade.uf);
      });
      citySearchResults.appendChild(li);
    });
  }
  citySearchResults.classList.add("is-open");
}

citySearchInput.addEventListener("input", () => renderizarResultadosCidade(citySearchInput.value));
citySearchInput.addEventListener("focus", () => {
  if (citySearchInput.value.trim().length >= 2) renderizarResultadosCidade(citySearchInput.value);
});
document.addEventListener("click", (e) => {
  if (!e.target.closest(".city-search")) citySearchResults.classList.remove("is-open");
});

carregarTudo();