// Slides compartilhados pelas duas versoes do deck.
// Injetado nos templates no lugar do marcador /*COMUM*/, antes de build().
// Depende de DADOS, esc(), cor() e rodape(), definidos no proprio template.

function slideContexto(n, t) {
  const c = DADOS.contexto;
  return `<section class="slide" style="--cat:${cor("major")}">
    <div class="slide-head"><span class="eyebrow">Contexto</span></div>
    <h2 class="slide-title">${esc(c.titulo)}</h2>
    <p class="extenso" style="max-width:74ch">${esc(c.nome_oficial)}</p>
    <div class="ctx">
      <div>
        <p class="resumo" style="max-width:56ch;margin-top:1.2rem">${esc(c.resumo)}</p>
        <div class="ctx-fichas">
          ${c.fichas.map(([k, v]) => `<div class="ficha"><span class="k">${esc(k)}</span><span class="v">${esc(v)}</span></div>`).join("")}
        </div>
      </div>
      <div>
        <div class="bloco-lat">
          <h3>As quatro etapas do plano</h3>
          <ol class="etapas">
            ${c.etapas.map(([num, nome, prazo]) => `<li><span class="et-n">${esc(num)}</span><span class="et-nome">${esc(nome)}</span><span class="et-prazo">${esc(prazo)}</span></li>`).join("")}
          </ol>
        </div>
        <div class="bloco-lat">
          <h3>O que o plano persegue</h3>
          <ul class="lista-simples">
            ${c.objetivos.map((o) => `<li>${esc(o)}</li>`).join("")}
          </ul>
        </div>
      </div>
    </div>
    ${rodape("pli.semil.sp.gov.br", "Contexto", n, t)}
  </section>`;
}

function slideCadeia(n, t) {
  const c = DADOS.cadeia;
  return `<section class="slide" style="--cat:${cor("tools")}">
    <div class="slide-head"><span class="eyebrow">Cadeia de produtos do TDR</span></div>
    <h2 class="slide-title">${esc(c.titulo)}</h2>
    <p class="extenso" style="max-width:82ch">${esc(c.resumo)}</p>
    <div class="cadeia">
      ${c.elos.map(([cod, nome, desc, apps]) => `<div class="elo">
        <div class="elo-cod">${esc(cod)}</div>
        <div class="elo-nome">${esc(nome)}</div>
        <p class="elo-desc">${esc(desc)}</p>
        <div class="elo-apps">${apps.map((a) => `<span class="chip">${esc(a)}</span>`).join("")}</div>
      </div>`).join("")}
    </div>
    ${rodape("Termo de Referência do PLI-SP", "Cadeia de produtos", n, t)}
  </section>`;
}

function slideFontes(n, t) {
  return `<section class="slide" style="--cat:${cor("major")}">
    <div class="slide-head"><span class="eyebrow">Procedência</span></div>
    <h2 class="slide-title">De onde vem cada afirmação</h2>
    <p class="extenso" style="max-width:78ch">Nada neste documento foi escrito de memória. O enquadramento institucional vem das fontes oficiais do plano; o detalhamento técnico, da documentação das próprias aplicações.</p>
    <div class="fontes">
      ${DADOS.fontes.map(([nome, desc], i) => `<div class="fonte">
        <span class="f-n">${String(i + 1).padStart(2, "0")}</span>
        <div>
          <div class="f-nome">${esc(nome)}</div>
          <div class="f-desc">${esc(desc)}</div>
        </div>
      </div>`).join("")}
    </div>
    ${rodape(DADOS.hub_url, DADOS.org + " · " + DADOS.data, n, t)}
  </section>`;
}
