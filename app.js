(function(){
  // Helpers
  const qs = (sel, ctx=document) => ctx.querySelector(sel);
  const qsa = (sel, ctx=document) => Array.from(ctx.querySelectorAll(sel));
  const toastEl = qs('#toast');
  function toast(msg, type='success'){
    if(!toastEl) return;
    toastEl.textContent = msg;
    toastEl.className = `toast show ${type}`;
    setTimeout(()=>{ toastEl.className = 'toast'; }, 3000);
  }

  // Tabs navegação
  const tabs = qsa('.tab');
  const panels = qsa('.tab-panel');
  tabs.forEach(btn => btn.addEventListener('click', () => {
    tabs.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    panels.forEach(p => p.classList.remove('active'));
    const id = btn.dataset.tab;
    const panel = qs(`#tab-${id}`);
    if(panel) panel.classList.add('active');
  }));

  // Dropdown de opções ao clicar nos ícones das abas
  let dropdown;
  function showDropdown(targetBtn, items){
    hideDropdown();
    dropdown = document.createElement('div');
    dropdown.className = 'dropdown-menu';
    const rect = targetBtn.getBoundingClientRect();
    dropdown.style.top = `${rect.bottom + 8 + window.scrollY}px`;
    dropdown.style.left = `${rect.left + window.scrollX}px`;
    dropdown.innerHTML = items.map(i => `<div class="dropdown-item" data-id="${i.id}">${i.icon||''}${i.label}</div>`).join('') + '<div class="dropdown-sep"></div>';
    document.body.appendChild(dropdown);
    dropdown.addEventListener('click', (e)=>{
      const itemEl = e.target.closest('.dropdown-item');
      if(!itemEl) return;
      const id = itemEl.getAttribute('data-id');
      iHandlers[id] && iHandlers[id]();
      hideDropdown();
    });
    document.addEventListener('click', onDocClick);
  }
  function hideDropdown(){
    if(dropdown){ dropdown.remove(); dropdown=null; document.removeEventListener('click', onDocClick);}  
  }
  function onDocClick(e){
    if(dropdown && !dropdown.contains(e.target)) hideDropdown();
  }

  const iHandlers = {
    'cad-evento': () => setActiveTab('eventos'),
    'cad-ministerio': () => setActiveTab('ministerio'),
    'cad-congregacao': () => setActiveTab('congregacoes'),
    'ver-relatorios': () => { setActiveTab('relatorios'); renderRelatorios(); },
    'rel-resumo-eventos': () => { setActiveTab('relatorios'); renderRelatorioEventos(); },
    'rel-ministerio-funcao': () => { setActiveTab('relatorios'); renderRelatorioMinisterio(); },
    'rel-total-congregacoes': () => { setActiveTab('relatorios'); renderRelatorioCongregacoes(); },
  };
  function setActiveTab(id){
    tabs.forEach(b => b.classList.remove('active'));
    panels.forEach(p => p.classList.remove('active'));
    const btn = qs(`#tab-btn-${id}`);
    const panel = qs(`#tab-${id}`);
    if(btn) btn.classList.add('active');
    if(panel) panel.classList.add('active');
  }

  const eventsMenu = [
    {id:'cad-evento', label:'Marcar Atendimento'},
  ];
  const minMenu = [
    {id:'cad-ministerio', label:'Cadastrar Irmão do Ministério'},
  ];
  const congMenu = [
    {id:'cad-congregacao', label:'Cadastrar Congregação'},
  ];
  const relMenu = [
    {id:'ver-relatorios', label:'Ver Relatórios'},
    {id:'rel-resumo-eventos', label:'Resumo de Atendimentos'},
    {id:'rel-ministerio-funcao', label:'Ministério por Função'},
    {id:'rel-total-congregacoes', label:'Total de Congregações'},
  ];

  const btnEventos = qs('#tab-btn-eventos');
  const btnMinisterio = qs('#tab-btn-ministerio');
  const btnCongregacoes = qs('#tab-btn-congregacoes');
  const btnRelatorios = qs('#tab-btn-relatorios');

  [
    [btnEventos, eventsMenu],
    [btnMinisterio, minMenu],
    [btnCongregacoes, congMenu],
    [btnRelatorios, relMenu]
  ].forEach(([btn, items]) => {
    if(!btn) return;
    btn.addEventListener('contextmenu', (e)=>{ e.preventDefault(); showDropdown(btn, items); });
    btn.addEventListener('click', (e)=>{
      // Clique normal alterna aba; clique com Shift abre opções
      if(e.shiftKey){ e.preventDefault(); showDropdown(btn, items); }
    });
  });

  // Firebase init
  let db = null;
  const configAlert = qs('#config-alert');
  try{
    const cfg = window.FIREBASE_CONFIG || {};
    if(!cfg.apiKey || !cfg.databaseURL){
      configAlert && configAlert.classList.remove('hidden');
    } else {
      firebase.initializeApp(cfg);
      db = firebase.database();
      configAlert && configAlert.classList.add('hidden');
      toast('Firebase conectado');
    }
  }catch(err){ console.error(err); }

  // CRUD helpers
  function write(path, obj){
    if(!db){ toast('Firebase não configurado', 'error'); return Promise.resolve(null); }
    const ref = db.ref(path).push();
    const data = { id: ref.key, createdAt: Date.now(), ...obj };
    return ref.set(data).then(()=> data);
  }
  function update(path, id, obj){
    if(!db){ toast('Firebase não configurado', 'error'); return Promise.resolve(null); }
    const ref = db.ref(`${path}/${id}`);
    const data = { ...obj, updatedAt: Date.now() };
    return ref.update(data).then(()=> ({ id, ...data }));
  }
  function readList(path, cb){
    if(!db) return;
    db.ref(path).on('value', snap => {
      const val = snap.val() || {};
      const list = Object.values(val).sort((a,b)=>a.createdAt-b.createdAt);
      cb(list);
    });
  }
  function remove(path, id){
    if(!db){ toast('Firebase não configurado', 'error'); return Promise.resolve(null); }
    const ref = db.ref(`${path}/${id}`);
    return ref.remove().then(()=> true);
  }

  // Eventos
  const formEvento = qs('#form-evento');
  const listaEventos = qs('#lista-eventos');
  const eventoCong = qs('#evento-congregacao');
  const eventoAtendente = qs('#evento-atendente');
  const eventoTipoSel = qs('select[name="tipo"]');
  const eventoEnsaioTipoWrap = qs('#evento-ensaio-tipo-wrapper');
  const eventoEnsaioTipoSel = qs('#evento-ensaio-tipo');
  const outraRegiaoBtn = qs('#atendente-outra-regiao-btn');
  const atendenteManualInput = qs('#evento-atendente-manual');
  const eventoCultosBody = qs('#evento-cultos-body');
  const tabelaReforcosBody = qs('#tabela-reforcos-body');
  const btnImprimirEventos = qs('#btn-imprimir-eventos');
  const btnExportarPdfEventos = qs('#btn-exportar-pdf-eventos');
  const btnExportarXlsEventos = qs('#btn-exportar-xls-eventos');
  // Atendimentos (nova página)
  const formAtendimento = qs('#form-atendimento');
  const atendCongSel = qs('#atend-congregacao');
  const atendTipoSel = qs('#atend-tipo');
  const atendAnciaoSel = qs('#atend-anciao');
  const atendAnciaoExternoChk = qs('#atend-anciao-externo');
  const atendAnciaoNomeInput = qs('#atend-anciao-nome');
  const atendEncWrap = qs('#atend-enc-wrapper');
  const atendEncRegSel = qs('#atend-enc-regional');
  const atendEncExternoChk = qs('#atend-enc-externo');
  const atendEncNomeInput = qs('#atend-enc-nome');
  const atendCultosBody = qs('#atend-cultos-body');
const tabelaAtendBatBody = qs('#tabela-atend-batismo-body');
const tabelaAtendCeiaBody = qs('#tabela-atend-ceia-body');
const tabelaAtendEnsaioBody = qs('#tabela-atend-ensaio-body');
const tabelaAtendMocBody = qs('#tabela-atend-mocidade-body');
  const listaAtendimentos = qs('#lista-atendimentos');
  // Filtro de tabela (Index): Ano/Mês
  const indexYearSel = qs('#index-year');
  const indexMonthSel = qs('#index-month');
  // Barra de pendências (Ticker do próximo mês)
  const tickerSemReforcoEl = qs('#ticker-sem-reforco');

  // Ticker de pendências de Reforço (próximo mês)
  function getNextMonthYear(){
    const now = new Date();
    const currentMonth = now.getMonth()+1; // 1-12
    const nextMonth = currentMonth===12 ? 1 : currentMonth+1;
    const nextYear = currentMonth===12 ? (now.getFullYear()+1) : now.getFullYear();
    return { year: nextYear, month: nextMonth };
  }
  function escapeHtml(str){ return (str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  let _tickerTimer = null, _tickerMsgs = [], _tickerIndex = 0;
  function startTicker(msgs){
    if(!tickerSemReforcoEl) return;
    _tickerMsgs = (msgs && msgs.length) ? msgs : ['Sem pendências para o próximo mês'];
    _tickerIndex = 0;
    showTickerItem();
    if(_tickerTimer) clearInterval(_tickerTimer);
    _tickerTimer = setInterval(()=>{
      _tickerIndex = (_tickerIndex+1) % _tickerMsgs.length;
      showTickerItem();
    }, 4000);
  }
  function showTickerItem(){
    if(!tickerSemReforcoEl) return;
    const msg = _tickerMsgs[_tickerIndex] || '';
    tickerSemReforcoEl.innerHTML = `<div class="ticker-item">${escapeHtml(msg)}</div>`;
  }
  function renderTickerSemReforco(){
    if(!tickerSemReforcoEl) return;
    const { year, month } = getNextMonthYear();
    const msgs = [];
    const listCong = congregacoesCacheEvents || [];
    const evs = eventosCache || [];
    const hasEvFor = (congId, tipoLabel) => {
      return evs.some(ev => {
        if(!ev || ev.congregacaoId!==congId) return false;
        if(!(ev.tipo=== 'Culto Reforço de Coletas' || ev.tipo=== 'RJM com Reforço de Coletas')) return false;
        if(tipoLabel==='CO' && ev.tipo!=='Culto Reforço de Coletas') return false;
        if(tipoLabel==='RJM' && ev.tipo!=='RJM com Reforço de Coletas') return false;
        const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
        return d && !isNaN(d.getTime()) && d.getFullYear()===year && (d.getMonth()+1)===month;
      });
    };
    listCong.forEach(c => {
      const label = c.nomeFormatado || (c.cidade && c.bairro ? `${c.cidade} - ${c.bairro}` : (c.nome||c.id));
      const cultos = Array.isArray(c.cultos) ? c.cultos : [];
      const temCO = cultos.some(ct => ct && ct.tipo==='Culto Oficial');
      const temRJM = cultos.some(ct => ct && ct.tipo==='RJM');
      const pend = [];
      if(temCO && !hasEvFor(c.id, 'CO')) pend.push('CO');
      if(temRJM && !hasEvFor(c.id, 'RJM')) pend.push('RJM');
      if(pend.length){
        msgs.push(`${label}: sem reforço ${pend.join(' e ')}`);
      }
    });
    startTicker(msgs);
  }
  // Relatórios
const relEventosEl = qs('#relatorio-eventos');
const relMinisterioEl = qs('#relatorio-ministerio');
const relCongregacoesEl = qs('#relatorio-congregacoes');
const relCongEnsaiosEl = qs('#relatorio-cong-ensaios');
// Filtros e botões de Relatórios
const relYearSel = qs('#rel-year');
const relMonthSel = qs('#rel-month');
const relCidadeSel = qs('#rel-cidade');
const relTipoSel = qs('#rel-tipo');
  const btnRelApply = qs('#rel-apply');
  const btnRelClear = qs('#rel-clear');
  const relAutoApplyChk = qs('#rel-auto-apply');
  const relLoading = qs('#rel-loading');
const btnRelPrint = qs('#rel-print');
const btnRelPdf = qs('#rel-print-pdf');
const btnRelXls = qs('#rel-export-xls');
const relEnsaiosSortSel = qs('#rel-ensaios-sort');
const btnEnsaiosExportCsv = qs('#rel-ensaios-export-csv');
const btnEnsaiosExportXls = qs('#rel-ensaios-export-xls');
const btnEnsaiosCopy = qs('#rel-ensaios-copy');
const relBairroSel = qs('#rel-bairro');
const relEnsaiosNextOnly = qs('#rel-ensaios-next-only');
const btnRelBackfillCidade = qs('#rel-backfill-cidade');
// Export de Serviços (Relatórios)
const relServicoTipoSel = qs('#rel-servico-tipo');
const btnRelServicoPdf = qs('#rel-servico-export-pdf');
const btnRelServicoCsv = qs('#rel-servico-export-csv');
const btnRelServicoXls = qs('#rel-servico-export-xls');
const btnRelServicoImport = qs('#rel-servico-import');
const fileRelServicoImport = qs('#file-rel-servico-import');
const btnRelServicoImportTest = qs('#rel-servico-import-test');

  // Seletores Musical
  const formEnsaioMusical = qs('#form-ensaio-musical');
  const musCongSel = qs('#mus-congregacao');
  const musEnsaioTipoSel = qs('#mus-ensaio-tipo');
  const musEnsaioAddTypeBtn = qs('#mus-ensaio-addtype-btn');
  const musEnsaioDataInp = qs('#mus-ensaio-data');
  const musEnsaioHoraInp = qs('#mus-ensaio-hora');
  const musEnsaioObsTxt = qs('#mus-ensaio-obs');

  // Seletores da Agenda Musical (declarar ANTES de uso)
  const musAgendaYearSel = qs('#mus-agenda-year');
  const musAgendaMonthSel = qs('#mus-agenda-month');
  const musAgendaBody = qs('#mus-agenda-body');
  const btnMusImprimirCalendario = qs('#btn-musical-imprimir-calendario');

  // Seletores do formulário de Pessoas (declarar ANTES de uso)

  // Caches globais (declarar ANTES de uso)
  let eventosCache = [];
  let congregacoesCacheEvents = [];
  let congregacoesByIdEvents = {};
  let resultadosCache = [];



  // Novos seletores de recorrência (agenda de ensaios)
  const musRecAnoSel = qs('#mus-rec-ano');
  const musRecDiaSel = qs('#mus-rec-dia');
  const musRecSemanaSel = qs('#mus-rec-semana');
  const musRecMesesWrap = qs('#mus-rec-meses');

  // Helpers de data para recorrência de ensaios
  function toYmdLocal(d){
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const dd = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${dd}`;
  }
  function getNthWeekdayOfMonth(year, month/*1-12*/, weekday/*0-6*/, nth/*1-5*/){
    const first = new Date(year, month-1, 1);
    const offset = (weekday - first.getDay() + 7) % 7;
    const day = 1 + offset + (nth-1)*7;
    const cand = new Date(year, month-1, day);
    if (cand.getMonth() !== (month-1)) return null;
    return toYmdLocal(cand);
  }
  function getLastWeekdayOfMonth(year, month/*1-12*/, weekday/*0-6*/){
    const last = new Date(year, month, 0); // último dia do mês
    const offset = (last.getDay() - weekday + 7) % 7;
    const day = last.getDate() - offset;
    const cand = new Date(year, month-1, day);
    return toYmdLocal(cand);
  }

  // Tipos de Ensaio customizados (localStorage) para Musical
  function getCustomEnsaioTypesMus(){
    try{ return JSON.parse(localStorage.getItem('ensaioTiposCustom')||'[]'); }catch{ return []; }
  }
  function setCustomEnsaioTypesMus(arr){
    try{ localStorage.setItem('ensaioTiposCustom', JSON.stringify(arr||[])); }catch{}
  }
  function renderCustomEnsaioTypesMus(){
    if(!musEnsaioTipoSel) return;
    const existing = new Set(Array.from(musEnsaioTipoSel.querySelectorAll('option')).map(o=>o.value.trim().toLowerCase()));
    getCustomEnsaioTypesMus().forEach(t=>{
      const v = String(t||'').trim();
      if(!v) return;
      if(existing.has(v.toLowerCase())) return;
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      musEnsaioTipoSel.appendChild(opt);
    });
  }
  function setupMusEnsaioTipoAdder(){
    if(!musEnsaioAddTypeBtn || !musEnsaioTipoSel) return;
    musEnsaioAddTypeBtn.addEventListener('click', ()=>{
      const name = (prompt('Nome do novo tipo de ensaio:')||'').trim();
      if(!name) return;
      const existing = new Set(Array.from(musEnsaioTipoSel.querySelectorAll('option')).map(o=>o.value.trim().toLowerCase()));
      if(existing.has(name.toLowerCase())){ toast('Tipo de ensaio já existe', 'warning'); return; }
      const updated = Array.from(new Set([...getCustomEnsaioTypesMus(), name].map(s=>String(s||'').trim()).filter(Boolean)));
      setCustomEnsaioTypesMus(updated);
      renderCustomEnsaioTypesMus();
      musEnsaioTipoSel.value = name;
      toast('Tipo de ensaio adicionado', 'success');
    });
  }
  // Inicialização dos tipos customizados para Musical
  renderCustomEnsaioTypesMus();
  setupMusEnsaioTipoAdder();

  // Handler de submit do formulário de Cadastro de Ensaios (Musical)
  if (formEnsaioMusical){
    formEnsaioMusical.addEventListener('submit', async (e)=>{
      e.preventDefault();
      try{
        const congregacaoId = musCongSel ? musCongSel.value : '';
        if(!congregacaoId){ toast('Selecione uma congregação', 'error'); return; }
        const congregacaoNome = (congregacoesByIdEvents && congregacoesByIdEvents[congregacaoId] && (congregacoesByIdEvents[congregacaoId].nome || '')) || (musCongSel && musCongSel.options && musCongSel.options[musCongSel.selectedIndex] && musCongSel.options[musCongSel.selectedIndex].text) || '';
        const ensaioTipo = musEnsaioTipoSel ? musEnsaioTipoSel.value : '';
        if(!ensaioTipo){ toast('Selecione o tipo de ensaio', 'error'); return; }

        const hora = musEnsaioHoraInp ? (musEnsaioHoraInp.value || '') : '';
        const obs = musEnsaioObsTxt ? (musEnsaioObsTxt.value || '') : '';
        const dataAvulsa = musEnsaioDataInp ? (musEnsaioDataInp.value || '') : '';

        // Recorrência
        const recAno = musRecAnoSel && musRecAnoSel.value ? parseInt(musRecAnoSel.value,10) : null;
        const recDia = musRecDiaSel && musRecDiaSel.value ? parseInt(musRecDiaSel.value,10) : null; // 0..6
        const recSemana = musRecSemanaSel && musRecSemanaSel.value ? musRecSemanaSel.value : null; // '1'..'5' ou 'last'
        const recMeses = musRecMesesWrap ? Array.from(musRecMesesWrap.querySelectorAll('input[name="mus-rec-meses"]:checked')).map(i=>parseInt(i.value,10)) : [];

        let datasGeradas = [];

        if (recAno && recDia!==null && recSemana && recMeses.length){
          datasGeradas = recMeses.map(m => {
            if(recSemana === 'last') return getLastWeekdayOfMonth(recAno, m, recDia);
            const nth = parseInt(recSemana,10);
            return getNthWeekdayOfMonth(recAno, m, recDia, nth);
          }).filter(Boolean);
          if (!datasGeradas.length){ toast('Nenhuma data encontrada para a recorrência selecionada', 'error'); return; }
        } else if (dataAvulsa){
          datasGeradas = [dataAvulsa];
        } else {
          toast('Informe data avulsa ou complete os campos de recorrência', 'error');
          return;
        }

        // Gravação: se em edição, atualiza apenas um; caso contrário, cria todos
        const editId = formEnsaioMusical.dataset.editId;
        if (editId){
          const evento = {
            tipo: 'Ensaio',
            ensaioTipo,
            congregacaoId,
            congregacaoNome,
            data: datasGeradas[0],
            hora,
            observacoes: obs
          };
          await update('eventos', editId, evento);
          delete formEnsaioMusical.dataset.editId;
          const submitBtn = formEnsaioMusical.querySelector('button[type="submit"]');
          const resetBtn = formEnsaioMusical.querySelector('button[type="reset"]');
          if(submitBtn) submitBtn.textContent = 'Salvar Ensaio';
          if(resetBtn) resetBtn.textContent = 'Limpar';
          toast('Ensaio atualizado');
        } else {
          await Promise.all(datasGeradas.map(data => {
            const evento = {
              tipo: 'Ensaio',
              ensaioTipo,
              congregacaoId,
              congregacaoNome,
              data,
              hora,
              observacoes: obs
            };
            return write('eventos', evento);
          }));
          toast(`${datasGeradas.length} ensaio(s) cadastrado(s)`);
        }

        // Re-renderiza agenda pós-salvar
        try{ renderMusAgenda(); }catch{}

        // Reset do formulário
        formEnsaioMusical.reset();
      }catch(err){
        console.error(err);
        toast('Falha ao salvar ensaio', 'error');
      }
    });
  }

  // Agenda de Ensaios (listar, filtrar e imprimir)
  function renderMusAgenda(){
    if(!musAgendaBody) return;
    const ySel = musAgendaYearSel && musAgendaYearSel.value ? parseInt(musAgendaYearSel.value,10) : null;
    const mSel = musAgendaMonthSel && musAgendaMonthSel.value ? parseInt(musAgendaMonthSel.value,10) : null;
    const ensaiosFiltrados = (eventosCache||[])
      .filter(ev => ev && ev.tipo==='Ensaio')
      .filter(ev=>{
        const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
        if(!d || isNaN(d.getTime())) return false;
        if(ySel && d.getFullYear()!==ySel) return false;
        if(mSel && (d.getMonth()+1)!==mSel) return false;
        return true;
      })
      .slice()
      .sort((a,b)=>{
        const ad = parseDateYmdLocal(a.data) || new Date(a.data);
        const bd = parseDateYmdLocal(b.data) || new Date(b.data);
        return ad - bd;
      });
    if(!ensaiosFiltrados.length){ musAgendaBody.innerHTML = '<tr><td colspan="5" class="text-muted">Nenhum ensaio cadastrado</td></tr>'; return; }

    const rows = ensaiosFiltrados.map(e => {
      const localNome = e.congregacaoNome || ((congregacoesByIdEvents && congregacoesByIdEvents[e.congregacaoId] && (congregacoesByIdEvents[e.congregacaoId].nome || '')) || '');
      return `<tr data-id="${e.id}">
        <td>${e.data || ''}</td>
        <td>${e.hora || '-'}</td>
        <td>${localNome}</td>
        <td>${e.ensaioTipo || '-'}</td>
        <td>
          <button class="btn btn-sm btn-outline-primary btn-edit" data-action="edit-ensaio" data-id="${e.id}">Editar</button>
          <button class="btn btn-sm btn-danger ms-2 btn-del" data-action="delete-ensaio" data-id="${e.id}">Excluir</button>
        </td>
      </tr>`;
    }).join('');

    musAgendaBody.innerHTML = rows;
  }
  if(musAgendaYearSel){ musAgendaYearSel.addEventListener('change', ()=>{ try{ renderMusAgenda(); }catch{} }); }
  if(musAgendaMonthSel){ musAgendaMonthSel.addEventListener('change', ()=>{ try{ renderMusAgenda(); }catch{} }); }
  // Atualizar agenda quando eventos mudam (listener dedicado)
  readList('eventos', list => { eventosCache = list; try{ renderMusAgenda(); }catch{} });
  // Ações de editar/excluir ensaio na agenda
  if(musAgendaBody){
    musAgendaBody.addEventListener('click', async (e)=>{
      const btnEdit = e.target.closest('button[data-action="edit-ensaio"]');
      const btnDel = e.target.closest('button[data-action="delete-ensaio"]');
      if(btnEdit){
        const id = btnEdit.getAttribute('data-id');
        const ev = (eventosCache||[]).find(x=>x.id===id);
        if(!ev || !formEnsaioMusical) return;
        try{
          if(musCongSel) musCongSel.value = ev.congregacaoId||'';
          if(musEnsaioTipoSel) musEnsaioTipoSel.value = ev.ensaioTipo||'';
          if(musEnsaioDataInp) musEnsaioDataInp.value = ev.data||'';
          if(musEnsaioHoraInp) musEnsaioHoraInp.value = ev.hora||'';
          if(musEnsaioObsTxt) musEnsaioObsTxt.value = ev.observacoes||'';
          formEnsaioMusical.dataset.editId = ev.id;
          const submitBtn = formEnsaioMusical.querySelector('button[type="submit"]');
          const resetBtn = formEnsaioMusical.querySelector('button[type="reset"]');
          if(submitBtn) submitBtn.textContent = 'Atualizar Ensaio';
          if(resetBtn) resetBtn.textContent = 'Cancelar Edição';
          formEnsaioMusical.scrollIntoView({ behavior:'smooth', block:'center' });
        }catch{}
        return;
      }
      if(btnDel){
        const id = btnDel.getAttribute('data-id');
        try{
          const ok = await remove('eventos', id);
          if(ok){ toast('Ensaio removido'); renderMusAgenda(); }
        }catch(err){ console.error(err); toast('Falha ao excluir ensaio', 'error'); }
        return;
      }
    });
  }

  // Listagem de Pessoas do Musical
  function renderMusPessoas(){
    if(!listaMusicalPessoas) return;
    readList('musicalPessoas', list => {
      const items = (list||[]).slice().sort((a,b)=> (a.nome||'').localeCompare(b.nome||'', undefined, { sensitivity:'base' }));
      if(!items.length){ listaMusicalPessoas.innerHTML = '<div class="text-muted">Nenhuma pessoa cadastrada</div>'; return; }
      listaMusicalPessoas.innerHTML = items.map(p => {
        const cat = p.categoria || '-';
        const inst = p.categoria==='Instrutor' ? ` • ${p.instrumento||'-'} (${p.tonalidade||'-'})` : '';
        return `<div class="list-item"><div class="list-item-info"><strong>${p.nome||''}</strong> — ${cat}${inst}</div></div>`;
      }).join('');
    });
  }


  const formMusPessoa = qs('#form-musical-pessoa');
  const musPessoaNomeInp = qs('#mus-pessoa-nome');
  const musPessoaCategoriaSel = qs('#mus-pessoa-categoria');
  const musPessoaCongSel = qs('#mus-pessoa-congregacao');
  const musInstrumentoWrap = qs('#mus-instrumento-wrap');
  const musTonalidadeWrap = qs('#mus-tonalidade-wrap');
  const musInstrumentoInp = qs('#mus-instrumento');
  const musTonalidadeInp = qs('#mus-tonalidade');
  const listaMusicalPessoas = qs('#lista-musical-pessoas');

  try{ renderMusPessoas(); }catch{}

  // Cadastro de Pessoas do Musical
  if(formMusPessoa){
    formMusPessoa.addEventListener('submit', async (e)=>{
      e.preventDefault();
      try{
        const nome = musPessoaNomeInp ? (musPessoaNomeInp.value||'') : '';
        const categoria = musPessoaCategoriaSel ? (musPessoaCategoriaSel.value||'') : '';
        const congregacaoId = musPessoaCongSel ? (musPessoaCongSel.value||'') : '';
        const instrumento = musInstrumentoInp ? (musInstrumentoInp.value||'') : '';
        const tonalidade = musTonalidadeInp ? (musTonalidadeInp.value||'') : '';
        if(!nome || !categoria){ toast('Preencha Nome e Categoria', 'error'); return; }
        const saved = await write('musicalPessoas', { nome, categoria, congregacaoId, instrumento, tonalidade });
        if(saved){ toast('Pessoa salva'); formMusPessoa.reset && formMusPessoa.reset(); try{ renderMusPessoas(); }catch{} }
      }catch(err){ console.error(err); toast('Falha ao salvar pessoa do musical', 'error'); }
    });
  }

  // Inicialização da aba Musical: preencher congregações e selects
  (function initMusical(){
    if (musCongSel || musPessoaCongSel) {
      readList('congregacoes', list => {
        // Atualiza caches globais usados em eventos e agenda musical
        congregacoesCacheEvents = (list || []).slice().sort((a, b) => (a && a.nome ? a.nome : '').localeCompare(b && b.nome ? b.nome : '', undefined, { sensitivity: 'base' }));
        congregacoesByIdEvents = {};
        congregacoesCacheEvents.forEach(c => { if (c && c.id) congregacoesByIdEvents[c.id] = c; });

        const optsHtml = ['<option value="">Selecione</option>']
          .concat(congregacoesCacheEvents.map(c => `<option value="${c.id}">${c.nome}</option>`))
          .join('');

        if (musCongSel) musCongSel.innerHTML = optsHtml;
        if (musPessoaCongSel) musPessoaCongSel.innerHTML = optsHtml;

        // Re-renderiza a agenda após o carregamento das congregações
        if (musAgendaBody) renderMusAgenda();
      });
    }
    if(musEnsaioTipoSel){ renderCustomEnsaioTypesMus(); setupMusEnsaioTipoAdder(); }
    if(musPessoaCategoriaSel){
      const toggleInstrutorFields = ()=>{
        const isInstrutor = (musPessoaCategoriaSel.value||'')==='Instrutor';
        musInstrumentoWrap && musInstrumentoWrap.classList.toggle('hidden', !isInstrutor);
        musTonalidadeWrap && musTonalidadeWrap.classList.toggle('hidden', !isInstrutor);
        if(!isInstrutor){ if(musInstrumentoInp) musInstrumentoInp.value=''; if(musTonalidadeInp) musTonalidadeInp.value=''; }
      };
      musPessoaCategoriaSel.addEventListener('change', toggleInstrutorFields);
      toggleInstrutorFields();
    }
    // Ano/Mês padrão na agenda
    if(musAgendaYearSel){
      const now = new Date(); const y = now.getFullYear();
      musAgendaYearSel.innerHTML = [y-1, y, y+1].map(yy=> `<option value="${yy}" ${yy===y?'selected':''}>${yy}</option>`).join('');
    }
    if(musAgendaMonthSel){
      const months = ['Todos','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
      const now = new Date(); const m = now.getMonth()+1;
      musAgendaMonthSel.innerHTML = months.map((label,i)=> i===0 ? `<option value="">${label}</option>` : `<option value="${i}" ${i===m?'selected':''}>${label}</option>`).join('');
    }

    // Preencher anos e meses (recorrência)
    if (musRecAnoSel){
      const now = new Date();
      const base = now.getFullYear();
      const anos = [base-1, base, base+1, base+2];
      musRecAnoSel.innerHTML = ['<option value="">Selecionar...</option>']
        .concat(anos.map(a => `<option value="${a}">${a}</option>`))
        .join('');
      musRecAnoSel.value = String(base);
    }
    if (musRecMesesWrap){
      const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
      musRecMesesWrap.innerHTML = meses.map((label, idx) => {
        const val = idx+1;
        return `<label class="form-check form-check-inline">
          <input type="checkbox" class="form-check-input" name="mus-rec-meses" value="${val}">
          <span class="form-check-label">${label}</span>
        </label>`;
      }).join('');
    }
  })();

  // Resultados: Santa Ceia e Batismos
  const resultadosForm = qs('#form-resultados');
  const resultadoTipoSel = qs('#resultado-tipo');
  const resultadoCongSel = qs('#resultado-congregacao');
  const resultadoDataInput = qs('#resultado-data');
  const resultadoAtendenteInput = qs('#resultado-atendente');
  const resultadoIrmaosInput = qs('#resultado-irmaos');
  const resultadoIrmasInput = qs('#resultado-irmas');
  const resultadoTotalInput = qs('#resultado-total');
  const resultadosListEl = qs('#resultados-list');
  let relEventosFilteredCache = null;
  async function applyRelFiltersFetch(){
    try{
      if(relLoading){ relLoading.classList.remove('d-none'); }
      if(btnRelApply){ btnRelApply.disabled = true; }
      if(!db){ if(typeof renderRelatorios==='function') renderRelatorios(); return; }
      const y = relYearSel && relYearSel.value ? parseInt(relYearSel.value,10) : null;
      const m = relMonthSel && relMonthSel.value ? parseInt(relMonthSel.value,10) : null;
      const cidade = relCidadeSel && relCidadeSel.value ? relCidadeSel.value : '';
      const tipo = relTipoSel && relTipoSel.value ? relTipoSel.value : '';
      let snap;
      if(cidade){
        snap = await db.ref('eventos').orderByChild('cidade').equalTo(cidade).once('value');
      } else if(tipo){
        snap = await db.ref('eventos').orderByChild('tipo').equalTo(tipo).once('value');
      } else if(y && m){
        const mm = `${m}`.padStart(2,'0');
        const daysInMonth = new Date(y, m, 0).getDate();
        const start = `${y}-${mm}-01`;
        const end = `${y}-${mm}-${String(daysInMonth).padStart(2,'0')}`;
        snap = await db.ref('eventos').orderByChild('data').startAt(start).endAt(end).once('value');
      } else if(y){
        const start = `${y}-01-01`;
        const end = `${y}-12-31`;
        snap = await db.ref('eventos').orderByChild('data').startAt(start).endAt(end).once('value');
      } else {
        snap = await db.ref('eventos').once('value');
      }
      const val = snap.val() || {};
      relEventosFilteredCache = Object.values(val).sort((a,b)=>{
        const ad = parseDateYmdLocal(a.data) || new Date(a.data);
        const bd = parseDateYmdLocal(b.data) || new Date(b.data);
        return ad - bd;
      });
      // Filtro adicional por bairro (cliente) se selecionado
      const bairroSel = relBairroSel && relBairroSel.value ? String(relBairroSel.value) : '';
      if(bairroSel){
        relEventosFilteredCache = (relEventosFilteredCache||[]).filter(ev => {
          try{
            const cong = congregacoesByIdEvents && congregacoesByIdEvents[ev.congregacaoId];
            return (cong && String(cong.bairro||'') === bairroSel);
          }catch{}
          return false;
        });
      }
    }catch(err){ console.error(err); toast('Falha ao buscar eventos no Firebase', 'error'); relEventosFilteredCache = null; }
    finally{
      if(relLoading){ relLoading.classList.add('d-none'); }
      if(btnRelApply){ btnRelApply.disabled = false; }
      try{ if(typeof renderRelatorios==='function') renderRelatorios(); }catch{}
    }
  }
  btnRelApply && btnRelApply.addEventListener('click', (e)=>{ e.preventDefault(); applyRelFiltersFetch(); });
btnRelClear && btnRelClear.addEventListener('click', ()=>{ relEventosFilteredCache = null; });
btnRelClear && btnRelClear.addEventListener('click', ()=>{ if(relYearSel) relYearSel.value=''; if(relMonthSel) relMonthSel.value=''; if(relCidadeSel) relCidadeSel.value=''; if(relBairroSel) relBairroSel.value=''; if(relEnsaiosSortSel) relEnsaiosSortSel.value='cidade'; if(relEnsaiosNextOnly) relEnsaiosNextOnly.checked=false; if(relTipoSel) relTipoSel.value=''; if(typeof renderRelatorios==='function') renderRelatorios(); });

  function toggleAtendenteManual(enable){
    const on = !!enable;
    if(atendenteManualInput){
      atendenteManualInput.classList.toggle('hidden', !on);
      atendenteManualInput.required = on;
      if(!on){ atendenteManualInput.value = ''; }
    }
    if(eventoAtendente){
      eventoAtendente.required = !on;
      if(on){ eventoAtendente.value = ''; }
    }
    if(outraRegiaoBtn){
      outraRegiaoBtn.textContent = on ? 'Usar lista do Ministério' : 'Irmão de outra região';
    }
  }

  if(outraRegiaoBtn){
    outraRegiaoBtn.addEventListener('click', ()=>{
      const hidden = atendenteManualInput ? atendenteManualInput.classList.contains('hidden') : true;
      toggleAtendenteManual(hidden);
    });
  }
  // Toggle de subtipo de Ensaio conforme o tipo selecionado
  function toggleEnsaioSubtype(){
    const isEnsaio = eventoTipoSel && (eventoTipoSel.value === 'Ensaio');
    if(eventoEnsaioTipoWrap){
      eventoEnsaioTipoWrap.classList.toggle('hidden', !isEnsaio);
    }
    if(eventoEnsaioTipoSel){
      eventoEnsaioTipoSel.required = !!isEnsaio;
      if(!isEnsaio){ eventoEnsaioTipoSel.value = ''; }
    }
  }
  if(eventoTipoSel){
    eventoTipoSel.addEventListener('change', toggleEnsaioSubtype);
    // inicializa estado
    toggleEnsaioSubtype();
  }
  function renderEventoCultosPreview(congId){
    if(!eventoCultosBody) return;
    if(!congId){
      eventoCultosBody.innerHTML = '<tr><td colspan="3" class="text-muted">Selecione uma congregação</td></tr>';
      return;
    }
    const c = congregacoesByIdEvents[congId];
    const cultos = (c && c.cultos) ? c.cultos : [];
    if(!cultos.length){
      eventoCultosBody.innerHTML = '<tr><td colspan="3" class="text-muted">Sem cultos cadastrados para esta congregação</td></tr>';
      return;
    }
    eventoCultosBody.innerHTML = cultos.map(ct => {
      return `<tr>
        <td>${ct.tipo}</td>
        <td>${ct.dia}</td>
        <td>${ct.horario}</td>
      </tr>`;
    }).join('');
  }
  const minCongSel = qs('#min-congregacao');
  if(formEvento){
    formEvento.addEventListener('submit', async (e)=>{
      e.preventDefault();
      const fd = new FormData(formEvento);
      const data = Object.fromEntries(fd.entries());
      const editId = formEvento.dataset.editId;
      const manualNome = (data.atendenteNomeManual||'').trim();
      if(!data.tipo || !data.data || !data.congregacaoId || (!data.atendenteId && !manualNome)){
        toast('Preencha tipo, data, atendente e congregação', 'error');
        return;
      }
      if(data.tipo === 'Ensaio'){
        const st = (data.ensaioTipo||'').trim();
        if(!st){ toast('Selecione o tipo de Ensaio', 'error'); return; }
      }
      // Nova regra: permitir coletas por tipo em todas as congregações;
      // bloquear apenas duplicidade de CO na MESMA congregação e mês.
      try{
        const newDate = parseDateYmdLocal(data.data);
        // Bloqueio: se já houver qualquer atendimento nesta congregação nesta data
        const sameDaySameCong = (eventosCache||[]).some(ev => {
          if(editId && ev.id === editId) return false;
          return ev.congregacaoId === data.congregacaoId && ev.data === data.data;
        });
        if(sameDaySameCong){
          toast('ALERTA: já existe atendimento nesta congregação nesta data.', 'error');
          return;
        }
        // Bloqueio: RJM somente se a congregação possuir RJM cadastrado
        if(data.tipo === 'RJM com Reforço de Coletas'){
          const cong = congregacoesByIdEvents[data.congregacaoId];
          const hasRjm = cong && Array.isArray(cong.cultos) && cong.cultos.some(ct => ct.tipo === 'RJM');
          if(!hasRjm){
            toast('Esta congregação não tem RJM.', 'error');
            return;
          }
        }
        const sameMonthSameCong = (eventosCache||[]).filter(ev => {
          if(editId && ev.id === editId) return false;
          if(ev.congregacaoId !== data.congregacaoId) return false;
          const d = parseDateYmdLocal(ev.data);
          return d && newDate && d.getFullYear() === newDate.getFullYear() && d.getMonth() === newDate.getMonth();
        });
        if(data.tipo === 'Culto Reforço de Coletas'){
          const cSel = congregacoesByIdEvents[data.congregacaoId];
          const nomeFmt = ((cSel && cSel.nomeFormatado)||'').toLowerCase();
          const cidade = ((cSel && cSel.cidade)||'').toLowerCase();
          const bairro = ((cSel && cSel.bairro)||'').toLowerCase();
          const isItuiCentro = (cidade === 'ituiutaba' && bairro === 'centro') || (nomeFmt.includes('ituiutaba') && nomeFmt.includes('centro'));
          const coList = sameMonthSameCong.filter(ev => ev.tipo === 'Culto Reforço de Coletas');
          if(isItuiCentro){
            if(coList.length >= 2){
              toast('Ituiutaba Centro: máximo de 2 coletas em Culto Oficial no mês.', 'error');
              return;
            }
            const isThu = newDate && newDate.getDay() === 4;
            if(!isThu){
              const nonThuAlready = coList.some(ev=>{
                const dEv = parseDateYmdLocal(ev.data);
                return dEv && dEv.getDay() !== 4;
              });
              if(nonThuAlready){
                toast('Ituiutaba Centro: dos 2 reforços, pelo menos 1 deve ser na quinta-feira (Culto Oficial).', 'error');
                return;
              }
            }
          } else {
            if(coList.length >= 1){
              toast('Já existe coleta em Culto Oficial nesta congregação neste mês.', 'error');
              return;
            }
          }
        }
        // RJM liberado por mês em todas as congregações (sem bloqueio por congregação)
      }catch{}
      // Encontrar nome do atendente selecionado (denormalização para exibir na lista)
      let atendenteNome = '';
      if(eventoAtendente){
        const opt = eventoAtendente.querySelector(`option[value="${data.atendenteId}"]`);
        atendenteNome = opt ? opt.textContent : '';
      }
      if(manualNome){
        atendenteNome = manualNome;
        data.atendenteId = '';
      }
      // Encontrar nome formatado da congregação selecionada (denormalização)
      let congregacaoNome = '';
      if(eventoCong){
        const congOpt = eventoCong.querySelector(`option[value="${data.congregacaoId}"]`);
        congregacaoNome = congOpt ? congOpt.textContent : '';
      }
      // Cidade da congregação (denormalizado)
      const cidadeEvento = (congregacoesByIdEvents && congregacoesByIdEvents[data.congregacaoId] && congregacoesByIdEvents[data.congregacaoId].cidade) ? (congregacoesByIdEvents[data.congregacaoId].cidade||'') : '';
      try{
        if(editId){
          const res = await update('eventos', editId, {
            tipo: data.tipo,
            data: data.data,
            congregacaoId: data.congregacaoId,
            congregacaoNome: congregacaoNome,
            cidade: cidadeEvento,
            atendenteId: data.atendenteId,
            atendenteNome: atendenteNome,
            ensaioTipo: data.ensaioTipo||'',
            observacoes: data.observacoes||''
          });
          if(res){ toast('Atendimento atualizado'); formEvento.reset(); }
        } else {
          const saved = await write('eventos', {
            tipo: data.tipo,
            data: data.data,
            congregacaoId: data.congregacaoId,
            congregacaoNome: congregacaoNome,
            cidade: cidadeEvento,
            atendenteId: data.atendenteId,
            atendenteNome: atendenteNome,
            ensaioTipo: data.ensaioTipo||'',
            observacoes: data.observacoes||''
          });
          if(saved){ toast('Atendimento salvo'); formEvento.reset(); }
        }
      }catch(err){
        console.error(err);
        const msg = (err && (err.code||err.message)) || 'Falha ao salvar atendimento';
        toast(msg, 'error');
      }
    });
    formEvento.addEventListener('reset', ()=>{
      delete formEvento.dataset.editId;
      const submitBtn = formEvento.querySelector('button[type="submit"]');
      const resetBtn = formEvento.querySelector('button[type="reset"]');
      if(submitBtn) submitBtn.textContent = 'Salvar Reforço de Coletas';
      if(resetBtn) resetBtn.textContent = 'Limpar';
    });
    formEvento.addEventListener('reset', ()=>{ toggleAtendenteManual(false); });

  readList('eventos', list => {
      eventosCache = list;
      if(relEventosEl || btnRelPrint || relYearSel || relCidadeSel || relTipoSel){
        try{ if(typeof initRelatoriosFilters==='function') initRelatoriosFilters(); }catch{}
        try{ if(typeof renderRelatorios==='function') renderRelatorios(); }catch{}
      }
      if(!listaEventos) return;
      // Agrupar por congregação e renderizar somente os nomes (apenas reforços de coletas)
      const byCong = {};
      (list||[]).forEach(ev => {
        if(!ev || !ev.congregacaoId) return;
        const isReforco = ev.tipo === 'Culto Reforço de Coletas' || ev.tipo === 'RJM com Reforço de Coletas';
        if(!isReforco) return;
        (byCong[ev.congregacaoId] = byCong[ev.congregacaoId] || []).push(ev);
      });
      const items = Object.keys(byCong).map(congId => {
        const arr = byCong[congId] || [];
        // Label da congregação
        const sample = arr[0];
        const congLabel = labelCong(sample || { congregacaoId: congId, congregacaoNome: '' });
        return { congId, congLabel, events: arr };
      }).sort((a,b)=> a.congLabel.localeCompare(b.congLabel));
      if(!items.length){
        listaEventos.innerHTML = '<div class="list-item"><div class="list-item-info"><span class="text-muted">Nenhum reforço de coletas cadastrado</span></div></div>';
      } else {
        listaEventos.innerHTML = items.map(item => {
          return `
          <div class="list-item" data-congid="${item.congId}">
            <div class="list-item-info">
              <button class="congregacao-link" data-congid="${item.congId}" style="background:none;border:none;padding:0;font-size:1.05rem;color:var(--primary-color);font-weight:600;cursor:pointer;">${item.congLabel}</button>
              <div class="list-details hidden" aria-live="polite"></div>
            </div>
          </div>
          `;
        }).join('');
      }
      renderTabelaReforcos();
      // Atualiza ticker (próximo mês)
      try{ renderTickerSemReforco(); }catch{}
    });
  // Inicializar selects de Ano/Mês na Index
  (function initIndexMonthYear(){
    try{
      const now = new Date();
      const currentYear = now.getFullYear();
      const currentMonth = now.getMonth()+1; // 1-12
      const selectedMonth = currentMonth === 12 ? 1 : (currentMonth + 1); // mês subsequente
      const selectedYear = currentMonth === 12 ? (currentYear + 1) : currentYear; // ajusta ano se dezembro

      if(indexYearSel){
        const years = [currentYear-1, currentYear, currentYear+1];
        indexYearSel.innerHTML = years.map(y=>`<option value="${y}" ${y===selectedYear?'selected':''}>${y}</option>`).join('');
        indexYearSel.addEventListener('change', ()=>{ try{ renderTabelaReforcos(); }catch{} });
      }
      if(indexMonthSel){
        const months = ['Todos','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
        indexMonthSel.innerHTML = months.map((label, i)=>{
          if(i===0) return `<option value="">${label}</option>`; // Todos
          return `<option value="${i}" ${i===selectedMonth?'selected':''}>${label}</option>`;
        }).join('');
        indexMonthSel.addEventListener('change', ()=>{ try{ renderTabelaReforcos(); }catch{} });
      }
      // Re-render com filtro padrão do mês subsequente
      try{ renderTabelaReforcos(); }catch{}
    }catch{}
  }());
  if(listaEventos){
    listaEventos.addEventListener('click', async (e)=>{
        const congBtn = e.target.closest('.congregacao-link');

        // Helper para renderizar todos os reforços (CO/RJM) da congregação
        const renderCongReforcosDetails = (congId, details) => {
          const diasSemanaPt = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
          const getTime = (ev) => {
            try{
              const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
              return d ? d.getTime() : 0;
            }catch{ return 0; }
          };
          const arr = (eventosCache||[])
            .filter(ev => ev && ev.congregacaoId === congId && (ev.tipo === 'Culto Reforço de Coletas' || ev.tipo === 'RJM com Reforço de Coletas'))
            .sort((a,b)=> getTime(a) - getTime(b));
          if(!arr.length){
            details.innerHTML = '<div class="meta text-muted">Nenhum reforço encontrado para esta congregação.</div>';
            details.classList.remove('hidden');
            return;
          }
          const rows = arr.map(ev => {
            const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
            const diaNome = diasSemanaPt[d.getDay()];
            const dataFmt = `${diaNome} - ${formatDate(ev.data)}`;
            const hora = horaDoEvento(ev);
            const atendente = ev.atendenteNome || '-';
            const tipo = ev.tipo + (ev.ensaioTipo ? ` - ${ev.ensaioTipo}` : '');
            return `
              <tr>
                <td>${dataFmt}</td>
                <td>${hora||'-'}</td>
                <td>${atendente}${ev.atendenteOutraRegiao ? ' <span class="flag-outra-regiao" title="Irmão de outra região"><svg viewBox="0 0 24 24"><path d="M12 2c-4.4 0-8 3.1-8 7 0 5 8 13 8 13s8-8 8-13c0-3.9-3.6-7-8-7zm0 9a2 2 0 1 1 0-4 2 2 0 0 1 0 4z" fill="currentColor"/></svg></span>' : ''}</td>
                <td>${tipo}</td>
                <td>
                  <button class="btn btn-sm btn-outline-primary" data-action="edit-ev" data-id="${ev.id}">Editar</button>
                  <button class="btn btn-sm btn-danger ms-2" data-action="delete-ev" data-id="${ev.id}">Excluir</button>
                </td>
              </tr>
            `;
          }).join('');
          details.innerHTML = `
            <table class="table table-sm mt-2">
              <thead>
                <tr>
                  <th>Data</th>
                  <th>Hora</th>
                  <th>Quem atende</th>
                  <th>Tipo</th>
                  <th>Ações</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          `;
          details.classList.remove('hidden');
        };

        if(congBtn){
          const congId = congBtn.getAttribute('data-congid');
          const container = congBtn.closest('.list-item');
          const details = container ? container.querySelector('.list-details') : null;
          if(!details){ return; }
          if(details.classList.contains('hidden')){
            renderCongReforcosDetails(congId, details);
          } else {
            details.classList.add('hidden');
          }
          return;
        }
        const btnEdit = e.target.closest('button[data-action="edit-ev"]');
        const btnDel = e.target.closest('button[data-action="delete-ev"]');
        const btnConfirmDel = e.target.closest('button[data-action="confirm-delete"]');
        const btnCancelDel = e.target.closest('button[data-action="cancel-delete"]');
        if(btnEdit){
          const id = btnEdit.getAttribute('data-id');
          const ev = (eventosCache||[]).find(x=>x.id===id);
          if(!ev || !formEvento) return;
          formEvento.dataset.editId = ev.id;
          const submitBtn = formEvento.querySelector('button[type="submit"]');
          const resetBtn = formEvento.querySelector('button[type="reset"]');
          if(submitBtn) submitBtn.textContent = 'Atualizar Reforço de Coletas';
          if(resetBtn) resetBtn.textContent = 'Cancelar Edição';
          const tipoSel = formEvento.querySelector('select[name="tipo"]');
          const dataInp = formEvento.querySelector('input[name="data"]');
          const obsTxt = formEvento.querySelector('textarea[name="observacoes"]');
          if(tipoSel) tipoSel.value = ev.tipo||'';
          // Ajusta subtipo de Ensaio ao editar
          if(eventoTipoSel){ toggleEnsaioSubtype(); }
          if(ev.tipo === 'Ensaio' && eventoEnsaioTipoSel){ eventoEnsaioTipoSel.value = ev.ensaioTipo||''; }
          if(dataInp) dataInp.value = ev.data||'';
          if(eventoCong) { eventoCong.value = ev.congregacaoId||''; renderEventoCultosPreview(ev.congregacaoId||''); }
          if(eventoAtendente) eventoAtendente.value = ev.atendenteId||'';
          if((!ev.atendenteId || ev.atendenteId==='') && ev.atendenteNome){
            toggleAtendenteManual(true);
            if(atendenteManualInput) atendenteManualInput.value = ev.atendenteNome;
          } else {
            toggleAtendenteManual(false);
          }
          if(obsTxt) obsTxt.value = ev.observacoes||'';

          // Mostrar o formulário sem precisar rolar manualmente: centraliza, foca e destaca
          try {
            // Garante que a aba de Reforço esteja visível (neste arquivo já é a ativa)
            const firstField = dataInp || formEvento.querySelector('input, select, textarea');
            formEvento.classList.add('form-highlight');
            formEvento.scrollIntoView({ behavior: 'smooth', block: 'center' });
            setTimeout(() => {
              if(firstField && typeof firstField.focus === 'function') { firstField.focus(); }
            }, 250);
            const handleFormGlowEnd = () => {
              formEvento.classList.remove('form-highlight');
              formEvento.removeEventListener('animationend', handleFormGlowEnd);
            };
            formEvento.addEventListener('animationend', handleFormGlowEnd);
          } catch {}
        }
        if(btnDel){
          const id = btnDel.getAttribute('data-id');
          const container = btnDel.closest('.list-item');
          const details = container ? container.querySelector('.list-details') : null;
          if(!details) return;
          // Inserir alerta vermelho de confirmação
          const existing = details.querySelector('.delete-confirm');
          if(existing){ existing.remove(); }
          const div = document.createElement('div');
          div.className = 'delete-confirm alert alert-danger mt-2';
          div.innerHTML = `
            <div><strong>Confirma a exclusão deste reforço?</strong></div>
            <div class="mt-2">
              <button class="btn btn-sm btn-danger" data-action="confirm-delete" data-id="${id}">Confirmar Exclusão</button>
              <button class="btn btn-sm btn-outline-secondary ms-2" data-action="cancel-delete">Cancelar</button>
            </div>
          `;
          details.appendChild(div);
          details.classList.remove('hidden');
          return;
        }
        if(btnConfirmDel){
          const id = btnConfirmDel.getAttribute('data-id');
          try{
            const removed = await remove('eventos', id);
            if(removed){
              toast('Reforço removido');
              // Atualiza cache e re-renderiza detalhes da congregação
              eventosCache = (eventosCache||[]).filter(x=>x.id===undefined || x.id!==id);
              try{
                const container = e.target.closest('.list-item');
                const congId = container ? (container.getAttribute('data-congid') || (container.querySelector('.congregacao-link') && container.querySelector('.congregacao-link').getAttribute('data-congid')) || '') : '';
                const details = container ? container.querySelector('.list-details') : null;
                if(congId && details){ renderCongReforcosDetails(congId, details); }
                renderTabelaReforcos();
              }catch{}
            }
          }catch(err){
            console.error(err);
            const msg = (err && (err.code||err.message)) || 'Falha ao excluir atendimento';
            toast(msg, 'error');
          }
          return;
        }
        if(btnCancelDel){
          const confirmEl = e.target.closest('.delete-confirm');
          if(confirmEl) confirmEl.remove();
          return;
        }
      });
    }
  }

  // Geração de arquivo HTML para impressão dos eventos
  if(btnImprimirEventos){
    btnImprimirEventos.addEventListener('click', ()=>{
      try{
        // Somente reforços (CO/RJM)
        const reforcos = (eventosCache||[])
          .filter(ev => ev.tipo==='Culto Reforço de Coletas' || ev.tipo==='RJM com Reforço de Coletas')
          .slice()
          .sort((a,b)=>{
            const ad = parseDateYmdLocal(a.data) || new Date(a.data);
            const bd = parseDateYmdLocal(b.data) || new Date(b.data);
            return ad - bd;
          });
        if(!reforcos.length){
          toast('Nenhum reforço cadastrado para imprimir', 'error');
          return;
        }
        const diasSemanaPt = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
        const rowsHtml = reforcos.map(ev => {
          const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
          const diaNome = diasSemanaPt[d.getDay()];
          const cong = congregacoesByIdEvents[ev.congregacaoId];
          const localLabel = ev.congregacaoNome || (cong ? (cong.nomeFormatado || (cong.cidade && cong.bairro ? `${cong.cidade} - ${cong.bairro}` : (cong.nome||ev.congregacaoId))) : ev.congregacaoId);
          const tipoCulto = ev.tipo==='Culto Reforço de Coletas' ? 'Culto Oficial' : (ev.tipo==='RJM com Reforço de Coletas' ? 'RJM' : '');
          let hora = '-';
          if(cong && Array.isArray(cong.cultos)){
            const match = cong.cultos.find(ct => ct.tipo===tipoCulto && ct.dia===diaNome);
            if(match && match.horario) hora = match.horario;
          }
          return `<tr>
            <td>${diaNome} - ${formatDate(ev.data)}</td>
            <td>${hora}</td>
            <td>${localLabel}</td>
            <td>${ev.atendenteNome||'-'}${ev.atendenteOutraRegiao ? ' <span class=\"flag-outra-regiao\" title=\"Irmão de outra região\"><svg viewBox=\"0 0 24 24\"><path d=\"M12 2c-4.4 0-8 3.1-8 7 0 5 8 13 8 13s8-8 8-13c0-3.9-3.6-7-8-7zm0 9a2 2 0 1 1 0-4 2 2 0 0 1 0 4z\" fill=\"currentColor\"/></svg></span>' : ''}</td>
            <td>${ev.tipo}${ev.ensaioTipo?` - ${ev.ensaioTipo}`:''}</td>
          </tr>`;
        }).join('');

        const now = new Date();
        const dd = String(now.getDate()).padStart(2,'0');
        const mm = String(now.getMonth()+1).padStart(2,'0');
        const yyyy = now.getFullYear();
        const hh = String(now.getHours()).padStart(2,'0');
        const mi = String(now.getMinutes()).padStart(2,'0');
        const geradoEm = `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
        const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <title>Atendimentos Cadastrados</title>
  <style>
    body{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; color:#000; margin:24px; }
    h1{ font-size: 18px; margin: 0 0 8px; }
    p{ font-size: 12px; margin: 0 0 12px; color:#333; }
    table{ width:100%; border-collapse: collapse; }
    th, td{ border:1px solid #000; padding:6px; font-size:12px; }
    th{ background:#f2f2f2; }
    @media print{ .no-print{ display:none; } }
  </style>
  </head>
<body>
  <h1>Atendimentos Cadastrados</h1>
  <p>Gerado em: ${geradoEm}</p>
  <table>
    <thead>
      <tr>
        <th>Data</th>
        <th>Hora</th>
        <th>Local</th>
        <th>Quem atende</th>
        <th>Tipo</th>
      </tr>
    </thead>
    <tbody>
      ${rowsHtml}
    </tbody>
  </table>
</body>
</html>`;
        const blob = new Blob([html], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'atendimentos-cadastrados.html';
        document.body.appendChild(a);
        a.click();
        setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
        toast('Arquivo gerado para impressão');
      } catch (err) {
        console.error(err);
        toast('Falha ao gerar arquivo de impressão', 'error');
      }
    });
  }
})();