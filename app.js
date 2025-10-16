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
    'ver-relatorios': () => setActiveTab('relatorios'),
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
    {id:'cad-evento', label:'Cadastrar Evento'},
  ];
  const minMenu = [
    {id:'cad-ministerio', label:'Cadastrar Irmão do Ministério'},
  ];
  const congMenu = [
    {id:'cad-congregacao', label:'Cadastrar Congregação'},
  ];
  const relMenu = [
    {id:'ver-relatorios', label:'Ver Relatórios'},
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
  function readList(path, cb){
    if(!db) return;
    db.ref(path).on('value', snap => {
      const val = snap.val() || {};
      const list = Object.values(val).sort((a,b)=>a.createdAt-b.createdAt);
      cb(list);
    });
  }

  // Eventos
  const formEvento = qs('#form-evento');
  const listaEventos = qs('#lista-eventos');
  const eventoCong = qs('#evento-congregacao');
  const eventoAtendente = qs('#evento-atendente');
  const minCongSel = qs('#min-congregacao');
  if(formEvento){
    formEvento.addEventListener('submit', async (e)=>{
      e.preventDefault();
      const fd = new FormData(formEvento);
      const data = Object.fromEntries(fd.entries());
      if(!data.tipo || !data.data || !data.congregacaoId || !data.atendenteId){
        toast('Preencha tipo, data, atendente e congregação', 'error');
        return;
      }
      // Encontrar nome do atendente selecionado (denormalização para exibir na lista)
      let atendenteNome = '';
      if(eventoAtendente){
        const opt = eventoAtendente.querySelector(`option[value="${data.atendenteId}"]`);
        atendenteNome = opt ? opt.textContent : '';
      }
      try{
        const saved = await write('eventos', {
          tipo: data.tipo,
          data: data.data,
          congregacaoId: data.congregacaoId,
          atendenteId: data.atendenteId,
          atendenteNome: atendenteNome,
          observacoes: data.observacoes||''
        });
        if(saved){ toast('Evento salvo'); formEvento.reset(); }
      }catch(err){
        console.error(err);
        const msg = (err && (err.code||err.message)) || 'Falha ao salvar evento';
        toast(msg, 'error');
      }
    });

    readList('eventos', list => {
      if(!listaEventos) return;
      listaEventos.innerHTML = list.map(ev => `
        <div class="item">
          <div>
            <strong>${ev.tipo}</strong>
            <div class="meta">Data: ${formatDate(ev.data)}</div>
            <div class="meta">Congregação: ${ev.congregacaoId}</div>
            <div class="meta">Atendido por: ${ev.atendenteNome||'-'}</div>
            <div class="meta">Opções: clique com botão direito na aba para mais ações</div>
            ${ev.observacoes?`<div class="meta">Obs: ${ev.observacoes}</div>`:''}
          </div>
        </div>
      `).join('');
    });
  }

  // Popular congregações no select de evento
  function fillSelect(sel, list, labelKey){
    if(!sel) return;
    sel.innerHTML = '<option value="">Selecionar...</option>' +
      list.map(x => {
        const label = x[labelKey] || x.nomeFormatado || (x.cidade && x.bairro ? `${x.cidade} - ${x.bairro}` : (x.nome||''));
        return `<option value="${x.id}">${label}</option>`;
      }).join('');
  }
  if(eventoCong){
    readList('congregacoes', list => {
      fillSelect(eventoCong, list, 'nomeFormatado');
    });
  }
  if(minCongSel){
    readList('congregacoes', list => {
      fillSelect(minCongSel, list, 'nomeFormatado');
    });
  }

  // Popular atendentes (Ministério) no select de evento
  if(eventoAtendente){
    readList('ministerio', list => {
      // Ordenar por nome
      list.sort((a,b)=> (a.nome||'').localeCompare(b.nome||''));
      fillSelect(eventoAtendente, list, 'nome');
    });
  }

  function formatDate(str){
    try{ const d = new Date(str); return d.toLocaleDateString('pt-BR'); }
    catch{ return str; }
  }

  const formCong = qs('#form-congregacao');
  const listaCong = qs('#lista-congregacoes');
  const congAnciaoSel = qs('#cong-anciao');
  const congDiaconoSel = qs('#cong-diacono');
  const badgeCO = qs('#badge-cooperador-oficial');
  const badgeCJ = qs('#badge-cooperador-jovens');
  const listaMin = qs('#lista-ministerio');
  let congLabelById = {};
  // Cache de rótulos de congregações para exibir nos listados
  readList('congregacoes', list => {
    congLabelById = {};
    list.forEach(c => {
      const label = c.nomeFormatado || (c.cidade && c.bairro ? `${c.cidade} - ${c.bairro}` : (c.nome||''));
      congLabelById[c.id] = label;
    });
  });

  // Popular selects de Ancião/Diácono com base no Ministério
  function populateMinistrySelect(sel, list, filterFn){
    if(!sel) return;
    const filtered = list.filter(filterFn);
    sel.innerHTML = '<option value="">Selecionar...</option>' + filtered.map(m=>`<option value="${m.id}">${m.nome}</option>`).join('');
  }

  let ministryCache = [];
  readList('ministerio', list => {
    ministryCache = list;
    // Atualizar contadores de marcações especiais (para badges)
    const counts = {
      cooperadorOficial: list.filter(m=>m.funcao==='Cooperador Oficial').length,
      cooperadorJovens: list.filter(m=>m.funcao==='Cooperador de Jovens').length,
    };
    if(badgeCO){ badgeCO.textContent = counts.cooperadorOficial; badgeCO.classList.toggle('hidden', counts.cooperadorOficial===0); }
    if(badgeCJ){ badgeCJ.textContent = counts.cooperadorJovens; badgeCJ.classList.toggle('hidden', counts.cooperadorJovens===0); }

    populateMinistrySelect(congAnciaoSel, list, m=>m.funcao==='Ancião');
    populateMinistrySelect(congDiaconoSel, list, m=>m.funcao==='Diácono');

    // Listar Ministério
    if(listaMin){
      // Ordenar por nome
      const sorted = [...list].sort((a,b)=> (a.nome||'').localeCompare(b.nome||''));
      listaMin.innerHTML = sorted.map(m => {
        const congLabel = m.congregacaoId ? (congLabelById[m.congregacaoId] || m.congregacaoId) : '';
        return `
          <div class="item">
            <div>
              <strong>${m.nome}</strong>
              <div class="meta">Função: ${m.funcao||'-'}</div>
              ${congLabel?`<div class="meta">Congregação: ${congLabel}</div>`:''}
              ${m.telefone?`<div class="meta">Tel: ${m.telefone}</div>`:''}
              ${m.email?`<div class="meta">E-mail: ${m.email}</div>`:''}
            </div>
          </div>
        `;
      }).join('');
    }
  });

  // Salvar Congregação
  if(formCong){
    formCong.addEventListener('submit', async (e)=>{
      e.preventDefault();
      const fd = new FormData(formCong);
      const data = Object.fromEntries(fd.entries());
      if(!data.cidade || !data.bairro || !data.endereco){ toast('Preencha cidade, bairro e endereço', 'error'); return; }
      const anciaoId = data.anciaoId||''; const anciaoTipo = data.anciaoTipo||'';
      const diaconoId = data.diaconoId||''; const diaconoTipo = data.diaconoTipo||'';

      const anciaoNome = ministryCache.find(m=>m.id===anciaoId)?.nome || '';
      const diaconoNome = ministryCache.find(m=>m.id===diaconoId)?.nome || '';

      try{
        const saved = await write('congregacoes', {
          cidade: data.cidade,
          bairro: data.bairro,
          endereco: data.endereco,
          anciaoId, anciaoTipo, anciaoNome,
          diaconoId, diaconoTipo, diaconoNome,
        });
        if(saved){ toast('Congregação salva'); formCong.reset(); }
      }catch(err){
        console.error(err);
        const msg = (err && (err.code||err.message)) || 'Falha ao salvar congregação';
        toast(msg, 'error');
      }
    });

    readList('congregacoes', list => {
      if(!listaCong) return;
      listaCong.innerHTML = list.map(c => `
        <div class="item">
          <div>
            <strong>${c.cidade} - ${c.bairro}</strong>
            <div class="meta">Endereço: ${c.endereco}</div>
            <div class="meta">Ancião: ${c.anciaoNome||'-'} ${c.anciaoTipo?`(${c.anciaoTipo})`:''}</div>
            <div class="meta">Diácono: ${c.diaconoNome||'-'} ${c.diaconoTipo?`(${c.diaconoTipo})`:''}</div>
          </div>
        </div>
      `).join('');
    });
  }

  // Salvar Ministério
  const formMin = qs('#form-ministerio');
  if(formMin){
    formMin.addEventListener('submit', async (e)=>{
      e.preventDefault();
      const fd = new FormData(formMin);
      const data = Object.fromEntries(fd.entries());
      if(!data.nome || !data.funcao){ toast('Preencha nome e função', 'error'); return; }
      const payload = {
        nome: data.nome,
        funcao: data.funcao,
        congregacaoId: data.congregacaoId||'',
        telefone: data.telefone||'',
        email: data.email||'',
        anciaoLocal: !!data.anciaoLocal,
        anciaoResponsavel: !!data.anciaoResponsavel,
        diaconoLocal: !!data.diaconoLocal,
        diaconoResponsavel: !!data.diaconoResponsavel,
      };
      try{
        const saved = await write('ministerio', payload);
        if(saved){ toast('Irmão do Ministério salvo'); formMin.reset(); }
      }catch(err){
        console.error(err);
        const msg = (err && (err.code||err.message)) || 'Falha ao salvar Ministério';
        toast(msg, 'error');
      }
    });
  }
}());
