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
  const outraRegiaoBtn = qs('#atendente-outra-regiao-btn');
  const atendenteManualInput = qs('#evento-atendente-manual');
  const eventoCultosBody = qs('#evento-cultos-body');
  const tabelaReforcosBody = qs('#tabela-reforcos-body');
  const btnImprimirEventos = qs('#btn-imprimir-eventos');
  // Relatórios
  const relEventosEl = qs('#relatorio-eventos');
  const relMinisterioEl = qs('#relatorio-ministerio');
  const relCongregacoesEl = qs('#relatorio-congregacoes');
  let eventosCache = [];
  let congregacoesCacheEvents = [];
  let congregacoesByIdEvents = {};

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
      // Nova regra: permitir coletas por tipo em todas as congregações;
      // bloquear apenas duplicidade de CO na MESMA congregação e mês.
      try{
        const newDate = parseDateYmdLocal(data.data);
        const sameMonthSameCong = (eventosCache||[]).filter(ev => {
          if(editId && ev.id === editId) return false;
          if(ev.congregacaoId !== data.congregacaoId) return false;
          const d = parseDateYmdLocal(ev.data);
          return d && newDate && d.getFullYear() === newDate.getFullYear() && d.getMonth() === newDate.getMonth();
        });
        if(data.tipo === 'Culto Reforço de Coletas'){
          const existsCO = sameMonthSameCong.some(ev => ev.tipo === 'Culto Reforço de Coletas');
          if(existsCO){
            toast('Já existe coleta em Culto Oficial nesta congregação neste mês.', 'error');
            return;
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
      try{
        if(editId){
          const res = await update('eventos', editId, {
            tipo: data.tipo,
            data: data.data,
            congregacaoId: data.congregacaoId,
            congregacaoNome: congregacaoNome,
            atendenteId: data.atendenteId,
            atendenteNome: atendenteNome,
            observacoes: data.observacoes||''
          });
          if(res){ toast('Atendimento atualizado'); formEvento.reset(); }
        } else {
          const saved = await write('eventos', {
            tipo: data.tipo,
            data: data.data,
            congregacaoId: data.congregacaoId,
            congregacaoNome: congregacaoNome,
            atendenteId: data.atendenteId,
            atendenteNome: atendenteNome,
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
      if(submitBtn) submitBtn.textContent = 'Salvar Atendimento';
      if(resetBtn) resetBtn.textContent = 'Limpar';
    });
    formEvento.addEventListener('reset', ()=>{ toggleAtendenteManual(false); });

    readList('eventos', list => {
      eventosCache = list;
      if(!listaEventos) return;
      const sorted = [...list].sort((a,b)=>{
        const ad = parseDateYmdLocal(a.data) || new Date(a.data);
        const bd = parseDateYmdLocal(b.data) || new Date(b.data);
        return ad - bd; // crescente: mais próximo primeiro
      });
      listaEventos.innerHTML = sorted.map(ev => {
        const congObj = congregacoesByIdEvents[ev.congregacaoId];
        const congLabel = ev.congregacaoNome || (congObj ? (congObj.nomeFormatado || (congObj.cidade && congObj.bairro ? `${congObj.cidade} - ${congObj.bairro}` : (congObj.nome||ev.congregacaoId))) : ev.congregacaoId);
        return `
        <div class="item">
          <div>
            <strong>${ev.tipo}</strong>
            <div class="meta">Data: ${formatDate(ev.data)}</div>
            <div class="meta">Congregação: ${congLabel}</div>
            <div class="meta">Atendido por: ${ev.atendenteNome||'-'}</div>
            ${ev.observacoes?`<div class="meta">Obs: ${ev.observacoes}</div>`:''}
          </div>
          <div>
            <button class="btn btn-sm btn-outline-secondary" data-action="edit-ev" data-id="${ev.id}">Editar</button>
            <button class="btn btn-sm btn-outline-danger" data-action="delete-ev" data-id="${ev.id}">Excluir</button>
          </div>
        </div>
      `}).join('');
      renderTabelaReforcos();
    });
  if(listaEventos){
    listaEventos.addEventListener('click', async (e)=>{
        const btnEdit = e.target.closest('button[data-action="edit-ev"]');
        const btnDel = e.target.closest('button[data-action="delete-ev"]');
        if(btnEdit){
          const id = btnEdit.getAttribute('data-id');
          const ev = (eventosCache||[]).find(x=>x.id===id);
          if(!ev || !formEvento) return;
          formEvento.dataset.editId = ev.id;
          const submitBtn = formEvento.querySelector('button[type="submit"]');
          const resetBtn = formEvento.querySelector('button[type="reset"]');
          if(submitBtn) submitBtn.textContent = 'Atualizar Atendimento';
          if(resetBtn) resetBtn.textContent = 'Cancelar Edição';
          const tipoSel = formEvento.querySelector('select[name="tipo"]');
          const dataInp = formEvento.querySelector('input[name="data"]');
          const obsTxt = formEvento.querySelector('textarea[name="observacoes"]');
          if(tipoSel) tipoSel.value = ev.tipo||'';
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
        }
        if(btnDel){
          const id = btnDel.getAttribute('data-id');
          const ok = window.confirm('Excluir este atendimento?');
          if(!ok) return;
          try{
            const removed = await remove('eventos', id);
            if(removed){ toast('Atendimento removido'); }
          }catch(err){
            console.error(err);
            const msg = (err && (err.code||err.message)) || 'Falha ao excluir atendimento';
            toast(msg, 'error');
          }
        }
      });
    }
  }

  // Geração de arquivo HTML para impressão dos eventos
  if(btnImprimirEventos){
    btnImprimirEventos.addEventListener('click', ()=>{
      try{
        const reforcos = (eventosCache||[]).slice().sort((a,b)=>{
          const ad = parseDateYmdLocal(a.data) || new Date(a.data);
          const bd = parseDateYmdLocal(b.data) || new Date(b.data);
          return ad - bd;
        });
        if(!reforcos.length){
          toast('Nenhum atendimento cadastrado para imprimir', 'error');
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
            <td>${ev.tipo}</td>
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
      }catch(err){
        console.error(err);
        toast('Falha ao gerar arquivo de impressão', 'error');
      }
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
      congregacoesCacheEvents = list;
      congregacoesByIdEvents = {};
      list.forEach(c => { congregacoesByIdEvents[c.id] = c; });
      fillSelect(eventoCong, list, 'nomeFormatado');
      renderEventoCultosPreview(eventoCong.value||'');
      renderTabelaReforcos();
    });
    eventoCong.addEventListener('change', (e)=>{
      const congId = e.target.value || '';
      renderEventoCultosPreview(congId);
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

  function parseDateYmdLocal(str){
    if(!str || !/^\d{4}-\d{2}-\d{2}$/.test(str)) return null;
    const [y,m,d] = str.split('-').map(n=>parseInt(n,10));
    return new Date(y, (m||1)-1, d||1);
  }
  function formatDate(str){
    try{
      const d = parseDateYmdLocal(str) || new Date(str);
      return d.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    }catch{ return str; }
  }

  const formCong = qs('#form-congregacao');
  const listaCong = qs('#lista-congregacoes');
  const congAnciaoSel = qs('#cong-anciao');
  const congDiaconoSel = qs('#cong-diacono');
  const badgeCO = qs('#badge-cooperador-oficial');
  const badgeCJ = qs('#badge-cooperador-jovens');
  const listaMin = qs('#lista-ministerio');
  // Horários de Cultos (CO/RJM)
  const cultosWrapper = qs('#cultos-wrapper');
  const addCultoBtn = qs('#add-culto');
  const diasSemana = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
  const diasIndice = { 'Domingo':0,'Segunda':1,'Terça':2,'Quarta':3,'Quinta':4,'Sexta':5,'Sábado':6 };
  function nextOccurrence(dia, horario){
    try{
      const targetDow = diasIndice[dia];
      if(targetDow===undefined || !horario) return null;
      const [h,m] = horario.split(':').map(n=>parseInt(n,10));
      const now = new Date();
      const base = new Date(now);
      base.setHours(h||0, m||0, 0, 0);
      const delta = (targetDow - now.getDay() + 7) % 7;
      base.setDate(now.getDate() + delta);
      if(delta===0 && base <= now){ base.setDate(base.getDate()+7); }
      return base;
    }catch{ return null; }
  }
  function formatDateTime(d){
    if(!d) return '-';
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yyyy = d.getFullYear();
    const hh = String(d.getHours()).padStart(2,'0');
    const mi = String(d.getMinutes()).padStart(2,'0');
    return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
  }
  function renderCultosPreview(){
    const tbody = qs('#cultos-preview-body');
    if(!tbody) return;
    const items = collectCultos();
    if(!items.length){
      tbody.innerHTML = '<tr><td colspan="3" class="text-muted">Nenhum horário adicionado</td></tr>';
      return;
    }
    tbody.innerHTML = items.map(it=>{
      return `<tr>
        <td>${it.tipo}</td>
        <td>${it.dia}</td>
        <td>${it.horario}</td>
      </tr>`;
    }).join('');
  }
  function createCultoRow(initial={}){
    const { tipo='Culto Oficial', dia='Domingo', horario='' } = initial;
    const row = document.createElement('div');
    row.className = 'culto-row grid-3';
    row.innerHTML = `
      <label>
        Tipo
        <select name="cultoTipo">
          <option value="Culto Oficial">Culto Oficial</option>
          <option value="RJM">RJM</option>
        </select>
      </label>
      <label>
        Dia da semana
        <select name="cultoDia">
          ${diasSemana.map(d=>`<option value="${d}">${d}</option>`).join('')}
        </select>
      </label>
      <label>
        Horário
        <input type="time" name="cultoHorario" />
      </label>
      <div class="form-actions">
        <button type="button" class="btn btn-sm btn-outline-danger culto-remove" title="Remover horário">Remover</button>
      </div>
    `;
    const selTipo = row.querySelector('select[name="cultoTipo"]');
    const selDia = row.querySelector('select[name="cultoDia"]');
    const inpHor = row.querySelector('input[name="cultoHorario"]');
    if(selTipo) selTipo.value = tipo;
    if(selDia) selDia.value = dia;
    if(inpHor) inpHor.value = horario;
    const rm = row.querySelector('.culto-remove');
    rm && rm.addEventListener('click', ()=> { row.remove(); renderCultosPreview(); });
    selTipo && selTipo.addEventListener('change', renderCultosPreview);
    selDia && selDia.addEventListener('change', renderCultosPreview);
    inpHor && inpHor.addEventListener('input', renderCultosPreview);
    return row;
  }
  function collectCultos(){
    if(!cultosWrapper) return [];
    return Array.from(cultosWrapper.querySelectorAll('.culto-row')).map(row=>{
      const tipo = row.querySelector('select[name="cultoTipo"]')?.value||'';
      const dia = row.querySelector('select[name="cultoDia"]')?.value||'';
      const horario = row.querySelector('input[name="cultoHorario"]')?.value||'';
      return { tipo, dia, horario };
    }).filter(c=> c.tipo && c.dia && c.horario);
  }
  if(addCultoBtn && cultosWrapper){
    addCultoBtn.addEventListener('click', ()=>{
      cultosWrapper.appendChild(createCultoRow({ tipo:'Culto Oficial', dia:'Domingo', horario:'' }));
      renderCultosPreview();
    });
  }

  // UI repeatable de vínculos de Ministério (Ancião/Diácono)
  const anciaosWrapper = qs('#anciaos-wrapper');
  const addAnciaoBtn = qs('#add-anciao');
  const diaconosWrapper = qs('#diaconos-wrapper');
  const addDiaconoBtn = qs('#add-diacono');

  function createAnciaoRow(initial={}){
    const { anciaoId='', anciaoTipo='' } = initial;
    const row = document.createElement('div');
    row.className = 'vinculo-row grid-3';
    const options = (ministryCache||[])
      .filter(m=>m.funcao==='Ancião')
      .map(m=>`<option value="${m.id}">${m.nome}</option>`)
      .join('');
    row.innerHTML = `
      <label>
        Ancião
        <select name="anciaoIdRow">
          <option value="">Selecionar...</option>
          ${options}
        </select>
      </label>
      <label>
        Tipo
        <select name="anciaoTipoRow">
          <option value="">Selecionar...</option>
          <option value="local">Local</option>
          <option value="responsavel">Responsável</option>
        </select>
      </label>
      <div class="form-actions">
        <button type="button" class="btn btn-sm btn-outline-danger vinculo-remove" title="Remover vínculo">Remover</button>
      </div>
    `;
    const selId = row.querySelector('select[name="anciaoIdRow"]');
    const selTipo = row.querySelector('select[name="anciaoTipoRow"]');
    if(selId) selId.value = anciaoId;
    if(selTipo) selTipo.value = anciaoTipo;
    const rm = row.querySelector('.vinculo-remove');
    rm && rm.addEventListener('click', ()=> row.remove());
    return row;
  }
  function createDiaconoRow(initial={}){
    const { diaconoId='', diaconoTipo='' } = initial;
    const row = document.createElement('div');
    row.className = 'vinculo-row grid-3';
    const options = (ministryCache||[])
      .filter(m=>m.funcao==='Diácono')
      .map(m=>`<option value="${m.id}">${m.nome}</option>`)
      .join('');
    row.innerHTML = `
      <label>
        Diácono
        <select name="diaconoIdRow">
          <option value="">Selecionar...</option>
          ${options}
        </select>
      </label>
      <label>
        Tipo
        <select name="diaconoTipoRow">
          <option value="">Selecionar...</option>
          <option value="local">Local</option>
          <option value="responsavel">Responsável</option>
        </select>
      </label>
      <div class="form-actions">
        <button type="button" class="btn btn-sm btn-outline-danger vinculo-remove" title="Remover vínculo">Remover</button>
      </div>
    `;
    const selId = row.querySelector('select[name="diaconoIdRow"]');
    const selTipo = row.querySelector('select[name="diaconoTipoRow"]');
    if(selId) selId.value = diaconoId;
    if(selTipo) selTipo.value = diaconoTipo;
    const rm = row.querySelector('.vinculo-remove');
    rm && rm.addEventListener('click', ()=> row.remove());
    return row;
  }
  function collectVinculosAnciao(){
    if(!anciaosWrapper) return [];
    return Array.from(anciaosWrapper.querySelectorAll('.vinculo-row')).map(row=>{
      const id = row.querySelector('select[name="anciaoIdRow"]')?.value||'';
      const tipo = row.querySelector('select[name="anciaoTipoRow"]')?.value||'';
      return { id, tipo };
    }).filter(x=> x.id);
  }
  function collectVinculosDiacono(){
    if(!diaconosWrapper) return [];
    return Array.from(diaconosWrapper.querySelectorAll('.vinculo-row')).map(row=>{
      const id = row.querySelector('select[name="diaconoIdRow"]')?.value||'';
      const tipo = row.querySelector('select[name="diaconoTipoRow"]')?.value||'';
      return { id, tipo };
    }).filter(x=> x.id);
  }
  if(addAnciaoBtn && anciaosWrapper){
    addAnciaoBtn.addEventListener('click', ()=>{
      anciaosWrapper.appendChild(createAnciaoRow());
    });
  }
  if(addDiaconoBtn && diaconosWrapper){
    addDiaconoBtn.addEventListener('click', ()=>{
      diaconosWrapper.appendChild(createDiaconoRow());
    });
  }
  async function linkMultipleMinisterioForCongregacao(congId, elders=[], deacons=[]){
    try{
      const ops = [];
      elders.forEach(el=>{
        if(!el.id) return;
        ops.push(update('ministerio', el.id, {
          congregacaoId: congId,
          anciaoLocal: el.tipo==='local',
          anciaoResponsavel: el.tipo==='responsavel',
        }));
      });
      deacons.forEach(dc=>{
        if(!dc.id) return;
        ops.push(update('ministerio', dc.id, {
          congregacaoId: congId,
          diaconoLocal: dc.tipo==='local',
          diaconoResponsavel: dc.tipo==='responsavel',
        }));
      });
      if(ops.length){ await Promise.all(ops); toast('Múltiplos vínculos de Ministério atualizados'); }
    }catch(err){ console.error(err); const msg = (err && (err.code||err.message)) || 'Falha ao vincular múltiplos'; toast(msg, 'error'); }
  }
  let congLabelById = {};
  let congregacoesCache = [];
  // Cache de rótulos de congregações para exibir nos listados
  readList('congregacoes', list => {
    congregacoesCache = list;
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
            <div>
              <button class="btn btn-sm btn-outline-secondary" data-action="edit-min" data-id="${m.id}">Editar</button>
            </div>
          </div>
        `;
      }).join('');
    }
    // Atualizar a lista de congregações para refletir cooperadores por congregação
    if(listaCong){ renderCongList(); }
  });

  // Vincular automaticamente Congregação quando salvar/editar Ministério
  async function autoLinkCongregacaoWithMinisterio(m){
    try{
      if(!m || !m.congregacaoId) return;
      const updates = {};
      if(m.funcao === 'Ancião'){
        const tipo = m.anciaoResponsavel ? 'responsavel' : (m.anciaoLocal ? 'local' : '');
        updates.anciaoId = m.id;
        updates.anciaoNome = m.nome || '';
        updates.anciaoTipo = tipo;
      } else if(m.funcao === 'Diácono'){
        const tipo = m.diaconoResponsavel ? 'responsavel' : (m.diaconoLocal ? 'local' : '');
        updates.diaconoId = m.id;
        updates.diaconoNome = m.nome || '';
        updates.diaconoTipo = tipo;
      } else {
        return;
      }
      const res = await update('congregacoes', m.congregacaoId, updates);
      if(res){ toast('Congregação vinculada ao Ministério'); }
    }catch(err){
      console.error(err);
      const msg = (err && (err.code||err.message)) || 'Falha ao vincular congregação';
      toast(msg, 'error');
    }
  }

  // Vincular automaticamente Ministério quando salvar/editar Congregação
  async function autoLinkMinisterioFromCongregacao(c){
    try{
      if(!c || !c.id) return;
      const ops = [];
      if(c.anciaoId){
        ops.push(update('ministerio', c.anciaoId, {
          congregacaoId: c.id,
          anciaoLocal: c.anciaoTipo === 'local',
          anciaoResponsavel: c.anciaoTipo === 'responsavel',
        }));
      }
      if(c.diaconoId){
        ops.push(update('ministerio', c.diaconoId, {
          congregacaoId: c.id,
          diaconoLocal: c.diaconoTipo === 'local',
          diaconoResponsavel: c.diaconoTipo === 'responsavel',
        }));
      }
      if(ops.length){
        await Promise.all(ops);
        toast('Ministério vinculado à congregação');
      }
    }catch(err){
      console.error(err);
      const msg = (err && (err.code||err.message)) || 'Falha ao vincular Ministério';
      toast(msg, 'error');
    }
  }

  // Relatórios: render helpers
  function renderRelatorios(){
    renderRelatorioEventos();
    renderRelatorioMinisterio();
    renderRelatorioCongregacoes();
  }
  function renderRelatorioEventos(){
    if(!relEventosEl) return;
    const list = eventosCache || [];
    if(!list.length){
      relEventosEl.innerHTML = '<li class="text-muted">Nenhum atendimento cadastrado</li>';
      return;
    }
    const porTipo = list.reduce((acc, ev)=>{ acc[ev.tipo] = (acc[ev.tipo]||0)+1; return acc; }, {});
    const porCong = list.reduce((acc, ev)=>{
      const cong = congregacoesByIdEvents[ev.congregacaoId];
      const label = ev.congregacaoNome || (cong ? (cong.nomeFormatado || (cong.cidade && cong.bairro ? `${cong.cidade} - ${cong.bairro}` : (cong.nome||ev.congregacaoId))) : ev.congregacaoId);
      acc[label] = (acc[label]||0)+1;
      return acc;
    }, {});
    const lines = [];
    lines.push(`Total de atendimentos: ${list.length}`);
    Object.keys(porTipo).sort().forEach(k=> { lines.push(`${k}: ${porTipo[k]}`); });
    Object.keys(porCong).sort().forEach(k=> { lines.push(`${k}: ${porCong[k]}`); });
    relEventosEl.innerHTML = lines.map(s=> `<li>${s}</li>`).join('');
  }
  function renderRelatorioMinisterio(){
    if(!relMinisterioEl) return;
    const list = ministryCache || [];
    if(!list.length){
      relMinisterioEl.innerHTML = '<li class="text-muted">Nenhum irmão cadastrado</li>';
      return;
    }
    const counts = list.reduce((acc, m)=>{ const f=m.funcao||'Sem função'; acc[f]= (acc[f]||0)+1; return acc; }, {});
    const order = ['Ancião','Diácono','Cooperador Oficial','Cooperador de Jovens','Encarregado Regional'];
    const lines = [];
    lines.push(`Total no Ministério: ${list.length}`);
    order.forEach(f=> { if(counts[f]) lines.push(`${f}: ${counts[f]}`); });
    Object.keys(counts).filter(f=> !order.includes(f)).sort().forEach(f=> { lines.push(`${f}: ${counts[f]}`); });
    relMinisterioEl.innerHTML = lines.map(s=> `<li>${s}</li>`).join('');
  }
  function renderRelatorioCongregacoes(){
    if(!relCongregacoesEl) return;
    const list = congregacoesCache || [];
    if(!list.length){
      relCongregacoesEl.textContent = 'Nenhuma congregação cadastrada';
      return;
    }
    relCongregacoesEl.textContent = `Total de congregações cadastradas: ${list.length}`;
  }

  // Editar Ministério: delegar clique e carregar no formulário
  if(listaMin){
    listaMin.addEventListener('click', (e)=>{
      const btn = e.target.closest('button[data-action="edit-min"]');
      if(!btn) return;
      const id = btn.getAttribute('data-id');
      const m = ministryCache.find(x=>x.id===id);
      if(!m) return;
      const formMinEl = qs('#form-ministerio');
      if(!formMinEl) return;
      formMinEl.dataset.editId = m.id;
      const submitBtn = formMinEl.querySelector('button[type="submit"]');
      const resetBtn = formMinEl.querySelector('button[type="reset"]');
      if(submitBtn) submitBtn.textContent = 'Atualizar Irmão';
      if(resetBtn) resetBtn.textContent = 'Cancelar Edição';
      formMinEl.querySelector('input[name="nome"]').value = m.nome||'';
      const funcSel = formMinEl.querySelector('select[name="funcao"]');
      if(funcSel) funcSel.value = m.funcao||'';
      const congSel = formMinEl.querySelector('select[name="congregacaoId"]');
      if(congSel) congSel.value = m.congregacaoId||'';
      formMinEl.querySelector('input[name="telefone"]').value = m.telefone||'';
      formMinEl.querySelector('input[name="email"]').value = m.email||'';
      const al = formMinEl.querySelector('input[name="anciaoLocal"]');
      const ar = formMinEl.querySelector('input[name="anciaoResponsavel"]');
      const dl = formMinEl.querySelector('input[name="diaconoLocal"]');
      const dr = formMinEl.querySelector('input[name="diaconoResponsavel"]');
      if(al) al.checked = !!m.anciaoLocal;
      if(ar) ar.checked = !!m.anciaoResponsavel;
      if(dl) dl.checked = !!m.diaconoLocal;
      if(dr) dr.checked = !!m.diaconoResponsavel;
    });
    const formMinResetEl = qs('#form-ministerio');
    formMinResetEl && formMinResetEl.addEventListener('reset', ()=>{
      delete formMinResetEl.dataset.editId;
      const submitBtn = formMinResetEl.querySelector('button[type="submit"]');
      const resetBtn = formMinResetEl.querySelector('button[type="reset"]');
      if(submitBtn) submitBtn.textContent = 'Salvar Irmão do Ministério';
      if(resetBtn) resetBtn.textContent = 'Limpar';
    });
  }

  // Salvar Congregação
  if(formCong){
    formCong.addEventListener('submit', async (e)=>{
      e.preventDefault();
      const fd = new FormData(formCong);
      const data = Object.fromEntries(fd.entries());
      if(!data.cidade || !data.bairro || !data.endereco){ toast('Preencha cidade, bairro e endereço', 'error'); return; }
      const cultos = collectCultos();
      const elders = collectVinculosAnciao();
      const deacons = collectVinculosDiacono();

      try{
        const editId = formCong.dataset.editId;
        if(editId){
          const updated = await update('congregacoes', editId, {
            cidade: data.cidade,
            bairro: data.bairro,
            endereco: data.endereco,
            cultos,
          });
          if(updated){
            toast('Congregação atualizada');
            await linkMultipleMinisterioForCongregacao(editId, elders, deacons);
            formCong.reset();
          }
        } else {
          const saved = await write('congregacoes', {
            cidade: data.cidade,
            bairro: data.bairro,
            endereco: data.endereco,
            cultos,
          });
          if(saved){
            toast('Congregação salva');
            await linkMultipleMinisterioForCongregacao(saved.id, elders, deacons);
            formCong.reset();
          }
        }
      }catch(err){
        console.error(err);
        const msg = (err && (err.code||err.message)) || 'Falha ao salvar congregação';
        toast(msg, 'error');
      }
    });

    // Render da lista de congregações incluindo Cooperadores
    function renderCongList(){
      if(!listaCong) return;
      const list = congregacoesCache || [];
      const labelFor = (c) => c.nomeFormatado || (c.cidade && c.bairro ? `${c.cidade} - ${c.bairro}` : (c.nome||''));
      const sorted = [...list].sort((a,b)=> (labelFor(a)||'').localeCompare(labelFor(b)||''));
      listaCong.innerHTML = sorted.map(c => {
        // Agregar Anciães e Diáconos: Ministério + campos da congregação, sem duplicar
        const eldersMin = (ministryCache||[])
          .filter(m => m.congregacaoId===c.id && m.funcao==='Ancião');
        const eldersNameSet = new Set(eldersMin.map(m => (m.nome||'').trim()));
        const eldersFormattedMin = eldersMin.map(m => {
          const name = (m.nome||'').trim();
          return m.anciaoLocal ? `<strong>${name}</strong>` : (m.anciaoResponsavel ? `<i>${name}</i>` : name);
        });
        const extraAnciaoName = (c.anciaoNome||'').trim();
        const extraAnciaoTipo = (c.anciaoTipo||'').trim();
        if(extraAnciaoName){
          const formattedExtra = extraAnciaoTipo==='Local' ? `<strong>${extraAnciaoName}</strong>` : (extraAnciaoTipo==='Responsável' ? `<i>${extraAnciaoName}</i>` : extraAnciaoName);
          if(!eldersNameSet.has(extraAnciaoName)) eldersFormattedMin.push(formattedExtra);
        }
        const eldersAllFormatted = eldersFormattedMin.join(', ');
        const eldersCount = eldersFormattedMin.length;
        const eldersLabel = eldersCount>1 ? 'Anciães' : 'Ancião';

        const deaconsMin = (ministryCache||[])
          .filter(m => m.congregacaoId===c.id && m.funcao==='Diácono');
        const deaconsNameSet = new Set(deaconsMin.map(m => (m.nome||'').trim()));
        const deaconsFormattedMin = deaconsMin.map(m => {
          const name = (m.nome||'').trim();
          return `${name}${m.diaconoResponsavel?' (responsável)': (m.diaconoLocal?' (local)':'')}`;
        });
        const extraDiaconoName = (c.diaconoNome||'').trim();
        const extraDiaconoTipo = (c.diaconoTipo||'').trim();
        if(extraDiaconoName){
          const formattedExtra = `${extraDiaconoName}${extraDiaconoTipo==='Responsável'?' (responsável)': (extraDiaconoTipo==='Local'?' (local)':'')}`;
          if(!deaconsNameSet.has(extraDiaconoName)) deaconsFormattedMin.push(formattedExtra);
        }
        const deaconsAllFormatted = deaconsFormattedMin.join(', ');
        const deaconsCount = deaconsFormattedMin.length;
        const deaconsLabel = deaconsCount>1 ? 'Diáconos' : 'Diácono';

        const coNames = (ministryCache||[])
          .filter(m => m.congregacaoId===c.id && m.funcao==='Cooperador Oficial')
          .map(m => m.nome).join(', ');
        const cjNames = (ministryCache||[])
          .filter(m => m.congregacaoId===c.id && m.funcao==='Cooperador de Jovens')
          .map(m => m.nome).join(', ');
        const coLine = coNames ? `<div class="meta">Cooperadores Oficiais: ${coNames}</div>` : '';
        const cjLine = cjNames ? `<div class="meta">Cooperadores de Jovens: ${cjNames}</div>` : '';
        const cultosOficiais = ((c.cultos||[]).filter(x=>x.tipo==='Culto Oficial').length)
          ? `<div class="meta">Cultos Oficiais: ${(c.cultos||[]).filter(x=>x.tipo==='Culto Oficial').map(x=>`${x.dia} ${x.horario}`).join(', ')}</div>`
          : '';
        const cultosRjm = ((c.cultos||[]).filter(x=>x.tipo==='RJM').length)
          ? `<div class="meta">RJM: ${(c.cultos||[]).filter(x=>x.tipo==='RJM').map(x=>`${x.dia} ${x.horario}`).join(', ')}</div>`
          : '';
        return `
          <div class="item">
            <div>
              <strong>${labelFor(c)}</strong>
              <div class="meta">Endereço: ${c.endereco}</div>
              <div class="meta">${eldersLabel}: ${eldersAllFormatted || '-'}</div>
              <div class="meta">${deaconsLabel}: ${deaconsAllFormatted || '-'}</div>
              ${coLine}
              ${cjLine}
              ${cultosOficiais}
              ${cultosRjm}
            </div>
            <div>
              <button class="btn btn-sm btn-outline-secondary" data-action="edit-cong" data-id="${c.id}">Editar</button>
            </div>
          </div>
        `;
      }).join('');
    }

    readList('congregacoes', list => {
      if(!listaCong) return;
      congregacoesCache = list;
      renderCongList();
    });

    // Editar Congregação: carregar dados no formulário
    function startEditCongregacao(c){
      if(!formCong || !c) return;
      formCong.dataset.editId = c.id;
      const submitBtn = formCong.querySelector('button[type="submit"]');
      const resetBtn = formCong.querySelector('button[type="reset"]');
      if(submitBtn) submitBtn.textContent = 'Atualizar Congregação';
      if(resetBtn) resetBtn.textContent = 'Cancelar Edição';
      formCong.querySelector('input[name="cidade"]').value = c.cidade||'';
      formCong.querySelector('input[name="bairro"]').value = c.bairro||'';
      formCong.querySelector('input[name="endereco"]').value = c.endereco||'';
      if(anciaosWrapper){
        anciaosWrapper.innerHTML = '';
        const eldersMin = (ministryCache||[]).filter(m=>m.congregacaoId===c.id && m.funcao==='Ancião');
        if(eldersMin.length){
          eldersMin.forEach(m=>{
            const tipo = m.anciaoResponsavel ? 'responsavel' : (m.anciaoLocal ? 'local' : '');
            anciaosWrapper.appendChild(createAnciaoRow({ anciaoId: m.id, anciaoTipo: tipo }));
          });
        } else if(c.anciaoId){
          const tipo = (c.anciaoTipo||'').toLowerCase();
          anciaosWrapper.appendChild(createAnciaoRow({ anciaoId: c.anciaoId, anciaoTipo: tipo }));
        }
      }
      if(diaconosWrapper){
        diaconosWrapper.innerHTML = '';
        const deaconsMin = (ministryCache||[]).filter(m=>m.congregacaoId===c.id && m.funcao==='Diácono');
        if(deaconsMin.length){
          deaconsMin.forEach(m=>{
            const tipo = m.diaconoResponsavel ? 'responsavel' : (m.diaconoLocal ? 'local' : '');
            diaconosWrapper.appendChild(createDiaconoRow({ diaconoId: m.id, diaconoTipo: tipo }));
          });
        } else if(c.diaconoId){
          const tipo = (c.diaconoTipo||'').toLowerCase();
          diaconosWrapper.appendChild(createDiaconoRow({ diaconoId: c.diaconoId, diaconoTipo: tipo }));
        }
      }
      if(cultosWrapper){
        cultosWrapper.innerHTML = '';
        (c.cultos||[]).forEach(ct=> cultosWrapper.appendChild(createCultoRow(ct)));
        renderCultosPreview();
      }
    }
    if(listaCong){
      listaCong.addEventListener('click', (e)=>{
        const btn = e.target.closest('button[data-action="edit-cong"]');
        if(!btn) return;
        const id = btn.getAttribute('data-id');
        const c = congregacoesCache.find(x=>x.id===id);
        startEditCongregacao(c);
      });
      formCong && formCong.addEventListener('reset', ()=>{
        delete formCong.dataset.editId;
        const submitBtn = formCong.querySelector('button[type="submit"]');
        const resetBtn = formCong.querySelector('button[type="reset"]');
        if(submitBtn) submitBtn.textContent = 'Salvar Atendimento';
        if(resetBtn) resetBtn.textContent = 'Limpar';
        if(cultosWrapper) cultosWrapper.innerHTML = '';
        const tbody = qs('#cultos-preview-body');
        if(tbody) tbody.innerHTML = '<tr><td colspan="3" class="text-muted">Nenhum horário adicionado</td></tr>';
        if(anciaosWrapper) anciaosWrapper.innerHTML = '';
        if(diaconosWrapper) diaconosWrapper.innerHTML = '';
      });
    }
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
        const editId = formMin.dataset.editId;
        if(editId){
          const updated = await update('ministerio', editId, payload);
          if(updated){ toast('Irmão do Ministério atualizado'); await autoLinkCongregacaoWithMinisterio({ id: editId, ...payload }); formMin.reset(); }
        } else {
          const saved = await write('ministerio', payload);
          if(saved){ toast('Irmão do Ministério salvo'); await autoLinkCongregacaoWithMinisterio(saved); formMin.reset(); }
        }
      }catch(err){
        console.error(err);
        const msg = (err && (err.code||err.message)) || 'Falha ao salvar Ministério';
        toast(msg, 'error');
      }
    });
  }
  // Agendamento de Coletas (Página agendar.html)
  const agendarGrid = qs('#agendar-grid');
  const agendarForm = qs('#agendar-form');
  const agendarCongLabelEl = qs('#agendar-cong-label');
  const agendarCongIdInput = qs('#agendar-cong-id');
  const agendarTipoSel = qs('#agendar-tipo');
  const agendarDataInp = qs('#agendar-data');
  const agendarCancelBtn = qs('#agendar-cancelar');

  if(agendarGrid){
    // Carregar congregações e renderizar botões
    readList('congregacoes', list => {
      const labelFor = (c) => c.nomeFormatado || (c.cidade && c.bairro ? `${c.cidade} - ${c.bairro}` : (c.nome||c.id));
      agendarGrid.innerHTML = list.map(c => 
         `<button type="button" class="btn btn-outline-primary agendar-cong-btn" data-id="${c.id}">
           <span class="icon" aria-hidden="true">
             <svg viewBox="0 0 24 24"><path d="M12 3l9 7-1.5 2L12 7 4.5 12 3 10z" fill="currentColor"/><path d="M5 13v7h5v-4h4v4h5v-7l-7-5z" fill="currentColor"/></svg>
           </span>
           <span class="agendar-label">${labelFor(c)}</span>
         </button>`
       ).join('');
      congregacoesByIdEvents = {};
      list.forEach(c => { congregacoesByIdEvents[c.id] = c; });

      agendarGrid.addEventListener('click', (e)=>{
        const btn = e.target.closest('.agendar-cong-btn');
        if(!btn) return;
        const congId = btn.getAttribute('data-id');
        const c = list.find(x => x.id===congId);
        const label = labelFor(c);
        if(agendarCongIdInput) agendarCongIdInput.value = congId;
        if(agendarCongLabelEl) agendarCongLabelEl.textContent = label;
        agendarForm && agendarForm.classList.remove('hidden');
        agendarTipoSel && (agendarTipoSel.value = 'Culto Reforço de Coletas');
        agendarDataInp && (agendarDataInp.value = '');
        const hint = qs('#agendar-hint');
        if(hint){
          if(label==='Ituiutaba - Centro'){
            hint.textContent = 'Obs: Ituiutaba - Centro deve ser agendado na quinta-feira (Culto Oficial).';
            hint.classList.remove('hidden');
          } else {
            hint.textContent = '';
            hint.classList.add('hidden');
          }
        }
      });
    });

    // Carregar eventos para validação
    readList('eventos', list => { eventosCache = list; });

    if(agendarForm){
      agendarForm.addEventListener('submit', async (e)=>{
        e.preventDefault();
        const congId = agendarCongIdInput ? agendarCongIdInput.value : '';
        const tipo = agendarTipoSel ? agendarTipoSel.value : '';
        const data = agendarDataInp ? agendarDataInp.value : '';
        if(!congId || !tipo || !data){ toast('Selecione congregação, tipo e data', 'error'); return; }

        const c = congregacoesByIdEvents[congId];
        const label = c ? (c.nomeFormatado || (c.cidade && c.bairro ? `${c.cidade} - ${c.bairro}` : (c.nome||congId))) : congId;

        // Exceção: Ituiutaba - Centro -> CO deve ser na quinta-feira
        if(label==='Ituiutaba - Centro' && tipo==='Culto Reforço de Coletas'){
          const d = parseDateYmdLocal(data);
          if(!d || d.getDay() !== 4){ toast('Ituiutaba - Centro: Culto Oficial deve ser na quinta-feira.', 'error'); return; }
        }

        const dNew = parseDateYmdLocal(data);
        const sameMonthSameCong = (eventosCache||[]).filter(ev=>{
          const dEv = parseDateYmdLocal(ev.data);
          return ev.congregacaoId===congId && dEv && dNew && dEv.getFullYear()===dNew.getFullYear() && dEv.getMonth()===dNew.getMonth();
        });
        // Regra: permitir múltiplas coletas, mas bloquear duplicidade por tipo (CO/RJM) na mesma congregação/mês
        const typeAlready = sameMonthSameCong.some(ev => ev.tipo===tipo);
        if(typeAlready){
          toast('Já existe este tipo agendado nesta congregação neste mês.', 'error');
          return;
        }

        // Atendente: lista do Ministério ou outra região (manual)
        const atendenteId = agendarAtendenteSel ? agendarAtendenteSel.value : '';
        const manualNome = agendarAtendenteManualInp ? agendarAtendenteManualInp.value.trim() : '';
        if(!atendenteId && !manualNome){
          toast('Informe quem vai atender (lista ou outra região).', 'error');
          return;
        }
        let atendenteNome = '';
        if(atendenteId && agendarAtendenteSel){
          const opt = agendarAtendenteSel.querySelector(`option[value="${atendenteId}"]`);
          atendenteNome = opt ? opt.textContent : '';
        }
        if(manualNome){
          atendenteNome = manualNome;
        }

        try{
          const saved = await write('eventos', {
            tipo,
            data,
            congregacaoId: congId,
            congregacaoNome: label,
            atendenteId: manualNome ? '' : (atendenteId||''),
            atendenteNome: atendenteNome,
            atendenteOutraRegiao: !!manualNome,
            observacoes: ''
          });
          if(saved){
            toast('Agendamento salvo');
            agendarForm.reset();
            toggleAgendarAtendenteManual(false);
            maybeShowAgendarAtendente();
            agendarForm.classList.add('hidden');
          }
        }catch(err){
          console.error(err);
          toast('Falha ao salvar agendamento', 'error');
        }
      });

      agendarCancelBtn && agendarCancelBtn.addEventListener('click', (e)=>{
        e.preventDefault();
        agendarForm.reset();
        toggleAgendarAtendenteManual(false);
        maybeShowAgendarAtendente();
        agendarForm.classList.add('hidden');
      });
    }
  }

  window._renderTabelaReforcosInner = function(){
    if(!tabelaReforcosBody) return;
    const tipoPriority = (t) => t==='Culto Reforço de Coletas' ? 0 : (t==='RJM com Reforço de Coletas' ? 1 : 2);
     const reforcos = (eventosCache||[]).slice().sort((a,b)=>{
       const pa = tipoPriority(a.tipo);
       const pb = tipoPriority(b.tipo);
       if (pa !== pb) return pa - pb; // CO primeiro, depois RJM
       const ad = parseDateYmdLocal(a.data) || new Date(a.data);
       const bd = parseDateYmdLocal(b.data) || new Date(b.data);
       return ad - bd; // crescente por data
     });
    if(!reforcos.length){
      tabelaReforcosBody.innerHTML = '<tr><td colspan="5" class="text-muted">Nenhum atendimento cadastrado</td></tr>';
      return;
    }
    const diasSemanaPt = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
    tabelaReforcosBody.innerHTML = reforcos.map(ev => {
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
      <td>${ev.tipo}</td>
    </tr>`;
    }).join('');
  };

}());
  function renderTabelaReforcos(){
    if (typeof window !== 'undefined' && typeof window._renderTabelaReforcosInner === 'function') {
      return window._renderTabelaReforcosInner();
    }
  }
