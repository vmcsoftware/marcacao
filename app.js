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
  const btnExportarPdfEventos = qs('#btn-exportar-pdf-eventos');
  const btnExportarXlsEventos = qs('#btn-exportar-xls-eventos');
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
const btnRelPrint = qs('#rel-print');
const btnRelPdf = qs('#rel-print-pdf');
const btnRelXls = qs('#rel-export-xls');
const relEnsaiosSortSel = qs('#rel-ensaios-sort');
const btnEnsaiosExportCsv = qs('#rel-ensaios-export-csv');
const btnEnsaiosExportXls = qs('#rel-ensaios-export-xls');
const btnEnsaiosCopy = qs('#rel-ensaios-copy');
const relBairroSel = qs('#rel-bairro');
const relEnsaiosNextOnly = qs('#rel-ensaios-next-only');
  let relEventosFilteredCache = null;
  async function applyRelFiltersFetch(){
    try{
      if(!db){ if(typeof renderRelatorios==='function') renderRelatorios(); return; }
      const y = relYearSel && relYearSel.value ? parseInt(relYearSel.value,10) : null;
      const m = relMonthSel && relMonthSel.value ? parseInt(relMonthSel.value,10) : null;
      const tipo = relTipoSel && relTipoSel.value ? relTipoSel.value : '';
      let snap;
      if(tipo){
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
    }catch(err){ console.error(err); toast('Falha ao buscar eventos no Firebase', 'error'); relEventosFilteredCache = null; }
    finally{
      try{ if(typeof renderRelatorios==='function') renderRelatorios(); }catch{}
    }
  }
  btnRelApply && btnRelApply.addEventListener('click', (e)=>{ e.preventDefault(); applyRelFiltersFetch(); });
  btnRelClear && btnRelClear.addEventListener('click', ()=>{ relEventosFilteredCache = null; });
btnRelClear && btnRelClear.addEventListener('click', ()=>{ if(relYearSel) relYearSel.value=''; if(relMonthSel) relMonthSel.value=''; if(relCidadeSel) relCidadeSel.value=''; if(relBairroSel) relBairroSel.value=''; if(relEnsaiosSortSel) relEnsaiosSortSel.value='cidade'; if(relEnsaiosNextOnly) relEnsaiosNextOnly.checked=false; if(relTipoSel) relTipoSel.value=''; if(typeof renderRelatorios==='function') renderRelatorios(); });
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
      if(relEventosEl || btnRelPrint || relYearSel || relCidadeSel || relTipoSel){
        try{ if(typeof initRelatoriosFilters==='function') initRelatoriosFilters(); }catch{}
        try{ if(typeof renderRelatorios==='function') renderRelatorios(); }catch{}
      }
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
            <div class="meta">Congregação: <span class="congregacao-highlight">${congLabel}</span></div>
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

  // Exportação em PDF (salvar via diálogo de impressão)
  if(btnExportarPdfEventos){
    btnExportarPdfEventos.addEventListener('click', ()=>{
      try{
        const reforcos = (eventosCache||[]).slice().sort((a,b)=>{
          const ad = parseDateYmdLocal(a.data) || new Date(a.data);
          const bd = parseDateYmdLocal(b.data) || new Date(b.data);
          return ad - bd;
        });
        if(!reforcos.length){ toast('Nenhum atendimento para exportar', 'error'); return; }
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
          const atendente = (ev.atendenteNome||'-').replace(/</g,'&lt;').replace(/>/g,'&gt;');
          return `<tr>
            <td>${diaNome} - ${formatDate(ev.data)}</td>
            <td>${hora}</td>
            <td>${localLabel}</td>
            <td>${atendente}${ev.atendenteOutraRegiao ? ' <span class="flag-outra-regiao" title="Irmão de outra região"><svg viewBox="0 0 24 24"><path d="M12 2c-4.4 0-8 3.1-8 7 0 5 8 13 8 13s8-8 8-13c0-3.9-3.6-7-8-7zm0 9a2 2 0 1 1 0-4 2 2 0 0 1 0 4z" fill="currentColor"/></svg></span>' : ''}</td>
            <td>${ev.tipo}</td>
          </tr>`;
        }).join('');
        const html = `<!DOCTYPE html>
<html lang='pt-BR'>
<head>
  <meta charset='UTF-8' />
  <title>Atendimentos - PDF</title>
  <style>
    body{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; color:#000; margin:24px; }
    h1{ font-size: 18px; margin: 0 0 8px; }
    table{ width:100%; border-collapse: collapse; }
    th, td{ border:1px solid #000; padding:6px; font-size:12px; }
    th{ background:#f2f2f2; }
    .no-print{ margin:12px 0; }
    @media print{ .no-print{ display:none; } }
  </style>
</head>
<body>
  <h1>Atendimentos Cadastrados</h1>
  <div class='no-print'>Use Ctrl+P e escolha "Salvar como PDF".</div>
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
    <tbody>${rowsHtml}</tbody>
  </table>
  <script>window.addEventListener('load', function(){ setTimeout(function(){ window.print(); }, 200); });</script>
</body>
</html>`;
        const w = window.open('', '_blank');
        if(!w){ toast('Não foi possível abrir a janela de impressão (pop-up bloqueado)', 'error'); return; }
        w.document.open();
        w.document.write(html);
        w.document.close();
        toast('Abra a caixa de impressão e salve em PDF');
      }catch(err){ console.error(err); toast('Falha ao exportar em PDF', 'error'); }
    });
  }

  // Exportação em XLS (Excel)
  if(btnExportarXlsEventos){
    btnExportarXlsEventos.addEventListener('click', ()=>{
      try{
        const reforcos = (eventosCache||[]).slice().sort((a,b)=>{
          const ad = parseDateYmdLocal(a.data) || new Date(a.data);
          const bd = parseDateYmdLocal(b.data) || new Date(b.data);
          return ad - bd;
        });
        if(!reforcos.length){ toast('Nenhum atendimento para exportar', 'error'); return; }
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
          const atendente = (ev.atendenteNome||'-').replace(/</g,'&lt;').replace(/>/g,'&gt;');
          return `<tr>
            <td>${diaNome} - ${formatDate(ev.data)}</td>
            <td>${hora}</td>
            <td>${localLabel}</td>
            <td>${atendente}</td>
            <td>${ev.tipo}</td>
          </tr>`;
        }).join('');
        const xlsHtml = `<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>
          <table border='1'>
            <thead><tr><th>Data</th><th>Hora</th><th>Local</th><th>Quem atende</th><th>Tipo</th></tr></thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </body></html>`;
        const blob = new Blob([xlsHtml], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'atendimentos.xls';
        document.body.appendChild(a);
        a.click();
        setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
        toast('Arquivo XLS gerado');
      }catch(err){ console.error(err); toast('Falha ao exportar XLS', 'error'); }
    });
  }

  // Popular congregações no select de evento
  function fillSelect(sel, list, labelKey){
    if(!sel) return;
    const makeLabel = (x) => x[labelKey] || x.nomeFormatado || (x.cidade && x.bairro ? `${x.cidade} - ${x.bairro}` : (x.nome||''));
    const sorted = [...(list||[])].sort((a,b)=> (makeLabel(a)||'').localeCompare(makeLabel(b)||'', undefined, { sensitivity: 'base' }));
    sel.innerHTML = '<option value="">Selecionar...</option>' +
      sorted.map(x => {
        const label = makeLabel(x);
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
      const day = String(d.getDate()).padStart(2,'0');
      const month = String(d.getMonth()+1).padStart(2,'0');
      const year = d.getFullYear();
      return `${day}/${month}/${year}`;
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
  // Utilitário: obter número de dias no mês
  function daysInMonth(year, monthIndex){
    return new Date(year, monthIndex+1, 0).getDate();
  }
  // Utilitário: encontrar a data da Nª ocorrência de um dia da semana em um mês
  // weekdayIndex: 0=Domingo..6=Sábado; monthIndex: 0..11; nth: 1..5
  function nthWeekdayOfMonth(year, monthIndex, weekdayIndex, nth){
    if(!year || year<1900 || monthIndex<0 || monthIndex>11 || weekdayIndex<0 || weekdayIndex>6 || nth<1) return null;
    const firstDay = new Date(year, monthIndex, 1).getDay();
    const firstOccurrence = 1 + ((weekdayIndex - firstDay + 7) % 7);
    const date = firstOccurrence + 7 * (nth - 1);
    const dim = daysInMonth(year, monthIndex);
    if(date > dim) return null;
    return new Date(year, monthIndex, date);
  }
  function formatYmd(date){
    if(!(date instanceof Date) || isNaN(date)) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth()+1).padStart(2,'0');
    const d = String(date.getDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }
  // Render preview de datas calculadas do Ensaio
  function renderEnsaioPreview(){
    const body = qs('#ensaio-preview-body');
    if(!body) return;
    const tipoEl = qs('#ensaio-tipo');
    const horarioEl = qs('#ensaio-horario');
    const diaEl = qs('#ensaio-dia');
    const semanaEl = qs('#ensaio-semana');
    const anoEl = qs('#ensaio-ano');
    const mesesWrap = qs('#ensaio-meses');
    const tipo = tipoEl ? (tipoEl.value||'') : '';
    const horario = horarioEl ? (horarioEl.value||'') : '';
    const diaSemana = diaEl ? (diaEl.value||'') : '';
    const semana = semanaEl ? parseInt(semanaEl.value||'',10) : NaN;
    const ano = anoEl ? parseInt(anoEl.value||'',10) : NaN;
    const meses = mesesWrap ? Array.from(mesesWrap.querySelectorAll('input[name="ensaioMes"]:checked')).map(inp=>parseInt(inp.value,10)).filter(n=>!isNaN(n)) : [];
    if(!tipo || !horario || !diaSemana || !semana || !ano || !meses.length){
      body.innerHTML = '<tr><td colspan="2" class="text-muted">Selecione dia, semana, ano e meses para calcular</td></tr>';
      return;
    }
    const wd = diasIndice[diaSemana];
    const mesesNomes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    body.innerHTML = meses.map(m => {
      const date = nthWeekdayOfMonth(ano, m-1, wd, semana);
      const ymd = date ? formatYmd(date) : '—';
      return `<tr><td>${mesesNomes[(m||1)-1]||m}</td><td>${ymd}</td></tr>`;
    }).join('');
  }
  // Coleta dados de Ensaio (Regional/Local, Meses, Horário, Dia/Semana/Ano e datas calculadas)
  function collectEnsaio(){
    const tipoEl = qs('#ensaio-tipo');
    const horarioEl = qs('#ensaio-horario');
    const diaEl = qs('#ensaio-dia');
    const semanaEl = qs('#ensaio-semana');
    const anoEl = qs('#ensaio-ano');
    const mesesWrap = qs('#ensaio-meses');
    const tipo = tipoEl ? (tipoEl.value||'') : '';
    const horario = horarioEl ? (horarioEl.value||'') : '';
    const diaSemana = diaEl ? (diaEl.value||'') : '';
    const semana = semanaEl ? parseInt(semanaEl.value||'',10) : NaN;
    const ano = anoEl ? parseInt(anoEl.value||'',10) : NaN;
    const meses = mesesWrap ? Array.from(mesesWrap.querySelectorAll('input[name="ensaioMes"]:checked')).map(inp=>parseInt(inp.value,10)).filter(n=>!isNaN(n)) : [];
    if(!tipo) return null;
    const datas = [];
    if(diaSemana && semana && ano && meses.length){
      const wd = diasIndice[diaSemana];
      meses.forEach(m => {
        const dt = nthWeekdayOfMonth(ano, m-1, wd, semana);
        if(dt){ datas.push({ mes:m, data: formatYmd(dt) }); }
      });
    }
    return { tipo, meses, horario, diaSemana, semana: (isNaN(semana)?undefined:semana), ano: (isNaN(ano)?undefined:ano), datas };
  }
  if(addCultoBtn && cultosWrapper){
    addCultoBtn.addEventListener('click', ()=>{
      cultosWrapper.appendChild(createCultoRow({ tipo:'Culto Oficial', dia:'Domingo', horario:'' }));
      renderCultosPreview();
    });
  }
  // Eventos para atualizar preview do Ensaio
  ['ensaio-tipo','ensaio-horario','ensaio-dia','ensaio-semana','ensaio-ano'].forEach(id => {
    const el = qs(`#${id}`);
    if(el){ el.addEventListener('change', renderEnsaioPreview); el.addEventListener('input', renderEnsaioPreview); }
  });
  const ensMesesWrap = qs('#ensaio-meses');
  if(ensMesesWrap){ ensMesesWrap.addEventListener('change', renderEnsaioPreview); }

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
              <button class="btn btn-sm btn-outline-danger" data-action="delete-min" data-id="${m.id}">Excluir</button>
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
  function cidadeDoEventoRel(ev){
    const cong = (typeof congregacoesByIdEvents!=='undefined' && congregacoesByIdEvents) ? congregacoesByIdEvents[ev.congregacaoId] : null;
    if(cong && cong.cidade) return cong.cidade;
    const label = ev.congregacaoNome || (cong ? (cong.nomeFormatado || (cong.cidade && cong.bairro ? `${cong.cidade} - ${cong.bairro}` : (cong.nome||ev.congregacaoId))) : ev.congregacaoId);
    const parts = (label||'').split(' - ');
    return parts[0]||'';
  }
  function getRelEventosFiltrados(){
    const all = (relEventosFilteredCache || eventosCache || []).slice();
    const y = relYearSel && relYearSel.value ? parseInt(relYearSel.value,10) : null;
    const m = relMonthSel && relMonthSel.value ? parseInt(relMonthSel.value,10) : null;
    const cidade = relCidadeSel && relCidadeSel.value ? relCidadeSel.value : '';
    const tipo = relTipoSel && relTipoSel.value ? relTipoSel.value : '';
    return all.filter(ev=>{
      const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
      if(y && (!d || d.getFullYear() !== y)) return false;
      if(m && (!d || (d.getMonth()+1) !== m)) return false;
      if(cidade && cidadeDoEventoRel(ev) !== cidade) return false;
      if(tipo && ev.tipo !== tipo) return false;
      return true;
    }).sort((a,b)=>{
      const ad = parseDateYmdLocal(a.data) || new Date(a.data);
      const bd = parseDateYmdLocal(b.data) || new Date(b.data);
      return ad - bd;
    });
  }
  function initRelatoriosFilters(){
    if(!(relYearSel||relMonthSel||relCidadeSel||relTipoSel)) return;
    const list = eventosCache || [];
    const anos = new Set();
    const cidades = new Set();
    const bairros = new Set();
    const tipos = new Set();
    list.forEach(ev=>{
      const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
      if(d && !isNaN(d)) anos.add(d.getFullYear());
      const c = cidadeDoEventoRel(ev);
      if(c) cidades.add(c);
      if(ev.tipo) tipos.add(ev.tipo);
    });
    // Também considerar anos de ensaio das congregações
    (congregacoesCache||[]).forEach(c=>{ const a = c.ensaio && parseInt(c.ensaio.ano,10); if(a) anos.add(a); });
    // Também considerar cidades das congregações
    (congregacoesCache||[]).forEach(c=>{ const city = (c.cidade||'').trim(); if(city) cidades.add(city); });
    // Também considerar bairros das congregações
    (congregacoesCache||[]).forEach(c=>{ const bairro = (c.bairro||'').trim(); if(bairro) bairros.add(bairro); });
    if(relYearSel){
      const sel = relYearSel.value;
      relYearSel.innerHTML = '<option value="">Todos</option>' + Array.from(anos).sort((a,b)=>a-b).map(y=> `<option value="${y}">${y}</option>`).join('');
      if(sel) relYearSel.value = sel;
    }
    if(relMonthSel){
      const sel = relMonthSel.value;
      const mesesNomes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
      relMonthSel.innerHTML = '<option value="">Todos</option>' + mesesNomes.map((n,i)=> `<option value="${i+1}">${n}</option>`).join('');
      if(sel) relMonthSel.value = sel;
    }
    if(relCidadeSel){
      const sel = relCidadeSel.value;
      relCidadeSel.innerHTML = '<option value="">Todas</option>' + Array.from(cidades).sort((a,b)=> a.localeCompare(b,'pt-BR')).map(c=> `<option value="${c}">${c}</option>`).join('');
      if(sel) relCidadeSel.value = sel;
    }
    if(relBairroSel){
      const sel = relBairroSel.value;
      relBairroSel.innerHTML = '<option value="">Todos</option>' + Array.from(bairros).sort((a,b)=> a.localeCompare(b,'pt-BR')).map(b=> `<option value="${b}">${b}</option>`).join('');
      if(sel) relBairroSel.value = sel;
    }
    if(relTipoSel){
      const sel = relTipoSel.value;
      relTipoSel.innerHTML = '<option value="">Todos</option>' + Array.from(tipos).sort((a,b)=> a.localeCompare(b,'pt-BR')).map(t=> `<option value="${t}">${t}</option>`).join('');
      if(sel) relTipoSel.value = sel;
    }
  }
  function renderRelatorios(){
    renderRelatorioEventos();
    renderRelatorioMinisterio();
    renderRelatorioCongregacoes();
    renderRelatorioCongEnsaios();
  }
  function renderRelatorioEventos(){
    if(!relEventosEl) return;
    const list = (typeof getRelEventosFiltrados === 'function') ? getRelEventosFiltrados() : (eventosCache || []);
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

  function renderRelatorioCongEnsaios(){
    if(!relCongEnsaiosEl) return;
    let list = congregacoesCache || [];
    if(!list.length){
      relCongEnsaiosEl.innerHTML = '<li class="text-muted">Nenhuma congregação cadastrada</li>';
      return;
    }
    const labelFor = (c) => c.nomeFormatado || (c.cidade && c.bairro ? `${c.cidade} - ${c.bairro}` : (c.nome||c.id));
    const mesesNomes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    const fmtDate = (ymd)=> (ymd && /^\d{4}-\d{2}-\d{2}$/.test(ymd)) ? formatDate(ymd) : (ymd||'-');
    const capitalizeFirst = (s)=> !s ? '' : (s.charAt(0).toUpperCase() + s.slice(1));
    const dowName = (ymd)=>{
      const d = parseDateYmdLocal(ymd);
      if(!d) return '';
      const w = d.toLocaleDateString('pt-BR', { weekday: 'long' });
      return capitalizeFirst(w);
    };
    const ySel = (relYearSel && relYearSel.value) ? parseInt(relYearSel.value,10) : null;
    const mSel = (relMonthSel && relMonthSel.value) ? parseInt(relMonthSel.value,10) : null;
    const citySel = (relCidadeSel && relCidadeSel.value) ? String(relCidadeSel.value) : '';
    const bairroSel = (relBairroSel && relBairroSel.value) ? String(relBairroSel.value) : '';
    const sortMode = (relEnsaiosSortSel && relEnsaiosSortSel.value) ? relEnsaiosSortSel.value : 'cidade';
    const nextOnly = !!(relEnsaiosNextOnly && relEnsaiosNextOnly.checked);

    // Filtrar por ano do ensaio, se selecionado
    if(ySel){ list = list.filter(c => c.ensaio && parseInt(c.ensaio.ano,10) === ySel); }
    // Filtrar por cidade, se selecionado
    if(citySel){ list = list.filter(c => (c.cidade||'') === citySel); }
    // Filtrar por bairro, se selecionado
    if(bairroSel){ list = list.filter(c => (c.bairro||'') === bairroSel); }

    const now = new Date();
    const nextDate = (e) => {
      const arr = (Array.isArray(e && e.datas ? e.datas : []) ? e.datas : [])
        .map(d=> parseDateYmdLocal(d.data))
        .filter(d=> d && d.getTime() >= now.getTime())
        .sort((a,b)=> a - b);
      return arr[0] || null;
    };

    const sorted = [...list].sort((a,b)=>{
      if(sortMode === 'proximo'){
        const na = nextDate(a.ensaio);
        const nb = nextDate(b.ensaio);
        const va = na ? na.getTime() : Number.POSITIVE_INFINITY;
        const vb = nb ? nb.getTime() : Number.POSITIVE_INFINITY;
        if(va === vb) return (labelFor(a)||'').localeCompare(labelFor(b)||'', undefined, { sensitivity: 'base' });
        return va - vb;
      }
      return (labelFor(a)||'').localeCompare(labelFor(b)||'', undefined, { sensitivity: 'base' });
    });

    const lines = sorted.map(c => {
      const e = c.ensaio || {};
      const datas = Array.isArray(e.datas) ? e.datas.slice() : [];
      let filtered = datas
        .filter(d=> !mSel || parseInt(d.mes,10) === mSel)
        .sort((a,b)=> (a.mes||0)-(b.mes||0));
      if(nextOnly){
        const arr = filtered
          .map(d=> ({...d, _date: parseDateYmdLocal(d.data)}))
          .filter(x=> x._date && x._date.getTime() >= now.getTime())
          .sort((a,b)=> a._date - b._date);
        filtered = arr.length ? [arr[0]] : [];
      }
      const horarioStr = e.horario ? ` às ${e.horario}` : '';
      const items = filtered.length
        ? filtered.map(d=> `${mesesNomes[(d.mes||1)-1]||d.mes}: ${fmtDate(d.data)} (${dowName(d.data)})${horarioStr}`).join(', ')
        : '—';
      const tipo = e.tipo ? ` (${e.tipo})` : '';
      return `<li><strong>${labelFor(c)}</strong>${tipo}: ${items}</li>`;
    });
    relCongEnsaiosEl.innerHTML = lines.join('');

    // Export helpers
    const getExportRows = () => {
      const rows = [];
      sorted.forEach(c => {
        const e = c.ensaio || {};
        const datas = Array.isArray(e.datas) ? e.datas.slice() : [];
        let filtered = datas
          .filter(d=> !mSel || parseInt(d.mes,10) === mSel)
          .sort((a,b)=> (a.mes||0)-(b.mes||0));
        if(nextOnly){
          const arr = filtered
            .map(d=> ({...d, _date: parseDateYmdLocal(d.data)}))
            .filter(x=> x._date && x._date.getTime() >= now.getTime())
            .sort((a,b)=> a._date - b._date);
          filtered = arr.length ? [arr[0]] : [];
        }
        filtered.forEach(d => {
          rows.push({
            Congregacao: labelFor(c),
            Tipo: e.tipo||'',
            Ano: e.ano||'',
            Mes: d.mes||'',
            Data: fmtDate(d.data),
            DiaSemana: dowName(d.data),
            Horario: e.horario||'',
          });
        });
      });
      return rows;
    };
    const toCsv = (rows) => {
      if(!rows.length) return 'Congregacao,Tipo,Ano,Mes,Data,DiaSemana,Horario\n';
      const headers = Object.keys(rows[0]);
      const esc = (v)=> String(v||'').replace(/"/g,'""');
      const lines = [headers.join(',')].concat(rows.map(r=> headers.map(h=> `"${esc(r[h])}"`).join(',')));
      return lines.join('\n');
    };
    const toXls = (rows) => {
      // Simples TSV (compatível com Excel)
      if(!rows.length) return 'Congregacao\tTipo\tAno\tMes\tData\tDiaSemana\tHorario\n';
      const headers = Object.keys(rows[0]);
      const esc = (v)=> String(v||'').replace(/\t/g,' ');
      const lines = [headers.join('\t')].concat(rows.map(r=> headers.map(h=> esc(r[h])).join('\t')));
      return lines.join('\n');
    };
    const download = (filename, content, mime) => {
      const blob = new Blob([content], { type: mime || 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename; document.body.appendChild(a); a.click();
      setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
    };
    if(btnEnsaiosExportCsv){
      btnEnsaiosExportCsv.onclick = () => {
        const rows = getExportRows();
        const csv = toCsv(rows);
        download('ensaios.csv', csv, 'text/csv;charset=utf-8');
      };
    }
    if(btnEnsaiosExportXls){
      btnEnsaiosExportXls.onclick = () => {
        const rows = getExportRows();
        const xls = toXls(rows);
        download('ensaios.xls', xls, 'application/vnd.ms-excel');
      };
    }
    if(btnEnsaiosCopy){
      btnEnsaiosCopy.onclick = async () => {
        const rows = getExportRows();
        const text = rows.map(r=> `${r.Congregacao} | ${r.Tipo} | ${r.Ano}-${r.Mes} | ${r.Data} (${r.DiaSemana}) ${r.Horario}`).join('\n');
        try{
          if(navigator.clipboard && navigator.clipboard.writeText){
            await navigator.clipboard.writeText(text);
            toast('Copiado para a área de transferência');
          } else {
            const ta = document.createElement('textarea');
            ta.value = text; document.body.appendChild(ta); ta.select();
            document.execCommand('copy'); ta.remove(); toast('Copiado');
          }
        }catch(err){ console.error(err); toast('Falha ao copiar', 'error'); }
      };
    }
  }

  // Relatórios: impressão e exportação com filtros aplicados
  function buildRowsHtmlFromEventos(list){
    const diasSemanaPt = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
    return (list||[]).map(ev => {
      const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
      const diaNome = d ? diasSemanaPt[d.getDay()] : '-';
      const cong = (typeof congregacoesByIdEvents!=='undefined' && congregacoesByIdEvents) ? congregacoesByIdEvents[ev.congregacaoId] : null;
      const localLabel = ev.congregacaoNome || (cong ? (cong.nomeFormatado || (cong.cidade && cong.bairro ? `${cong.cidade} - ${cong.bairro}` : (cong.nome||ev.congregacaoId))) : ev.congregacaoId);
      const hora = (typeof horaDoEvento==='function' ? (horaDoEvento(ev)||'-') : '-');
      const atendente = (ev.atendenteNome||'-').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      return `<tr>
        <td>${diaNome} - ${formatDate(ev.data)}</td>
        <td>${hora}</td>
        <td>${localLabel}</td>
        <td>${atendente}${ev.atendenteOutraRegiao ? ' <span class="flag-outra-regiao" title="Irmão de outra região"><svg viewBox="0 0 24 24"><path d="M12 2c-4.4 0-8 3.1-8 7 0 5 8 13 8 13s8-8 8-13c0-3.9-3.6-7-8-7zm0 9a2 2 0 1 1 0-4 2 2 0 0 1 0 4z" fill="currentColor"/></svg></span>' : ''}</td>
        <td>${ev.tipo||'-'}</td>
      </tr>`;
    }).join('');
  }
  if(btnRelPrint){
    btnRelPrint.addEventListener('click', ()=>{
      try{
        const list = (typeof getRelEventosFiltrados==='function') ? getRelEventosFiltrados() : (eventosCache||[]);
        if(!list.length){ toast('Nenhum atendimento para imprimir', 'error'); return; }
        const rowsHtml = buildRowsHtmlFromEventos(list);
        const now = new Date();
        const dd = String(now.getDate()).padStart(2,'0');
        const mm = String(now.getMonth()+1).padStart(2,'0');
        const yyyy = now.getFullYear();
        const hh = String(now.getHours()).padStart(2,'0');
        const mi = String(now.getMinutes()).padStart(2,'0');
        const geradoEm = `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
        const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"/><title>Relatório de Atendimentos</title>
<style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;margin:24px;color:#000}h1{font-size:18px;margin:0 0 8px}p{font-size:12px;margin:0 0 12px;color:#333}table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:6px;font-size:12px}th{background:#f2f2f2}@media print{.no-print{display:none}}</style>
</head><body>
<h1>Relatório de Atendimentos</h1>
<p>Gerado em: ${geradoEm}</p>
<table><thead><tr><th>Data</th><th>Hora</th><th>Local</th><th>Quem atende</th><th>Tipo</th></tr></thead><tbody>${rowsHtml}</tbody></table>
</body></html>`;
        const blob = new Blob([html], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = 'relatorio-eventos-filtrados.html'; document.body.appendChild(a); a.click();
        setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
        toast('Arquivo gerado para impressão');
      }catch(err){ console.error(err); toast('Falha ao gerar arquivo de impressão', 'error'); }
    });
  }
  if(btnRelPdf){
    btnRelPdf.addEventListener('click', ()=>{
      try{
        const list = (typeof getRelEventosFiltrados==='function') ? getRelEventosFiltrados() : (eventosCache||[]);
        if(!list.length){ toast('Nenhum atendimento para exportar', 'error'); return; }
        const rowsHtml = buildRowsHtmlFromEventos(list);
        const html = `<!DOCTYPE html><html lang='pt-BR'><head><meta charset='UTF-8'/><title>Relatórios - PDF</title>
<style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;color:#000;margin:24px}h1{font-size:18px;margin:0 0 8px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:6px;font-size:12px}th{background:#f2f2f2}.no-print{margin:12px 0}@media print{.no-print{display:none}}</style>
</head><body>
<h1>Relatórios - Atendimentos</h1>
<div class='no-print'>Use Ctrl+P e escolha "Salvar como PDF".</div>
<table><thead><tr><th>Data</th><th>Hora</th><th>Local</th><th>Quem atende</th><th>Tipo</th></tr></thead><tbody>${rowsHtml}</tbody></table>
<script>window.addEventListener('load', function(){ setTimeout(function(){ window.print(); }, 200); });</script>
</body></html>`;
        const w = window.open('', '_blank');
        if(!w){ toast('Não foi possível abrir a janela de impressão (pop-up bloqueado)', 'error'); return; }
        w.document.open(); w.document.write(html); w.document.close();
        toast('Abra a caixa de impressão e salve em PDF');
      }catch(err){ console.error(err); toast('Falha ao exportar em PDF', 'error'); }
    });
  }
  if(btnRelXls){
    btnRelXls.addEventListener('click', ()=>{
      try{
        const list = (typeof getRelEventosFiltrados==='function') ? getRelEventosFiltrados() : (eventosCache||[]);
        if(!list.length){ toast('Nenhum atendimento para exportar', 'error'); return; }
        const rowsHtml = buildRowsHtmlFromEventos(list);
        const xlsHtml = `<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>
          <table border='1'>
            <thead><tr><th>Data</th><th>Hora</th><th>Local</th><th>Quem atende</th><th>Tipo</th></tr></thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </body></html>`;
        const blob = new Blob([xlsHtml], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = 'relatorio-eventos-filtrados.xls'; document.body.appendChild(a); a.click();
        setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
        toast('Arquivo XLS gerado');
      }catch(err){ console.error(err); toast('Falha ao exportar XLS', 'error'); }
    });
  }

  // Editar/Excluir Ministério: delegar clique para editar ou remover
  if(listaMin){
    listaMin.addEventListener('click', async (e)=>{
      const btnDel = e.target.closest('button[data-action="delete-min"]');
      const btnEdit = e.target.closest('button[data-action="edit-min"]');
      if(btnDel){
        const id = btnDel.getAttribute('data-id');
        const m = ministryCache.find(x=>x.id===id);
        if(!m) return;
        const ok = confirm('Tem certeza que deseja excluir este irmão?');
        if(!ok) return;
        try{
          await remove('ministerio', id);
          toast('Irmão excluído com sucesso');
          // Se estiver editando este registro, reseta o formulário
          const formMinEl = qs('#form-ministerio');
          if(formMinEl && formMinEl.dataset.editId === id){
            delete formMinEl.dataset.editId;
            const submitBtn = formMinEl.querySelector('button[type="submit"]');
            const resetBtn = formMinEl.querySelector('button[type="reset"]');
            if(submitBtn) submitBtn.textContent = 'Salvar Irmão do Ministério';
            if(resetBtn) resetBtn.textContent = 'Limpar';
            formMinEl.reset && formMinEl.reset();
          }
        }catch(err){
          console.error(err);
          toast('Erro ao excluir. Tente novamente.', 'error');
        }
        return;
      }
      if(btnEdit){
        const id = btnEdit.getAttribute('data-id');
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
      }
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
        const ensaio = collectEnsaio();
        if(editId){
          const updated = await update('congregacoes', editId, {
            cidade: data.cidade,
            bairro: data.bairro,
            endereco: data.endereco,
            cultos,
            ...(ensaio ? { ensaio } : {})
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
            ...(ensaio ? { ensaio } : {})
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
        // Resumo de Ensaio (tipo, meses e horário)
        const ensaioLine = (function(){
          const e = c.ensaio;
          if(!e || !e.tipo) return '';
          const mesesNomes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
          const mesesStr = Array.isArray(e.meses) && e.meses.length
            ? e.meses.slice().sort((a,b)=>a-b).map(n=> mesesNomes[(n||1)-1] || n).join(', ')
            : '-';
          const horarioStr = e.horario ? ` às ${e.horario}` : '';
          const icon = `<span class="icon" aria-hidden="true"><svg viewBox="0 0 24 24" width="16" height="16"><circle cx="12" cy="12" r="9" fill="none" stroke="currentColor" stroke-width="2"/><path d="M12 7v5l4 2" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round"/></svg></span>`;
          return `<div class="meta ensaio-line">${icon}Ensaio ${e.tipo}: ${mesesStr}${horarioStr}</div>`;
        })();
        // Datas calculadas de Ensaio por mês (se disponíveis)
        const ensaioDatasLine = (function(){
          const e = c.ensaio;
          if(!e || !Array.isArray(e.datas) || !e.datas.length) return '';
          const mesesNomes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
          const fmt = (ymd)=> (ymd && /^\d{4}-\d{2}-\d{2}$/.test(ymd)) ? formatDate(ymd) : (ymd||'-');
          const items = e.datas.slice().sort((a,b)=> (a.mes||0)-(b.mes||0)).map(d=> `${mesesNomes[(d.mes||1)-1]||d.mes}: ${fmt(d.data)}`).join(', ');
          return `<div class="meta ensaio-line">Datas de Ensaio: ${items}</div>`;
        })();
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
              ${ensaioLine}
              ${ensaioDatasLine}
            </div>
            <div>
              <button class="btn btn-sm btn-outline-secondary" data-action="edit-cong" data-id="${c.id}">Editar</button>
              <button class="btn btn-sm btn-outline-danger" data-action="delete-cong" data-id="${c.id}">Excluir</button>
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
      // Pré-carregar Ensaio no formulário (tipo, meses e horário)
      const ensTipoSel = qs('#ensaio-tipo');
      const ensHorarioInp = qs('#ensaio-horario');
      const ensDiaSel = qs('#ensaio-dia');
      const ensSemanaSel = qs('#ensaio-semana');
      const ensAnoInp = qs('#ensaio-ano');
      const ensMesesWrap = qs('#ensaio-meses');
      const ens = c.ensaio || null;
      if(ensTipoSel) ensTipoSel.value = (ens && ens.tipo) ? ens.tipo : '';
      if(ensHorarioInp) ensHorarioInp.value = (ens && ens.horario) ? ens.horario : '';
      if(ensDiaSel) ensDiaSel.value = (ens && ens.diaSemana) ? ens.diaSemana : '';
      if(ensSemanaSel) ensSemanaSel.value = (ens && ens.semana) ? String(ens.semana) : '';
      if(ensAnoInp) ensAnoInp.value = (ens && ens.ano) ? String(ens.ano) : '';
      if(ensMesesWrap){
        const allChecks = Array.from(ensMesesWrap.querySelectorAll('input[name="ensaioMes"]'));
        allChecks.forEach(ch => { ch.checked = false; });
        const months = (ens && Array.isArray(ens.meses)) ? ens.meses : [];
        months.forEach(m => {
          const chk = ensMesesWrap.querySelector(`input[name="ensaioMes"][value="${m}"]`);
          if(chk) chk.checked = true;
        });
      }
      renderEnsaioPreview();
    }
    if(listaCong){
      listaCong.addEventListener('click', async (e)=>{
        const btnDel = e.target.closest('button[data-action="delete-cong"]');
        const btnEdit = e.target.closest('button[data-action=\"edit-cong\"]');
        if(btnDel){
          const id = btnDel.getAttribute('data-id');
          const c = congregacoesCache.find(x=>x.id===id);
          if(!c) return;
          const ok = confirm('Tem certeza que deseja excluir esta congregação?');
          if(!ok) return;
          try{
            await remove('congregacoes', id);
            toast('Congregação excluída com sucesso');
            if(formCong && formCong.dataset.editId === id){
              delete formCong.dataset.editId;
              const submitBtn = formCong.querySelector('button[type="submit"]');
              const resetBtn = formCong.querySelector('button[type="reset"]');
              if(submitBtn) submitBtn.textContent = 'Salvar Atendimento';
              if(resetBtn) resetBtn.textContent = 'Limpar';
              formCong.reset && formCong.reset();
              if(cultosWrapper) cultosWrapper.innerHTML = '';
              const tbody = qs('#cultos-preview-body');
              if(tbody) tbody.innerHTML = '<tr><td colspan="3" class="text-muted">Nenhum horário adicionado</td></tr>';
              if(anciaosWrapper) anciaosWrapper.innerHTML = '';
              if(diaconosWrapper) diaconosWrapper.innerHTML = '';
            }
          }catch(err){
            console.error(err);
            toast('Erro ao excluir. Tente novamente.', 'error');
          }
          return;
        }
        if(btnEdit){
          const id = btnEdit.getAttribute('data-id');
          const c = congregacoesCache.find(x=>x.id===id);
          startEditCongregacao(c);
        }
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
      const sorted = [...(list||[])].sort((a,b)=> (labelFor(a)||'').localeCompare(labelFor(b)||'', undefined, { sensitivity: 'base' }));
      agendarGrid.innerHTML = sorted.map(c => 
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
            hint.textContent = 'Obs: Ituiutaba - Centro pode agendar 2 coletas por mês; uma deve ser na quinta-feira (Culto Oficial).';
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

        // Bloqueio: RJM somente se a congregação possuir RJM cadastrado
        if(tipo === 'RJM com Reforço de Coletas'){
          const hasRjm = c && Array.isArray(c.cultos) && c.cultos.some(ct => ct.tipo === 'RJM');
          if(!hasRjm){ toast('Esta congregação não tem RJM.', 'error'); return; }
        }
        // Bloqueio: se já houver qualquer atendimento nesta congregação nesta data
        const sameDaySameCong = (eventosCache||[]).some(ev => ev.congregacaoId === congId && ev.data === data);
        if(sameDaySameCong){ toast('ALERTA: já existe atendimento nesta congregação nesta data.', 'error'); return; }

        // Exceção anterior substituída: ver regra abaixo que permite até 2 CO/mês em Ituiutaba Centro com exigência de ao menos uma quinta-feira.

        const dNew = parseDateYmdLocal(data);
        const sameMonthSameCong = (eventosCache||[]).filter(ev=>{
          const dEv = parseDateYmdLocal(ev.data);
          return ev.congregacaoId===congId && dEv && dNew && dEv.getFullYear()===dNew.getFullYear() && dEv.getMonth()===dNew.getMonth();
        });
        // Regra: permitir múltiplas coletas; duplicidade por tipo bloqueada,
        // exceto para Ituiutaba Centro em 'Culto Reforço de Coletas' (máx. 2 no mês, com ao menos um na quinta)
        if(tipo==='Culto Reforço de Coletas'){
          const cSel = congregacoesByIdEvents[congId];
          const nomeFmt = ((cSel && cSel.nomeFormatado)||'').toLowerCase();
          const cidade = ((cSel && cSel.cidade)||'').toLowerCase();
          const bairro = ((cSel && cSel.bairro)||'').toLowerCase();
          const isItuiCentro = (cidade === 'ituiutaba' && bairro === 'centro') || (nomeFmt.includes('ituiutaba') && nomeFmt.includes('centro'));
          const coList = sameMonthSameCong.filter(ev=>ev.tipo==='Culto Reforço de Coletas');
          if(isItuiCentro){
            if(coList.length >= 2){
              toast('Ituiutaba Centro: máximo de 2 coletas em Culto Oficial no mês.', 'error');
              return;
            }
            const isThu = dNew && dNew.getDay() === 4;
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
        } else {
          const typeAlready = sameMonthSameCong.some(ev => ev.tipo===tipo);
          if(typeAlready){
            toast('Já existe este tipo agendado nesta congregação neste mês.', 'error');
            return;
          }
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
      <td><span class="congregacao-highlight">${localLabel}</span></td>
      <td>${ev.atendenteNome||'-'}${ev.atendenteOutraRegiao ? ' <span class=\"flag-outra-regiao\" title=\"Irmão de outra região\"><svg viewBox=\"0 0 24 24\"><path d=\"M12 2c-4.4 0-8 3.1-8 7 0 5 8 13 8 13s8-8 8-13c0-3.9-3.6-7-8-7zm0 9a2 2 0 1 1 0-4 2 2 0 0 1 0 4z\" fill=\"currentColor\"/></svg></span>' : ''}</td>
      <td>${ev.tipo}</td>
    </tr>`;
    }).join('');
  };

  // Agenda 2026 (Página agenda2026.html)
  const agenda2026Body = qs('#agenda2026-body');
  const btnImprimirAgenda2026 = qs('#btn-imprimir-agenda2026');
  const btnExportarPdfAgenda2026 = qs('#btn-exportar-pdf-agenda2026');
  const btnExportarXlsAgenda2026 = qs('#btn-exportar-xls-agenda2026');
  const btnImportarAgenda2026 = qs('#btn-importar-agenda2026');
  const fileImportAgenda2026 = qs('#file-import-agenda2026');
  const btnNovoAgenda2026 = qs('#btn-novo-agenda2026');
  const agendaForm = qs('#agenda2026-form');
  const agendaSalvarBtn = qs('#agenda2026-salvar');
  const agendaCancelarBtn = qs('#agenda2026-cancelar');
  const agendaLimparBtn = qs('#agenda2026-limpar');
  const agendaDataInp = qs('#agenda2026-data');
  const agendaHoraInp = qs('#agenda2026-hora');
  const agendaTipoSel = qs('#agenda2026-tipo');
  const agendaSetorSel = qs('#agenda2026-setor');
  const agendaDescInp = qs('#agenda2026-descricao');
  const agendaCidadeInp = qs('#agenda2026-cidade');
  const agendaCongInp = qs('#agenda2026-congregacao');
  const agendaLocalSel = qs('#agenda2026-local');
  const agendaPublicoInp = qs('#agenda2026-publico');
  const agendaRespInp = qs('#agenda2026-responsavel');
  // Botões de adicionar opção (+) e datalist de congregações
  const btnAddTipo = qs('#btn-add-tipo');
  const btnAddSetor = qs('#btn-add-setor');
  const btnAddCidade = qs('#btn-add-cidade');
  const agendaCongDatalist = qs('#agenda2026-congregacoes-list');
  let agenda2026Cache = [];
  let agendaEditId = null;

  function formatAgendaDate(d){
    if(!d) return '-';
    const monthsAbbr = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
    return `${String(d.getDate()).padStart(2,'0')}/${monthsAbbr[d.getMonth()]}`;
  }
  function parseDataAgendaValor(val){
    try{
      if(typeof val === 'number' && window.XLSX && XLSX.SSF){
        const dc = XLSX.SSF.parse_date_code(val);
        if(dc){ return new Date(2026, (dc.m||1)-1, dc.d||1); }
      }
      if(typeof val === 'string'){
        const s = val.trim().toLowerCase();
        const map = { jan:0, fev:1, mar:2, abr:3, mai:4, jun:5, jul:6, ago:7, set:8, out:9, nov:10, dez:11 };
        let m=null, d=null;
        let m1 = s.match(/^(\d{1,2})\s*\/?\s*([a-zç]{3})$/i);
        if(m1){ d=parseInt(m1[1],10); const mm=map[m1[2].replace('ç','c')]; if(mm!==undefined) m=mm; }
        if(m===null){
          let m2 = s.match(/^(\d{1,2})\s*\/\s*(\d{1,2})(?:\/(\d{2,4}))?$/);
          if(m2){ d=parseInt(m2[1],10); m=parseInt(m2[2],10)-1; }
        }
        if(m!==null && d!==null){ return new Date(2026, m, d); }
      }
    }catch{}
    return null;
  }
  function horaDoEvento(ev){
    try{
      const d = parseDateYmdLocal(ev.data);
      const cong = congregacoesByIdEvents[ev.congregacaoId];
      const diasSemanaPt = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
      const diaNome = d ? diasSemanaPt[d.getDay()] : '';
      const tipoCulto = ev.tipo==='Culto Reforço de Coletas' ? 'Culto Oficial' : (ev.tipo==='RJM com Reforço de Coletas' ? 'RJM' : '');
      if(cong && Array.isArray(cong.cultos)){
        const match = cong.cultos.find(ct => ct.tipo===tipoCulto && ct.dia===diaNome);
        if(match && match.horario) return match.horario;
      }
    }catch{}
    return '-';
  }
  function labelCong(ev){
    const cong = congregacoesByIdEvents[ev.congregacaoId];
    return ev.congregacaoNome || (cong ? (cong.nomeFormatado || (cong.cidade && cong.bairro ? `${cong.cidade} - ${cong.bairro}` : (cong.nome||ev.congregacaoId))) : ev.congregacaoId);
  }

  function renderAgenda2026(){
    if(!agenda2026Body) return;
    const list = (agenda2026Cache||[]).slice().sort((a,b)=>{
      const ad = parseDateYmdLocal(a.data) || parseDataAgendaValor(a.data) || new Date(2026,0,1);
      const bd = parseDateYmdLocal(b.data) || parseDataAgendaValor(b.data) || new Date(2026,0,1);
      return ad - bd;
    });
    if(!list.length){
      agenda2026Body.innerHTML = '<tr><td colspan="8" class="text-muted">Nenhum serviço agendado para 2026</td></tr>';
      return;
    }
    const daysAbbr = ['dom','seg','ter','qua','qui','sex','sáb'];
    let n=1;
    const daysFull = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
    agenda2026Body.innerHTML = list.map(ev=>{
      const d = parseDateYmdLocal(ev.data) || parseDataAgendaValor(ev.data);
      const dataFmt = d ? formatAgendaDate(d) : '-';
      const di = d ? daysAbbr[d.getDay()] : '-';
      const diaTitle = d ? daysFull[d.getDay()] : '';
      const hora = ev.hora || '-';
      const cidade = ev.cidade || '-';
      const localLabel = ev.congregacao || '-';
      const today = new Date();
      const dOnly = d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null;
      const tOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const rowClass = dOnly && dOnly < tOnly ? 'event-past' : 'event-future';
      const svcTipo = ev.tipo || ((ev.servico||'').split(' - ')[0] || '');
      const svcDesc = ev.descricao || (function(){ const parts=(ev.servico||'').split(' - '); parts.shift(); return parts.join(' - '); })();
      const slug = (svcTipo||'').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/(^-|-$)/g,'') || 'outros';
      return `<tr class="${rowClass}">
        <td class="index-col">${n++}</td>
        <td class="date-col"><span class="date">${dataFmt}</span></td>
        <td class="day-col"><span class="badge-day" title="${diaTitle}">${di}</span></td>
        <td class="time-col"><span class="time">${hora}</span></td>
        <td class="service-col"><span class="servico"><span class="pill-servico servico-${slug}" title="${svcTipo||'-'}">${svcTipo||'-'}</span>${svcDesc?`<span class="servico-desc" title="${svcDesc}"> — ${svcDesc}</span>`:''}</span></td>
        <td class="city-col"><span class="chip chip-cidade">${cidade}</span></td>
        <td class="cong-col"><span class="congregacao-highlight">${localLabel}</span></td>
        <td class="actions">
          <button type="button" class="btn btn-sm btn-outline-primary btn-icon btn-edit-agenda2026" data-id="${ev.id}" title="Editar">
            <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zm15.71-9.04a1.003 1.003 0 0 0 0-1.42l-1.5-1.5a1.003 1.003 0 0 0-1.42 0l-1.83 1.83 3.75 3.75 1.99-1.66z" fill="currentColor"/></svg>
            <span class="visually-hidden">Editar</span>
          </button>
          <button type="button" class="btn btn-sm btn-outline-danger btn-icon ms-1 btn-del-agenda2026" data-id="${ev.id}" title="Excluir">
            <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><path d="M6 19a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z" fill="currentColor"/></svg>
            <span class="visually-hidden">Excluir</span>
          </button>
        </td>
      </tr>`;
    }).join('');
  }

  function buildAgenda2026RowsHtml(list, withIndex=true){
    const daysAbbr = ['dom','seg','ter','qua','qui','sex','sáb'];
    let n=1;
    return list.map(ev=>{
      const d = parseDateYmdLocal(ev.data) || parseDataAgendaValor(ev.data);
      const dataFmt = d ? formatAgendaDate(d) : '-';
      const di = d ? daysAbbr[d.getDay()] : '-';
      const hora = ev.hora || '-';
      const cidade = ev.cidade || '-';
      const localLabel = ev.congregacao || '-';
      return `<tr>
        ${withIndex?`<td>${n++}</td>`:''}
        <td>${dataFmt}</td>
        <td>${di}</td>
        <td>${hora}</td>
        <td>${ev.servico||ev.tipo||'-'}</td>
        <td>${cidade}</td>
        <td>${localLabel}</td>
      </tr>`;
    }).join('');
  }

  if(agenda2026Body){
    // Carregar dados e renderizar
    readList('agenda2026', list => { agenda2026Cache = list; renderAgenda2026(); });

    // Preencher sugestões (datalist) de Congregações a partir do cadastro existente
    if(agendaCongDatalist){
      readList('congregacoes', list => {
        const labelFor = (c) => c.nomeFormatado || (c.cidade && c.bairro ? `${c.cidade} - ${c.bairro}` : (c.nome||c.id));
        const sorted = [...(list||[])].sort((a,b)=> (labelFor(a)||'').localeCompare(labelFor(b)||'', undefined, { sensitivity: 'base' }));
        agendaCongDatalist.innerHTML = '';
        sorted.forEach(c => {
          const opt = document.createElement('option');
          opt.value = labelFor(c);
          agendaCongDatalist.appendChild(opt);
        });
      });
    }

    // Função auxiliar para adicionar opção aos selects
    function addOptionToSelect(selectEl, label){
      if(!selectEl) return;
      const norm = (s)=> String(s||'').trim();
      const val = norm(label);
      if(!val) return;
      const exists = Array.from(selectEl.options||[]).some(o => norm(o.value).toLowerCase() === val.toLowerCase());
      if(!exists){
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        selectEl.appendChild(opt);
      }
      selectEl.value = val;
      if(typeof toast==='function') toast('Opção adicionada');
    }

    // Eventos dos botões "+" para adicionar novas opções
    btnAddTipo && btnAddTipo.addEventListener('click', ()=>{
      const v = prompt('Nova opção para Tipo de Serviço:');
      if(v!=null) addOptionToSelect(agendaTipoSel, v);
    });
    btnAddSetor && btnAddSetor.addEventListener('click', ()=>{
      const v = prompt('Nova opção para Setor:');
      if(v!=null) addOptionToSelect(agendaSetorSel, v);
    });
    btnAddCidade && btnAddCidade.addEventListener('click', ()=>{
      const v = prompt('Nova opção para Cidade:');
      if(v!=null) addOptionToSelect(agendaCidadeInp, v);
    });

    // Formulário Novo agendamento
    if(btnNovoAgenda2026 && agendaForm){
      btnNovoAgenda2026.addEventListener('click', ()=>{
        agendaForm.classList.toggle('hidden');
        if(!agendaForm.classList.contains('hidden')){
          agendaEditId = null;
          (agendaForm && agendaForm.reset && agendaForm.reset());
          if(!agendaDataInp.value){
            // se já estamos em 2026, mantém hoje; senão padroniza 2026-01-01
            const now = new Date();
            const y = now.getFullYear()===2026 ? now.getFullYear() : 2026;
            const m = String((now.getFullYear()===2026 ? now.getMonth()+1 : 1)).padStart(2,'0');
            const d = String((now.getFullYear()===2026 ? now.getDate() : 1)).padStart(2,'0');
            agendaDataInp.value = `${y}-${m}-${d}`;
          }
        }
      });
    }
    if(agendaCancelarBtn && agendaForm){
      agendaCancelarBtn.addEventListener('click', ()=>{
        agendaEditId = null;
        agendaForm.classList.add('hidden');
      });
    }
    if(agendaLimparBtn){
      agendaLimparBtn.addEventListener('click', ()=>{
        agendaEditId = null;
      });
    }
    if(agendaSalvarBtn){
      agendaSalvarBtn.addEventListener('click', async ()=>{
        try{
          const data = (agendaDataInp && agendaDataInp.value)||'';
          const hora = (agendaHoraInp && agendaHoraInp.value)||'';
          const tipo = (agendaTipoSel && agendaTipoSel.value)||'';
          const setor = (agendaSetorSel && agendaSetorSel.value)||'';
          const descricao = (agendaDescInp && agendaDescInp.value)||'';
          const cidade = (agendaCidadeInp && agendaCidadeInp.value)||'';
          const congregacao = (agendaCongInp && agendaCongInp.value)||'';
          const local = (agendaLocalSel && agendaLocalSel.value)||'';
          const publico = (agendaPublicoInp && agendaPublicoInp.value)||'';
          const responsavel = (agendaRespInp && agendaRespInp.value)||'';
          if(!data || !hora || !tipo){ toast('Preencha Data, Hora e Tipo.', 'error'); return; }
          const servico = tipo + (descricao ? ` - ${descricao}` : '');
          if(agendaEditId){
            await update('agenda2026', agendaEditId, { data, hora, tipo, setor, descricao, cidade, congregacao, local, publico, responsavel, servico });
            toast('Registro atualizado');
          }else{
            await write('agenda2026', { data, hora, tipo, setor, descricao, cidade, congregacao, local, publico, responsavel, servico });
            toast('Agendamento salvo');
          }
          (agendaForm && agendaForm.reset && agendaForm.reset());
          agendaEditId = null;
          agendaForm && agendaForm.classList.add('hidden');
        }catch(err){ console.error(err); toast('Falha ao salvar agendamento', 'error'); }
      });
    }

    // Ações Editar/Excluir na tabela
    if(agenda2026Body){
      agenda2026Body.addEventListener('click', async (e)=>{
        const btn = e.target.closest('button');
        if(!btn) return;
        const id = btn.getAttribute('data-id');
        if(!id) return;
        if(btn.classList.contains('btn-edit-agenda2026')){
          e.preventDefault();
          try{
            const ev = (agenda2026Cache||[]).find(x=>x.id===id);
            if(!ev){ toast('Registro não encontrado', 'error'); return; }
            agendaEditId = id;
            if(agendaForm && agendaForm.reset){ agendaForm.reset(); }
            agendaDataInp && (agendaDataInp.value = ev.data || '');
            agendaHoraInp && (agendaHoraInp.value = ev.hora || '');
            if(agendaTipoSel){
              const val = ev.tipo || '';
              if(val){
                const exists = Array.from(agendaTipoSel.options||[]).some(o=>o.value===val);
                if(!exists){ const opt=document.createElement('option'); opt.value=val; opt.textContent=val; agendaTipoSel.appendChild(opt); }
              }
              agendaTipoSel.value = val;
            }
            if(agendaSetorSel){
              const val = ev.setor || '';
              if(val){
                const exists = Array.from(agendaSetorSel.options||[]).some(o=>o.value===val);
                if(!exists){ const opt=document.createElement('option'); opt.value=val; opt.textContent=val; agendaSetorSel.appendChild(opt); }
              }
              agendaSetorSel.value = val;
            }
            agendaDescInp && (agendaDescInp.value = ev.descricao || '');
            if(agendaCidadeInp){
              const val = ev.cidade || '';
              if(val && agendaCidadeInp.tagName==='SELECT'){
                const exists = Array.from(agendaCidadeInp.options||[]).some(o=>o.value===val);
                if(!exists){ const opt=document.createElement('option'); opt.value=val; opt.textContent=val; agendaCidadeInp.appendChild(opt); }
              }
              agendaCidadeInp.value = val;
            }
            agendaCongInp && (agendaCongInp.value = ev.congregacao || '');
            agendaLocalSel && (agendaLocalSel.value = ev.local || '');
            agendaPublicoInp && (agendaPublicoInp.value = ev.publico || '');
            agendaRespInp && (agendaRespInp.value = ev.responsavel || '');
            agendaForm && agendaForm.classList.remove('hidden');
            toast('Editando registro...');
          }catch(err){ console.error(err); toast('Falha ao carregar registro', 'error'); }
        }else if(btn.classList.contains('btn-del-agenda2026')){
          e.preventDefault();
          if(!confirm('Deseja excluir este registro?')) return;
          try{
            await remove('agenda2026', id);
            toast('Registro excluído');
          }catch(err){ console.error(err); toast('Falha ao excluir', 'error'); }
        }
      });
    }

    // Importação XLS
    if(btnImportarAgenda2026 && fileImportAgenda2026){
      btnImportarAgenda2026.addEventListener('click', ()=> fileImportAgenda2026.click());
      fileImportAgenda2026.addEventListener('change', async (e)=>{
        const file = e.target.files && e.target.files[0];
        e.target.value = '';
        if(!file) return;
        try{
          if(!window.XLSX){ toast('Biblioteca XLSX não carregada', 'error'); return; }
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(buf, { type:'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:true });
          if(!rows || !rows.length){ toast('Planilha vazia', 'error'); return; }
          let hIdx = 0;
          for(let i=0;i<Math.min(10, rows.length); i++){
            const r = (rows[i]||[]).map(v=>String(v||'').toLowerCase());
            if(r.some(x=>x.includes('data')) && r.some(x=>x.includes('serv'))){ hIdx=i; break; }
          }
          const header = (rows[hIdx]||[]).map(v=>String(v||'').toLowerCase());
          const col = {
            data: header.findIndex(h=>h.includes('data')),
            hora: header.findIndex(h=>h.startsWith('hor')),
            serv: header.findIndex(h=>h.includes('serv')),
            cid: header.findIndex(h=>h.includes('cidade')),
            cong: header.findIndex(h=>h.includes('congreg')),
          };
          if(col.data<0 || col.serv<0 || col.cid<0 || col.cong<0){ toast('Cabeçalho não reconhecido. Verifique as colunas.', 'error'); return; }
          const items = [];
          for(let i=hIdx+1;i<rows.length;i++){
            const r = rows[i]||[];
            const vData = r[col.data];
            const vServ = r[col.serv];
            const vCid  = r[col.cid];
            const vCong = r[col.cong];
            const vHora = col.hora>=0? r[col.hora] : '';
            if(!vData || !vServ) continue;
            const d = parseDataAgendaValor(vData);
            if(!d) continue;
            const ymd = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
            items.push({ data: ymd, servico: String(vServ||'').trim(), cidade: String(vCid||'').trim(), congregacao: String(vCong||'').trim(), hora: String(vHora||'').trim() });
          }
          if(!items.length){ toast('Nenhuma linha válida encontrada.', 'error'); return; }
          const overwrite = confirm('Substituir registros atuais da Agenda 2026 pelo arquivo importado?');
          if(overwrite){
            for(const it of (agenda2026Cache||[])){
              try{ await remove('agenda2026', it.id); }catch{}
            }
          }
          for(const it of items){ await write('agenda2026', it); }
          toast(`Importação concluída: ${items.length} itens`);
        }catch(err){ console.error(err); toast('Falha ao importar XLS', 'error'); }
      });
    }

    // Exportações
    function getAgendaListSorted(){
      return (agenda2026Cache||[]).slice().sort((a,b)=>{
        const ad = parseDateYmdLocal(a.data) || parseDataAgendaValor(a.data) || new Date(2026,0,1);
        const bd = parseDateYmdLocal(b.data) || parseDataAgendaValor(b.data) || new Date(2026,0,1);
        return ad-bd;
      });
    }

    btnImprimirAgenda2026 && btnImprimirAgenda2026.addEventListener('click', ()=>{
      try{
        const list = getAgendaListSorted();
        if(!list.length){ toast('Sem registros de 2026 para imprimir', 'error'); return; }
        const rows = buildAgenda2026RowsHtml(list, true);
        const now = new Date();
        const gerado = `${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
        const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"/><title>Agenda 2026</title>
          <style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;margin:24px;color:#000}h1{font-size:18px;margin:0 0 4px}h2{font-size:14px;margin:0 0 12px;color:#333}table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:6px;font-size:12px}th{background:#f2f2f2}@media print{.no-print{display:none}}</style>
        </head><body>
          <h1>CONGREGAÇÃO CRISTÃ NO BRASIL</h1>
          <h2>AGENDA DE REUNIÕES E SERVIÇOS 2026 — Administração | Ituiutaba</h2>
          <p class="no-print">Gerado em: ${gerado}</p>
          <table><thead><tr><th>Nº</th><th>Data</th><th>Di</th><th>Hor.</th><th>Serviço</th><th>Cidade</th><th>Congregação / Salão</th></tr></thead>
          <tbody>${rows}</tbody></table>
        </body></html>`;
        const blob = new Blob([html], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a=document.createElement('a'); a.href=url; a.download='agenda-2026.html'; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(url); a.remove();},0);
        toast('Arquivo gerado para impressão');
      }catch(err){ console.error(err); toast('Falha ao gerar impressão', 'error'); }
    });

    btnExportarPdfAgenda2026 && btnExportarPdfAgenda2026.addEventListener('click', ()=>{
      try{
        const list = getAgendaListSorted();
        if(!list.length){ toast('Sem registros de 2026 para exportar', 'error'); return; }
        const rows = buildAgenda2026RowsHtml(list, true);
        const html = `<!DOCTYPE html><html lang='pt-BR'><head><meta charset='UTF-8'/><title>Agenda 2026 - PDF</title>
          <style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;margin:24px;color:#000}h1{font-size:18px;margin:0 0 4px}h2{font-size:14px;margin:0 0 12px;color:#333}table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:6px;font-size:12px}th{background:#f2f2f2}.no-print{margin:12px 0}@media print{.no-print{display:none}}</style>
        </head><body>
          <h1>CONGREGAÇÃO CRISTÃ NO BRASIL</h1>
          <h2>AGENDA DE REUNIÕES E SERVIÇOS 2026 — Administração | Ituiutaba</h2>
          <div class='no-print'>Use Ctrl+P e escolha "Salvar como PDF".</div>
          <table><thead><tr><th>Nº</th><th>Data</th><th>Di</th><th>Hor.</th><th>Serviço</th><th>Cidade</th><th>Congregação / Salão</th></tr></thead>
          <tbody>${rows}</tbody></table>
          <script>window.addEventListener('load', function(){ setTimeout(function(){ window.print(); }, 200); });</script>
        </body></html>`;
        const w = window.open('', '_blank');
        if(!w){ toast('Popup bloqueado. Permita pop-ups.', 'error'); return; }
        w.document.open(); w.document.write(html); w.document.close();
      }catch(err){ console.error(err); toast('Falha ao exportar PDF', 'error'); }
    });

    btnExportarXlsAgenda2026 && btnExportarXlsAgenda2026.addEventListener('click', ()=>{
      try{
        const list = getAgendaListSorted();
        if(!list.length){ toast('Sem registros de 2026 para exportar', 'error'); return; }
        const rows = buildAgenda2026RowsHtml(list, true);
        const xls = `<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>
          <table border='1'><thead><tr><th>Nº</th><th>Data</th><th>Di</th><th>Hor.</th><th>Serviço</th><th>Cidade</th><th>Congregação / Salão</th></tr></thead>
          <tbody>${rows}</tbody></table></body></html>`;
        const blob = new Blob([xls], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a=document.createElement('a'); a.href=url; a.download='agenda-2026.xls'; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(url); a.remove();},0);
        toast('Arquivo XLS gerado');
      }catch(err){ console.error(err); toast('Falha ao exportar XLS', 'error'); }
    });
  }

  // ===================== Agenda Mensal (Página agenda.html) =====================
  const agendaMonthRoot = qs('#agenda-month-root');
  if (agendaMonthRoot) {
    const agendaMonthLabel = qs('#agenda-month-label');
    const agendaMonthGrid = qs('#agenda-month-grid');
    const agendaTodayBtn = qs('#agenda-today-btn');

    let viewDate = new Date();
    let eventosCal = [];
    let agenda2026Cal = [];
    let congregacoesReady = false;

    function monthLabel(d) {
      const meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
      return `${meses[d.getMonth()]} de ${d.getFullYear()}`;
    }

    function yyyyMmDd(d) {
      const m = `${d.getMonth()+1}`.padStart(2,'0');
      const day = `${d.getDate()}`.padStart(2,'0');
      return `${d.getFullYear()}-${m}-${day}`;
    }

    function sameYmd(a, b) {
      return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate();
    }

    function normalizeEventosForMonth(list) {
      return (list || []).map(ev => {
        const d = parseDateYmdLocal(ev.data);
        const hora = horaDoEvento(ev) || '';
        const titulo = labelCong(ev) || (ev.congregacaoNome || '');
        const tipo = ev.tipo || 'Evento';
        return { fonte: 'atendimento', data: d, ymd: yyyyMmDd(d), hora, titulo, tipo };
      });
    }

    function normalizeAgenda2026ForMonth(list) {
      return (list || []).map(item => {
        const d = parseDataAgendaValor(item.data);
        if (!d || isNaN(d)) return null;
        const hora = item.hora || horaDoEvento(item) || '';
        const titulo = labelCong(item) || (item.congregacaoNome || '');
        const tipo = item.tipo || item.servico || 'Agenda';
        return { fonte: 'agenda2026', data: d, ymd: yyyyMmDd(d), hora, titulo, tipo };
      }).filter(Boolean);
    }

    function eventsByDayInView() {
      const y = viewDate.getFullYear();
      const m = viewDate.getMonth();
      const map = new Map();
      const push = (evt) => {
        if (!evt || !evt.data || isNaN(evt.data)) return;
        if (evt.data.getFullYear() !== y || evt.data.getMonth() !== m) return;
        const key = evt.ymd;
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(evt);
      };
      eventosCal.forEach(push);
      agenda2026Cal.forEach(push);
      // ordenar por hora, quando houver
      for (const arr of map.values()) {
        arr.sort((a,b) => (a.hora||'').localeCompare(b.hora||''));
      }
      return map;
    }

    function renderCalendar() {
      if (!agendaMonthLabel || !agendaMonthGrid) return;
      agendaMonthLabel.textContent = monthLabel(viewDate);

      const first = new Date(viewDate.getFullYear(), viewDate.getMonth(), 1);
      const start = new Date(first);
      // domingo como início da semana
      start.setDate(first.getDate() - first.getDay());
      const today = new Date();
      const byDay = eventsByDayInView();

      let html = '';
      for (let i = 0; i < 42; i++) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        const inMonth = d.getMonth() === viewDate.getMonth();
        const isToday = sameYmd(d, today);
        const key = yyyyMmDd(d);
        const dayEvents = byDay.get(key) || [];

        html += '<div class="cal-cell">';
        html += '<div class="cal-day">';
        if (isToday) {
          html += `<span class="today">${d.getDate()}</span>`;
        } else {
          html += `<span class="${inMonth ? '' : 'muted'}">${d.getDate()}</span>`;
        }
        html += '</div>';

        if (dayEvents.length) {
          for (const evt of dayEvents.slice(0, 4)) {
            const hora = evt.hora ? `<span class="meta">${evt.hora}</span>` : '';
            const tipo = `<span class="pill">${evt.tipo}</span>`;
            html += `<div class="event">${tipo}<span class="title">${evt.titulo}</span>${hora}</div>`;
          }
          if (dayEvents.length > 4) {
            html += `<div class="event"><span class="meta">+${dayEvents.length-4} mais…</span></div>`;
          }
        }

        html += '</div>';
      }
      agendaMonthGrid.innerHTML = html;
    }

    // Carregamentos
    readList('congregacoes', (list) => {
      congregacoesByIdEvents = {};
      (list || []).forEach(c => { congregacoesByIdEvents[c.id] = c; });
      congregacoesReady = true;
      renderCalendar();
    });

    readList('eventos', (list) => {
      eventosCal = normalizeEventosForMonth(list);
      renderCalendar();
    });

    readList('agenda2026', (list) => {
      agenda2026Cal = normalizeAgenda2026ForMonth(list);
      renderCalendar();
    });

    if (agendaTodayBtn) {
      agendaTodayBtn.addEventListener('click', () => {
        viewDate = new Date();
        renderCalendar();
      });
    }

    // primeira render após DOM
    renderCalendar();
  }

// ===================== Dashboard (Página dashboard.html) =====================
  const dashRoot = qs('#dashboard-root');
  if(dashRoot){
    const elYear = qs('#dash-year');
    const elMonth = qs('#dash-month');
    const elCidade = qs('#dash-cidade');
    const elTipo = qs('#dash-tipo');
    const elSetor = qs('#dash-setor');
    const elDate = qs('#dash-date');
    const elUpcomingOnly = qs('#dash-upcoming-only');
    const btnSearch = qs('#dash-search');
    const btnClear = qs('#dash-clear');

    const sumTotal = qs('#dash-sum-total');
    const sumCidades = qs('#dash-sum-cidades');
    const sumTipos = qs('#dash-sum-tipos');
    const countInfo = qs('#dash-count');

    const tableBody = qs('#dash-table-body');

    const btnImprimirDash = qs('#btn-imprimir-dashboard');
    const btnExportarPdfDash = qs('#btn-exportar-pdf-dashboard');
    const btnExportarXlsDash = qs('#btn-exportar-xls-dashboard');

    const cardDaily = qs('#dash-card-daily');
    const ctxMonthly = qs('#dashChartMonthly') ? qs('#dashChartMonthly').getContext('2d') : null;
    const ctxTipos = qs('#dashChartTipos') ? qs('#dashChartTipos').getContext('2d') : null;
    const ctxDaily = qs('#dashChartDaily') ? qs('#dashChartDaily').getContext('2d') : null;
    const ctxCidades = qs('#dashChartCidades') ? qs('#dashChartCidades').getContext('2d') : null;

    const monthNames = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    const fullWeek = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];

    let all = [];
    let filtered = [];
    let chartMonthly = null, chartTipos = null, chartDaily = null, chartCidades = null;

    function slug(str){
      return (String(str||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/(^-|-$)/g,'')) || 'na';
    }
    function norm(str){ return String(str||'').trim(); }

    function getTipoServicoItem(it){
      const t = norm(it && it.tipo);
      const s = norm(it && it.servico);
      if(t) return t;
      if(s && s.includes(' - ')) return s.split(' - ')[0];
      return s || 'Outros';
    }
    function getServicoDescItem(it){
      const d = norm(it && it.descricao);
      const s = norm(it && it.servico);
      if(d) return d;
      if(s && s.includes(' - ')) return s.split(' - ').slice(1).join(' - ');
      return '';
    }

    function parseDateYmdLocal(ymd){
      if(!ymd) return null;
      const m = String(ymd).match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if(!m) return null;
      const y = Number(m[1]), mo = Number(m[2])-1, d = Number(m[3]);
      return new Date(y, mo, d);
    }

    function fillSelect(el, items, firstLabel='Todos'){
      if(!el) return;
      const opts = [`<option value="">${firstLabel}</option>`].concat(items.map(v=>`<option value="${v}">${v}</option>`));
      el.innerHTML = opts.join('');
    }

    function initFilters(list){
      const years = Array.from(new Set(list.map(it => (parseDateYmdLocal(it.data)||{}).getFullYear && parseDateYmdLocal(it.data).getFullYear()).filter(Boolean))).sort((a,b)=>a-b);
      const tipos = Array.from(new Set(list.map(getTipoServicoItem).map(norm).filter(Boolean))).sort((a,b)=>a.localeCompare(b,'pt'));
      const setores = Array.from(new Set(list.map(it => norm(it.setor)).filter(Boolean))).sort((a,b)=>a.localeCompare(b,'pt'));
      const cidades = Array.from(new Set(list.map(it => norm(it.cidade)).filter(Boolean))).sort((a,b)=>a.localeCompare(b,'pt'));

      fillSelect(elYear, years.map(String),'Todos');
      if(years.includes(2026)){ elYear.value = '2026'; }

      fillSelect(elMonth, monthNames.map((n,i)=>String(i+1).padStart(2,'0')+': '+n), 'Todos');
      // Normalizar valores do mês no formato MM na UI, mas guardamos apenas MM
      if(elMonth){
        // converter value exibido "MM: Nome" para apenas MM
        const opts = Array.from(elMonth.options);
        opts.forEach(o=>{ if(o.value){ o.value = o.value.slice(0,2); o.textContent = o.textContent.slice(4); }});
      }

      fillSelect(elTipo, tipos, 'Todos');
      fillSelect(elSetor, setores, 'Todos');
      fillSelect(elCidade, cidades, 'Todas');
    }

    function applyFilters(){
      const y = elYear && elYear.value ? Number(elYear.value) : null;
      const m = elMonth && elMonth.value ? Number(elMonth.value) : null;
      const cid = norm(elCidade && elCidade.value);
      const tp = norm(elTipo && elTipo.value);
      const st = norm(elSetor && elSetor.value);
      const dStr = elDate && elDate.value ? elDate.value : '';
      const upcomingOnly = !!(elUpcomingOnly && elUpcomingOnly.checked);
      const today = new Date(); today.setHours(0,0,0,0);

      const by = all.filter(it => {
        const d = parseDateYmdLocal(it.data);
        if(!d) return false;
        if(upcomingOnly && d.getTime() < today.getTime()) return false; // apenas próximos
        const yy = d.getFullYear();
        const mm = d.getMonth()+1;
        if(y && yy !== y) return false;
        if(m && mm !== m) return false;
        if(dStr){
          const f = parseDateYmdLocal(dStr);
          if(!f) return false;
          if(d.getFullYear()!==f.getFullYear() || d.getMonth()!==f.getMonth() || d.getDate()!==f.getDate()) return false;
        }
        if(cid && norm(it.cidade)!==cid) return false;
        if(tp && norm(getTipoServicoItem(it))!==tp) return false;
        if(st && norm(it.setor)!==st) return false;
        return true;
      });
      return by;
    }

    function updateSummary(list){
      if(sumTotal) sumTotal.textContent = String(list.length);
      if(sumCidades) sumCidades.textContent = String(new Set(list.map(i=>norm(i.cidade)).filter(Boolean)).size);
      if(sumTipos) sumTipos.textContent = String(new Set(list.map(i=>norm(getTipoServicoItem(i))).filter(Boolean)).size);
      if(countInfo){
        const y = elYear && elYear.value ? `Ano ${elYear.value}` : 'Todos os anos';
        const m = elMonth && elMonth.value ? `, ${monthNames[Number(elMonth.value)-1]}` : '';
        countInfo.textContent = `${list.length} registro(s) — ${y}${m}`;
      }
    }

    function pad2(n){ return String(n).padStart(2,'0'); }

    function renderTable(list){
      if(!tableBody) return;
      if(!list.length){ tableBody.innerHTML = '<tr><td colspan="7" class="text-muted">Nenhum dado para exibir</td></tr>'; return; }
      const days = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
      const rows = list.slice().sort((a,b)=>{
        const ad = parseDateYmdLocal(a.data) || new Date(2100,0,1);
        const bd = parseDateYmdLocal(b.data) || new Date(2100,0,1);
        if(ad-bd!==0) return ad-bd;
        const ah = a.hora||''; const bh = b.hora||'';
        return ah.localeCompare(bh);
      }).map((it,idx)=>{
        const d = parseDateYmdLocal(it.data);
        const day = d ? days[d.getDay()] : '';
        const ymd = d ? `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}` : it.data;
        const hora = norm(it.hora);
        const svcTipo = getTipoServicoItem(it);
        const svcDesc = getServicoDescItem(it);
        return `<tr>
          <td>${idx+1}</td>
          <td>${ymd}</td>
          <td>${day}</td>
          <td>${hora?`<span class="time">${hora}</span>`:''}</td>
          <td><span class="pill-servico servico-${slug(svcTipo)}">${svcTipo}</span>${svcDesc?` <span class="servico-desc">— ${svcDesc}</span>`:''}</td>
          <td>${norm(it.cidade)}</td>
          <td>${norm(it.congregacao)}</td>
        </tr>`;
      }).join('');
      tableBody.innerHTML = rows;
    }

    function ensureChart(inst, ctx, type, data, options){
      if(!ctx || typeof Chart==='undefined') return inst;
      if(inst){ inst.data = data; inst.options = options||{}; inst.update(); return inst; }
      return new Chart(ctx, { type, data, options: options||{} });
    }

    function renderCharts(list){
      // Monthly
      const y = elYear && elYear.value ? Number(elYear.value) : null;
      const monthsCount = new Array(12).fill(0);
      list.forEach(it=>{
        const d = parseDateYmdLocal(it.data);
        if(!d) return;
        if(y && d.getFullYear()!==y) return; // se ano filtrado, manter aderência do gráfico mensal
        monthsCount[d.getMonth()]++;
      });
      chartMonthly = ensureChart(chartMonthly, ctxMonthly, 'bar', {
        labels: monthNames,
        datasets: [{ label:'Eventos/mês', data: monthsCount, backgroundColor:'#3b82f6' }]
      }, { plugins:{ legend:{ display:false }}, scales:{ y:{ beginAtZero:true, ticks:{ precision:0 } } } });

      // Tipos
      const tiposMap = {};
      list.forEach(it=>{ const k = norm(getTipoServicoItem(it))||'Outros'; tiposMap[k]=(tiposMap[k]||0)+1; });
      const tiposLabels = Object.keys(tiposMap).sort((a,b)=>tiposMap[b]-tiposMap[a]);
      const tiposData = tiposLabels.map(l=>tiposMap[l]);
      const colors = ['#0ea5e9','#22c55e','#f59e0b','#ef4444','#8b5cf6','#14b8a6','#eab308','#06b6d4','#f97316','#84cc16'];
      chartTipos = ensureChart(chartTipos, ctxTipos, 'doughnut', {
        labels: tiposLabels,
        datasets: [{ data: tiposData, backgroundColor: tiposLabels.map((_,i)=>colors[i%colors.length]) }]
      }, { plugins:{ legend:{ position:'bottom' } } });

      // Cidades (Top 12)
      const cidMap = {};
      list.forEach(it=>{ const k = norm(it.cidade)||'—'; cidMap[k]=(cidMap[k]||0)+1; });
      const cidEntries = Object.entries(cidMap).sort((a,b)=>b[1]-a[1]).slice(0,12);
      chartCidades = ensureChart(chartCidades, ctxCidades, 'bar', {
        labels: cidEntries.map(x=>x[0]),
        datasets: [{ label:'Eventos', data: cidEntries.map(x=>x[1]), backgroundColor:'#10b981' }]
      }, { indexAxis:'y', plugins:{ legend:{ display:false }}, scales:{ x:{ beginAtZero:true, ticks:{ precision:0 } } } });

      // Daily (se um mês estiver selecionado)
      const m = elMonth && elMonth.value ? Number(elMonth.value) : null;
      if(m){
        if(cardDaily) cardDaily.classList.remove('d-none');
        const daily = new Array(31).fill(0);
        const yearRef = y || (list[0] && (parseDateYmdLocal(list[0].data)||{}).getFullYear && parseDateYmdLocal(list[0].data).getFullYear());
        list.forEach(it=>{
          const d = parseDateYmdLocal(it.data); if(!d) return;
          if(d.getMonth()+1!==m) return; if(y && d.getFullYear()!==y) return; if(!y && yearRef && d.getFullYear()!==yearRef) return;
          daily[d.getDate()-1]++;
        });
        const daysLabels = daily.map((_,i)=>String(i+1));
        chartDaily = ensureChart(chartDaily, ctxDaily, 'line', {
          labels: daysLabels,
          datasets: [{ label:'Eventos/dia', data: daily, borderColor:'#6366f1', backgroundColor:'rgba(99,102,241,.25)', tension:.2, fill:true }]
        }, { plugins:{ legend:{ display:false }}, scales:{ y:{ beginAtZero:true, ticks:{ precision:0 } } } });
      } else {
        if(cardDaily) cardDaily.classList.add('d-none');
        if(chartDaily){ chartDaily.destroy(); chartDaily = null; }
      }
    }

    function buildRowsHtml(list){
      const days = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
      function pad2(n){ return String(n).padStart(2,'0'); }
      return list.slice().sort((a,b)=>{
        const ad = parseDateYmdLocal(a.data) || new Date(2100,0,1);
        const bd = parseDateYmdLocal(b.data) || new Date(2100,0,1);
        if(ad-bd!==0) return ad-bd;
        const ah = a.hora||''; const bh = b.hora||'';
        return ah.localeCompare(bh);
      }).map((it,idx)=>{
        const d = parseDateYmdLocal(it.data);
        const day = d ? days[d.getDay()] : '';
        const ymd = d ? `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}` : it.data;
        const hora = norm(it.hora);
         const svcTipo = getTipoServicoItem(it);
         return `<tr><td>${idx+1}</td><td>${ymd}</td><td>${day}</td><td>${hora||''}</td><td>${svcTipo}</td><td>${norm(it.cidade)}</td><td>${norm(it.congregacao)}</td></tr>`;
      }).join('');
    }

    function render(){
      filtered = applyFilters();
      updateSummary(filtered);
      renderTable(filtered);
      renderCharts(filtered);
    }

    // Exportações do Dashboard
    if(btnImprimirDash){
      btnImprimirDash.addEventListener('click', ()=>{
        try{
          const list = filtered || [];
          if(!list.length){ toast('Sem registros para imprimir', 'error'); return; }
          const rows = buildRowsHtml(list);
          const now = new Date();
          const gerado = `${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
          const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"/><title>Dashboard - Eventos</title>
            <style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;margin:24px;color:#000}h1{font-size:18px;margin:0 0 4px}h2{font-size:14px;margin:0 0 12px;color:#333}table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:6px;font-size:12px}th{background:#f2f2f2}@media print{.no-print{display:none}}</style>
          </head><body>
            <h1>CONGREGAÇÃO CRISTÃ NO BRASIL</h1>
            <h2>Eventos (Filtros do Dashboard)</h2>
            <p class="no-print">Gerado em: ${gerado}</p>
            <table><thead><tr><th>Nº</th><th>Data</th><th>Di</th><th>Hor.</th><th>Serviço</th><th>Cidade</th><th>Congregação / Salão</th></tr></thead>
            <tbody>${rows}</tbody></table>
          </body></html>`;
          const w = window.open('', '_blank');
          if(!w){ toast('Popup bloqueado. Permita pop-ups.', 'error'); return; }
          w.document.open(); w.document.write(html); w.document.close();
        }catch(err){ console.error(err); toast('Falha ao gerar impressão', 'error'); }
      });
    }
    if(btnExportarPdfDash){
      btnExportarPdfDash.addEventListener('click', ()=>{
        try{
          const list = filtered || [];
          if(!list.length){ toast('Sem registros para exportar', 'error'); return; }
          const rows = buildRowsHtml(list);
          const html = `<!DOCTYPE html><html lang='pt-BR'><head><meta charset='UTF-8'/><title>Dashboard - PDF</title>
            <style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;margin:24px;color:#000}h1{font-size:18px;margin:0 0 4px}h2{font-size:14px;margin:0 0 12px;color:#333}table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:6px;font-size:12px}th{background:#f2f2f2}.no-print{margin:12px 0}@media print{.no-print{display:none}}</style>
          </head><body>
            <h1>CONGREGAÇÃO CRISTÃ NO BRASIL</h1>
            <h2>Eventos — Dashboard</h2>
            <div class='no-print'>Use Ctrl+P e escolha "Salvar como PDF".</div>
            <table><thead><tr><th>Nº</th><th>Data</th><th>Di</th><th>Hor.</th><th>Serviço</th><th>Cidade</th><th>Congregação / Salão</th></tr></thead>
            <tbody>${rows}</tbody></table>
            <script>window.addEventListener('load', function(){ setTimeout(function(){ window.print(); }, 200); });</script>
          </body></html>`;
          const w = window.open('', '_blank');
          if(!w){ toast('Popup bloqueado. Permita pop-ups.', 'error'); return; }
          w.document.open(); w.document.write(html); w.document.close();
        }catch(err){ console.error(err); toast('Falha ao exportar PDF', 'error'); }
      });
    }
    if(btnExportarXlsDash){
      btnExportarXlsDash.addEventListener('click', ()=>{
        try{
          const list = filtered || [];
          if(!list.length){ toast('Sem registros para exportar', 'error'); return; }
          const rows = buildRowsHtml(list);
          const xls = `<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>
            <table border='1'><thead><tr><th>Nº</th><th>Data</th><th>Di</th><th>Hor.</th><th>Serviço</th><th>Cidade</th><th>Congregação / Salão</th></tr></thead>
            <tbody>${rows}</tbody></table></body></html>`;
          const blob = new Blob([xls], { type: 'application/vnd.ms-excel' });
          const url = URL.createObjectURL(blob);
          const a=document.createElement('a'); a.href=url; a.download='dashboard-eventos.xls'; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(url); a.remove();},0);
          toast('Arquivo XLS gerado');
        }catch(err){ console.error(err); toast('Falha ao exportar XLS', 'error'); }
      });
    }

    [elYear, elMonth, elCidade, elTipo, elSetor, elDate].forEach(el=>{ if(el) el.addEventListener('input', render); });
      if(elUpcomingOnly) elUpcomingOnly.addEventListener('change', render);
    if(btnSearch) btnSearch.addEventListener('click', render);
    if(btnClear){ btnClear.addEventListener('click', ()=>{
      if(elYear) elYear.value='';
      if(elMonth) elMonth.value='';
      if(elCidade) elCidade.value='';
      if(elTipo) elTipo.value='';
      if(elSetor) elSetor.value='';
      if(elDate) elDate.value='';
      render();
    }); }

    let filtersInitialized = false;
    let agendaData = [];
    let eventosData = [];
    function recomputeAll(){
      const merged = ([]).concat(agendaData||[], eventosData||[]);
      all = merged;
      try{ if(!filtersInitialized){ initFilters(all); filtersInitialized = true; } }catch(err){ console.error(err); }
      render();
    }
    function cidadeDoEvento(ev){
      try{
        const c = congregacoesByIdEvents && congregacoesByIdEvents[ev.congregacaoId];
        if(c && c.cidade) return c.cidade;
        const nome = String(ev.congregacaoNome||'');
        if(nome.includes(' - ')) return nome.split(' - ')[0];
      }catch{}
      return '';
    }
    function normalizeEventosForDashboard(list){
      return (list||[]).map(ev=>{
        return {
          data: ev.data,
          hora: horaDoEvento(ev),
          tipo: ev.tipo,
          cidade: cidadeDoEvento(ev),
          congregacao: labelCong(ev)
        };
      });
    }
    // Carregar congregações para enriquecer eventos (cidade/hora)
    readList('congregacoes', list => {
      congregacoesByIdEvents = {};
      (list||[]).forEach(c => { congregacoesByIdEvents[c.id] = c; });
      recomputeAll();
    });
    // Carregar eventos (Atendimentos) e normalizar para o Dashboard
    readList('eventos', list => {
      eventosCache = list || [];
      eventosData = normalizeEventosForDashboard(eventosCache);
      recomputeAll();
    });
    // Carregar Agenda 2026
    readList('agenda2026', list => {
      agendaData = list || [];
      recomputeAll();
    });
  }

}());
  function renderTabelaReforcos(){
    if (typeof window !== 'undefined' && typeof window._renderTabelaReforcosInner === 'function') {
      return window._renderTabelaReforcosInner();
    }
  }

  // Relatórios: Calendário de eventos
  function openEventosCalendar(){
    try{
      const list = getRelEventosFiltrados();
      if(!list || !list.length){
        toast('Nenhum evento encontrado para gerar o calendário');
        return;
      }
      const y = (relYearSel && relYearSel.value) ? parseInt(relYearSel.value,10) : (parseDateYmdLocal(list[0].data)||new Date()).getFullYear();
      const m = (relMonthSel && relMonthSel.value) ? parseInt(relMonthSel.value,10) : ((parseDateYmdLocal(list[0].data)||new Date()).getMonth()+1);
      // filtrar lista para o mês/ano escolhido
      const monthList = list.filter(ev=>{
        const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
        return d && d.getFullYear()===y && (d.getMonth()+1)===m;
      });
      if(!monthList.length){
        toast('Nenhum evento no mês selecionado');
        return;
      }
      // agrupar por dia
      const byDay = {};
      monthList.forEach(ev=>{
        const d = parseDateYmdLocal(ev.data) || new Date(ev.data);
        const day = d.getDate();
        (byDay[day] = byDay[day] || []).push(ev);
      });
      Object.keys(byDay).forEach(k=>{
        byDay[k].sort((a,b)=>{
          const ha = horaDoEvento(a)||''; const hb = horaDoEvento(b)||'';
          return ha.localeCompare(hb);
        });
      });

      // construir calendário mensal
      const first = new Date(y, m-1, 1);
      const title = monthLabel(first);
      const startWeekday = first.getDay();
      const daysInMonth = new Date(y, m, 0).getDate();
      const cells = [];
      for(let i=0;i<startWeekday;i++){ cells.push({ empty:true }); }
      for(let d=1; d<=daysInMonth; d++){
        cells.push({ day:d, events: byDay[d]||[] });
      }
      while(cells.length % 7 !== 0){ cells.push({ empty:true }); }

      const html = `<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Calendário de Eventos - ${title}</title>
  <style>
    body{ font-family: Arial, Helvetica, sans-serif; margin:20px; }
    .topbar{ display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; }
    .title{ font-size:20px; font-weight:bold; }
    .print-btn{ padding:6px 10px; border:1px solid #0d6efd; background:#e7f1ff; color:#0d6efd; border-radius:4px; cursor:pointer; }
    .grid{ display:grid; grid-template-columns: repeat(7, 1fr); gap:8px; }
    .cell{ border:1px solid #ddd; min-height:120px; padding:6px; border-radius:6px; }
    .cell.empty{ background:#fafafa; }
    .day{ font-weight:bold; margin-bottom:4px; color:#333; }
    .event{ margin:4px 0; padding:4px; border-left:3px solid #0d6efd; background:#f5f9ff; border-radius:4px; }
    .event .meta{ font-size:12px; color:#555; }
    .event .title{ font-size:13px; color:#111; }
    .legend{ margin-top:12px; font-size:12px; color:#666; }
    @media print{ .print-btn{ display:none; } .cell{ break-inside: avoid; } }
  </style>
  <script>
    function imprimir(){ window.print(); }
  </script>
</head>
<body>
  <div class="topbar">
    <div class="title">Calendário de Eventos — ${title}</div>
    <button class="print-btn" onclick="imprimir()">Imprimir</button>
  </div>
  <div class="grid">
    ${cells.map(c=>{
      if(c.empty) return '<div class="cell empty"></div>';
      const itens = (c.events||[]).map(ev=>{
        const hora = (ev.hora || horaDoEvento(ev) || '');
        const tt = ev.tipo || '';
        const cong = labelCong(ev) || (ev.congregacaoNome || '');
        return `<div class="event"><div class="meta">${tt}${hora? ' • ' + hora : ''}</div><div class="title">${cong}</div></div>`;
      }).join('');
      return `<div class="cell"><div class="day">${c.day}</div>${itens||''}</div>`;
    }).join('')}
  </div>
  <div class="legend">Eventos mostrados conforme filtros de Relatórios.</div>
</body>
</html>`;

      const w = window.open('', '_blank');
      if(!w){ toast('Bloqueado pelo navegador: permita pop-ups para visualizar o calendário'); return; }
      w.document.open();
      w.document.write(html);
      w.document.close();
    }catch(err){
      console.error(err);
      const msg = (err && (err.message||err.code)) || 'Falha ao gerar calendário';
      toast(msg, 'error');
    }
  }

  // Bind do botão de calendário em Relatórios
  document.addEventListener('DOMContentLoaded', function(){
    try{
      const btnCal = qs('#rel-eventos-calendar');
      if(btnCal) btnCal.addEventListener('click', openEventosCalendar);
    }catch{}
  });