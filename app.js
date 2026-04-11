`// =====================================================================
// DANTHI DASHBOARD (v5.0 - Absolute Zero)
// A modular, strict-mode architecture for stability.
// =====================================================================

const app = {
  state: {
    dark: true,
    allOS: [],
    rotaSelection: new Set(),
    rotaSequence: [],
    draggedIdx: null,
    agendaSelection: null,
    agendaTemplate: 'vistoria',
    contatoCache: {} // local cache for Drive lookups: { itemID: { nome, tel } }
  },

  init() {
    this.ui.init();
    this.gapiService.loadScripts();
  },

  switchTab(tabId) {
    document.querySelectorAll('.tab-pane').forEach(el => el.classList.remove('active'));
    document.getElementById(\`tab-\${tabId}\`).classList.add('active');
    
    document.querySelectorAll('.nav-btn').forEach(btn => {
      btn.classList.remove('active');
      if(btn.getAttribute('onclick').includes(tabId)) {
        btn.classList.add('active');
      }
    });
    
    if (tabId === 'rota') this.features.rota.renderSidebar();
    if (tabId === 'agenda') this.features.agenda.renderSidebar();
  },

  toggleTheme() {
    this.state.dark = !this.state.dark;
    document.body.classList.toggle('dark-mode', this.state.dark);
    document.body.classList.toggle('light-mode', !this.state.dark);
  },

  showToast(msg, type='success') {
    const t = document.createElement('div');
    t.className = \`toast \${type}\`;
    t.innerText = msg;
    document.getElementById('toast-container').appendChild(t);
    setTimeout(() => t.classList.add('show'), 10);
    setTimeout(() => { t.classList.remove('show'); setTimeout(()=>t.remove(),300); }, 3000);
  },

  // Helpers
  isNaoVistoriada(r) {
    return ['OS em Andamento sem vistoria', 'OS não Vistoriada com PEPT', 'OS em Andamento'].includes(r.st) && !r.vis;
  },
  
  parseDriveFolderDate(mesStr) {
    const p = mesStr.split('-');
    if(p.length !== 2) return '';
    const nameMap = { 'jan':'01-JANEIRO','fev':'02-FEVEREIRO','mar':'03-MARÇO','abr':'04-ABRIL','mai':'05-MAIO','jun':'06-JUNHO','jul':'07-JULHO','ago':'08-AGOSTO','set':'09-SETEMBRO','out':'10-OUTUBRO','nov':'11-NOVEMBRO','dez':'12-DEZEMBRO' };
    return nameMap[p[0].toLowerCase()] || '';
  },

  folderName(item, atv, osStr) {
    return \`\${item} - \${atv} - \${osStr}\`;
  },

  extractContactFromTxt(text, r) {
    const m = {nome: '', tel: ''};
    const rxName = /Nome.*?:\\s*(.+)|Contato.*?:\\s*(.+)/gi;
    const rxPhone = /Telefone.*?:\\s*([\\d\\s\\-\\(\\)]+)|Celular.*?:\\s*([\\d\\s\\-\\(\\)]+)/gi;
    let n1 = rxName.exec(text); if(n1) m.nome = (n1[1]||n1[2]||'').trim();
    let t1 = rxPhone.exec(text); if(t1) m.tel = (t1[1]||t1[2]||'').replace(/\\D/g,'');
    
    // Hard fallback if empty but exists deep down:
    if(!m.nome) { const arr = text.split('\\n'); if(arr.length > 5) m.nome = arr[5].substring(0,30); }
    return m;
  }
};

// =====================================================================
// UI MANAGER
// =====================================================================
app.ui = {
  init() {
    const savedDark = localStorage.getItem('danthi-dark');
    if (savedDark === 'light') app.toggleTheme();
  },

  renderOS() {
    this.filterOS();
    document.getElementById('kpi-total').textContent = app.state.allOS.length;
    document.getElementById('kpi-vist').textContent = app.state.allOS.filter(r => r.vis).length;
    document.getElementById('kpi-pend').textContent = app.state.allOS.filter(r => app.isNaoVistoriada(r)).length;
    
    app.features.rota.populateMuns();
  },

  filterOS() {
    const q = document.getElementById('os-search').value.toLowerCase();
    const st = document.getElementById('os-status').value;
    const grid = document.getElementById('os-grid');
    grid.innerHTML = '';
    
    const filtered = app.state.allOS.filter(r => {
      let match = true;
      if (q && ![r.end, r.mun, r.item].join(' ').toLowerCase().includes(q)) match = false;
      if (st && r.st !== st) match = false;
      return match;
    });
    
    document.getElementById('os-count').textContent = \`\${filtered.length} OS\`;
    filtered.forEach(r => {
      const c = STCFG[r.st] || {bg: '#475569', lbl: r.st || '—'};
      const card = document.createElement('div');
      card.className = 'os-card';
      card.innerHTML = \`
        <div class="os-card-header">
          <div class="os-num">#\${r.item}</div>
          <div class="badge" style="background: \${c.bg}">\${c.lbl}</div>
        </div>
        <div class="os-end">\${r.end}</div>
        <div class="os-meta">📍 \${r.mun}</div>
      \`;
      card.onclick = () => this.openDrawer(r);
      grid.appendChild(card);
    });
  },

  filterRotaSidebar() {
    app.features.rota.renderSidebar();
  },

  openDrawer(r) {
    const c = STCFG[r.st] || {bg: '#475569', lbl: r.st || '—'};
    document.getElementById('d-badge').textContent = c.lbl;
    document.getElementById('d-badge').style.background = c.bg;
    document.getElementById('d-title').textContent = \`#\${r.item}\`;
    document.getElementById('d-end').textContent = r.end;
    document.getElementById('d-bairro').textContent = r.bairro;
    document.getElementById('d-mun').textContent = r.mun;
    
    document.getElementById('d-nome').innerHTML = \`<span style="color:var(--accent);">Localizando...</span>\`;
    document.getElementById('d-tel').innerHTML = \`<span style="color:var(--accent);">Localizando...</span>\`;
    
    document.getElementById('drawer-overlay').classList.add('open');
    document.getElementById('drawer').classList.add('open');

    // Trigger Drive Lookup if authenticated
    app.gapiService.findContactData(r).then(ctt => {
      if(ctt.nome || ctt.tel) {
        document.getElementById('d-nome').textContent = ctt.nome || 'N/D';
        document.getElementById('d-tel').textContent = ctt.tel || 'N/D';
      } else {
        document.getElementById('d-nome').textContent = '— Não encontrado —';
        document.getElementById('d-tel').textContent = '— Não encontrado —';
      }
    });
  },

  closeDrawer() {
    document.getElementById('drawer-overlay').classList.remove('open');
    document.getElementById('drawer').classList.remove('open');
  }
};

// =====================================================================
// FEATURES (Rotas & Agenda)
// =====================================================================
app.features = {
  rota: {
    populateMuns() {
      const sel = document.getElementById('rota-mun');
      sel.innerHTML = '<option value="">Filtrar município...</option>';
      const muns = [...new Set(app.state.allOS.filter(r=>app.isNaoVistoriada(r)).map(r=>r.mun))].sort();
      muns.forEach(m => {
        if(m) { const o = document.createElement('option'); o.value = m; o.textContent = m; sel.appendChild(o); }
      });
    },

    renderSidebar() {
      const list = document.getElementById('rota-sidebar-list');
      const munFilt = document.getElementById('rota-mun').value;
      list.innerHTML = '';
      
      const toShow = app.state.allOS.filter(r => app.isNaoVistoriada(r) && r.end && (!munFilt || r.mun === munFilt));
      toShow.forEach(r => {
        const item = document.createElement('div');
        item.style.cssText = 'padding: 8px; border-bottom: 1px solid var(--border-color); display:flex; gap:8px; align-items:start;';
        const isChecked = app.state.rotaSelection.has(r.item) ? 'checked' : '';
        item.innerHTML = \`<input type="checkbox" \${isChecked} style="margin-top:4px; cursor:pointer;">
                          <div style="flex:1">
                            <div style="font-weight:700; font-size:12px;">#\${r.item}</div>
                            <div style="font-size:11px; color:var(--text-muted); line-height:1.3;">\${r.end} <br>\${r.mun}</div>
                          </div>\`;
        item.querySelector('input').onchange = (e) => {
          if (e.target.checked) app.state.rotaSelection.add(r);
          else {
            for (let obj of app.state.rotaSelection) {
              if (obj.item === r.item) app.state.rotaSelection.delete(obj);
            }
          }
        };
        list.appendChild(item);
      });
    },

    generate() {
      if(app.state.rotaSelection.size === 0) return app.showToast('Selecione pelo menos 1 OS', 'error');
      
      // Attempt Optimization
      app.state.rotaSequence = Array.from(app.state.rotaSelection);
      if(typeof google === 'undefined' || !google.maps || !google.maps.DirectionsService) {
        app.showToast('G-Maps fall-back. Ordem manual.', 'warning');
        this.renderSequence();
        return;
      }

      app.showToast('Otimizando rota com IA...');
      const ds = new google.maps.DirectionsService();
      const origin = document.getElementById('rota-partida-val') ? document.getElementById('rota-partida-val').value : ORIGEM;
      const waypoints = app.state.rotaSequence.map(r => ({ location: \`\${r.end}, \${r.mun}, RJ\`, stopover: true }));
      
      ds.route({ origin, destination: origin, waypoints, optimizeWaypoints: true, travelMode: 'DRIVING' }, (res, status) => {
        if(status === 'OK') {
          const order = res.routes[0].waypoint_order;
          app.state.rotaSequence = order.map(i => Array.from(app.state.rotaSelection)[i]);
          this.renderSequence();
        } else {
          app.showToast('Falha Maps API: '+status+'. Ordem manual ativa.');
          this.renderSequence();
        }
      });
    },

    renderSequence() {
      const container = document.getElementById('rota-result');
      container.innerHTML = '';
      const origin = ORIGEM; // Simplification, could be dynamic

      app.state.rotaSequence.forEach((r, i) => {
        const el = document.createElement('div');
        el.className = 'r-stop';
        el.setAttribute('draggable', 'true');
        el.dataset.idx = i;
        el.innerHTML = \`
          <div class="r-handle" title="Arraste para reordenar">⠿</div>
          <div class="r-num">\${i + 1}</div>
          <div class="r-info">
            <div class="r-info-title">#\${r.item}</div>
            <div class="r-info-sub">\${r.end}</div>
          </div>
        \`;

        // Physical pure HTML5 Drag and drop
        el.addEventListener('dragstart', (e) => {
          app.state.draggedIdx = i;
          e.dataTransfer.effectAllowed = 'move';
          el.style.opacity = '0.5';
        });
        el.addEventListener('dragend', () => {
          el.style.opacity = '1';
          document.querySelectorAll('.r-stop').forEach(x => x.classList.remove('drag-over'));
        });
        el.addEventListener('dragenter', e => { e.preventDefault(); el.classList.add('drag-over'); });
        el.addEventListener('dragover', e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; });
        el.addEventListener('dragleave', () => el.classList.remove('drag-over'));
        el.addEventListener('drop', e => {
          e.preventDefault();
          e.stopPropagation();
          el.classList.remove('drag-over');
          
          const from = app.state.draggedIdx;
          const to = parseInt(el.dataset.idx);
          if (from !== null && from !== to && !isNaN(from) && !isNaN(to)) {
            // Arrays magic
            const moved = app.state.rotaSequence.splice(from, 1)[0];
            app.state.rotaSequence.splice(to, 0, moved);
            app.features.rota.renderSequence(); // Recursively re-render correctly
          }
        });
        container.appendChild(el);
      });

      this.updateIFrame(origin);
    },

    reverse() {
      if(!app.state.rotaSequence.length) return;
      app.state.rotaSequence.reverse();
      this.renderSequence();
    },

    updateIFrame(origin) {
      if(!app.state.rotaSequence.length) return;
      const originUrl = encodeURIComponent(origin);
      const wayptsUrl = app.state.rotaSequence.map(r => encodeURIComponent(\`\${r.end}, \${r.mun}, RJ\`)).join('|');
      const url = \`https://www.google.com/maps/embed/v1/directions?key=\${CFG.KEY}&origin=\${originUrl}&destination=\${originUrl}&waypoints=\${wayptsUrl}&mode=driving&language=pt-BR\`;
      document.getElementById('rota-map-container').innerHTML = \`<iframe src="\${url}" style="width:100%;height:100%;border:none" allowfullscreen loading="lazy"></iframe>\`;
    }
  },

  agenda: {
    renderSidebar() {
      const list = document.getElementById('agenda-sidebar-list');
      list.innerHTML = '';
      app.state.allOS.filter(r => app.isNaoVistoriada(r)).forEach(r => {
         const item = document.createElement('div');
         item.style.cssText = 'padding: 8px; border-bottom: 1px solid var(--border-color); cursor:pointer; display:flex; flex-direction:column; gap:4px;';
         item.innerHTML = \`<div style="font-weight:700; font-size:13px; color:\${app.state.agendaSelection?.item === r.item ? 'var(--accent)' : 'var(--text-primary)'}">#\${r.item}</div>
                           <div style="font-size:11px; color:var(--text-muted);">\${r.end}</div>\`;
         item.onclick = () => this.selectForAgenda(r);
         list.appendChild(item);
      });
    },

    setTemplate(type) {
      app.state.agendaTemplate = type;
      document.getElementById('btn-tmpl-v').classList.toggle('btn-primary', type === 'vistoria');
      document.getElementById('btn-tmpl-m').classList.toggle('btn-primary', type === 'medicao');
      this.updatePreview();
    },

    async selectForAgenda(r) {
      app.state.agendaSelection = r;
      this.renderSidebar(); // refresh exact highlight UI
      
      const spn = document.getElementById('agenda-spinner');
      const inN = document.getElementById('agenda-nome');
      const inT = document.getElementById('agenda-tel');
      
      inN.value = ''; inT.value = '';
      spn.style.display = 'inline';

      // Use strictly resilient Drive resolver
      const ctt = await app.gapiService.findContactData(r);
      spn.style.display = 'none';

      if (ctt.nome || ctt.tel) {
        inN.value = ctt.nome || 'Não detectado';
        inT.value = ctt.tel || '';
      } else {
        inN.placeholder = 'Digite manualmente (Não achou no Drive)';
      }
      this.updatePreview();
    },

    updatePreview() {
      const r = app.state.agendaSelection;
      const el = document.getElementById('agenda-preview');
      if (!r) { el.textContent = 'Selecione uma OS na lateral para gerar a mensagem...'; return; }

      const dtInp = document.getElementById('agenda-date').value;
      let dStr = '[DATA]';
      if (dtInp) {
        const d = new Date(dtInp + 'T00:00:00');
        dStr = d.toLocaleDateString('pt-BR');
      }

      const hI = document.getElementById('agenda-h-ini').value || '10';
      const hF = document.getElementById('agenda-h-fim').value || '12';
      const nome = document.getElementById('agenda-nome').value || 'Cliente';
      
      let msg = '';
      if(app.state.agendaTemplate === 'vistoria') {
        msg = \`Olá \${nome},\\n\\nAqui é da Danthi Engenharia, prestando serviço para a Caixa Econômica.\\nEntramos em contato para verificar a viabilidade de agendamento de uma *Vistoria Técnica* (Avaliação do Imóvel).\\n\\n📌 Nosso engenheiro estará na região no dia *\${dStr}* (entre \${hI}h e \${hF}h).\\n\\nPodemos confirmar essa data?\`;
      } else {
         msg = \`Olá \${nome},\\n\\nAqui é da Danthi Engenharia.\\nAgendamento de medição de obra CAIXA CEF para o dia *\${dStr}* (entre \${hI}h e \${hF}h).\\n\\nPode nos receber na obra?\`;
      }

      el.textContent = msg;
    },

    sendWhatsApp() {
      if(!app.state.agendaSelection) return app.showToast('Nenhuma OS selecionada', 'error');
      let tel = document.getElementById('agenda-tel').value.replace(/\\D/g, '');
      if (tel.length < 10) return app.showToast('Número inválido ou muito curto', 'error');
      
      const msg = document.getElementById('agenda-preview').textContent;
      window.open(\`https://wa.me/55\${tel}?text=\${encodeURIComponent(msg)}\`, '_blank');
    }
  }
};

// =====================================================================
// GAPI SERVICES (Auth, Sheets, Drive)
// =====================================================================
app.gapiService = {
  loadScripts() {
    const s1 = document.createElement('script');
    s1.src = 'https://apis.google.com/js/api.js';
    s1.onload = () => gapi.load('client', () => this.initGapiClient());
    document.head.appendChild(s1);

    const s2 = document.createElement('script');
    s2.src = 'https://accounts.google.com/gsi/client';
    s2.onload = () => this.initGoogleIdentity();
    document.head.appendChild(s2);
    
    // Inject Maps automatically using API config
    const s3 = document.createElement('script');
    s3.src = \`https://maps.googleapis.com/maps/api/js?key=\${CFG.KEY}\`;
    document.head.appendChild(s3);
  },

  async initGapiClient() {
    await gapi.client.init({
      apiKey: CFG.KEY,
      discoveryDocs: [
        'https://sheets.googleapis.com/$discovery/rest?version=v4',
        'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
      ]
    });
    this.fetchDataGrid(); // Automatically public fetch if sheet is open to readers
  },

  initGoogleIdentity() {
    this.tc = google.accounts.oauth2.initTokenClient({
      client_id: CFG.CID,
      scope: CFG.SC,
      callback: (resp) => {
        if(resp.access_token) {
          app.gapiService.token = resp.access_token;
          app.gapiService.auth = true;
          gapi.client.setToken({access_token: resp.access_token});
          document.getElementById('auth-status').textContent = 'Conectado (Drive ATIVO)';
          document.getElementById('auth-status').classList.add('online');
          document.getElementById('btn-auth').style.display = 'none';
        }
      }
    });
    this.tc.requestAccessToken({prompt:''}); // Request on Boot silently
  },

  login() {
    this.tc.requestAccessToken({prompt:''});
  },

  // DATA LOAD AND PARSE
  async fetchDataGrid() {
    try {
      const res = await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId: CFG.SID, range: \`\${CFG.SHT}!A2:AG\` });
      const rows = res.result.values || [];
      app.state.allOS = rows.map((r, i) => {
        // Raw parsing based strictly on column positions assigned in previous implementation
        const p = v => (v === null || v === undefined) ? '' : String(v).trim();
        const b = v => { const s = p(v).toUpperCase(); return s === 'TRUE' || s === 'VERDADEIRO'; };
        return {
          ri: i + 2, item: p(r[0]), os: p(r[1]), osp: p(r[2]), atv: p(r[4]), 
          end: p(r[6]), bairro: p(r[7]), mun: p(r[8]), cep: p(r[10]), tip: p(r[11]), 
          mes: p(r[13]).replace('.',''), st: p(r[23]), vis: b(r[17])
        };
      });
      console.log('Dados importados:', app.state.allOS.length);
      app.ui.renderOS();
    } catch(e) {
      console.error(e);
      app.showToast('Erro ao carregar Sheets. Verifique console.', 'error');
    }
  },

  // DRIVE RESOLVER (The absolute zero robust tracker)
  async findContactData(objOS) {
    if(!this.auth || !this.token) return {nome: null, tel: null}; // Cannot bypass OAuth
    if(app.state.contatoCache[objOS.item]) return app.state.contatoCache[objOS.item];

    try {
      // 1. Find Mes folder inside master root CFG.DID
      const folderMesName = app.parseDriveFolderDate(objOS.mes); 
      const res1 = await gapi.client.drive.files.list({q: \`name='\${folderMesName}' and '\${CFG.DID}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'\`, fields: 'files(id)', pageSize: 1});
      if(res1.result.files.length === 0) throw new Error('Mes não localizado');
      const mesId = res1.result.files[0].id;

      // 2. Find Item Folder
      const osFolderName = app.folderName(objOS.item, objOS.atv, objOS.osp);
      const res2 = await gapi.client.drive.files.list({q: \`name='\${osFolderName}' and '\${mesId}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'\`, fields: 'files(id)', pageSize: 1});
      if(res2.result.files.length === 0) throw new Error('OS não localizada no Drive');
      const osId = res2.result.files[0].id;

      // 3. Find 01 Convocacao
      const res3 = await gapi.client.drive.files.list({q: \`name='01 - Convocação' and '\${osId}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'\`, fields: 'files(id)', pageSize: 1});
      if(res3.result.files.length === 0) throw new Error('Pasta 01 não existe');
      const subId = res3.result.files[0].id;

      // 4. Extract TXT File Media content
      const resF = await gapi.client.drive.files.list({q: \`'\${subId}' in parents and trashed=false and mimeType='text/plain'\`, fields: 'files(id)', pageSize: 1});
      if(resF.result.files.length === 0) throw new Error('Sem TXT de convocação');
      
      const fileId = resF.result.files[0].id;
      const respTxt = await fetch(\`https://www.googleapis.com/drive/v3/files/\${fileId}?alt=media\`, {headers: {Authorization: \`Bearer \${this.token}\`}});
      
      if(respTxt.ok) {
        const text = await respTxt.text();
        const extracted = app.extractContactFromTxt(text, objOS);
        app.state.contatoCache[objOS.item] = extracted; // Memoize
        return extracted;
      }
    } catch(e) {
      console.warn("Drive Tree Traversal Broken Early:", e.message);
    }
    return {nome: null, tel: null};
  }
};

window.onload = () => app.init();
`
