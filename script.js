/* =========================================
    CORE E UTILITÁRIOS
    ========================================= */
const webAppURL = "https://script.google.com/macros/s/AKfycby3AVvBD0F86UMaFNtc5bXe5Kf1tDtuLlem7Z1-HOG8K0XVClInwn2h9DJ1AIa-mDg6/exec";

let currentCloudXMLData = null; 
let currentXmlFile = null; 
let myChart = null; 
let currentModalData = null; 
let chartsInstance = {}; 

/* --- DARK MODE LOGIC --- */
function toggleDarkMode() {
    const html = document.querySelector('html');
    html.classList.toggle('dark');
    const isDark = html.classList.contains('dark');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
    document.getElementById('themeText').innerText = isDark ? "Modo Claro" : "Modo Escuro";
    updateChartTheme();
}
function initTheme() {
    const saved = localStorage.getItem('theme');
    if (saved === 'dark' || (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
        document.querySelector('html').classList.add('dark');
        document.getElementById('themeText').innerText = "Modo Claro";
    }
}

/* --- NAVIGATION --- */
function switchModule(moduleName) {
    document.querySelectorAll('.module-content').forEach(el => el.classList.add('hidden-module'));
    const target = document.getElementById('module-' + moduleName);
    if(target) target.classList.remove('hidden-module');
    
    const buttons = ['xml', 'form', 'view'];
    buttons.forEach(b => {
        const btn = document.getElementById('nav-' + b);
        if(!btn) return;
        
        const isActive = 'nav-' + b === 'nav-' + (moduleName.replace('import', 'xml'));
        if(isActive) {
            btn.className = "w-full text-left py-2.5 px-4 rounded-xl bg-gradient-to-r from-indigo-600 to-violet-600 text-white shadow-lg shadow-indigo-500/30 transition-all duration-200 flex items-center gap-3 text-sm font-bold group border border-black";
            btn.querySelector('i').classList.remove('text-indigo-500', 'dark:text-indigo-400');
        } else {
            btn.className = "w-full text-left py-2.5 px-4 rounded-xl transition-all duration-200 flex items-center gap-3 text-sm font-medium group text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-700 hover:text-indigo-600 dark:hover:text-indigo-400";
        }
    });
    if(moduleName === 'cloud-view') { 
        carregarNFs(); 
        setTimeout(updateCharts, 100); 
    }
}

function mostrarFeedback(txt) {
    const f = document.getElementById('feedback');
    document.getElementById('feedbackText').innerText = txt;
    f.classList.remove('hidden', 'translate-y-[-20px]', 'opacity-0');
    setTimeout(() => f.classList.add('hidden'), 4000);
}

function fmt(val) { return val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }

function downloadTextFile(filename, content) {
    const blob = new Blob([content], { type: 'text/xml' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); window.URL.revokeObjectURL(url);
}

function initMasks() {
    document.querySelectorAll('.numeric-input').forEach(input => {
        new Cleave(input, { blocks: [44], numericOnly: true });
    });
}

/* =========================================
    PARSER XML 
    ========================================= */
function extractSpecialCode(text) {
    if (!text) return null;
    const regex = /\b(EMB|MKT)[-.\w]*\b/gi;
    const found = text.match(regex);
    return found ? found[0].toUpperCase() : null;
}

function parseNFeData(xmlString) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "text/xml");
    const txt = (el, tag) => el?.getElementsByTagName(tag)[0]?.textContent || '';
    const float = (el, tag) => parseFloat(txt(el, tag)) || 0;

    const infNFe = xmlDoc.getElementsByTagName("infNFe")[0];
    if (!infNFe) return null;

    const ide = infNFe.getElementsByTagName("ide")[0];
    const emit = infNFe.getElementsByTagName("emit")[0];
    const dest = infNFe.getElementsByTagName("dest")[0];
    const total = infNFe.getElementsByTagName("ICMSTot")[0]; 
    const infAdic = infNFe.getElementsByTagName("infAdic")[0];
    const protNFe = xmlDoc.getElementsByTagName("protNFe")[0];

    const enderEmit = emit.getElementsByTagName("enderEmit")[0];
    const enderDest = dest.getElementsByTagName("enderDest")[0];

    const getAddr = (ender) => {
        if(!ender) return '';
        return `${txt(ender, "xLgr")}, ${txt(ender, "nro")} - ${txt(ender, "xBairro")} - ${txt(ender, "xMun")}/${txt(ender, "UF")}`;
    };

    const refs = [];
    const nfRefTags = infNFe.getElementsByTagName("NFref");
    for(let r of nfRefTags) { const key = txt(r, "refNFe"); if(key) refs.push(`Ref: ${key}`); }
    const infCpl = txt(infAdic, "infCpl");

    const items = [];
    const detTags = infNFe.getElementsByTagName("det");
    for (let det of detTags) {
        const prod = det.getElementsByTagName("prod")[0];
        const imposto = det.getElementsByTagName("imposto")[0];
        const icms = imposto?.getElementsByTagName("ICMS")[0]; 
        const icmsTag = icms ? icms.children[0] : null;

        const xProd = txt(prod, "xProd");
        const cProd = txt(prod, "cProd");
        const infItem = txt(det, "infAdProd");
        const special = extractSpecialCode(xProd) || extractSpecialCode(infItem);
        
        items.push({
            codigo: special ? special : cProd,
            desc: xProd,
            ncm: txt(prod, "NCM"),
            cst: icmsTag ? (txt(icmsTag, "CST") || txt(icmsTag, "CSOSN")) : '',
            cfop: txt(prod, "CFOP"),
            un: txt(prod, "uCom"),
            qtd: float(prod, "qCom"),
            vUnit: float(prod, "vUnCom"),
            vTotal: float(prod, "vProd"),
            vBc: float(icmsTag, "vBC"),
            vIcms: float(icmsTag, "vICMS"),
            pIcms: float(icmsTag, "pICMS"),
            vIpi: float(imposto?.getElementsByTagName("IPI")[0], "vIPI"),
            pIpi: float(imposto?.getElementsByTagName("IPI")[0], "pIPI")
        });
    }

    let chave = infNFe.getAttribute("Id");
    if(chave && chave.startsWith("NFe")) chave = chave.substring(3);

    return {
        id: chave,
        chaveFormatada: chave ? chave.replace(/(\d{4})/g, '$1 ').trim() : '',
        natOp: txt(ide, "natOp"),
        nNF: txt(ide, "nNF"), 
        serie: txt(ide, "serie"), 
        dhEmi: txt(ide, "dhEmi"),
        dhSaiEnt: txt(ide, "dhSaiEnt"),
        tpNF: txt(ide, "tpNF"), 
        emit: txt(emit, "xNome"), 
        cnpjEmit: txt(emit, "CNPJ"),
        ieEmit: txt(emit, "IE"),
        endEmit: getAddr(enderEmit),
        dest: txt(dest, "xNome"),
        cnpjDest: txt(dest, "CNPJ") || txt(dest, "CPF"),
        ieDest: txt(dest, "IE"),
        endDest: getAddr(enderDest),
        vBC: float(total, "vBC"),
        vICMS: float(total, "vICMS"),
        vBCST: float(total, "vBCST"),
        vST: float(total, "vST"),
        vProd: float(total, "vProd"),
        vFrete: float(total, "vFrete"),
        vSeg: float(total, "vSeg"),
        vDesc: float(total, "vDesc"),
        vIPI: float(total, "vIPI"),
        vPIS: float(total, "vPIS"),
        vCOFINS: float(total, "vCOFINS"),
        vOutro: float(total, "vOutro"),
        vNF: float(total, "vNF"), 
        vTax: float(total, "vICMS") + float(total, "vIPI"),
        nProt: txt(protNFe?.getElementsByTagName("infProt")[0], "nProt"),
        itens: items, 
        referencias: refs,
        infCpl: infCpl,
        rawXml: xmlString 
    };
}

/* =========================================
    GERADOR DE DANFE VISUAL
    ========================================= */
function gerarDanfe(nfData) {
    if(!nfData) return alert("Dados indisponíveis para gerar DANFE.");

    const win = window.open('', '_blank');
    const fmtMoeda = (v) => Number(v).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    const fmtData = (d) => d ? new Date(d).toLocaleDateString('pt-BR') : '';
    const fmtQtd = (v) => Number(v).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 4});

    const css = `
        @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700&display=swap');
        body { margin: 0; padding: 10mm; font-family: 'Roboto Condensed', 'Arial Narrow', sans-serif; font-size: 10px; -webkit-print-color-adjust: exact; }
        .danfe-container { border: 1px solid #000; width: 100%; max-width: 210mm; margin: 0 auto; }
        .row { display: flex; width: 100%; }
        .col { border-right: 1px solid #000; border-bottom: 1px solid #000; padding: 2px; box-sizing: border-box; display: flex; flex-direction: column; overflow: hidden; }
        .col:last-child { border-right: none; }
        .label { font-size: 7px; text-transform: uppercase; color: #444; margin-bottom: 1px; }
        .value { font-size: 10px; font-weight: bold; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .title { font-weight: bold; text-align: center; background: #eee; border-bottom: 1px solid #000; padding: 2px; text-transform: uppercase; font-size: 9px; }
        .center { text-align: center; }
        .right { text-align: right; }
        .bold { font-weight: bold; }
        .wrap { white-space: normal; }
        .canhoto { border-bottom: 1px dashed #000; padding-bottom: 5px; margin-bottom: 5px; }
        .header-main { display: flex; border-bottom: 1px solid #000; }
        .logo-area { width: 100px; display: flex; align-items: center; justify-content: center; border-right: 1px solid #000; border-bottom: 1px solid #000; }
        .danfe-ident { flex: 1; border-right: 1px solid #000; border-bottom: 1px solid #000; padding: 5px; text-align: center; }
        .barcode-area { width: 280px; padding: 5px; border-bottom: 1px solid #000; }
        .big-number { font-size: 14px; }
        .items-table { width: 100%; border-collapse: collapse; font-size: 9px; margin-top: 5px; }
        .items-table th { border: 1px solid #000; background: #ddd; padding: 2px; font-weight: bold; font-size: 7px; }
        .items-table td { border: 1px solid #000; padding: 2px; }
        .no-border-bottom { border-bottom: none !important; }
        .no-border-right { border-right: none !important; }
        .h-25 { height: 25px; }
        @media print { body { padding: 0; margin: 5mm; } .danfe-container { border: 2px solid #000; } button { display: none; } }
    `;

    const itensHtml = nfData.itens.map(i => `
        <tr>
            <td>${i.codigo}</td><td>${i.desc}</td><td class="center">${i.ncm || ''}</td><td class="center">${i.cst || ''}</td>
            <td class="center">${i.cfop || ''}</td><td class="center">${i.un}</td><td class="right">${fmtQtd(i.qtd)}</td>
            <td class="right">${fmtMoeda(i.vUnit)}</td><td class="right bold">${fmtMoeda(i.vTotal)}</td>
            <td class="right">${fmtMoeda(i.vBc)}</td><td class="right">${fmtMoeda(i.vIcms)}</td><td class="right">${fmtMoeda(i.vIpi)}</td>
            <td class="right">${Number(i.pIcms||0).toFixed(0)}%</td><td class="right">${Number(i.pIpi||0).toFixed(0)}%</td>
        </tr>
    `).join('');

    const emptyRows = Math.max(0, 12 - nfData.itens.length);
    const emptyHtml = Array(emptyRows).fill('<tr><td colspan="14" style="border:none; border-left:1px solid #000; border-right:1px solid #000; color:transparent;">.</td></tr>').join('');

    const html = `
    <!DOCTYPE html><html><head><title>DANFE - ${nfData.nNF}</title><style>${css}</style>
    <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.5/dist/JsBarcode.all.min.js"><\/script></head>
    <body onload="initBarcode()"><div style="text-align:center; margin-bottom:10px;"><button onclick="window.print()" style="padding:10px 20px; font-size:16px; cursor:pointer;">IMPRIMIR DANFE</button></div>
        <div class="danfe-container canhoto no-border-bottom">
            <div class="row"><div class="col" style="flex:4"><span class="label">RECEBEMOS DE ${nfData.emit.substring(0,40)}</span><div style="margin-top:10px; font-size:11px; text-align:center;">DATA DE RECEBIMENTO ______/______/______</div></div>
            <div class="col center" style="flex:1"><span class="label">NF-e</span><span class="value big-number">Nº ${nfData.nNF}</span><span class="label">SÉRIE ${nfData.serie}</span></div></div>
        </div>
        <div class="danfe-container">
            <div class="header-main"><div class="logo-area"><span style="font-size:12px; font-weight:bold; color:#ccc;">LOGO</span></div>
                <div class="col" style="flex:3; border-bottom:0; border-right:1px solid #000;"><div style="padding:5px;"><div class="value big-number wrap">${nfData.emit}</div><div class="label wrap" style="margin-top:2px;">${nfData.endEmit}</div><div class="label" style="margin-top:2px;">CNPJ: ${nfData.cnpjEmit} IE: ${nfData.ieEmit}</div></div></div>
                <div class="danfe-ident" style="border-bottom:0;"><div class="big-number bold">DANFE</div><div class="label">DOCUMENTO AUXILIAR DA<br>NOTA FISCAL ELETRÔNICA</div>
                <div class="row" style="margin-top:5px; justify-content:center;"><div style="text-align:left; margin-right:10px;"><span class="value">0 - ENTRADA</span><br><span class="value">1 - SAÍDA</span></div><div style="border:1px solid #000; padding:2px 6px; height:20px; font-size:14px; font-weight:bold;">${nfData.tpNF || '1'}</div></div>
                <div class="value bold" style="margin-top:5px;">Nº ${nfData.nNF}</div><div class="value">SÉRIE ${nfData.serie}</div></div>
                <div class="col" style="width:310px; border-bottom:0; border-right:0;"><div class="barcode-area center" style="border-right:0;"><svg id="barcode"></svg></div>
                <div class="col" style="border-right:0;"><span class="label">CHAVE DE ACESSO</span><span class="value center" style="letter-spacing:1px;">${nfData.chaveFormatada}</span></div>
                <div class="col" style="border-right:0; border-bottom:0;"><span class="label">CONSULTA NO PORTAL NACIONAL DA NF-E</span><span class="value center">www.nfe.fazenda.gov.br/portal</span></div></div>
            </div>
            <div class="row"><div class="col" style="flex:3"><span class="label">NATUREZA DA OPERAÇÃO</span><span class="value">${nfData.natOp || 'Venda de Mercadorias'}</span></div><div class="col" style="flex:2; border-right:0;"><span class="label">PROTOCOLO</span><span class="value">${nfData.nProt || '-'}</span></div></div>
            <div class="title">DESTINATÁRIO / REMETENTE</div>
            <div class="row"><div class="col" style="flex:4"><span class="label">NOME / RAZÃO SOCIAL</span><span class="value">${nfData.dest}</span></div><div class="col" style="flex:2"><span class="label">CNPJ / CPF</span><span class="value">${nfData.cnpjDest || ''}</span></div><div class="col" style="flex:1; border-right:0;"><span class="label">DATA EMISSÃO</span><span class="value center">${fmtData(nfData.dhEmi)}</span></div></div>
            <div class="row"><div class="col" style="flex:3"><span class="label">ENDEREÇO</span><span class="value">${nfData.endDest || ''}</span></div><div class="col" style="flex:2"><span class="label">INSCRIÇÃO ESTADUAL</span><span class="value">${nfData.ieDest || ''}</span></div><div class="col" style="flex:1; border-right:0;"><span class="label">DATA SAÍDA</span><span class="value center">${fmtData(nfData.dhSaiEnt)}</span></div></div>
            <div class="title">CÁLCULO DO IMPOSTO</div>
            <div class="row"><div class="col" style="flex:1"><span class="label">BC ICMS</span><span class="value right">${fmtMoeda(nfData.vBC)}</span></div><div class="col" style="flex:1"><span class="label">VALOR ICMS</span><span class="value right">${fmtMoeda(nfData.vICMS)}</span></div><div class="col" style="flex:1"><span class="label">BC ICMS ST</span><span class="value right">${fmtMoeda(nfData.vBCST)}</span></div><div class="col" style="flex:1"><span class="label">VALOR ICMS ST</span><span class="value right">${fmtMoeda(nfData.vST)}</span></div><div class="col" style="flex:1"><span class="label">V. PRODUTOS</span><span class="value right">${fmtMoeda(nfData.vProd)}</span></div><div class="col" style="flex:1; border-right:0;"><span class="label">V. TOTAL NOTA</span><span class="value right bold">${fmtMoeda(nfData.vNF)}</span></div></div>
            <div class="title" style="margin-top:0;">DADOS DO PRODUTO / SERVIÇO</div>
            <table class="items-table"><thead><tr><th style="width:50px">CÓDIGO</th><th>DESCRIÇÃO</th><th style="width:40px">NCM</th><th style="width:30px">CST</th><th style="width:30px">CFOP</th><th style="width:30px">UN</th><th style="width:40px">QTD</th><th style="width:50px">V.UNIT</th><th style="width:50px">V.TOTAL</th><th style="width:50px">BC ICMS</th><th style="width:40px">V.ICMS</th><th style="width:40px">V.IPI</th><th style="width:30px">%ICMS</th><th style="width:30px">%IPI</th></tr></thead><tbody>${itensHtml}${emptyHtml}</tbody></table>
            <div class="title" style="margin-top:5px;">DADOS ADICIONAIS</div>
            <div class="row" style="height:100px; border-bottom:0;"><div class="col" style="flex:1; border-right:0; border-bottom:0;"><span class="label">INFORMAÇÕES COMPLEMENTARES</span><span class="value wrap" style="font-weight:normal; font-size:9px;">${nfData.infCpl || ''} <br> ${nfData.referencias.join(' - ')}</span></div></div>
        </div>
        <script>function initBarcode() { try { JsBarcode("#barcode", "${nfData.id}", { format: "CODE128", displayValue: false, height: 35, width: 1, margin: 0 }); } catch(e) { console.error("Erro barcode", e); } }<\/script>
    </body></html>`;
    win.document.write(html); win.document.close();
}

/* =========================================
    MÓDULO 1: IMPORTADOR XML LOCAL
    ========================================= */
let localInvoices = [];

function initXMLModule() {
    loadFromStorage();
    const dz = document.getElementById("dropZone");
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('active'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('active'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('active'); handleFileUpload(e.dataTransfer.files); });
}

function handleFileUpload(files) {
    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = parseNFeData(e.target.result);
            if(data) {
                localInvoices.push(data);
                saveToStorage();
                renderLocalTable();
            } else { alert('XML Inválido: ' + file.name); }
        };
        reader.readAsText(file);
    });
}

function renderLocalTable() {
    const tbody = document.getElementById("invoiceTableBody");
    tbody.innerHTML = "";
    let totalVal = 0, totalTax = 0;
    localInvoices.sort((a,b) => new Date(b.dhEmi) - new Date(a.dhEmi));

    localInvoices.forEach((nf, index) => {
        totalVal += nf.vNF; totalTax += nf.vTax;
        tbody.innerHTML += `
            <tr>
                <td class="px-6 py-4 text-slate-500 dark:text-slate-400 text-sm border-l-4 border-l-transparent hover:border-l-indigo-500 transition-all">${new Date(nf.dhEmi).toLocaleDateString('pt-BR')}</td>
                <td class="px-6 py-4 font-bold text-slate-700 dark:text-slate-300 text-sm">${nf.emit.substring(0,25)}...</td>
                <td class="px-6 py-4 font-mono text-xs text-slate-500 dark:text-slate-400 bg-slate-50 dark:bg-slate-700 rounded w-fit">#${nf.nNF}</td>
                <td class="px-6 py-4 text-center"><span class="bg-indigo-50 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-300 px-2 py-1 rounded-md text-xs font-bold">${nf.itens.length}</span></td>
                <td class="px-6 py-4 text-right font-bold text-slate-700 dark:text-slate-300">${fmt(nf.vNF)}</td>
                <td class="px-6 py-4 text-center">
                    <div class="flex justify-center gap-2">
                        <button title="Ver Detalhes" onclick="openModalLocal(${index})" class="w-8 h-8 rounded-full bg-slate-100 dark:bg-slate-700 hover:bg-blue-100 dark:hover:bg-blue-900/50 text-slate-400 hover:text-blue-600 transition flex items-center justify-center border border-slate-200/50 dark:border-slate-600"><i class="fas fa-eye"></i></button>
                        <button title="Baixar XML" onclick="downloadLocalXML(${index})" class="w-8 h-8 rounded-full bg-slate-100 dark:bg-slate-700 hover:bg-emerald-100 dark:hover:bg-emerald-900/50 text-slate-400 hover:text-emerald-600 transition flex items-center justify-center border border-slate-200/50 dark:border-slate-600"><i class="fas fa-download"></i></button>
                        <button title="Excluir" onclick="removeInvoiceLocal(${index})" class="w-8 h-8 rounded-full bg-slate-100 dark:bg-slate-700 hover:bg-red-100 dark:hover:bg-red-900/50 text-slate-400 hover:text-red-600 transition flex items-center justify-center border border-slate-200/50 dark:border-slate-600"><i class="fas fa-trash"></i></button>
                    </div>
                </td>
            </tr>`;
    });

    document.getElementById("kpiTotal").innerText = fmt(totalVal);
    document.getElementById("kpiCount").innerText = localInvoices.length;
    document.getElementById("kpiTax").innerText = fmt(totalTax);
    if(localInvoices.length === 0) tbody.innerHTML = `<tr><td colspan="6" class="text-center py-10 text-slate-400 dark:text-slate-500 font-medium">Nenhum XML importado ainda.</td></tr>`;
}

function openModalLocal(index) {
    const nf = localInvoices[index];
    if(!nf) return;
    populateAndShowModal(nf);
    
    document.getElementById('pdfContainer').classList.add('hidden');
    document.getElementById('pdfPlaceholder').innerHTML = '<i class="fas fa-file-invoice text-6xl mb-4"></i><p>XML Local (Sem PDF)</p>';
    
    const footer = document.getElementById('modalFooterActions');
    footer.innerHTML = `
        <button onclick="downloadLocalXML(${index})" class="text-emerald-600 hover:text-emerald-800 text-sm font-bold flex items-center bg-emerald-50 px-3 py-2 rounded-lg transition"><i class="fas fa-download mr-2"></i> Baixar XML</button>
        <button onclick="gerarDanfe(currentModalData)" class="text-slate-600 hover:text-slate-800 text-sm font-bold flex items-center bg-slate-100 px-3 py-2 rounded-lg transition"><i class="fas fa-print mr-2"></i> Visualizar DANFE</button>
    `;
}

function downloadLocalXML(index) {
    const nf = localInvoices[index];
    if(nf && nf.rawXml) downloadTextFile(`NFe-${nf.nNF}.xml`, nf.rawXml);
    else alert("Conteúdo XML não disponível.");
}

function removeInvoiceLocal(i) { if(confirm('Remover?')) { localInvoices.splice(i, 1); saveToStorage(); renderLocalTable(); } }
function clearData() { if(confirm("Limpar tudo?")) { localStorage.removeItem('nfe_v4_data'); localInvoices = []; renderLocalTable(); } }
function saveToStorage() { localStorage.setItem('nfe_v4_data', JSON.stringify(localInvoices)); }
function loadFromStorage() { const d = localStorage.getItem('nfe_v4_data'); if(d) { localInvoices = JSON.parse(d); renderLocalTable(); } }


/* =========================================
    MÓDULO 2 & 3: CLOUD LOGIC
    ========================================= */
let allData = [], workingData = [];
let currentPage = 1, rowsPerPage = 10;
let deleteIdCloud = null;

function autofillFromXML(files) {
    if(files.length === 0) return;
    
    currentXmlFile = files[0]; 

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = parseNFeData(e.target.result);
        if(data) {
            document.getElementById('razao').value = data.emit.substring(0, 50); 
            document.getElementById('nfVenda').value = data.nNF;
            document.getElementById('nfRemessa').value = ""; 
            
            const emitUpper = data.emit.toUpperCase();
            const sel = document.getElementById('fabrica');
            for(let i=0; i<sel.options.length; i++){
                if(emitUpper.includes(sel.options[i].value) && sel.options[i].value !== ""){ sel.selectedIndex = i; break; }
            }

            currentCloudXMLData = data;
            document.getElementById('xmlLinkedBanner').classList.remove('hidden');
            document.getElementById('xmlItemCount').innerText = data.itens.length;
            atualizarPreview();
        } else { alert("XML Inválido"); }
    };
    reader.readAsText(files[0]);
}

function showLoadingBtn() { 
    const btn = document.getElementById('btnEnviarNF');
    btn.disabled = true; btn.querySelector('.dots').classList.add('active'); 
    btn.querySelector('.btn-text').innerText = "Enviando...";
}
function hideLoadingBtn() { 
    const btn = document.getElementById('btnEnviarNF');
    btn.disabled = false; btn.querySelector('.dots').classList.remove('active'); 
    btn.querySelector('.btn-text').innerText = "Enviar Nota Fiscal";
}

function atualizarPreview() {
    const r = document.getElementById('razao').value;
    const v = document.getElementById('nfVenda').value;
    const rem = document.getElementById('nfRemessa').value;
    document.getElementById('previewTextoCobertura').textContent = 
        `segue NF de armazenagem referente a NF (Nº NF: ${v || '...'}) - (Razão Social: ${r || '...'}) NF de remessa (Nº NF Remessa (Fornecedor): ${rem || '...'})`;
}
['razao', 'nfVenda', 'nfRemessa'].forEach(id => {
    document.getElementById(id)?.addEventListener('input', atualizarPreview);
});

async function enviarNF(){
    const razao = document.getElementById('razao').value.trim();
    const nfVenda = document.getElementById('nfVenda').value.trim();
    const nfRemessa = document.getElementById('nfRemessa').value.trim();
    const nfCobertura = document.getElementById('nfCobertura').value.trim();
    const fabrica = document.getElementById('fabrica').value;
    let observacao = document.getElementById('observacao').value.trim();
    const arquivoPdf = document.getElementById('arquivo').files[0];

    if(!fabrica || !arquivoPdf){ alert('Preencha Fábrica e Arquivo PDF Principal.'); return; }

    if(currentCloudXMLData) {
        const xmlSummary = JSON.stringify({
            emit: currentCloudXMLData.emit,
            cnpj: currentCloudXMLData.cnpjEmit,
            vNF: currentCloudXMLData.vNF,
            itens: currentCloudXMLData.itens.map(i => ({
                c: i.codigo, d: i.desc, q: i.qtd, u: i.un,
                v: i.vTotal, vu: i.vUnit 
            }))
        });
        observacao += `|||XMLDATA:${xmlSummary}`;
    }

    showLoadingBtn();

    const readFileAsBase64 = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = error => reject(error);
            reader.readAsDataURL(file);
        });
    };

    try {
        const ret1 = document.getElementById('retorno1') ? document.getElementById('retorno1').files[0] : null;
        const ret2 = document.getElementById('retorno2') ? document.getElementById('retorno2').files[0] : null;
        const ret3 = document.getElementById('retorno3') ? document.getElementById('retorno3').files[0] : null;

        const leituras = [readFileAsBase64(arquivoPdf)];
        if(currentXmlFile) leituras.push(readFileAsBase64(currentXmlFile));
        if(ret1) leituras.push(readFileAsBase64(ret1));
        if(ret2) leituras.push(readFileAsBase64(ret2));
        if(ret3) leituras.push(readFileAsBase64(ret3));
        
        const resultados = await Promise.all(leituras);
        
        let resultIndex = 0;
        const pdfBase64 = resultados[resultIndex++];
        const xmlBase64 = currentXmlFile ? resultados[resultIndex++] : null;
        const ret1Base64 = ret1 ? resultados[resultIndex++] : null;
        const ret2Base64 = ret2 ? resultados[resultIndex++] : null;
        const ret3Base64 = ret3 ? resultados[resultIndex++] : null;

        const payload = { 
            razao, nfVenda, nfRemessa, nfCobertura, 
            fabrica, observacao, 
            pdfBase64: pdfBase64, 
            pdfName: arquivoPdf.name,
            xmlBase64: xmlBase64,
            xmlName: currentXmlFile ? currentXmlFile.name : null,
            retorno1Base64: ret1Base64, retorno1Name: ret1 ? ret1.name : null,
            retorno2Base64: ret2Base64, retorno2Name: ret2 ? ret2.name : null,
            retorno3Base64: ret3Base64, retorno3Name: ret3 ? ret3.name : null
        };

        const res = await fetch(webAppURL, { method:'POST', body: JSON.stringify(payload)});
        const result = await res.json();
        
        if(result.status === 'success'){
            mostrarFeedback('NF Enviada com sucesso!');
            document.getElementById('nfForm').reset();
            document.getElementById('xmlLinkedBanner').classList.add('hidden');
            currentCloudXMLData = null; 
            currentXmlFile = null; 
            atualizarPreview();
        } else { 
            alert('Erro no script Google: ' + (result.message || 'Desconhecido')); 
        }

    } catch(e) { 
        console.error(e);
        alert('Erro ao processar arquivos ou conexão.'); 
    } finally { 
        hideLoadingBtn(); 
    }
}

async function carregarNFs(){
    const tbody = document.querySelector('#nfTable tbody');
    if(!tbody) return;
    
    tbody.innerHTML = '<tr><td colspan="10" class="text-center py-8"><i class="fas fa-spinner fa-spin text-indigo-500 text-2xl"></i><br><span class="text-xs text-slate-400 mt-2 block">Sincronizando dados...</span></td></tr>';
    
    try {
        const res = await fetch(webAppURL + '?action=get');
        allData = await res.json();
        if(!Array.isArray(allData)) allData = [];
        allData.sort((a,b)=> (new Date(b.dataHora||0) - new Date(a.dataHora||0)));
        aplicarFiltroPesquisa();
    } catch(e) { 
        console.error(e); 
        // Adicionei o motivo do erro na tela para te ajudar caso continue!
        tbody.innerHTML = `<tr><td colspan="10" class="text-center text-red-500 py-4 font-bold"><i class="fas fa-exclamation-triangle"></i> Falha: ${e.message}</td></tr>`;
    }
}

function extractValueFromObs(obs) {
    if (!obs || !obs.includes('|||XMLDATA:')) return 0;
    try {
        const parts = obs.split('|||XMLDATA:');
        const jsonStr = parts[1];
        const data = JSON.parse(jsonStr);
        return Number(data.vNF) || 0;
    } catch (e) {
        return 0;
    }
}

function aplicarFiltroPesquisa(){
    const filtro = document.getElementById('filtroFabrica')?.value || '';
    const termo = (document.getElementById('searchInput')?.value || '').toLowerCase().trim();
    
    const dtStartVal = document.getElementById('dateStart')?.value || '';
    const dtEndVal = document.getElementById('dateEnd')?.value || '';
    const dtStart = dtStartVal ? new Date(dtStartVal) : null;
    const dtEnd = dtEndVal ? new Date(dtEndVal) : null;
    
    if(dtEnd) dtEnd.setHours(23,59,59,999);

    workingData = allData.filter(r => {
        if(filtro && String(r.fabrica) !== filtro) return false;
        
        if (r.dataHora) {
            const rDate = new Date(r.dataHora);
            if (dtStart && rDate < dtStart) return false;
            if (dtEnd && rDate > dtEnd) return false;
        }

        if(!termo) return true;
        const combined = (String(r.razao)+String(r.nfVenda)+String(r.nfCobertura)+String(r.observacao)).toLowerCase();
        return combined.includes(termo);
    });

    let totalVal = 0;
    let totalQtd = workingData.length;
    
    workingData.forEach(r => {
        totalVal += extractValueFromObs(r.observacao);
    });

    const elDashVal = document.getElementById('dashTotalValor');
    const elDashQtd = document.getElementById('dashTotalQtd');
    const elDashTkt = document.getElementById('dashTicketMedio');
    
    if(elDashVal) elDashVal.innerText = fmt(totalVal);
    if(elDashQtd) elDashQtd.innerText = totalQtd;
    if(elDashTkt) elDashTkt.innerText = totalQtd > 0 ? fmt(totalVal / totalQtd) : 'R$ 0,00';

    currentPage = 1; 
    renderCloudTable(); 
    updateCharts();
}

function updateCharts() {
    const htmlEl = document.querySelector('html');
    const isDark = htmlEl ? htmlEl.classList.contains('dark') : false;
    const textColor = isDark ? '#cbd5e1' : '#64748b';
    const gridColor = isDark ? '#334155' : '#f1f5f9';

    const statsQtd = {};
    const statsVal = {};

    workingData.forEach(d => {
        const fab = d.fabrica || 'Outros';
        const val = extractValueFromObs(d.observacao);
        
        statsQtd[fab] = (statsQtd[fab] || 0) + 1;
        statsVal[fab] = (statsVal[fab] || 0) + val;
    });

    createOrUpdateChart('chartValorFabrica', 'bar', {
        labels: Object.keys(statsVal),
        datasets: [{
            label: 'Valor Total (R$)',
            data: Object.values(statsVal),
            backgroundColor: ['#10b981cc', '#3b82f6cc', '#6366f1cc', '#8b5cf6cc', '#ec4899cc'],
            borderRadius: 4,
            indexAxis: 'y'
        }]
    }, {
        indexAxis: 'y',
        scales: { 
            x: { ticks: { color: textColor, callback: (v) => 'R$ ' + v/1000 + 'k' }, grid: { color: gridColor } }, 
            y: { ticks: { color: textColor }, grid: { display:false } } 
        }
    });

    createOrUpdateChart('chartQtdFabrica', 'doughnut', {
        labels: Object.keys(statsQtd),
        datasets: [{
            label: 'Qtd Notas',
            data: Object.values(statsQtd),
            backgroundColor: ['#3b82f6cc', '#10b981cc', '#f59e0bcc', '#ef4444cc', '#6366f1cc'],
            borderColor: isDark ? '#1e293b' : '#ffffff',
            borderWidth: 2
        }]
    }, {
        cutout: '60%',
        plugins: { legend: { position: 'right', labels: { color: textColor, boxWidth: 12 } } },
        scales: { x: { display: false }, y: { display: false } }
    });
}

function createOrUpdateChart(canvasId, type, data, extraOptions = {}) {
    const canvas = document.getElementById(canvasId);
    if(!canvas) return; // Trava de segurança para impedir o erro!

    const ctx = canvas.getContext('2d');
    
    if (chartsInstance[canvasId]) {
        chartsInstance[canvasId].destroy();
    }

    const defaultOptions = {
        responsive: true, 
        maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false } }
    };

    const options = { ...defaultOptions, ...extraOptions };
    if(extraOptions.plugins) options.plugins = { ...defaultOptions.plugins, ...extraOptions.plugins };
    if(extraOptions.scales) options.scales = extraOptions.scales;

    chartsInstance[canvasId] = new Chart(ctx, { type, data, options });
}

function updateChartTheme() { 
    if(Object.keys(chartsInstance).length > 0) updateCharts(); 
}

function renderCloudTable(){
    const tbody = document.querySelector('#nfTable tbody');
    if(!tbody) return;
    
    const totalItems = workingData.length;
    const totalPages = Math.ceil(totalItems/rowsPerPage) || 1;
    if(currentPage > totalPages) currentPage = totalPages;
    const start = (currentPage-1)*rowsPerPage;
    const pageData = workingData.slice(start, start+rowsPerPage);

    tbody.innerHTML = '';
    if(pageData.length === 0) tbody.innerHTML = '<tr><td colspan="10" class="text-center py-8 text-slate-400 font-medium">Nenhum registro encontrado.</td></tr>';
    else pageData.forEach(r => {
        const rawObs = r.observacao || '';
        const cleanObs = rawObs.split('|||XMLDATA:')[0];
        const hasXmlData = rawObs.includes('|||XMLDATA:');
        
        let htmlRetorno = '<div class="flex gap-1 justify-center">';
        
        for(let i=1; i<=3; i++) {
            const prop = i === 1 ? 'pdfRetorno' : `pdfRetorno${i}`;
            const link = r[prop];

            if(link) {
                 htmlRetorno += `<a href="${link}" target="_blank" class="w-6 h-6 rounded bg-emerald-50 dark:bg-emerald-900/20 text-emerald-500 hover:bg-emerald-100 dark:hover:bg-emerald-900/40 hover:text-emerald-700 inline-flex items-center justify-center transition border border-emerald-200/50 dark:border-emerald-900/50 text-[10px]" title="Ver Retorno ${i}"><i class="fas fa-file-pdf"></i><span class="ml-0.5 font-bold">${i}</span></a>`;
            } else {
                 htmlRetorno += `
                <label class="cursor-pointer w-6 h-6 rounded bg-slate-100 dark:bg-slate-700 hover:bg-indigo-100 dark:hover:bg-indigo-900/50 text-slate-400 hover:text-indigo-600 transition flex items-center justify-center border border-slate-200/50 dark:border-slate-600 text-[10px]" title="Upload Retorno ${i}">
                    <i class="fas fa-cloud-upload-alt"></i>
                    <input type="file" class="hidden" accept="application/pdf" onchange="uploadRetorno('${r.id}', this, ${i})">
                </label>`;
            }
        }
        htmlRetorno += '</div>';

        let htmlXmlLink = '';
        if(r.xmlLink) {
           htmlXmlLink = `<a href="${r.xmlLink}" target="_blank" class="w-7 h-7 rounded bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 hover:bg-emerald-100 dark:hover:bg-emerald-900/50 transition flex items-center justify-center" title="Baixar XML Salvo"><i class="fas fa-code text-xs"></i></a>`;
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="px-4 py-4 text-slate-500 dark:text-slate-400 font-mono text-xs border-l-4 border-l-transparent hover:border-l-indigo-500 transition-all">${r.dataHora ? new Date(r.dataHora).toLocaleDateString() : '-'}</td>
            <td contenteditable="true" onblur="updateCloud('${r.id}','razao',this.innerText)" class="px-4 py-4 font-bold text-slate-700 dark:text-slate-300 text-sm focus:bg-indigo-50 dark:focus:bg-indigo-900 focus:outline-none rounded transition">${r.razao||''}</td>
            <td contenteditable="true" onblur="updateCloud('${r.id}','nfVenda',this.innerText)" class="px-4 py-4 font-mono text-xs text-slate-600 dark:text-slate-400 focus:bg-indigo-50 dark:focus:bg-indigo-900 focus:outline-none rounded transition">${r.nfVenda||''}</td>
            <td contenteditable="true" onblur="updateCloud('${r.id}','nfRemessa',this.innerText)" class="px-4 py-4 font-mono text-xs text-slate-600 dark:text-slate-400 focus:bg-indigo-50 dark:focus:bg-indigo-900 focus:outline-none rounded transition">${r.nfRemessa||''}</td>
            <td contenteditable="true" onblur="updateCloud('${r.id}','nfCobertura',this.innerText)" class="px-4 py-4 font-mono text-xs text-slate-600 dark:text-slate-400 focus:bg-indigo-50 dark:focus:bg-indigo-900 focus:outline-none rounded transition">${r.nfCobertura||''}</td>
            <td contenteditable="true" onblur="updateCloud('${r.id}','fabrica',this.innerText)" class="px-4 py-4 text-xs font-bold text-slate-500 dark:text-slate-400 focus:bg-indigo-50 dark:focus:bg-indigo-900 focus:outline-none rounded transition">${r.fabrica||''}</td>
            <td class="px-4 py-4 text-center">${r.pdfLink ? `<a href="${r.pdfLink}" target="_blank" class="w-8 h-8 rounded-full bg-red-50 dark:bg-red-900/20 text-red-500 hover:bg-red-100 dark:hover:bg-red-900/40 hover:text-red-700 inline-flex items-center justify-center transition border border-red-200/50 dark:border-red-900/50"><i class="fas fa-file-pdf"></i></a>` : '-'}</td>
            <td class="px-4 py-4 text-center">${htmlRetorno}</td>
            <td title="${cleanObs}" class="px-4 py-4 max-w-[100px] truncate cursor-help text-xs text-slate-400 text-center" onclick="alert('${cleanObs}')">${cleanObs ? '<i class="fas fa-comment-alt"></i>' : '-'}</td>
            <td class="px-4 py-4">
                <div class="flex gap-1 justify-center">
                    <button title="Ver Itens e PDF" onclick="openModalCloud('${r.id}')" class="w-7 h-7 rounded bg-indigo-50 dark:bg-indigo-900/30 ${hasXmlData ? 'text-indigo-600 dark:text-indigo-400' : 'text-slate-300'} hover:bg-indigo-100 dark:hover:bg-indigo-900/50 transition"><i class="fas fa-list text-xs"></i></button>
                    ${htmlXmlLink}
                    <button title="Excluir" onclick="askDeleteCloud('${r.id}')" class="w-7 h-7 rounded bg-red-50 dark:bg-red-900/30 text-red-400 hover:bg-red-100 dark:hover:bg-red-900/50 hover:text-red-600 transition"><i class="fas fa-times text-xs"></i></button>
                </div>
            </td>`;
        tbody.appendChild(tr);
    });
    
    const elPrev = document.getElementById('prevBtn');
    const elNext = document.getElementById('nextBtn');
    const elInfo = document.getElementById('showingInfo');
    
    if(elPrev) elPrev.disabled = currentPage <= 1;
    if(elNext) elNext.disabled = currentPage >= totalPages;
    if(elInfo) elInfo.innerText = `${start+1}-${Math.min(start+rowsPerPage, totalItems)} de ${totalItems}`;
}

function uploadRetorno(id, input, slot) {
    const file = input.files[0];
    if(!file) return;

    if (file.size > 5 * 1024 * 1024) { 
        alert('O arquivo é muito grande (Máx 5MB).');
        return;
    }

    mostrarFeedback(`Enviando Retorno (Slot ${slot})...`);
    
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = async function(e) {
        const base64Data = e.target.result.split(',')[1]; 

        const payload = {
            action: 'uploadRetorno',
            id: id,
            slot: slot,
            fileName: `RETORNO_${slot}_` + file.name,
            mimeType: file.type,
            fileData: base64Data
        };

        try {
            const res = await fetch(webAppURL, { method: 'POST', body: JSON.stringify(payload) });
            const data = await res.json();
            
            if(data.status === 'success'){
                mostrarFeedback(`Retorno ${slot} Salvo!`);
                carregarNFs();
            } else {
                alert('Erro ao salvar: ' + data.message);
            }
        } catch(error) {
            console.error(error);
            alert('Erro na conexão.');
        }
    };
}

function openModalCloud(id) {
    const r = allData.find(x => String(x.id) === String(id));
    if(!r) return alert("Erro: Registro não encontrado.");
    
    const pdfContainer = document.getElementById('pdfContainer');
    const pdfFrame = document.getElementById('pdfFrame');
    const pdfPlaceholder = document.getElementById('pdfPlaceholder');
    const pdfLinkBtn = document.getElementById('pdfExternalLink');

    if(r.pdfLink) {
        pdfContainer.classList.remove('hidden'); pdfFrame.classList.remove('hidden'); pdfPlaceholder.classList.add('hidden');
        pdfFrame.src = r.pdfLink.replace('/view', '/preview'); pdfLinkBtn.href = r.pdfLink;
    } else {
        pdfFrame.classList.add('hidden'); pdfPlaceholder.classList.remove('hidden'); pdfLinkBtn.href = "#";
    }

    const rawObs = r.observacao || '';
    const footer = document.getElementById('modalFooterActions');
    footer.innerHTML = ``;

    if(!rawObs.includes('|||XMLDATA:')) {
        const simpleData = { nNF: r.nfVenda, emit: r.razao, cnpjEmit: '-', dest: 'NFS3', vNF: 0, referencias: [], itens: [] };
        populateAndShowModal(simpleData);
        return;
    }

    try {
        const splitData = rawObs.split('|||XMLDATA:');
        const jsonStr = splitData[1];
        const xmlData = JSON.parse(jsonStr);
        
        const modalData = {
            nNF: r.nfVenda, emit: xmlData.emit || "Desconhecido", cnpjEmit: xmlData.cnpj || "-", dest: "Armazenado na Base",
            vNF: Number(xmlData.vNF) || 0,
            referencias: ["Dados recuperados da Base"],
            itens: Array.isArray(xmlData.itens) ? xmlData.itens.map(i => ({
                codigo: i.c, desc: i.d, qtd: i.q, un: i.u,
                vTotal: i.v, 
                vUnit: i.vu || 0 
            })) : []
        };

        populateAndShowModal(modalData);
        footer.innerHTML = `<button onclick="gerarDanfe(currentModalData)" class="text-slate-600 hover:text-slate-800 text-sm font-bold flex items-center bg-slate-100 px-3 py-2 rounded-lg transition border border-slate-300"><i class="fas fa-print mr-2"></i> Visualizar DANFE</button>`;

    } catch(e) { alert("Erro ao ler os dados ocultos do XML."); }
}

function populateAndShowModal(nf) {
    currentModalData = nf; 

    document.getElementById("modalTitle").innerText = `Nota Fiscal ${nf.nNF}`;
    document.getElementById("modalEmitente").innerText = nf.emit;
    document.getElementById("modalCnpjEmit").innerText = nf.cnpjEmit || '-';
    document.getElementById("modalDestinatario").innerText = nf.dest;
    
    const refSec = document.getElementById("refSection");
    const refList = document.getElementById("refList");
    if (nf.referencias && nf.referencias.length > 0) {
        refSec.classList.remove("hidden-force");
        refList.innerHTML = nf.referencias.map(r => `<li>${r}</li>`).join('');
    } else { refSec.classList.add("hidden-force"); }

    const tbody = document.getElementById("modalItemsTable");
    tbody.innerHTML = "";
    nf.itens.forEach(item => {
        tbody.innerHTML += `
            <tr>
                <td class="px-4 py-2 font-mono text-xs text-slate-500 dark:text-slate-400">${item.codigo}</td>
                <td class="px-4 py-2 font-medium text-slate-700 dark:text-slate-300">${item.desc}</td>
                <td class="px-4 py-2 text-right text-slate-600 dark:text-slate-400">
                    ${item.qtd} <span class="text-[10px] font-bold text-slate-400 bg-slate-100 dark:bg-slate-700 px-1.5 py-0.5 rounded ml-1">${item.un || ''}</span>
                </td>
                <td class="px-4 py-2 text-right font-medium text-slate-600 dark:text-slate-400">${fmt(item.vUnit || 0)}</td>
                <td class="px-4 py-2 text-right font-bold text-slate-800 dark:text-slate-200">${fmt(item.vTotal)}</td>
            </tr>`;
    });
    document.getElementById("modalTotalNF").innerText = fmt(nf.vNF);
    document.getElementById("detailsModal").classList.remove("hidden");
}

function closeModal() { 
    document.getElementById("detailsModal").classList.add("hidden"); 
    document.getElementById("pdfFrame").src = "";
}
function changePage(dir) { currentPage += dir; renderCloudTable(); }
function changeRowsPerPage() { rowsPerPage = parseInt(document.getElementById('rowsPerPage').value); currentPage=1; renderCloudTable(); }

async function updateCloud(id, field, val) { try { await fetch(`${webAppURL}?action=update&id=${id}&${field}=${encodeURIComponent(val)}`); } catch(e){} }
function askDeleteCloud(id) { deleteIdCloud = id; document.getElementById('deleteModalCloud').classList.remove('hidden'); }
document.getElementById('confirmDeleteCloud').onclick = async () => {
    if(!deleteIdCloud) return;
    document.getElementById('deleteModalCloud').classList.add('hidden');
    await fetch(`${webAppURL}?action=delete&id=${deleteIdCloud}`);
    mostrarFeedback('Registro excluído'); carregarNFs();
};

function exportCSV() {
    if (!workingData || workingData.length === 0) {
        alert("Sem dados para exportar.");
        return;
    }

    const headers = ["Data", "Razão Social", "NF Venda", "NF Remessa", "NF Cobertura", "Fábrica", "Observação", "Link PDF", "Link Retorno 1", "Link Retorno 2", "Link Retorno 3"];
    const csvRows = [];
    
    csvRows.push('\uFEFF' + headers.join(";"));

    for (const row of workingData) {
        const values = [
            row.dataHora ? new Date(row.dataHora).toLocaleDateString() : "",
            `"${(row.razao || "").replace(/"/g, '""')}"`,
            row.nfVenda || "",
            row.nfRemessa || "",
            row.nfCobertura || "",
            row.fabrica || "",
            `"${(row.observacao || "").replace(/\n/g, " ").replace(/"/g, '""')}"`,
            row.pdfLink || "",
            row.pdfRetorno || "",
            row.pdfRetorno2 || "",
            row.pdfRetorno3 || ""
        ];
        csvRows.push(values.join(";"));
    }

    const csvString = csvRows.join("\n");
    const blob = new Blob([csvString], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `exportacao_nf_integra_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

document.addEventListener('DOMContentLoaded', () => {
    initTheme(); initMasks(); initXMLModule(); switchModule('xml-import');
});