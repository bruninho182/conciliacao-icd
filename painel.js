(() => {
    if (document.getElementById('icd-conciliacao')) return;

    const host = document.createElement('div');
    host.id = 'icd-conciliacao';
    Object.assign(host.style, {
        position: 'fixed', right: '24px', bottom: '24px', width: 'min(360px, 90vw)', zIndex: '2147483647'
    });
    const ui = host.attachShadow({mode: 'closed'});
    ui.innerHTML = `
        <style>
            section { background: #fff; color: #23183d; border: 1px solid #d8cdeb;
                border-radius: 12px; padding: 16px; font: 14px system-ui; box-shadow: 0 4px 24px #0003; }
            h2 { margin: 0 0 12px; font-size: 18px; }
            button, input { font: inherit; margin: 3px 2px; }
            button { border: 1px solid #6e4da4; border-radius: 6px; padding: 5px 9px;
                background: #f4effa; color: #301a50; cursor: pointer; }
            button:disabled { opacity: .5; cursor: default; }
            #close { float: right; }
            p { margin: 8px 0; }
        </style>
        <section aria-label="Conciliação ICD">
            <button id="close" aria-label="Fechar painel">✕</button>
            <h2>Conciliação ICD</h2>
            <label>Planilha Excel <input id="file" type="file" accept=".xlsx"></label>
            <p id="status" role="status">Selecione a planilha e deixe o campo Pesquisar visível em Reservas.</p>
            <button id="test" disabled>Testar 5 consultas</button>
            <button id="run" disabled>Consultar restantes</button>
            <button id="stop" disabled>Pausar</button>
            <button id="save" disabled>Baixar cópia</button>
            <button id="report" disabled>Baixar pendências</button>
        </section>`;
    document.documentElement.append(host);

    const get = (id) => ui.getElementById(id);
    const report = new Map();
    const matched = new Map();
    const uuid = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
    let book, fileName, running, stopped;
    let filled = 0;

    function refresh() {
        get('file').disabled = Boolean(running || book);
        get('test').disabled = !book || running;
        get('run').disabled = !book || running;
        get('stop').disabled = !running;
        get('save').disabled = !filled;
        get('report').disabled = !report.size;
        get('close').disabled = Boolean(running || filled);
    }

    function visible(element) {
        return Boolean(element?.getClientRects().length && getComputedStyle(element).visibility === 'visible');
    }

    function searchField() {
        const inputs = [...document.querySelectorAll('input:not([type=hidden])')].filter(visible);
        const fields = inputs.filter(input => {
            const label = [
                input.placeholder,
                input.getAttribute('aria-label'),
                ...[...input.labels ?? []].map(node => node.textContent),
                input.closest('.MuiFormControl-root')?.querySelector('label')?.textContent
            ];
            return label.some(text => /^pesquisar$/i.test(text?.trim()));
        });
        if (fields.length !== 1) throw new Error('Abra os filtros de Reservas e deixe somente o campo Pesquisar correspondente visível.');
        return fields[0];
    }

    function ensureAccess() {
        if (location.origin !== 'https://app.icdgrupo.com.br' ||
            location.pathname.replace(/\/$/, '') !== '/apps/sales/bookings' ||
            document.querySelector('iframe[src*="challenges.cloudflare.com"], #challenge-running, #challenge-form') ||
            /just a moment|attention required|access denied|forbidden|too many requests/i.test(document.title) ||
            /falha na verificação|verifique se você é humano|acesso negado|muitas solicitações|verify (?:that )?you are human|checking your browser|too many requests|access denied/i.test(document.body?.innerText ?? '')) {
            throw new Error('A página mudou ou exibiu uma verificação ou restrição. O lote foi pausado.');
        }
    }

    const pause = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    function readCode(cell) {
        const isCode = (value) => /^[A-Z0-9]{8}$/i.test(value);
        const exact = [...cell.querySelectorAll('strong')].filter(node =>
            node.parentElement.querySelector('button[aria-label="Copiar"], button[title="Copiar"]'));
        const candidates = exact.map(node => node.textContent.trim()).filter(isCode);
        if (exact.length) return candidates.length === 1 ? candidates[0] : null;
        const label = cell.cloneNode(true);
        label.querySelectorAll('button, [role="button"], svg, [aria-hidden="true"]')
            .forEach(node => node.remove());
        const codes = [...new Set(label.textContent.match(/\b[A-Z0-9]{8}\b/gi) ?? [])];
        return codes.length === 1 ? codes[0] : null;
    }

    function readResult(match) {
        const tables = [...document.querySelectorAll('table, [role="table"], [role="grid"]')].filter(visible);
        const single = /(?:^|\D)1\s*[-\u2010-\u2015\u2212]\s*1\s*(?:of|de|\/)\s*1(?![\d.,])/i;
        const pagination = [...document.querySelectorAll(
            '.MuiTablePagination-displayedRows, .MuiDataGrid-footerContainer, [role="navigation"]'
        )].filter(visible);
        const counts = pagination.map(node => node.innerText || node.textContent);
        if (!counts.some(text => single.test(text)) && !single.test(document.body.innerText)) {
            return {reason: 'A paginação ainda não confirmou um único resultado.'};
        }
        for (const table of tables) {
            const headers = [...table.querySelectorAll('th, [role="columnheader"]')]
                .map(node => (node.innerText || node.textContent).trim().toLowerCase());
            const codeColumn = headers.indexOf('código');
            const supplierColumn = headers.indexOf('fornecedor');
            if (codeColumn < 0 || supplierColumn < 0) continue;
            const rows = [...table.querySelectorAll('tr, [role="row"]')].filter(row =>
                visible(row) && row.querySelector('td, [role="cell"], [role="gridcell"]'));
            if (rows.length !== 1) continue;
            const cells = [...rows[0].querySelectorAll('td, [role="cell"], [role="gridcell"]')];
            const supplier = cells[supplierColumn];
            if (!supplier || !match.test(supplier.innerText || supplier.textContent)) {
                return {reason: 'O resultado exibido ainda não corresponde ao localizador completo.'};
            }
            const cell = cells[codeColumn];
            if (!cell) continue;
            const code = readCode(cell);
            if (code) return {code};
            return {reason: 'O código da reserva não pôde ser lido na célula Código.'};
        }
        return {reason: 'A tabela de resultados não foi identificada.'};
    }

    async function lookup(locator) {
        ensureAccess();
        const field = searchField();
        const suffix = locator.slice(-12);
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(field, suffix);
        field.dispatchEvent(new Event('input', {bubbles: true}));
        field.dispatchEvent(new Event('change', {bubbles: true}));
        field.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', code: 'Enter', bubbles: true}));
        const match = new RegExp('(?<![0-9a-f])' + locator + '(?![0-9a-f])', 'i');
        const deadline = Date.now() + 8000;
        let result;
        await pause(900);
        while (Date.now() < deadline && !stopped) {
            ensureAccess();
            if (!field.isConnected || field.value !== suffix) throw new Error('O campo Pesquisar mudou durante a consulta.');
            result = readResult(match);
            if (result.code) return result;
            await pause(250);
        }
        return result ?? {reason: 'Consulta interrompida.'};
    }

    async function run(limit) {
        if (!book || running) return;
        running = true;
        stopped = false;
        refresh();
        let attempts = 0;
        let misses = 0;
        try {
            for (const row of book.rows) {
                if (stopped) break;
                if (row.done) continue;
                const key = row.locator.toLowerCase();
                if (!uuid.test(key)) {
                    report.set(row.number, 'Localizador fora do formato UUID');
                    row.done = true;
                    continue;
                }
                if (!matched.has(key)) {
                    if (attempts >= limit) break;
                    attempts++;
                    get('status').textContent = `Consultando linha ${row.number} · ${filled} preenchidas.`;
                    const result = await lookup(row.locator);
                    if (stopped) break;
                    matched.set(key, result);
                    misses = result.code ? 0 : misses + 1;
                }
                const {code, reason} = matched.get(key);
                if (code) {
                    book.setVoucher(row, code);
                    filled++;
                } else {
                    report.set(row.number, reason);
                }
                row.done = true;
                refresh();
                if (misses >= 3) throw new Error('Três consultas seguidas sem confirmação. ' + reason);
            }
            const remaining = book.rows.filter(row => !row.done).length;
            get('status').textContent = `${filled} preenchidas · ${report.size} pendências · ${remaining} a consultar.`;
        } catch (error) {
            get('status').textContent = `${error.message} ${filled} preenchidas até aqui.`;
        } finally {
            running = false;
            refresh();
        }
    }

    function download(blob, name) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = name;
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    get('file').addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        get('file').disabled = true;
        try {
            book = await globalThis.icdXlsx.open(file);
            fileName = file.name.replace(/\.xlsx$/i, '');
            get('status').textContent = `${book.rows.length} linhas elegíveis. Confira os filtros do ICD e inicie pelo teste.`;
        } catch (error) {
            get('status').textContent = error.message;
            get('file').value = '';
        }
        refresh();
    });
    get('test').onclick = () => run(5);
    get('run').onclick = () => run(Infinity);
    get('stop').onclick = () => { stopped = true; };
    get('save').onclick = () => download(book.exportFile(), fileName + ' - ICD.xlsx');
    get('report').onclick = () => {
        const csv = [...report].map(([line, reason]) => `${line};${reason}`).join('\r\n');
        download(new Blob(['\ufeffLinha;Motivo\r\n' + csv], {type: 'text/csv;charset=utf-8'}),
            'Pendências ICD.csv');
    };
    get('close').onclick = () => host.remove();
})();
