globalThis.icdXlsx = (() => {
    const decode = new TextDecoder();
    const encode = new TextEncoder();
    const xml = (bytes) => {
        const doc = new DOMParser().parseFromString(decode.decode(bytes), 'application/xml');
        if (doc.querySelector('parsererror')) throw new Error('XML inválido na planilha.');
        return doc;
    };
    const tags = (node, name) => [...node.getElementsByTagNameNS('*', name)];
    const crcTable = Uint32Array.from({length: 256}, (_, i) => {
        for (let bit = 0; bit < 8; bit++) i = (i >>> 1) ^ (i & 1 ? 0xedb88320 : 0);
        return i;
    });
    function crc32(bytes) {
        let crc = 0xffffffff;
        for (const byte of bytes) crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 255];
        return (crc ^ 0xffffffff) >>> 0;
    }
    function readZip(buffer) {
        const view = new DataView(buffer);
        const bytes = new Uint8Array(buffer);
        const u16 = (offset) => view.getUint16(offset, true);
        const u32 = (offset) => view.getUint32(offset, true);
        let end = bytes.length - 22;
        while (end >= Math.max(0, bytes.length - 65557)) {
            if (u32(end) === 0x06054b50 && end + 22 + u16(end + 20) === bytes.length) break;
            end--;
        }
        if (end < Math.max(0, bytes.length - 65557) || u16(end + 4) || u16(end + 6) ||
            u16(end + 10) === 65535) throw new Error('Arquivo XLSX inválido ou não compatível.');
        const files = new Map();
        let offset = u32(end + 16);
        for (let i = 0; i < u16(end + 10); i++) {
            if (u32(offset) !== 0x02014b50 || (u16(offset + 8) & 1)) {
                throw new Error('Arquivo XLSX inválido ou protegido por senha.');
            }
            const length = u16(offset + 28);
            const name = bytes.slice(offset + 46, offset + 46 + length);
            const local = u32(offset + 42);
            const start = local + 30 + u16(local + 26) + u16(local + 28);
            const size = u32(offset + 24);
            const data = bytes.subarray(start, start + u32(offset + 20));
            files.set(decode.decode(name), {
                name, size, data, method: u16(offset + 10), crc: u32(offset + 16),
                time: u16(offset + 12), date: u16(offset + 14)
            });
            offset += 46 + length + u16(offset + 30) + u16(offset + 32);
        }
        return files;
    }
    function writeZip(files) {
        const localParts = [];
        const centralParts = [];
        let offset = 0;
        let centralSize = 0;
        for (const file of files.values()) {
            const local = new DataView(new ArrayBuffer(30));
            const central = new DataView(new ArrayBuffer(46));
            local.setUint32(0, 0x04034b50, true);
            central.setUint32(0, 0x02014b50, true);
            central.setUint16(4, 20, true);
            central.setUint32(42, offset, true);
            for (const [header, shift] of [[local, 0], [central, 2]]) {
                header.setUint16(4 + shift, 20, true);
                header.setUint16(6 + shift, 0x800, true);
                header.setUint16(8 + shift, file.method, true);
                header.setUint16(10 + shift, file.time, true);
                header.setUint16(12 + shift, file.date, true);
                header.setUint32(14 + shift, file.crc, true);
                header.setUint32(18 + shift, file.data.length, true);
                header.setUint32(22 + shift, file.size, true);
                header.setUint16(26 + shift, file.name.length, true);
            }
            localParts.push(local.buffer, file.name, file.data);
            centralParts.push(central.buffer, file.name);
            offset += 30 + file.name.length + file.data.length;
            centralSize += 46 + file.name.length;
        }
        const end = new DataView(new ArrayBuffer(22));
        end.setUint32(0, 0x06054b50, true);
        end.setUint16(8, files.size, true);
        end.setUint16(10, files.size, true);
        end.setUint32(12, centralSize, true);
        end.setUint32(16, offset, true);
        return new Blob([...localParts, ...centralParts, end.buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
    }
    async function open(file) {
        const files = readZip(await file.arrayBuffer());
        async function part(path) {
            const entry = files.get(path);
            if (!entry) throw new Error('Parte ausente na planilha: ' + path);
            if (entry.method === 0) return xml(entry.data);
            if (entry.method !== 8) throw new Error('Compactação da planilha não compatível.');
            const stream = new Blob([entry.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
            return xml(await new Response(stream).arrayBuffer());
        }
        const book = await part('xl/workbook.xml');
        const sheet = tags(book, 'sheet').find(node => node.getAttribute('name') === 'Relatório Imply');
        if (!sheet) throw new Error('A aba Relatório Imply não foi encontrada.');
        const id = sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        const relations = await part('xl/_rels/workbook.xml.rels');
        const target = tags(relations, 'Relationship').find(node => node.getAttribute('Id') === id)?.getAttribute('Target');
        if (!target) throw new Error('A aba Relatório Imply está sem referência.');
        const path = new URL(target, 'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1);
        const doc = await part(path);
        const strings = files.has('xl/sharedStrings.xml')
            ? tags(await part('xl/sharedStrings.xml'), 'si').map(node => tags(node, 't').map(t => t.textContent).join(''))
            : [];
        function value(cell) {
            if (!cell) return '';
            const raw = tags(cell, 'v')[0]?.textContent ?? '';
            if (cell.getAttribute('t') === 's') return (strings[Number(raw)] ?? '').trim();
            if (cell.getAttribute('t') === 'inlineStr') return tags(cell, 't').map(t => t.textContent).join('').trim();
            return raw.trim();
        }
        const rows = [];
        let headerFound = false;
        for (const row of tags(doc, 'row')) {
            const number = Number(row.getAttribute('r'));
            const cells = new Map(tags(row, 'c').map(node => [node.getAttribute('r'), node]));
            const a = cells.get('A' + number);
            const b = cells.get('B' + number);
            const c = cells.get('C' + number);
            if (number === 4) headerFound = value(a) === 'Status' && value(b) === 'Voucher' && value(c) === 'Localizador';
            if (number > 4 && !/^CA$|^Cancelad[oa](?:\s*\(CA\))?$/i.test(value(a)) &&
                !value(b) && !(b && tags(b, 'f').length) && value(c)) {
                rows.push({number, row, b, c, locator: value(c)});
            }
        }
        if (!headerFound) throw new Error('Confira os cabeçalhos Status, Voucher e Localizador na linha 4.');
        return {
            rows,
            setVoucher(entry, code) {
                const element = (name) => doc.createElementNS(entry.row.namespaceURI,
                    entry.row.prefix ? entry.row.prefix + ':' + name : name);
                if (!entry.b) {
                    entry.b = element('c');
                    entry.b.setAttribute('r', 'B' + entry.number);
                    entry.row.insertBefore(entry.b, entry.c);
                }
                const inline = element('is');
                const text = element('t');
                text.textContent = code;
                inline.append(text);
                for (const node of [...entry.b.children]) {
                    if (node.localName === 'v' || node.localName === 'is') node.remove();
                }
                entry.b.setAttribute('t', 'inlineStr');
                entry.b.prepend(inline);
            },
            exportFile() {
                const data = encode.encode(new XMLSerializer().serializeToString(doc));
                Object.assign(files.get(path), {data, size: data.length, method: 0, crc: crc32(data)});
                return writeZip(files);
            }
        };
    }
    return {open};
})();
