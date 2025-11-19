// ICS Parsing Functions
function unfoldICS(text){
    return text.replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '');
}

function parseVEvent(block){
    const b = unfoldICS(block);
    const getField = (name) => {
        const re = new RegExp(name + '(?:;[^:]*)?:([\\s\\S]*?)(?:\\r?\\n[A-Z]|$)', 'i');
        const m = b.match(re);
        return m ? m[1].trim() : null;
    };
    const dtstartLine = (b.match(/DTSTART(?:;[^:]*)?:[^\r\n]*/i) || [null])[0];
    const dtendLine = (b.match(/DTEND(?:;[^:]*)?:[^\r\n]*/i) || [null])[0];
    const summary = getField('SUMMARY') || '(senza titolo)';
    const description = getField('DESCRIPTION') || '';
    const location = getField('LOCATION') || '';
    const extractValue = (line) => {
        if (!line) return null;
        return line.split(':').slice(1).join(':').trim();
    };
    return {
        summary,
        description,
        location,
        dtstartRaw: extractValue(dtstartLine),
        dtendRaw: extractValue(dtendLine)
    };
}

function parseICSToDate(raw){
    if (!raw) return null;
    raw = raw.trim();
    let m = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s+([A-Za-z]+))?$/);
    if (m){
        const d = Number(m[1]), mo = Number(m[2]) - 1, y = Number(m[3]);
        const hh = Number(m[4]), mm = Number(m[5]), ss = m[6]?Number(m[6]):0;
        return new Date(y, mo, d, hh, mm, ss);
    }
    m = raw.match(/^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?Z?$/);
    if (m){
        const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
        const hh = m[4]?Number(m[4]):0, mm = m[5]?Number(m[5]):0, ss = m[6]?Number(m[6]):0;
        if (raw.endsWith("Z")){
            return new Date(Date.UTC(y, mo, d, hh, mm, ss));
        }
        return new Date(y, mo, d, hh, mm, ss);
    }
    const p = Date.parse(raw);
    if (!isNaN(p)) return new Date(p);
    return null;
}

function toISODate(d){
    if (!d) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${day}`;
}

async function handleFile(file, targetISO){
    const text = await file.text();
    const unfolded = unfoldICS(text);
    const vevents = [];
    const regex = /BEGIN:VEVENT([\s\S]*?)END:VEVENT/ig;
    let match;
    while ((match = regex.exec(unfolded)) !== null){
        vevents.push(match[1]);
    }
    const parsed = vevents.map(block => {
        const ev = parseVEvent(block);
        ev.dtstart = parseICSToDate(ev.dtstartRaw);
        ev.dtend = parseICSToDate(ev.dtendRaw);
        ev.isoDate = ev.dtstart ? toISODate(ev.dtstart) : null;
        return ev;
    });
    return parsed.filter(ev => ev.isoDate === targetISO);
}

function stripHTML(str){
    if (!str) return "";
    return str.replace(/<br\s*\/?>/gi, "\n").replace(/<\/?[^>]+>/gi, "").trim();
}

function extractEmail(text){
    if (!text) return "";
    const emailPattern = /([a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/;
    const match = text.match(emailPattern);
    return match ? match[1].replace(/^n/, "") : "";
}

function extractFields(description){
    const desc = stripHTML(description);
    let nome = "";
    let prenotatoDa = "";
    let mNome = desc.match(/Nome alunno\s*[:\-]?\s*(.*)/i);
    if(mNome) nome = mNome[1].split("\n")[0].trim().replace(/\n/g, "").replace(/\\n/g, "");
    let mPren = desc.match(/Prenotato da\s*[:\-]?\s*(.*)/i);
    if(mPren) {
        let fullText = mPren[1].split("\n")[0].trim();
        prenotatoDa = fullText.replace(/\s*[a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\s*/gi, "").trim();
        prenotatoDa = prenotatoDa.replace(/\n/g, "").replace(/\\n/g, "").replace(/\\/g, "").trim();
    }
    return {nome, prenotatoDa};
}

function generateXLSfromEvents(events, dateISO){
    const data = events.map(ev => {
        const {nome, prenotatoDa} = extractFields(ev.description);
        let timeStr = "Orario non definito";
        if(ev.dtstart){
            const s = ev.dtstart.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});
            if(ev.dtend){
                const e = ev.dtend.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});
                timeStr = `${s} - ${e}`;
            } else {
                timeStr = s;
            }
        }
        return {
            "Nome alunno": nome,
            "Orario": timeStr,
            "Classe": ev.location || "",
            "Prenotato da": prenotatoDa,
            "Mail": extractEmail(stripHTML(ev.description)) || ""
        };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Eventi " + dateISO);
    XLSX.writeFile(wb, `eventi_${dateISO}.xlsx`);
}

// UI Functions
function showFeedback(message, type) {
    const feedbackArea = document.getElementById('feedbackArea');
    feedbackArea.innerHTML = `<div class="feedback ${type}">${message}</div>`;
    setTimeout(() => { feedbackArea.innerHTML = ''; }, 5000);
}

// File handling
const fileInput = document.getElementById('icsFile');
const dropZone = document.getElementById('dropZone');
const fileName = document.getElementById('fileName');
const goBtn = document.getElementById('goBtn');

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        fileName.textContent = file.name;
        fileName.style.display = 'block';
    }
});

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        fileInput.files = files;
        fileName.textContent = files[0].name;
        fileName.style.display = 'block';
    }
});

goBtn.addEventListener('click', async () => {
    const f = fileInput.files[0];
    const dateVal = document.getElementById('dateInput').value;

    if (!f) {
        showFeedback('Seleziona un file .ics', 'error');
        return;
    }
    if (!dateVal) {
        showFeedback('Seleziona una data', 'error');
        return;
    }

    goBtn.classList.add('loading');
    goBtn.disabled = true;

    try {
        let events = await handleFile(f, dateVal);
        events.sort((a, b) => {
            if (!a.dtstart) return 1;
            if (!b.dtstart) return -1;
            return a.dtstart - b.dtstart;
        });
        generateXLSfromEvents(events, dateVal);
        showFeedback(`File Excel generato con successo (${events.length} eventi)`, 'success');
    } catch (err) {
        console.error(err);
        showFeedback('Errore nel parsing del file .ics', 'error');
    } finally {
        goBtn.classList.remove('loading');
        goBtn.disabled = false;
    }
});
