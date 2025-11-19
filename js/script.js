/* ---------------------------
   FUNZIONI DI SUPPORTO ICS
   --------------------------- */

function unfoldICS(text){
  return text
    .replace(/\r\n[ \t]/g, '')
    .replace(/\n[ \t]/g, '');
}

function parseVEvent(block){
  const b = unfoldICS(block);

  const getField = (name) => {
    const re = new RegExp(name + '(?:;[^:]*)?:([\\s\\S]*?)(?:\\r?\\n[A-Z]|$)', 'i');
    const m = b.match(re);
    return m ? m[1].trim() : null;
  };

  const dtstartLine = (b.match(/DTSTART(?:;[^:]*)?:[^\r\n]*/i) || [null])[0];
  const dtendLine   = (b.match(/DTEND(?:;[^:]*)?:[^\r\n]*/i) || [null])[0];

  const summary     = getField('SUMMARY') || '(senza titolo)';
  const description = getField('DESCRIPTION') || '';
  const location    = getField('LOCATION') || '';

  const extractValue = (line) => {
    if (!line) return null;
    return line.split(':').slice(1).join(':').trim();
  };

  return {
    summary,
    description,
    location,
    dtstartRaw: extractValue(dtstartLine),
    dtendRaw:   extractValue(dtendLine)
  };
}

function parseICSToDate(raw){
  if (!raw) return null;
  raw = raw.trim();

  // Formato: 01.12.2020 16:39 CET
  let m = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s+([A-Za-z]+))?$/);
  if (m){
    const d = Number(m[1]), mo = Number(m[2]) - 1, y = Number(m[3]);
    const hh = Number(m[4]), mm = Number(m[5]), ss = m[6]?Number(m[6]):0;
    return new Date(y, mo, d, hh, mm, ss);
  }

  // Formato ICS standard
  m = raw.match(/^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?Z?$/);
  if (m){
    const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
    const hh = m[4]?Number(m[4]):0, mm = m[5]?Number(m[5]):0, ss = m[6]?Number(m[6]):0;
    if (raw.endsWith("Z")){
      return new Date(Date.UTC(y, mo, d, hh, mm, ss));
    }
    return new Date(y, mo, d, hh, mm, ss);
  }

  // Fallback
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
    ev.dtend   = parseICSToDate(ev.dtendRaw);
    ev.isoDate = ev.dtstart ? toISODate(ev.dtstart) : null;
    return ev;
  });

  return parsed.filter(ev => ev.isoDate === targetISO);
}


/* ---------------------------
   RIPULITURA HTML DA DESCRIPTION
   --------------------------- */

function stripHTML(str){
  if (!str) return "";
  return str
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/?[^>]+>/gi, "")
    .trim();
}


/* ---------------------------
   GENERAZIONE PDF
   --------------------------- */

function generatePDFfromEvents(events, dateISO){
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ unit:'mm', format:'a4' });
  let y = 15;

  pdf.setFontSize(18);
  pdf.text(`Eventi del ${dateISO}`, 14, y);
  y += 10;

  pdf.setFontSize(11);

  if (!events.length){
    pdf.text("Nessun evento trovato.", 14, y);
  } else {
    events.forEach(ev => {
      pdf.setFontSize(13);
      pdf.text(ev.summary, 14, y);
      y += 6;

      pdf.setFontSize(10);

      let timeStr = "Orario non definito";
      if (ev.dtstart){
        const s = ev.dtstart.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});
        if (ev.dtend){
          const e = ev.dtend.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});
          timeStr = `${s} - ${e}`;
        } else {
          timeStr = s;
        }
      }

      pdf.text(timeStr, 14, y);
      y += 5;

      if (ev.location){
        const loc = pdf.splitTextToSize("Luogo: " + ev.location, 180);
        pdf.text(loc, 14, y);
        y += loc.length * 5;
      }

      if (ev.description){
        let desc = ev.description
          .replace(/\\n/g, "\n")
          .replace(/\\,/, ",");

        desc = stripHTML(desc);

        const lines = pdf.splitTextToSize(desc, 180);
        pdf.text(lines, 14, y);
        y += lines.length * 5;
      }

      y += 6;
      if (y > 280){
        pdf.addPage();
        y = 20;
      }
    });
  }

  pdf.save(`eventi_${dateISO}.pdf`);
}


/* ---------------------------
   EVENT HANDLER
   --------------------------- */

document.getElementById('go').addEventListener('click', async () => {
  const f = document.getElementById('icsFile').files[0];
  const dateVal = document.getElementById('dateInput').value;

  if (!f){ alert('Seleziona un file .ics'); return; }
  if (!dateVal){ alert('Seleziona una data'); return; }

  try {
    let events = await handleFile(f, dateVal);

    // Ordina per orario di inizio
    events.sort((a, b) => {
      if (!a.dtstart) return 1;
      if (!b.dtstart) return -1;
      return a.dtstart - b.dtstart;
    });

    generatePDFfromEvents(events, dateVal);

  } catch (err){
    console.error(err);
    alert("Errore nel parsing del file .ics. Controlla la console.");
  }
});
