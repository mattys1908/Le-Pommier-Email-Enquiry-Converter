let DB = JSON.parse(localStorage.getItem('enq_v2') || '[]');
let editingId = null;

function persist() { localStorage.setItem('enq_v2', JSON.stringify(DB)); }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }
function today() { return new Date().toISOString().split('T')[0]; }
function fmtDate(d) { if (!d) return '—'; const [y,m,day]=d.split('-'); return `${day}/${m}/${y}`; }
// Enquiry date export format: 1-Mar or 1-Mar-26 if year differs from current
function fmtEnquiryDate(d) {
  if (!d) return '';
  const [y,m,day] = d.split('-');
  const date = new Date(parseInt(y), parseInt(m)-1, parseInt(day));
  if (isNaN(date.getTime())) return '';
  const mon = date.toLocaleDateString('en-US', { month: 'short' });
  const currentYear = new Date().getFullYear();
  const yr = parseInt(y);
  if (yr !== currentYear) {
    return `${parseInt(day)}-${mon}-${String(yr).slice(-2)}`;
  }
  return `${parseInt(day)}-${mon}`;
}
// SharePoint "Date of function" format: Saturday, August 29, 2026
function fmtEventDate(d) {
  if (!d) return '';
  const [y,m,day] = d.split('-');
  const date = new Date(parseInt(y), parseInt(m)-1, parseInt(day));
  if (isNaN(date.getTime())) return '';
  return date.toLocaleDateString('en-US', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
}
function esc(s) { return s ? String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : ''; }
function showToast(msg) {
  const t=document.getElementById('toast');
  const isDelete = msg.toLowerCase().includes('delet');
  t.innerHTML = (isDelete ? '🗑 ' : '✓ ') + msg;
  t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'),2800);
}

// ── PASTE HANDLING ──
function handleInput() {
  const val = document.getElementById('paste-input').value.trim();
  document.getElementById('parse-btn').classList.toggle('hidden', val.length < 20);
}
function scheduleParse() { setTimeout(doParse, 100); }
function doParse() {
  const raw = document.getElementById('paste-input').value;
  if (!raw.trim()) return;
  const parsed = parseEmail(raw);
  document.getElementById('paste-input').value = '';
  document.getElementById('parse-btn').classList.add('hidden');
  openModal(parsed);
}

// ══════════════════════════════════════════════
// PARSER — handles 3 source formats:
//   A) Pink Book  — inline "Name: X  Email address: X …"
//   B) Plain / forwarded email — name signed at bottom, Tel: for phone
//   C) Le Pommier contact form — label line then value line
// ══════════════════════════════════════════════
function parseEmail(raw) {
  const lines = raw.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n').map(l=>l.trim());
  const text  = lines.join('\n');

  const MONTHS = {january:1,february:2,march:3,april:4,may:5,june:6,july:7,
                  august:8,september:9,october:10,november:11,december:12,
                  jan:1,feb:2,mar:3,apr:4,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  const INTERNAL = ['lepommier','digitalniche','pink-book','eventmanagementsolutions'];

  function nextVal(re) {
    for (let i=0;i<lines.length;i++) {
      const l = lines[i];
      // Mac format: label stands alone on its line, value on the next non-empty line
      if (re.test(l)) {
        const rest = l.replace(re,'').trim();
        if (!rest) {
          // pure label line — grab the next non-empty line as the value
          for (let j=i+1;j<lines.length;j++) { if (lines[j]) return lines[j]; }
        } else {
          // Windows format: label and value on same line, e.g. "Name Natalie" or "Name: Natalie"
          // rest already stripped the label; return it if it looks like a value (not another label)
          return rest.replace(/^:\s*/,'');
        }
      }
    }
    return '';
  }
  function inlineVal(re) { const m=text.match(re); return m?m[1].trim():''; }
  function cap(s) { return s ? s.charAt(0).toUpperCase()+s.slice(1) : ''; }

  function parseDate(s) {
    if (!s) return '';
    let m=s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
    if (m) { let y=m[3]; if(y.length===2)y='20'+y; return `${y}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
    m=s.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
    if (m) { const mn=MONTHS[m[2].toLowerCase()]; if(mn) return `${m[3]}-${String(mn).padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
    return '';
  }

  function findEventDate(src) {
    let m=src.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
    if (m) { let y=m[3]; if(y.length===2)y='20'+y; return `${y}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
    const MW=Object.keys(MONTHS).join('|');
    const r1=new RegExp(`(\\d{1,2})(?:st|nd|rd|th)?\\s+(?:of\\s+)?(${MW})(?:[,\\s]+(\\d{4}))?`,'i');
    const r2=new RegExp(`(${MW})\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:[,\\s]+(\\d{4}))?`,'i');
    m=src.match(r1);
    if (m) { const mn=MONTHS[m[2].toLowerCase()]; const y=m[3]||new Date().getFullYear(); if(mn) return `${y}-${String(mn).padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
    m=src.match(r2);
    if (m) { const mn=MONTHS[m[1].toLowerCase()]; const y=m[3]||new Date().getFullYear(); if(mn) return `${y}-${String(mn).padStart(2,'0')}-${m[2].padStart(2,'0')}`; }
    return '';
  }

  // Detect function type from free text (message body / subject)
  function detectFuncType(src) {
    if (!src) return '';
    // Months — so "21st of July" is not treated as a 21st birthday
    const mon = 'jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?';
    // Require ordinal suffix so "60 people" is not a 60th birthday
    const notDateOrd = (n) => new RegExp(
      String.raw`\b${n}(?:st|nd|rd|th|ste)\b(?!\s+(?:of\s+)?(?:${mon}))`,
      'i'
    );
    const funcMap = [
      // Weddings & related
      [/(troue|wedding|bruilof|huwelik)/i,                         'Wedding'],
      [/(engagement|verlowing)/i,                                  'Engagement'],
      [/(bridal\s*shower|bruidsaand|bachelorette|hen\s*party)/i,   'Bridal Shower'],
      [/(bachelor|stag\s*(?:do|party)|jonkmans\s*partytjie)/i,     'Bachelor Party'],
      [/(baby\s*shower)/i,                                         'Baby Shower'],
      // Birthdays — "21st party/birthday" ahead of the date; don't let "Party on the 21st of July" match
      [/(?:21st|21ste|een-en-twintigste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag|verjaardag)|(?:birthday|verjaarsdag|verjaardag).{0,30}(?:21st|21ste)|een-en-twintigste/i, '21st Birthday'],
      [notDateOrd(21),                                             '21st Birthday'],
      [/(?:30th|30ste|dertigste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag)|(?:birthday|verjaarsdag).{0,30}(?:30th|30ste)|dertigste/i, '30th Birthday'],
      [notDateOrd(30),                                             '30th Birthday'],
      [/(?:40th|40ste|veertigste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag)|(?:birthday|verjaarsdag).{0,30}(?:40th|40ste)/i, '40th Birthday'],
      [notDateOrd(40),                                             '40th Birthday'],
      [/(?:50th|50ste|vyftigste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag)|(?:birthday|verjaarsdag).{0,30}(?:50th|50ste)/i, '50th Birthday'],
      [notDateOrd(50),                                             '50th Birthday'],
      [/(?:60th|60ste|sestigste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag)|(?:birthday|verjaarsdag).{0,30}(?:60th|60ste)/i, '60th Birthday'],
      [notDateOrd(60),                                             '60th Birthday'],
      [/(?:70th|70ste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag)|(?:birthday|verjaarsdag).{0,30}(?:70th|70ste)/i, '70th Birthday'],
      [notDateOrd(70),                                             '70th Birthday'],
      [/(?:80th|80ste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag)|(?:birthday|verjaarsdag).{0,30}(?:80th|80ste)/i, '80th Birthday'],
      [notDateOrd(80),                                             '80th Birthday'],
      [/(?:90th|90ste)\s*(?:birthday|b[\s-]?day|party|verjaarsdag)|(?:birthday|verjaarsdag).{0,30}(?:90th|90ste)/i, '90th Birthday'],
      [notDateOrd(90),                                             '90th Birthday'],
      [/(birthday|verjaarsdag|verjaardag|b[\s-]?day)\b/i,          'Birthday'],
      [/(anniversary|herdenking|trou.?herdenking)/i,               'Anniversary'],
      // Faith / family
      [/(baptism|doop|christening|doopfees)/i,                     'Baptism'],
      [/(confirmation|belydenis)/i,                                'Confirmation'],
      [/(funeral|begrafnis|memorial\s*service|herdenkingsdiens)/i, 'Memorial'],
      // Corporate / formal
      [/(conference|konferensie)/i,                                'Conference'],
      [/(seminar|workshop|werkwinkel)/i,                           'Workshop'],
      [/(meeting|vergadering|board\s*meeting)/i,                   'Meeting'],
      [/(team[\s-]*building|spanbou)/i,                            'Team Building'],
      [/(corporate|korporatiewe\s*funksie|company\s*function)/i,   'Corporate'],
      [/(product\s*launch|bekendstelling|launch\s*event)/i,        'Product Launch'],
      [/(awards?\s*(?:evening|night|ceremony|dinner)|prysuitdeling)/i, 'Awards'],
      [/(fundrais(?:er|ing)|fondsinsameling)/i,                    'Fundraiser'],
      [/(gala)/i,                                                  'Gala'],
      // Year-end / seasonal
      [/(year[\s.-]*end|jaar[\s.-]*end|kersfees|christmas|nye|new\s*year)/i, 'Year-End'],
      // School / academic
      [/(matric|gr[. ]*?12|matriek)/i,                             'Matric Function'],
      [/(graduation|gradeer|graduandi)/i,                          'Graduation'],
      [/(reunion|re\u00FCnie|klasre[uü]nie)/i,                     'Reunion'],
      // Meals / social
      [/(braai|bbq|barbecue)/i,                                    'Braai'],
      [/(cocktail)/i,                                              'Cocktail Party'],
      [/(dinner|aandete|galadinee)/i,                              'Dinner'],
      [/(lunch|middagete|luncheon)/i,                              'Lunch'],
      [/(breakfast|ontbyt)/i,                                      'Breakfast'],
      [/(high\s*tea)/i,                                            'High Tea'],
      // Catch-alls (least specific — keep last)
      [/(party|partytjie|fees|feestjie|celebration|viering)/i,     'Party'],
    ];
    for (const [re, label] of funcMap) {
      if (re.test(src)) return label;
    }
    return '';
  }

  // Website "Enquiry" dropdown values that are not a real function type
  function isGenericEnquiry(s) {
    return !s || /^(event|general|other|contact)\s*enquir/i.test(s);
  }

  const isPinkBook    = /pink-book\.co\.za|Listing title:|Listing URL:/i.test(text);
  // Mac: label on its own line, value on next line. Windows: label and value collapsed on same line.
  const isContactForm = (lines.some(l=>/^Name$/i.test(l)) && lines.some(l=>/^(Last|Last Name)$/i.test(l)))
    || (lines.some(l=>/^Name\s+\S/i.test(l)) && lines.some(l=>/^(Last|Last Name|Surname)\s+\S/i.test(l)))
    || (lines.some(l=>/^Name:\s+\S/i.test(l)) && lines.some(l=>/^(Last|Last Name|Surname):\s+\S/i.test(l)));
  // A sign-off means it's a proper email regardless of length or missing headers
  const hasSignOff = /^(Kind Regards|Best Regards|Regards,?|Yours sincerely|Warm regards|Best wishes|Groete|Baie dankie)/im.test(text);
  // WhatsApp: short, conversational, no email structure and no formal sign-off
  const isWhatsApp    = !isPinkBook && !isContactForm && !hasSignOff
    && !text.includes('From:') && !text.includes('Subject:')
    && text.length < 600;

  let firstName='', lastName='', email='', phone='', enquiryDate=today(), eventDate='', message='', funcType='', pax='', heardFrom='';

  // ── FORMAT A: PINK BOOK ──
  if (isPinkBook) {
    const fullName = inlineVal(/Name:\s*([^\n\r]+)/i);
    const parts = fullName.trim().split(/\s+/);
    firstName = parts[0]||''; lastName = parts.slice(1).join(' ')||'';
    email = inlineVal(/Email\s*(?:address)?:\s*([\w.+-]+@[\w.-]+)/i);
    phone = inlineVal(/Contact\s*(?:number)?:\s*([+0-9][0-9 \-]{6,15})/i);
    const wd = inlineVal(/Wedding\s*Date:\s*([^\n\r,]+)/i);
    eventDate = parseDate(wd.trim());
    funcType = 'Wedding';
    heardFrom = 'Pink Book';
    // PAX from "Number of guests: 50"
    const guestM = text.match(/(?:number of guests?|pax)[:\s]*([\d]+)/i);
    if (guestM) pax = guestM[1];
    const det = inlineVal(/Details:\s*([\s\S]{10,})/i);
    message = det ? det.split(/Listing title:/i)[0].trim().slice(0,800) : '';
    const dh=text.match(/Date:\s*\w+,?\s+(\d{1,2})\s+(\w+)\s+(\d{4})/i);
    if (dh) { const mn=MONTHS[dh[2].toLowerCase()]; if(mn) enquiryDate=`${dh[3]}-${String(mn).padStart(2,'0')}-${dh[1].padStart(2,'0')}`; }

  // ── FORMAT C: LE POMMIER CONTACT FORM ──
  } else if (isContactForm) {
    firstName = nextVal(/^Name$/i);
    lastName  = nextVal(/^(Last|Last Name|Surname)$/i);
    email     = nextVal(/^(Email|E-mail|Email Address)$/i);
    phone     = nextVal(/^(Phone|Tel|Mobile|Contact Number|Number|Cell)$/i);
    const dh=text.match(/Date:\s*\w+,?\s+(\d{1,2})\s+(\w+)\s+(\d{4})/i);
    if (dh) { const mn=MONTHS[dh[2].toLowerCase()]; if(mn) enquiryDate=`${dh[3]}-${String(mn).padStart(2,'0')}-${dh[1].padStart(2,'0')}`; }
    // Enquiry label often includes "(Wedding enquiries…)" — don't treat that note as the value
    for (let i = 0; i < lines.length; i++) {
      const m = lines[i].match(/^Enquiry(?:\s*\([^)]*\))?\s*:?\s*(.*)$/i);
      if (!m) continue;
      const rest = m[1].trim();
      if (rest) { funcType = rest; break; }
      for (let j = i + 1; j < lines.length; j++) {
        if (lines[j]) { funcType = lines[j]; break; }
      }
      break;
    }
    funcType = funcType.replace(/\(.*?\)/g, '').trim();
    heardFrom = 'Website';
    // PAX sometimes mentioned in message
    const cpaxM = text.match(/(\d+)\s*(?:pax|people|persons?|guests?|mense|gaste)/i);
    if (cpaxM) pax = cpaxM[1];
    const mi=lines.findIndex(l=>/^Message$/i.test(l));
    if (mi!==-1) {
      const ml=[];
      for (let i=mi+1;i<lines.length;i++) {
        if (/^(Kind Regards|Regards|Sent from|--)/i.test(lines[i])) break;
        ml.push(lines[i]);
      }
      message=ml.join('\n').trim().slice(0,800);
    }
    // Windows: "Message" and body on same line, or "Message: body text..."
    if (!message) { const m=text.match(/^Message[:\s]+(.{10,})/im); if(m) message=m[1].trim().slice(0,800); }
    const msgPart=mi>=0?text.slice(text.indexOf(lines[mi])):text;
    eventDate=findEventDate(msgPart);
    // Prefer type found in the message over the website Enquiry dropdown
    // (dropdown is often just "Event enquiry"; labels also say "Wedding enquiries…")
    const fromMsg = detectFuncType(message);
    if (fromMsg) {
      funcType = fromMsg;
    } else if (isGenericEnquiry(funcType)) {
      funcType = '';
    } else {
      // Keep a real dropdown value like "Wedding" when the message has no keyword
      funcType = detectFuncType(funcType) || funcType;
    }

  // ── FORMAT D: WHATSAPP ──
  } else if (isWhatsApp) {
    heardFrom = 'WhatsApp';

    // Strip WhatsApp timestamp prefix e.g. [2026/03/29, 08:49:56] Ma:
    const cleanText = text.replace(/^\[[\d\/\-, :]+\]\s*\w+:\s*/gm, '').trim();
    message = cleanText.slice(0, 800);

    // Name detection — tries patterns in order, most specific first
    const waSignOffRe = /^(Kind Regards|Regards,?|Best Regards|Best wishes|Yours sincerely|Warm regards|Groete|Baie dankie|Thank you|Thanks)/i;
    const waGreetingWords = /^(Good|Hi|Hello|Dear|Goeie|Goedag|Hallo|Morning|Afternoon|Evening|Kind|Warm|Best|Yours)/i;
    const waCleanLines = cleanText.split('\n').map(l => l.trim()).filter(Boolean);

    let nameM =
      // "Dis Rozaan hier" (Afrikaans intro)
      cleanText.match(/Dis\s+([A-Z\u00C0-\u00FFa-z][a-z\u00C0-\u00FF'-]+(?:\s+[A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+)?)\s+hier/i) ||
      // "ek is / I'm / My name is X"
      cleanText.match(/ek\s+is\s+([A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+(?:\s+[A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+)?)/i) ||
      cleanText.match(/I(?:'m| am)\s+([A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+(?:\s+[A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+)?)/i) ||
      cleanText.match(/My name is\s+([A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+(?:\s+[A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+)?)/i) ||
      // "Thanks/Thank you Melanie" at end
      cleanText.match(/Thank(?:s| you)[,.]?\s+([A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+(?:\s+[A-Z\u00C0-\u00FF][a-z\u00C0-\u00FF'-]+)?)\s*$/im);

    // Sign-off → name (same logic as plain email parser)
    // e.g. "Kind regards\nMizelle"
    if (!nameM) {
      const soIdx = waCleanLines.findIndex(l => waSignOffRe.test(l));
      if (soIdx >= 0) {
        for (let i = soIdx + 1; i < Math.min(waCleanLines.length, soIdx + 4); i++) {
          const l = waCleanLines[i];
          if (l && /^[A-Z][a-zA-Z\u00C0-\u00FF'-]+(?:\s+[A-Z][a-zA-Z\u00C0-\u00FF'-]+)?$/.test(l) && !waGreetingWords.test(l)) {
            const p = l.split(/\s+/); firstName = cap(p[0]); lastName = cap(p.slice(1).join(' ')); break;
          }
        }
      }
    }

    // Name-first: capitalised word at very start, excluding greeting/sign-off words
    if (!nameM && !firstName) {
      nameM = cleanText.match(/^(?!(?:Good|Hi|Hello|Dear|Goeie|Goedag|Hallo|Morning|Afternoon|Evening|Kind|Warm|Best|Yours|Thank|Regards)\b)([A-Z][a-z\u00C0-\u00FF'-]{2,}(?:\s+[A-Z][a-z\u00C0-\u00FF'-]+)?)[,\s]/m);
    }

    // Initials at end e.g. "JR"
    if (!nameM && !firstName) {
      const ll = waCleanLines.slice(-1)[0]||'';
      const im = ll.match(/^([A-Z]{2,4})$/);
      if (im) nameM = im;
    }

    if (!firstName && nameM) { const p = nameM[1].trim().split(/\s+/); firstName = cap(p[0]); lastName = cap(p.slice(1).join(' ')); }

    // Email — grab any email in the message (no domain filtering for WhatsApp)
    const waEmailM = cleanText.match(/[\w.+-]+@[\w-]+\.[\w.]+/);
    if (waEmailM) email = waEmailM[0].trim();

    // Phone — SA number
    const phM = text.match(/(?:\+27|0)[0-9][\d\s\-]{7,12}/);
    if (phM) phone = phM[0].trim();

    // PAX — always use the UPPER bound of a range
    const rangeM = cleanText.match(/\+?-?\s*(\d+)\s*(?:tot|to|[-\u2013])\s*(\d+)/i);
    if (rangeM) {
      pax = String(Math.max(parseInt(rangeM[1]), parseInt(rangeM[2])));
    } else {
      const paxM =
        cleanText.match(/pax\s+of\s+(\d+)/i) ||
        cleanText.match(/(\d+)\s*(?:pax|people|persons?|guests?|delegates?|mense|gaste|persone|persona)/i) ||
        cleanText.match(/(?:approximately|roughly|about|omtrent|sowat|[±])\s*(\d+)/i);
      if (paxM) pax = paxM[1];
    }

    // Function type — Afrikaans + English
    funcType = detectFuncType(cleanText);

    // Event date — use cleanText so timestamps don't create false dates
    eventDate = findEventDate(cleanText);

  // FORMAT B: PLAIN / FORWARDED EMAIL
  } else {
    // Expanded sign-off patterns
    const signOffRe = /^(Kind Regards|Regards,?|Best Regards|Best wishes|Yours sincerely|Warm regards|Groete|Baie dankie)/i;

    // Name: scan forward from sign-off for first capitalised name-like line
    const signOffIdx = lines.findIndex(l => signOffRe.test(l));
    if (signOffIdx >= 0) {
      for (let i = signOffIdx + 1; i < Math.min(lines.length, signOffIdx + 6); i++) {
        const l = lines[i];
        if (l && /^[A-Z][a-zA-Z'-]+(?:\s+[A-Z][a-zA-Z'-]+){0,2}$/.test(l) && !/^(Tel|www|http)/i.test(l)) {
          const p = l.split(/\s+/); firstName = p[0]; lastName = p.slice(1).join(' '); break;
        }
      }
    }
    // Fallback: backwards from Tel: line
    if (!firstName) {
      const telIdx = lines.findIndex(l => /^Tel[:\s]/i.test(l));
      if (telIdx > 0) {
        for (let i = telIdx-1; i >= Math.max(0, telIdx-5); i--) {
          const l = lines[i];
          if (l && /^[A-Z][a-zA-Z'-]+(?: [A-Z][a-zA-Z'-]+)?$/.test(l)) {
            const p = l.split(/\s+/); firstName = p[0]; lastName = p.slice(1).join(' '); break;
          }
        }
      }
    }
    // Fallback: From: header innermost occurrence
    if (!firstName) {
      const froms = [...text.matchAll(/From:\s*([A-Z][a-zA-Z'-]+(?:\s+[A-Z][a-zA-Z'-]+)?)\s*[<\n]/g)];
      if (froms.length) { const last = froms[froms.length-1]; const p = last[1].trim().split(/\s+/); firstName = p[0]; lastName = p.slice(1).join(' '); }
    }
    // Email — exclude internal domains
    const allEmails = [...text.matchAll(/[\w.+-]+@[\w-]+\.[\w.]+/g)].map(m => m[0]);
    email = allEmails.find(e => !INTERNAL.some(d => e.includes(d))) || '';
    // Phone
    const telLine = lines.find(l => /^Tel[:\s]/i.test(l));
    if (telLine) phone = telLine.replace(/^Tel[:\s]*/i, '').trim();
    if (!phone) { const m = text.match(/(?:\+[\d\s\-]{8,15}|0[0-9][\d\s\-]{7,12})/); if (m) phone = m[0].trim(); }
    // Enquiry date from innermost Date: header
    const allDates = [...text.matchAll(/Date:\s*\w+,?\s+(\d{1,2})\s+(\w+)\s+(\d{4})/gi)];
    if (allDates.length) { const last = allDates[allDates.length-1]; const mn = MONTHS[last[2].toLowerCase()]; if (mn) enquiryDate = `${last[3]}-${String(mn).padStart(2,'0')}-${last[1].padStart(2,'0')}`; }
    // heardFrom; pax
    heardFrom = 'Website';
    const epaxM = text.match(/(\d+)\s*(?:pax|people|persons?|guests?|delegates?|mense|gaste)/i);
    if (epaxM) pax = epaxM[1];
    // Function type: try known keywords from subject, then message body
    const subjMatch = text.match(/Subject:\s*(?:FW:|RE:|FWD:)?\s*(.+)/i);
    if (subjMatch) {
      const subjText = subjMatch[1].trim();
      funcType = detectFuncType(subjText);
      // No keyword — use subject but strip trailing date/year noise
      if (!funcType) funcType = subjText.replace(/\s*\d{1,2}\s+\w+\s+\d{4}\s*$/, '').replace(/\s*\d{4}\s*$/, '').trim();
    }
    // Message body: greeting to sign-off
    const lastSubjIdx = lines.map((l,i) => /^Subject:/i.test(l) ? i : -1).filter(i => i >= 0).pop() || 0;
    const bodyLines = lines.slice(lastSubjIdx + 1).filter(l => l);
    const msgStart = bodyLines.findIndex(l => /^(Dear|Hello|Hi |Good (day|morning|afternoon|evening))/i.test(l));
    const bodyFrom = msgStart >= 0 ? bodyLines.slice(msgStart) : bodyLines;
    const msgEnd = bodyFrom.findIndex(l => signOffRe.test(l) || /^(Thank you|Tel:|www\.)/i.test(l));
    const bodySlice = msgEnd > 0 ? bodyFrom.slice(0, msgEnd) : bodyFrom;
    message = bodySlice.join('\n').trim().slice(0, 800);
    // If subject had no real function type (or was generic), scan the body
    if (!funcType || isGenericEnquiry(funcType) || /^Contact Us Form/i.test(funcType)) {
      const fromMsg = detectFuncType(message);
      if (fromMsg) funcType = fromMsg;
    }
    // Event date
    const searchText = (subjMatch ? subjMatch[1] : '') + '\n' + text;
    eventDate = findEventDate(searchText);
  }
  return { firstName:cap(firstName.trim()), lastName:cap(lastName.trim()), email, phone, enquiryDate, eventDate, message, funcType, pax, heardFrom };
}

// ── MODAL ──
function openModal(p={}) {
  editingId = null;
  document.getElementById('modal-title').textContent = '✦ New Enquiry';
  document.getElementById('save-btn').textContent = 'Save Enquiry';
  document.getElementById('f-first').value = p.firstName||'';
  document.getElementById('f-last').value  = p.lastName||'';
  document.getElementById('f-email').value = p.email||'';
  document.getElementById('f-phone').value = p.phone||'';
  document.getElementById('f-venue').value = p.venue||'';
  document.getElementById('f-type').value  = p.funcType||'';
  document.getElementById('f-pax').value   = p.pax||'';
  document.getElementById('f-heard').value = p.heardFrom||'';
  document.getElementById('f-enquiry-date').value = p.enquiryDate||today();
  document.getElementById('f-event-date').value   = p.eventDate||'';
  document.getElementById('f-notes').value = p.message||'';
  const mw=document.getElementById('msg-wrap');
  if (p.message&&p.message.length>5) { mw.classList.remove('hidden'); document.getElementById('msg-text').textContent=p.message; }
  else mw.classList.add('hidden');
  document.getElementById('modal-overlay').classList.add('active');
  setTimeout(()=>document.getElementById('f-first').focus(),250);
}

function openEdit(id) {
  const row=DB.find(r=>r.id===id); if(!row) return;
  editingId=id;
  document.getElementById('modal-title').textContent = 'Edit Enquiry';
  document.getElementById('save-btn').textContent = 'Update';
  document.getElementById('f-first').value = row.firstName||'';
  document.getElementById('f-last').value  = row.lastName||'';
  document.getElementById('f-email').value = row.email||'';
  document.getElementById('f-phone').value = row.phone||'';
  document.getElementById('f-venue').value = row.venue||'';
  document.getElementById('f-type').value  = row.funcType||'';
  document.getElementById('f-pax').value   = row.pax||'';
  document.getElementById('f-heard').value = row.heardFrom||'';
  document.getElementById('f-enquiry-date').value = row.enquiryDate||today();
  document.getElementById('f-event-date').value   = row.eventDate||'';
  document.getElementById('f-notes').value = row.notes||'';
  document.getElementById('msg-wrap').classList.add('hidden');
  document.getElementById('modal-overlay').classList.add('active');
}

function closeModal() { document.getElementById('modal-overlay').classList.remove('active'); }
function overlayClick(e) { if(e.target===document.getElementById('modal-overlay')) closeModal(); }

function getRowFromForm() {
  return {
    firstName: document.getElementById('f-first').value.trim(),
    lastName:  document.getElementById('f-last').value.trim(),
    email:     document.getElementById('f-email').value.trim(),
    phone:     document.getElementById('f-phone').value.trim(),
    venue:     document.getElementById('f-venue').value.trim(),
    funcType:  document.getElementById('f-type').value.trim(),
    pax:       document.getElementById('f-pax').value.trim(),
    heardFrom: document.getElementById('f-heard').value.trim(),
    enquiryDate: document.getElementById('f-enquiry-date').value || today(),
    eventDate:   document.getElementById('f-event-date').value,
    notes:     document.getElementById('f-notes').value.trim(),
  };
}

async function saveEntry() {
  const first = document.getElementById('f-first').value.trim();
  if (!first) { document.getElementById('f-first').style.borderColor='#e74c3c'; document.getElementById('f-first').focus(); return; }
  document.getElementById('f-first').style.borderColor='';
  const row = { id: editingId || uid(), ...getRowFromForm(), firstName: first };
  if (editingId) {
    const idx = DB.findIndex(r => r.id === editingId);
    if (idx !== -1) DB[idx] = row;
  } else {
    DB.unshift(row);
  }
  persist(); renderTable(); animateNewRows(); closeModal();
  const copied = await copyRowToClipboard(row);
  showToast(copied
    ? (editingId ? 'Updated & copied — paste into Excel at the next empty row' : 'Saved & copied — paste into Excel at the next empty row')
    : (editingId ? 'Updated — tap Copy on the row to copy for Excel' : 'Saved — tap Copy on the row to copy for Excel'));
}

async function copyFormForExcel() {
  const first = document.getElementById('f-first').value.trim();
  if (!first) { document.getElementById('f-first').style.borderColor='#e74c3c'; document.getElementById('f-first').focus(); return; }
  document.getElementById('f-first').style.borderColor='';
  const copied = await copyRowToClipboard(getRowFromForm());
  showToast(copied ? 'Copied for Excel — paste at the next empty row' : 'Select the text below and copy manually');
}

async function copyRowForExcel(id) {
  const row = DB.find(r => r.id === id);
  if (!row) return;
  const copied = await copyRowToClipboard(row);
  showToast(copied ? 'Copied for Excel — paste at the next empty row' : 'Select the text below and copy manually');
}

function deleteRow(id) {
  if (!confirm('Delete this enquiry?')) return;
  DB=DB.filter(r=>r.id!==id);
  persist(); renderTable(); showToast('Deleted');
}

// ── TABLE ──
function renderTable() {
  const search=(document.getElementById('search')?.value||'').toLowerCase();
  let rows=DB;
  if (search) rows=rows.filter(r=>((r.firstName||'')+' '+(r.lastName||'')+(r.email||'')+(r.phone||'')+(r.funcType||'')+(r.venue||'')).toLowerCase().includes(search));
  document.getElementById('record-count').textContent=DB.length+' record'+(DB.length!==1?'s':'');
  const dab = document.getElementById('delete-all-btn');
  if (dab) dab.style.display = DB.length > 0 ? 'inline-flex' : 'none';
  const cont=document.getElementById('table-body');
  if (!rows.length) {
    cont.innerHTML=`<div class="empty-state"><p>${DB.length===0?'No enquiries yet. Paste an email above or click "+ Add manually" to get started.':'No results match your search.'}</p></div>`;
    return;
  }
  let h=`<table><thead><tr><th>Name</th><th>Email</th><th>Phone</th><th>Venue</th><th>Type</th><th>PAX</th><th>Heard From</th><th>Enquiry Date</th><th>Event Date</th><th>Notes</th><th></th></tr></thead><tbody>`;
  rows.forEach(r=>{
    const name=[r.firstName,r.lastName].filter(Boolean).join(' ');
    const ns=r.notes?(r.notes.length>40?r.notes.slice(0,40)+'…':r.notes):'';
    h+=`<tr>
      <td><strong style="font-weight:500">${esc(name)}</strong></td>
      <td class="td-muted">${r.email?`<a href="mailto:${esc(r.email)}" style="color:inherit;text-decoration:none">${esc(r.email)}</a>`:'—'}</td>
      <td class="td-muted">${esc(r.phone)||'—'}</td>
      <td class="td-muted">${esc(r.venue)||'—'}</td>
      <td class="td-muted">${esc(r.funcType)||'—'}</td>
      <td class="td-muted" style="text-align:center">${esc(r.pax)||'—'}</td>
      <td class="td-muted">${esc(r.heardFrom)||'—'}</td>
      <td class="td-muted">${fmtDate(r.enquiryDate)}</td>
      <td class="td-muted">${fmtDate(r.eventDate)}</td>
      <td class="td-muted" style="max-width:120px;font-size:0.78rem">${esc(ns)||'—'}</td>
      <td style="white-space:nowrap;text-align:right;min-width:160px;width:160px">
        <button class="btn btn-copy" style="font-size:0.78rem;padding:4px 8px" onclick="copyRowForExcel('${r.id}')">Copy</button>
        <button class="btn" style="font-size:0.78rem;padding:4px 8px" onclick="openEdit('${r.id}')">Edit</button>
        <button class="btn btn-danger" onclick="deleteRow('${r.id}')">✕</button>
      </td>
    </tr>`;
  });
  h+=`</tbody></table>`;
  cont.innerHTML=h;
}

// ── SHAREPOINT ROW FORMAT (single source for clipboard + CSV) ──
const SHAREPOINT_COLUMNS = [
  { label: 'Date of function', get: r => fmtEventDate(r.eventDate) },
  { label: 'Name', get: r => [r.firstName, r.lastName].filter(Boolean).join(' ') },
  { label: 'Email address', get: r => r.email },
  { label: 'Venue', get: r => r.venue },
  { label: 'Date of enquiry', get: r => fmtEnquiryDate(r.enquiryDate) },
  { label: 'Where did you hear about us', get: r => r.heardFrom },
  { label: 'Type of function', get: r => r.funcType },
  { label: 'PAX', get: r => r.pax },
  { label: 'Site inspection', get: () => '' },
  { label: 'Reason for not confirming', get: () => '' },
  { label: 'Confirmed', get: () => '' },
  { label: 'Rate Quoted', get: () => '' },
  { label: 'Notes/Paid', get: r => r.notes },
];

function sanitizeCell(v) {
  return String(v ?? '')
    .replace(/[\r\n\u2028\u2029]+/g, ' ')
    .replace(/\t/g, ' ')
    .trim();
}

function rowToSharePointCells(row) {
  return SHAREPOINT_COLUMNS.map(col => sanitizeCell(col.get(row)));
}

function rowToTSV(row) {
  return rowToSharePointCells(row).join('\t');
}

function rowToHTMLTableRow(row) {
  const cells = rowToSharePointCells(row).map(c => `<td>${esc(c)}</td>`).join('');
  return `<table><tr>${cells}</tr></table>`;
}

async function copyRowToClipboard(row) {
  const text = rowToTSV(row);
  const html = rowToHTMLTableRow(row);
  if (navigator.clipboard?.write && typeof ClipboardItem !== 'undefined') {
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/plain': new Blob([text], { type: 'text/plain' }),
          'text/html': new Blob([html], { type: 'text/html' }),
        }),
      ]);
      return true;
    } catch (_) { /* fall through */ }
  }
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (_) { /* fall through */ }
  }
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0';
    document.body.appendChild(ta);
    ta.select();
    ta.setSelectionRange(0, text.length);
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    if (ok) return true;
  } catch (_) { /* fall through */ }
  openClipboardFallback(text);
  return false;
}

function openClipboardFallback(text) {
  const overlay = document.getElementById('clipboard-fallback-overlay');
  const field = document.getElementById('clipboard-fallback-text');
  if (!overlay || !field) return;
  field.value = text;
  overlay.classList.add('active');
  requestAnimationFrame(() => { field.focus(); field.select(); });
}

function closeClipboardFallback() {
  document.getElementById('clipboard-fallback-overlay')?.classList.remove('active');
}

function clipboardFallbackOverlayClick(e) {
  if (e.target === document.getElementById('clipboard-fallback-overlay')) closeClipboardFallback();
}

function selectClipboardFallback() {
  const field = document.getElementById('clipboard-fallback-text');
  field?.focus();
  field?.select();
}

// ── EXPORT CSV (matches SharePoint column order) ──
function exportCSV() {
  if (!DB.length) { showToast('Nothing to export yet'); return; }
  const hdr = SHAREPOINT_COLUMNS.map(c => c.label);
  const lines = [hdr.join(',')];
  DB.forEach(r => {
    lines.push(rowToSharePointCells(r).map(csv).join(','));
  });
  const blob = new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = 'le_pommier_enquiries_' + today() + '.csv'; a.click();
  URL.revokeObjectURL(url); showToast('CSV downloaded ✓');
}
function csv(v) { return v ? '"' + String(v).replace(/"/g, '""') + '"' : ''; }

document.addEventListener('keydown',e=>{ if(e.key==='Escape'){ closeModal(); closeHowTo(); closeClipboardFallback(); } if((e.ctrlKey||e.metaKey)&&e.key==='Enter') saveEntry(); });

// Header scroll shadow
window.addEventListener('scroll', () => {
  document.querySelector('header').classList.toggle('scrolled', window.scrollY > 10);
}, { passive: true });

// Animate new rows in on save
const _origSave = saveEntry;
function animateNewRows() {
  requestAnimationFrame(() => {
    document.querySelectorAll('tbody tr:not(.row-new)').forEach((tr, i) => {
      if (i === 0) {
        tr.classList.add('row-new');
        tr.style.animationDelay = '0ms';
        setTimeout(() => tr.classList.remove('row-new'), 600);
      }
    });
  });
}

// ── HOW TO ──
const HT_TOTAL = 4;
let htCurrent = 0;
function openHowTo() { htCurrent=0; htRender(); document.getElementById('ht-overlay').classList.add('active'); }
function closeHowTo() { document.getElementById('ht-overlay').classList.remove('active'); }
function htOverlayClick(e) { if(e.target===document.getElementById('ht-overlay')) closeHowTo(); }
function htGo(i) { htCurrent=i; htRender(); }
function htStep(dir) { htCurrent=Math.max(0,Math.min(HT_TOTAL-1,htCurrent+dir)); htRender(); }
function htRender() {
  document.querySelectorAll('.ht-panel').forEach((p,i)=>p.classList.toggle('active',i===htCurrent));
  document.querySelectorAll('.ht-tab').forEach((t,i)=>t.classList.toggle('active',i===htCurrent));
  const prog=document.getElementById('ht-progress');
  prog.innerHTML=Array.from({length:HT_TOTAL},(_,i)=>`<div class="ht-dot${i===htCurrent?' active':''}"></div>`).join('');
  document.getElementById('ht-prev').style.visibility=htCurrent===0?'hidden':'visible';
  document.getElementById('ht-next').textContent=htCurrent===HT_TOTAL-1?'Done':'Next →';
  document.getElementById('ht-next').onclick=htCurrent===HT_TOTAL-1?closeHowTo:()=>htStep(1);
}

renderTable();

// ── DELETE ALL ──
function openDeleteAll() {
  const count = DB.length;
  if (!count) return;
  document.getElementById('delete-all-count').textContent = count + ' enquiry record' + (count !== 1 ? 's' : '');
  document.getElementById('delete-all-overlay').classList.add('active');
}
function closeDeleteAll() { document.getElementById('delete-all-overlay').classList.remove('active'); }
function closeDeleteAllOverlay(e) { if (e.target === document.getElementById('delete-all-overlay')) closeDeleteAll(); }
function confirmDeleteAll() {
  DB = [];
  persist();
  renderTable();
  closeDeleteAll();
  showToast('All records deleted');
}
