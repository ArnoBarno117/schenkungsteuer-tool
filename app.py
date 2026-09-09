"""Schenkungsteuer: bewusst vereinfachtes Modell gemäß Excel-Vorgabe.

Start: streamlit run app.py. Rechenkern verwendet Decimal, keine KI-Aufrufe.
"""
from datetime import date
from decimal import Decimal as D, ROUND_FLOOR
from html import escape
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path
import json
import re
from urllib.parse import urlencode, urljoin, urlparse
from urllib.request import Request, urlopen, HTTPRedirectHandler, build_opener

VERSION = "1.1 / Rechtsstand geprüft am 09.09.2026"
BMF_ROOT = "https://www.bundesfinanzministerium.de/"
BMF_2026 = BMF_ROOT + "Content/DE/Downloads/BMF_Schreiben/Steuerarten/Erbschaft_Schenkungsteuerrecht/2025-10-21-bewert-lebensl-nutzung-leistung-1-1-26.pdf?__blob=publicationFile&v=5"
ALLOWANCES = {"Ehepartner": 500000, "Kind": 400000, "Enkel": 200000}
BRACKETS = [(D(str(limit)), D(str(rate))) for limit, rate in
            [(75000, '.07'), (300000, '.11'), (600000, '.15'),
             (6000000, '.19'), (13000000, '.23'), (26000000, '.27'), ('Infinity', '.30')]]
NOTES = [
    "Vereinfachte Berechnung für genau einen Erwerber der Steuerklasse I bei unbeschränkter Steuerpflicht.",
    "Nur der eingegebene offene Freibetrag wird abgezogen. Keine Zusammenrechnung von Vorerwerben, keine Progressionsberechnung und keine Steueranrechnung nach § 14 ErbStG. Das Ergebnis kann deshalb von einer Steuerfestsetzung abweichen.",
    "Annahme: Enkel sind nicht Kinder bereits verstorbener Kinder. Vorgabefreibeträge: Ehepartner 500.000 EUR, Kind 400.000 EUR, Enkel 200.000 EUR (§ 16 Abs. 1 ErbStG).",
    "85 % Regelverschonung wird vorausgesetzt. Die eingegebene nicht begünstigte Quote wird übernommen; insbesondere keine Prüfung des Verwaltungsvermögenstests, der Beteiligungsvoraussetzungen, Lohnsummen, Behaltensfristen oder früherer begünstigter Erwerbe. Die 26-Mio.-EUR-Grenze muss einschließlich einschlägiger Vorerwerbe eingehalten sein (§§ 13a, 13b ErbStG). Keine Berechnung nach §§ 13c, 28a ErbStG.",
    "Abzugsbetrag nach § 13a Abs. 2 ErbStG nur bei Auswahl 'Ja'; seine Verfügbarkeit ist vom Nutzer zu prüfen.",
    "Nießbrauch zu 100 % oder kein Nießbrauch. Jahreswert = Nominalbetrag × Ausschüttungsquote, nicht Kurswert × Ausschüttungsquote. Anteiliger Abzug entsprechend der Excel-Vorgabe (§ 10 Abs. 6 ErbStG). Keine Jahreswertbegrenzung nach § 16 BewG gemäß Modellvorgabe.",
    "Keine Berechnung für Steuerklassen II und III, insbesondere keine Ermittlung des Entlastungsbetrages nach § 19a ErbStG für diese Erwerber.",
    "Steuerpflichtiger Erwerb mindestens null, anschließend Abrundung auf volle 100 EUR (§ 10 Abs. 1 Satz 6 ErbStG). Tarif und Härteausgleich nach § 19 Abs. 1 und 3 ErbStG. Geldbeträge und Ergebnisquoten werden ohne Nachkommastellen angezeigt, intern aber ungerundet weitergerechnet. Kurs, Ausschüttung und BMF-Faktor behalten erforderliche Dezimalstellen. Gerundete Anzeigen können beim Nachrechnen abweichen.",
    "BMF-Tabellen werden jahresbezogen geladen. Das aktualisiert nicht automatisch den gesetzlichen Rechenkern. Für andere Rechtsstände ist eine fachliche Prüfung erforderlich.",
]


def num(value, places=0):
    return f"{D(str(value)):,.{places}f}".replace(',', '_').replace('.', ',').replace('_', '.')


def eur(value):
    return num(value) + " EUR"


def pct(value, places=0):
    return num(D(str(value)) * 100, places) + " %"


def exact_num(value):
    """Deutsche Schreibweise ohne angehängte Dezimalnullen, ohne Wertverlust."""
    value = D(str(value))
    places = max(0, -value.as_tuple().exponent)
    rendered = num(value, places)
    return rendered.rstrip('0').rstrip(',') if places else rendered


def parse_input(text, whole=False):
    text = str(text).strip().replace(' ', '').replace('\u00a0', '')
    if not re.fullmatch(r'(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?', text):
        raise ValueError('Bitte eine deutsche Zahl eingeben, z. B. 1.000.000 (Dezimaltrennzeichen: Komma).')
    value = D(text.replace('.', '').replace(',', '.'))
    if whole and value != value.to_integral_value():
        raise ValueError('Bitte einen Wert ohne Nachkommastellen eingeben.')
    return value


def german_input(label, default, key, maximum=None, whole=False, help=None):
    import streamlit as st
    def normalize():
        try:
            value = parse_input(st.session_state[key], whole)
            st.session_state[key] = exact_num(value)
        except ValueError:
            pass  # Ungültige Eingabe bleibt zur Korrektur sichtbar.
    raw = st.text_input(label, value=exact_num(default), key=key,
                        on_change=normalize, help=help)
    try:
        value = parse_input(raw, whole)
        if maximum is not None and value > maximum:
            raise ValueError(f'Der Wert darf höchstens {exact_num(maximum)} betragen.')
        return value
    except ValueError as error:
        st.error(str(error))
        st.stop()


def calculation_html(rows):
    """Drei explizite Spalten, ohne DataFrame-Index und ohne Ergebnisumbruch."""
    body = ''.join('<tr>'+''.join('<td>'+escape(str(row[k]))+'</td>'
                   for k in ('Rechenschritt','Berechnung','Ergebnis'))+'</tr>' for row in rows)
    return '''<style>
    .tax-scroll {overflow-x:auto; width:100%;}
    .tax-table {border-collapse:collapse; width:100%; min-width:650px; font-size:1rem;}
    .tax-table th,.tax-table td {padding:12px 14px; text-align:left; vertical-align:top;
      border-bottom:1px solid #8895a540; font-variant-numeric:tabular-nums;}
    .tax-table th {background:#7893b022;}
    .tax-table td:first-child {width:32%;}
    .tax-table td:nth-child(2) {overflow-wrap:break-word;}
    .tax-table th:last-child,.tax-table td:last-child {width:24%; min-width:175px;
      white-space:nowrap; text-align:right; font-weight:600;}
    </style><div class="tax-scroll"><table class="tax-table" aria-label="Rechenweg">
    <thead><tr><th scope="col">Rechenschritt</th><th scope="col">Berechnung</th>
    <th scope="col">Ergebnis</th></tr></thead><tbody>'''+body+'</tbody></table></div>'


def age_on(birth, day):
    if birth > day:
        raise ValueError("Das Geburtsdatum liegt nach dem Schenkungsdatum.")
    return day.year - birth.year - ((day.month, day.day) < (birth.month, birth.day))


def tax_class_one(base):
    base = D(str(base))
    if base < 0:
        raise ValueError("Negative Steuerbemessungsgrundlage.")
    previous_limit, previous_rate = D(0), D(0)
    for limit, rate in BRACKETS:
        if base <= limit:
            ordinary = base * rate
            cap = previous_limit * previous_rate + (base - previous_limit) / 2
            final = min(ordinary, cap) if previous_limit else ordinary
            return dict(rate=rate if base else D(0), ordinary=ordinary, cap=cap,
                        previous_limit=previous_limit, previous_rate=previous_rate,
                        relief=ordinary-final, tax=final)
        previous_limit, previous_rate = limit, rate


def calculate(nominal, course, nonfav_percent, deduction, usufruct, distribution_percent, factor, allowance):
    nominal, course, nonfav_percent, distribution_percent, factor, allowance = map(
        lambda x: D(str(x)), (nominal, course, nonfav_percent, distribution_percent, factor, allowance))
    if not all(x.is_finite() for x in (nominal, course, nonfav_percent, distribution_percent, factor, allowance)):
        raise ValueError("Bitte endliche Zahlen eingeben.")
    if nominal <= 0 or course <= 0:
        raise ValueError("Nominalbetrag und Kurs müssen größer als null sein.")
    if not D(0) <= nonfav_percent <= D(100) or not D(0) <= distribution_percent <= D(100):
        raise ValueError("Quoten müssen zwischen 0 und 100 % liegen.")
    if allowance < 0 or (usufruct and not D(0) < factor < D(19)):
        raise ValueError("Freibetrag oder Nießbrauchsfaktor ist ungültig.")
    rows = []
    def row(label, formula, value, percentage=False):
        rows.append({"Rechenschritt": label, "Berechnung": formula,
                     "Ergebnis": pct(value) if percentage else eur(value)})
        return value
    total = row("1. Steuerwert der Schenkung", f"{eur(nominal)} × {exact_num(course*100)} %", nominal * course)
    fav = row("2. Begünstigtes Vermögen", f"{eur(total)} × {pct(1-nonfav_percent/100)}", total*(1-nonfav_percent/100))
    if fav > 26000000:
        raise ValueError("Begünstigtes Vermögen über 26 Mio. EUR: außerhalb dieses vereinfachten Modells.")
    nonfav = row("3. Nicht begünstigtes Vermögen", f"{eur(total)} × {pct(nonfav_percent/100)}", total*nonfav_percent/100)
    exempt = row("4. Verschonungsabschlag (85 %)", f"{eur(fav)} × 85 %", fav*D('.85'))
    rest = row("5. Verbleibendes begünstigtes Vermögen", f"{eur(fav)} − {eur(exempt)}", fav-exempt)
    available = max(D(0), D(150000)-max(D(0), rest-D(150000))/2)
    applied = min(rest, available) if deduction else D(0)
    row("6. Abzugsbetrag (§ 13a Abs. 2)",
        f"min({eur(rest)}; max(0; 150.000 EUR − max(0; {eur(rest)} − 150.000 EUR) / 2))" if deduction else "Nicht verfügbar: 0 EUR", applied)
    taxed_fav = row("7. Zusätzlich zu versteuern", f"{eur(rest)} − {eur(applied)}", rest-applied)
    before = row("8. Erwerb vor Nießbrauch", f"{eur(nonfav)} + {eur(taxed_fav)}", nonfav+taxed_fav)
    ratio = row("9. Steuerpflichtiger Anteil am Erwerb", f"{eur(before)} / {eur(total)}", before/total, True)
    annual = row("10. Jahreswert des Nießbrauchs", f"{eur(nominal)} × {exact_num(distribution_percent)} %" if usufruct else "Kein Nießbrauch", nominal*distribution_percent/100 if usufruct else D(0))
    capital = row("11. Kapitalwert des Nießbrauchs", f"{eur(annual)} × {num(factor, 3)}" if usufruct else "Kein Nießbrauch", annual*factor if usufruct else D(0))
    allowed_capital = row("12. Abzugsfähiger Kapitalwert", f"{eur(capital)} × {pct(ratio)}", capital*ratio)
    net = row("13. Erwerb nach Nießbrauch", f"{eur(before)} − {eur(allowed_capital)}", before-allowed_capital)
    row("14. Offener persönlicher Freibetrag", "Nutzereingabe", allowance)
    raw = row("15. Erwerb nach Freibetrag", f"{eur(net)} − {eur(allowance)}", net-allowance)
    positive = row("16. Begrenzung auf mindestens null", f"max(0; {eur(raw)})", max(D(0), raw))
    base = row("17. Abrundung auf volle 100 EUR", f"Abrunden({eur(positive)} / 100) × 100", (positive/100).to_integral_value(rounding=ROUND_FLOOR)*100)
    tax = tax_class_one(base)
    row("18. Tarifsteuer (§ 19 Abs. 1)", f"{eur(base)} × {pct(tax['rate'], 0)}", tax['ordinary'])
    if tax['previous_limit']:
        row("19. Steuer an vorheriger Wertgrenze", f"{eur(tax['previous_limit'])} × {pct(tax['previous_rate'], 0)}", tax['previous_limit']*tax['previous_rate'])
        row("20. Höchstbetrag nach Härteausgleich", f"{eur(tax['previous_limit']*tax['previous_rate'])} + ({eur(base)} − {eur(tax['previous_limit'])}) / 2", tax['cap'])
    row("21. Ermäßigung durch Härteausgleich", "Tarifsteuer − min(Tarifsteuer; Höchstbetrag)" if tax['previous_limit'] else "Erste Tarifstufe: keine Ermäßigung", tax['relief'])
    row("22. Schenkungsteuer", f"{eur(tax['ordinary'])} − {eur(tax['relief'])}", tax['tax'])
    burden = row("23. Steuerbelastung auf geschenkten Anteil", f"{eur(tax['tax'])} / {eur(total)}", tax['tax']/total, True)
    return dict(rows=rows, total=total, net=net, base=base, burden=burden, **tax)


class Links(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []
    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            href = dict(attrs).get('href')
            if href:
                self.links.append(href)


def official_url(url):
    p = urlparse(url)
    if p.scheme != 'https' or p.hostname not in ('www.bundesfinanzministerium.de', 'bundesfinanzministerium.de') or p.port not in (None, 443) or p.username:
        raise ValueError("Bitte ausschließlich einen HTTPS-Link auf bundesfinanzministerium.de verwenden.")
    return url


class OfficialRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        official_url(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def fetch(url):
    official_url(url)
    request = Request(url, headers={'User-Agent': 'Schenkungsteuer-Rechner/1.0 (BMF-Tabellenabruf)'})
    with build_opener(OfficialRedirect()).open(request, timeout=12) as response:
        official_url(response.url)
        content = response.read(8_000_001)
        if len(content) > 8_000_000:
            raise ValueError("BMF-Datei überschreitet die Größenbegrenzung.")
        return content


def links_from_html(content):
    parser = Links()
    parser.feed(content.decode('utf-8', errors='replace'))
    return list(dict.fromkeys(urljoin(BMF_ROOT, x) for x in parser.links))


def parse_bmf_pdf(data, year, url):
    from pypdf import PdfReader
    if not data.startswith(b'%PDF'):
        raise ValueError("Der BMF-Abruf hat keine PDF-Datei geliefert.")
    reader = PdfReader(BytesIO(data))
    if len(reader.pages) > 12:
        raise ValueError("Unerwartete BMF-Datei.")
    text = '\n'.join(page.extract_text(extraction_mode='layout') or '' for page in reader.pages)
    if not re.search(r'für\s+(?:Bewertungsstichtage|Stichtage)\s+ab\s+1\.?\s*Januar\s+'+str(year), text, re.I):
        raise ValueError(f"PDF enthält keine eindeutige Gültigkeit ab 1. Januar {year}.")
    if 'Männer' not in text or 'Frauen' not in text or 'Vervielfältiger' not in text:
        raise ValueError("Tabellenüberschriften fehlen.")
    matches = re.findall(r'^\s*(\d{1,3})\s+(\d+,\d{2})\s+(\d+,\d{3})\s+(\d+,\d{2})\s+(\d+,\d{3})\s*$', text, re.M)
    factors = {}
    for age, _, male, _, female in matches:
        if age in factors:
            raise ValueError("Doppelte Alterszeile in BMF-Tabelle.")
        factors[age] = [male.replace(',', '.'), female.replace(',', '.')]
    if set(factors) != {str(x) for x in range(101)}:
        raise ValueError("BMF-Tabelle nicht vollständig lesbar (Alter 0 bis 100 erforderlich).")
    for sex in (0, 1):
        values = [D(factors[str(x)][sex]) for x in range(101)]
        if not all(D(0) < x < D(19) for x in values) or any(a < b for a, b in zip(values, values[1:])):
            raise ValueError("BMF-Faktoren bestehen die Plausibilitätsprüfung nicht.")
    issued = re.search(r'\b(\d{1,2}\.\s*(?:Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)\s+\d{4})\b', text.split('Betreff:')[0])
    if not issued:
        raise ValueError("Datum des BMF-Schreibens konnte nicht gelesen werden.")
    reference = re.search(r'GZ:\s*([^\n]+)', text)
    return dict(year=year, issued=' '.join(issued.group(1).split()),
                reference=reference.group(1).strip() if reference else '', url=url, factors=factors)


def load_from_url(url, year):
    data = fetch(url)
    if data.startswith(b'%PDF'):
        return parse_bmf_pdf(data, year, url)
    candidates = [x for x in links_from_html(data) if '.pdf' in x and ('nutzung' in x.lower() or 'leistung' in x.lower())]
    if not candidates:
        raise ValueError("Auf der BMF-Seite wurde kein passendes PDF gefunden.")
    return parse_bmf_pdf(fetch(candidates[0]), year, candidates[0])


def discover_bmf(year):
    query = urlencode({'resourceId':'240616', 'input_':'510386', 'pageLocale':'de',
                       'templateQueryString':f'lebenslänglichen Nutzung Leistung {year}', 'lang':'de'})
    search_url = BMF_ROOT+'SiteGlobals/Forms/Suche/Servicesuche_Formular.html?'+query
    overview = BMF_ROOT+'Web/DE/Themen/Steuern/Steuerarten/Erbschaft_und_Schenkungsteuer/erbschaft_schenkungsteuer.html'
    seen = set()
    # Die thematische Übersicht liefert aktuelle Veröffentlichungen wesentlich
    # gezielter als die allgemeine BMF-Suche. Letztere bleibt ein Zusatzweg.
    for index_url in (overview, search_url):
        try:
            links = links_from_html(fetch(index_url))
        except (ValueError, OSError):
            continue
        candidates = [x for x in links if '/BMF_Schreiben/' in x and ('nutzung' in x.lower() or 'leistung' in x.lower()) and '.pdf' in x]
        candidates = sorted(set(candidates), reverse=True)[:4]
        for url in candidates:
            if url in seen:
                continue
            seen.add(url)
            try:
                return load_from_url(url, year)
            except (ValueError, OSError):
                continue
    raise ValueError(f"Keine vollständig lesbare BMF-Tabelle für {year} gefunden. Bitte den offiziellen BMF-Link unten eintragen oder den Stichtag prüfen.")


def get_bmf(year, override=''):
    if override:
        return load_from_url(override, year), "Online vom BMF geladen; vollständig und auf Plausibilität geprüft."
    if year == 2026:
        try:
            return load_from_url(BMF_2026, year), "Online vom BMF geladen; vollständig und auf Plausibilität geprüft."
        except Exception:
            return json.loads(BMF_2026_SNAPSHOT), "Online-Abruf nicht möglich: mitgelieferte, anhand des BMF-Schreibens geprüfte Tabelle 2026 verwendet."
    return discover_bmf(year), "Über die BMF-Veröffentlichungsübersicht / Suche geladen; vollständig und auf Plausibilität geprüft."


def make_pdf(inputs, result, source, status):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, TableStyle, PageBreak
    output = BytesIO()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Cell', fontName='Helvetica', fontSize=8.2, leading=11, spaceAfter=0))
    styles['BodyText'].fontSize = 9
    styles['BodyText'].leading = 13
    def p(text, style='Cell'):
        return Paragraph(escape(str(text)), styles[style])
    story = [p('Schenkungsteuer · Berechnungsnachweis', 'Title'), p(VERSION, 'BodyText'), Spacer(1, 12),
             p('Schenkungsteuer: '+eur(result['tax']), 'Heading2'),
             p('Steuerbelastung auf den geschenkten Anteil: '+pct(result['burden']), 'BodyText'),
             p('Eingaben', 'Heading2')]
    def table(data, widths):
        t = LongTable([[p(c) for c in r] for r in data], colWidths=widths, repeatRows=1, hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e5edf6')),
            ('VALIGN',(0,0),(-1,-1),'TOP'), ('LINEBELOW',(0,0),(-1,0),.6,colors.HexColor('#234468')),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f5f7fa')]),
            ('TOPPADDING',(0,0),(-1,-1),7), ('BOTTOMPADDING',(0,0),(-1,-1),7)]))
        return t
    story += [table([['Parameter','Wert']]+list(inputs.items()), [190, 333]),
              p('Rechenweg', 'Heading2'),
              table([['Rechenschritt','Berechnung','Ergebnis']]+[list(r.values()) for r in result['rows']], [150,233,140]),
              p('BMF-Quelle', 'Heading2'), p(source, 'BodyText'), p(status, 'BodyText'),
              PageBreak(), p('Annahmen und Berechnungsumfang', 'Heading2')]
    story += [p(note, 'BodyText') for note in NOTES]
    def footer(canvas, doc):
        canvas.setFont('Helvetica',8)
        canvas.setFillColor(colors.HexColor('#536477'))
        canvas.drawString(36, 24, 'Schenkungsteuer | '+date.today().strftime('%d.%m.%Y'))
        canvas.drawRightString(A4[0]-36, 24, f'Seite {doc.page}')
    SimpleDocTemplate(output, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=34,
                      bottomMargin=40, title='Schenkungsteuer - Berechnungsnachweis', author='Schenkungsteuer-Rechner').build(story, onFirstPage=footer, onLaterPages=footer)
    return output.getvalue()


def main():
    import streamlit as st
    st.set_page_config(page_title='Schenkungsteuer-Rechner', page_icon='§', layout='wide')
    st.title('Schenkungsteuer-Rechner')
    st.caption('Anteilsübertragung · 85 % Regelverschonung · optionaler Vorbehaltsnießbrauch · ein Erwerber')
    st.info('Vereinfachtes Modell: nur offener Freibetrag; keine Zusammenrechnung oder Steueranrechnung aus Vorerwerben (§ 14 ErbStG).')
    left, right = st.columns([1, 2.2], gap='large')
    with left:
        st.subheader('Ihre Eingaben')
        day = st.date_input('Schenkungsdatum', date.today(), min_value=date(2026,1,1), max_value=date(2100,12,31), format='DD.MM.YYYY')
        relation = st.selectbox('Verwandtschaftsverhältnis', list(ALLOWANCES), index=1)
        if relation == 'Enkel':
            st.caption('Annahme: Enkel sind nicht Kinder bereits verstorbener Kinder')
        allowance = german_input('Offener Freibetrag (EUR)', ALLOWANCES[relation], 'allowance_de_'+relation,
                                  maximum=ALLOWANCES[relation], whole=True)
        nominal = german_input('Nominalbetrag der geschenkten Anteile (EUR)', 100000, 'nominal_de', whole=True)
        course_percent = german_input('Kurswert (%)', D('478.5'), 'course_percent_de',
                                      help='Beispiel: 560 bedeutet 560 % des Nominalbetrags. Dezimalwerte bitte mit Komma eingeben.')
        course = course_percent / 100
        quote = german_input('Nicht begünstigtes Vermögen (%)', 10, 'quote_de', maximum=100, whole=True)
        deduction = st.radio('Abzugsbetrag nach § 13a Abs. 2 verfügbar?', ['Ja','Nein'], horizontal=True) == 'Ja'
        usufruct = st.radio('Nießbrauch', ['Ja (100 %)','Nein'], horizontal=True) == 'Ja (100 %)'
        distribution, factor, birth, sex, age = 0.0, D(0), None, None, None
        source, status = 'Kein Nießbrauch: kein BMF-Faktor erforderlich.', ''
        bmf_error = None
        if usufruct:
            distribution = german_input('Durchschnittliche Ausschüttung (% des Nominalbetrags)', D('5.33'), 'distribution_de', maximum=100)
            sex = st.selectbox('Geschlecht des Nießbrauchers (BMF-Tabelle)', ['Mann','Frau'])
            birth = st.date_input('Geburtsdatum des Nießbrauchers', date(1972,1,1), min_value=date(1900,1,1), max_value=date(2100,12,31), format='DD.MM.YYYY')
            with st.expander('BMF-Abruf / abweichender offizieller Link'):
                st.caption('Automatischer Abruf für das Jahr der Schenkung. Falls die BMF-Suche geändert wurde, hier die offizielle Veröffentlichungsseite oder PDF-Adresse einsetzen.')
                override = st.text_input('BMF-Link (optional)', key='bmf_url')
                refresh = st.button('BMF-Tabelle erneut abrufen')
            cached_bmf = st.cache_data(ttl=86400, show_spinner=False)(get_bmf)
            if refresh:
                cached_bmf.clear()
            try:
                age = age_on(birth, day)
                with st.spinner('BMF-Tabelle wird geprüft …'):
                    bmf, status = cached_bmf(day.year, override.strip())
                factor = D(bmf['factors'][str(min(age,100))][0 if sex=='Mann' else 1])
                source = f"Gemäß BMF-Schreiben vom {bmf['issued']} ({bmf['reference']}), Bewertungsstichtage ab 1. Januar {bmf['year']}. Quelle: {bmf['url']}"
                st.caption(f'Vollendetes Lebensalter: {age} Jahre · Faktor: {num(factor,3)}'+(' (Tabellenzeile 100 und darüber)' if age>=100 else ''))
            except Exception as error:
                bmf_error = str(error)
        with st.expander('Annahmen und Berechnungsumfang'):
            for note in NOTES:
                st.write(note)
    with right:
        st.subheader('Ergebnis und Rechenweg')
        if bmf_error:
            st.error('Kein Ergebnis: '+bmf_error)
            st.stop()
        try:
            result = calculate(nominal, course, quote, deduction, usufruct, distribution, factor, allowance)
        except ValueError as error:
            st.error(str(error))
            st.stop()
        a,b = st.columns(2)
        a.metric('Schenkungsteuer', eur(result['tax']))
        b.metric('Belastung des geschenkten Anteils', pct(result['burden']))
        st.caption(f"Steuerpflichtiger Erwerb: {eur(result['base'])} · Tarif: {pct(result['rate'],0)} · Härteausgleich: {eur(result['relief'])}")
        inputs = {'Schenkungsdatum': day.strftime('%d.%m.%Y'), 'Erwerber': relation,
                  'Offener Freibetrag': eur(allowance), 'Nominalbetrag': eur(nominal),
                  'Kurswert': exact_num(course_percent)+' %', 'Nicht begünstigte Quote': num(quote)+' %',
                  'Regelverschonung': '85 %', 'Abzugsbetrag verfügbar': 'Ja' if deduction else 'Nein',
                  'Nießbrauch': 'Ja (100 %)' if usufruct else 'Nein'}
        if usufruct:
            inputs.update({'Ausschüttung auf Nominalbetrag':exact_num(distribution)+' %', 'Geschlecht': sex,
                           'Geburtsdatum':birth.strftime('%d.%m.%Y'), 'Vollendetes Lebensalter':str(age),
                           'BMF-Vervielfältiger':num(factor,3)})
        st.download_button('Berechnung als PDF speichern', make_pdf(inputs,result,source,status),
                           file_name=f'Schenkungsteuer_{day.isoformat()}.pdf', mime='application/pdf')
        st.html(calculation_html(result['rows']))
        st.caption('Geldbeträge und Ergebnisquoten werden ohne Nachkommastellen angezeigt. Intern wird ungerundet gerechnet. Kurs, Ausschüttung und BMF-Faktor bleiben mit ihren erforderlichen Dezimalstellen sichtbar.')
        st.write(source)
        if status:
            st.caption(status)
    st.divider()
    st.caption(VERSION+' · Keine dauerhafte Speicherung der Eingaben durch diese Anwendung.')


# Vollständiger, aus der gelieferten BMF-PDF extrahierter Offline-Datenstand.
BMF_2026_SNAPSHOT = '{"year": 2026, "issued": "21. Oktober 2025", "reference": "IV D 4 - S 3104/00002/013/003", "url": "https://www.bundesfinanzministerium.de/Content/DE/Downloads/BMF_Schreiben/Steuerarten/Erbschaft_Schenkungsteuerrecht/2025-10-21-bewert-lebensl-nutzung-leistung-1-1-26.pdf?__blob=publicationFile&v=5", "factors": {"0": ["18.402", "18.465"], "1": ["18.391", "18.456"], "2": ["18.375", "18.443"], "3": ["18.359", "18.430"], "4": ["18.341", "18.417"], "5": ["18.322", "18.402"], "6": ["18.303", "18.387"], "7": ["18.282", "18.371"], "8": ["18.260", "18.354"], "9": ["18.237", "18.336"], "10": ["18.213", "18.317"], "11": ["18.187", "18.297"], "12": ["18.160", "18.276"], "13": ["18.132", "18.254"], "14": ["18.101", "18.230"], "15": ["18.070", "18.206"], "16": ["18.037", "18.180"], "17": ["18.002", "18.152"], "18": ["17.965", "18.124"], "19": ["17.926", "18.093"], "20": ["17.886", "18.062"], "21": ["17.843", "18.028"], "22": ["17.799", "17.992"], "23": ["17.751", "17.955"], "24": ["17.701", "17.915"], "25": ["17.649", "17.873"], "26": ["17.593", "17.829"], "27": ["17.535", "17.783"], "28": ["17.473", "17.735"], "29": ["17.409", "17.683"], "30": ["17.340", "17.629"], "31": ["17.269", "17.572"], "32": ["17.193", "17.511"], "33": ["17.113", "17.448"], "34": ["17.030", "17.382"], "35": ["16.942", "17.312"], "36": ["16.850", "17.238"], "37": ["16.753", "17.160"], "38": ["16.652", "17.078"], "39": ["16.545", "16.993"], "40": ["16.434", "16.903"], "41": ["16.317", "16.808"], "42": ["16.193", "16.709"], "43": ["16.065", "16.604"], "44": ["15.930", "16.494"], "45": ["15.788", "16.379"], "46": ["15.640", "16.258"], "47": ["15.485", "16.130"], "48": ["15.323", "15.997"], "49": ["15.155", "15.856"], "50": ["14.977", "15.711"], "51": ["14.794", "15.557"], "52": ["14.605", "15.398"], "53": ["14.406", "15.230"], "54": ["14.199", "15.054"], "55": ["13.983", "14.873"], "56": ["13.762", "14.680"], "57": ["13.533", "14.481"], "58": ["13.293", "14.273"], "59": ["13.048", "14.058"], "60": ["12.798", "13.832"], "61": ["12.538", "13.601"], "62": ["12.272", "13.359"], "63": ["12.002", "13.108"], "64": ["11.725", "12.849"], "65": ["11.444", "12.583"], "66": ["11.155", "12.306"], "67": ["10.860", "12.020"], "68": ["10.561", "11.725"], "69": ["10.251", "11.421"], "70": ["9.938", "11.107"], "71": ["9.619", "10.780"], "72": ["9.293", "10.447"], "73": ["8.960", "10.105"], "74": ["8.627", "9.754"], "75": ["8.282", "9.393"], "76": ["7.936", "9.022"], "77": ["7.586", "8.648"], "78": ["7.230", "8.271"], "79": ["6.868", "7.885"], "80": ["6.509", "7.490"], "81": ["6.158", "7.100"], "82": ["5.798", "6.703"], "83": ["5.441", "6.305"], "84": ["5.089", "5.908"], "85": ["4.743", "5.512"], "86": ["4.411", "5.133"], "87": ["4.086", "4.765"], "88": ["3.778", "4.418"], "89": ["3.496", "4.086"], "90": ["3.234", "3.770"], "91": ["2.984", "3.480"], "92": ["2.764", "3.217"], "93": ["2.566", "2.975"], "94": ["2.393", "2.764"], "95": ["2.226", "2.558"], "96": ["2.076", "2.384"], "97": ["1.951", "2.226"], "98": ["1.843", "2.094"], "99": ["1.771", "1.987"], "100": ["1.680", "1.897"]}}'

if __name__ == '__main__':
    main()
