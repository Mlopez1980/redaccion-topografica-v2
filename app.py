from flask import Flask, render_template, request, send_file, jsonify
import re, os, json, traceback
from io import BytesIO
from datetime import datetime

APP_VERSION = "v3.3-versioncheck"

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "static", "logo_hc.png")
HEADER_TEXT = "Este programa fue creado por Honduras Constructores S de R L"

app = Flask(__name__)

UNIDADES = ["cero","uno","dos","tres","cuatro","cinco","seis","siete","ocho","nueve"]
ESPECIALES_10_19 = ["diez","once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve"]
VEINTI = ["veinte","veintiuno","veintidós","veintitrés","veinticuatro","veinticinco","veintiséis","veintisiete","veintiocho","veintinueve"]
DECENAS = [None,None,"veinte","treinta","cuarenta","cincuenta","sesenta","setenta","ochenta","noventa"]
CIENTOS = {100:"cien",200:"doscientos",300:"trescientos",400:"cuatrocientos",500:"quinientos",600:"seiscientos",700:"setecientos",800:"ochocientos",900:"novecientos"}

def numero_a_palabras(n:int)->str:
    if n<0 or n>999: raise ValueError("Solo se admite 0–999 para esta app")
    if n<10: return UNIDADES[n]
    if 10<=n<=19: return ESPECIALES_10_19[n-10]
    if 20<=n<=29: return VEINTI[n-20]
    if 30<=n<=99:
        d,u=divmod(n,10)
        return DECENAS[d] if u==0 else f"{DECENAS[d]} y {UNIDADES[u]}"
    c=(n//100)*100; r=n%100
    if n==100: return CIENTOS[100]
    pref=CIENTOS.get(c,"ciento")
    return pref if r==0 else f"{pref} {numero_a_palabras(r)}"

def forma_masculina(frase:str)->str:
    frase=re.sub(r"\bveintiuno\b","veintiún",frase)
    frase=re.sub(r" y uno\b"," y un",frase)
    frase=re.sub(r"\buno\b","un",frase)
    return frase

def plural_si_corresponde(v:int,sing:str)->str:
    return sing if v==1 else sing+"s"

ETIQUETA_RE=re.compile(r"^(?P<num>\d+)?(?P<letras>[A-Za-z]+)?$")

def etiqueta_a_texto(etq:str, convertir_numeros=True)->str:
    etq=(etq or "").strip()
    m=ETIQUETA_RE.match(etq)
    if not m: return etq
    num=m.group("num"); letras=(m.group("letras") or "").upper()
    partes=[]
    if num: partes.append(numero_a_palabras(int(num)) if convertir_numeros else num)
    if letras: partes.append(letras)
    return " ".join(partes) if partes else etq

CARD_WORD={"N":"Norte","S":"Sur","E":"Este","O":"Oeste","W":"Oeste"}

def parsear_rumbo_texto(raw:str):
    if not raw: return None
    t=raw.strip()
    if not t: return None
    # Normaliza palabras cardinales a letras
    t = re.sub(r'\bNORTE\b','N',t,flags=re.I)
    t = re.sub(r'\bSUR\b','S',t,flags=re.I)
    t = re.sub(r'\bESTE\b','E',t,flags=re.I)
    t = re.sub(r'\bOESTE\b','O',t,flags=re.I)
    norm=t.replace("°"," ").replace("º"," ").replace("’","'").replace("´","'")
    norm=re.sub(r"[;|/]+"," ",norm)
    norm = re.sub(r'\bW\b','O',norm,flags=re.I)

    # Formato con comas: N, 25, 35, 20, O
    partes=[p.strip() for p in norm.split(',') if p.strip()]
    if len(partes)>=5:
        c1=partes[0].upper()
        try:
            g=int(re.sub(r"\D","", partes[1] or "0"))
            m=int(re.sub(r"\D","", partes[2] or "0"))
            s=int(re.sub(r"\D","", partes[3] or "0"))
        except ValueError:
            return None
        c2=partes[4].upper()
        return (c1,g,m,s,c2)

    # Formato libre: N 25 35 20 O (o con símbolos mezclados)
    cards=re.findall(r"[NnSsEeOoWw]",norm)
    nums=re.findall(r"\d+",norm)
    if len(nums)>=3 and len(cards)>=2:
        c1=cards[0].upper(); c2=cards[-1].upper()
        if c1=='W': c1='O'
        if c2=='W': c2='O'
        g=int(nums[0]); m=int(nums[1]); s=int(nums[2])
        return (c1,g,m,s,c2)

    return None

def rumbo_texto(card1:str,g:int,m:int,s:int,card2:str)->str:
    g_w=forma_masculina(numero_a_palabras(g))
    m_w=forma_masculina(numero_a_palabras(m))
    s_w=forma_masculina(numero_a_palabras(s))
    return (f"{CARD_WORD.get(card1,card1)} {g_w} {plural_si_corresponde(g,'grado')}, "
            f"{m_w} {plural_si_corresponde(m,'minuto')}, "
            f"{s_w} {plural_si_corresponde(s,'segundo')} {CARD_WORD.get(card2,card2)}")

def rumbo_compacto_solicitado(card1:str,g:int,m:int,s:int,card2:str)->str:
    return f"{card1} ° {g} {m}´{s}´{card2}"

def rumbo_compacto_convencional(card1:str,g:int,m:int,s:int,card2:str)->str:
    return f"{card1} {g}° {m}' {s}'' {card2}"

@app.route('/_version')
def version():
    return jsonify({"version": APP_VERSION, "docx": DOCX_AVAILABLE})

@app.route('/', methods=['GET','POST'])
def index():
    errores=[]; resultado=None
    if request.method=='POST':
        convertir = request.form.get('convertir','on')=='on'
        est_ini_list = request.form.getlist('est_ini[]')
        est_fin_list = request.form.getlist('est_fin[]')
        rumbo_txt_list = request.form.getlist('rumbo_texto[]')
        distancia_list = request.form.getlist('distancia[]')
        colind_list = request.form.getlist('colindancia[]')

        n = max(len(est_ini_list), len(est_fin_list), len(rumbo_txt_list), len(distancia_list), len(colind_list))
        tramos = []
        prev_fin_raw = None

        for i in range(n):
            est_ini = (est_ini_list[i] if i < len(est_ini_list) else '').strip()
            est_fin = (est_fin_list[i] if i < len(est_fin_list) else '').strip()
            rumbo_raw = (rumbo_txt_list[i] if i < len(rumbo_txt_list) else '').strip()
            distancia_raw = (distancia_list[i] if i < len(distancia_list) else '').strip()
            colind = (colind_list[i] if i < len(colind_list) else '').strip()

            # Auto-encadenar: si no hay inicio y existe un fin previo, usarlo
            if not est_ini and prev_fin_raw:
                est_ini = prev_fin_raw

            # Saltar filas completamente vacías
            if not (est_ini or est_fin or rumbo_raw or distancia_raw or colind):
                continue

            if not est_ini or not est_fin:
                errores.append(f"Fila {i+1}: ingresa estación inicio y fin (auto-encadené inicio='{est_ini or '∅'}').")
                prev_fin_raw = est_fin or prev_fin_raw
                continue

            parsed = parsear_rumbo_texto(rumbo_raw)
            if not parsed:
                errores.append(f"Fila {i+1}: no pude interpretar el rumbo \"{rumbo_raw}\".")
                prev_fin_raw = est_fin
                continue
            c1,g,m,s,c2 = parsed

            try:
                distancia = float(distancia_raw) if distancia_raw else None
            except ValueError:
                errores.append(f"Fila {i+1}: distancia inválida.")
                distancia = None

            est_ini_txt = etiqueta_a_texto(est_ini, convertir_numeros=convertir)
            est_fin_txt = etiqueta_a_texto(est_fin, convertir_numeros=convertir)
            texto_rumbo = rumbo_texto(c1,g,m,s,c2)
            compacto_solicitado = rumbo_compacto_solicitado(c1,g,m,s,c2)
            compacto_convencional = rumbo_compacto_convencional(c1,g,m,s,c2)

            redaccion = f"De la estación {est_ini_txt} a la estación {est_fin_txt}, con rumbo {texto_rumbo}."
            if distancia is not None:
                redaccion += f" Distancia {distancia:.2f} m."
            if colind:
                redaccion += f" {colind}"

            tramos.append({
                "est_ini_txt": est_ini_txt,
                "est_fin_txt": est_fin_txt,
                "rumbo_texto": texto_rumbo,
                "rumbo_compacto_solicitado": compacto_solicitado,
                "rumbo_compacto_convencional": compacto_convencional,
                "distancia": distancia,
                "colindancia": colind,
                "redaccion": redaccion
            })

            # actualizar para encadenar siguiente inicio
            prev_fin_raw = est_fin

        if not tramos and not errores:
            errores.append("Agrega al menos un tramo.")

        if not errores:
            redaccion_total = "\n".join(f"{idx+1}) {t['redaccion']}" for idx, t in enumerate(tramos))
            resultado = {
                "tramos": tramos,
                "redaccion_total": redaccion_total
            }

    return render_template('formulario.html', errores=errores, resultado=resultado,
                           docx_ready=DOCX_AVAILABLE, app_version=APP_VERSION)

@app.route('/descargar', methods=['POST'])
def descargar():
    if not DOCX_AVAILABLE:
        return "La librería python-docx no está instalada en el servidor.", 500

    try:
        payload_json = request.form.get('payload_json', '')
        data = json.loads(payload_json)
        tramos = data.get("tramos", [])
    except Exception as e:
        return f"No pude leer los datos para el Word. Detalle: {e}", 400

    try:
        doc = Document()
        styles=doc.styles['Normal']; styles.font.name='Calibri'; styles.font.size=Pt(11)

        section=doc.sections[0]; header=section.header
        table=header.add_table(rows=1, cols=2); table.autofit=True
        try:
            if os.path.exists(LOGO_PATH):
                left_p=table.cell(0,0).paragraphs[0]
                left_p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                left_p.add_run().add_picture(LOGO_PATH, width=Inches(1.8))
        except Exception:
            pass
        right_p=table.cell(0,1).paragraphs[0]; right_p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
        rt=right_p.add_run(HEADER_TEXT); rt.bold=True; rt.font.size=Pt(10)

        doc.add_heading('Redacción topográfica', level=1)
        doc.add_paragraph(datetime.now().strftime("Generado el %Y-%m-%d %H:%M:%S"))

        doc.add_paragraph().add_run("Resumen de tramos:").bold = True
        for i, t in enumerate(tramos, start=1):
            doc.add_paragraph(f"{i}) {t['redaccion']}")

        doc.add_paragraph().add_run("Detalle:").bold = True
        tbl = doc.add_table(rows=1, cols=6)
        hdr = tbl.rows[0].cells
        hdr[0].text = "Est. Inicio"
        hdr[1].text = "Est. Fin"
        hdr[2].text = "Rumbo (texto)"
        hdr[3].text = "Rumbo compacto"
        hdr[4].text = "Distancia (m)"
        hdr[5].text = "Colindancia"

        for t in tramos:
            row = tbl.add_row().cells
            row[0].text = t["est_ini_txt"]
            row[1].text = t["est_fin_txt"]
            row[2].text = t["rumbo_texto"]
            row[3].text = t["rumbo_compacto_solicitado"]
            row[4].text = "" if t["distancia"] is None else f"{t['distancia']:.2f}"
            row[5].text = t["colindancia"] or ""

        bio = BytesIO(); doc.save(bio); bio.seek(0)
        return send_file(bio, as_attachment=True, download_name="Redaccion_Multitramos.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception:
        return "Error generando DOCX:\n" + traceback.format_exc(), 500

if __name__=='__main__':
    app.run(debug=True)
