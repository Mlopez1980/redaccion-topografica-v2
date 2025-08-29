from flask import Flask, render_template, request, send_file, jsonify
import re, os
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
@app.route('/_version')
def version():
    return jsonify({"version": APP_VERSION, "docx": DOCX_AVAILABLE})

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
    norm=t.replace("°"," ").replace("º"," ").replace("’","'").replace("´","'")
    norm=re.sub(r"[;|/]+"," ",norm)
    cards=re.findall(r"[NnSsEeOoWw]",norm)
    nums=re.findall(r"\d+",norm)
    if len(nums)<3 or not cards:
        partes=[p.strip() for p in t.split(',') if p.strip()]
        if len(partes)>=5:
            c1=partes[0].strip().upper()
            g=int(re.sub(r"\D","",partes[1]))
            m=int(re.sub(r"\D","",partes[2]))
            s=int(re.sub(r"\D","",partes[3]))
            c2=partes[4].strip().upper()
            c1='O' if c1=='W' else c1
            c2='O' if c2=='W' else c2
            return (c1,g,m,s,c2)
        return None
    c1=cards[0].upper(); c2=cards[-1].upper()
    c1='O' if c1=='W' else c1
    c2='O' if c2=='W' else c2
    g=int(nums[0]); m=int(nums[1]); s=int(nums[2])
    return (c1,g,m,s,c2)

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

from flask import jsonify

@app.route('/', methods=['GET','POST'])
def index():
    resultado=None; errores=[]
    if request.method=='POST':
        est_ini_raw=(request.form.get('est_ini') or '').strip()
        est_fin_raw=(request.form.get('est_fin') or '').strip()
        convertir = request.form.get('convertir','on')=='on'
        colindancia=(request.form.get('colindancia') or '').strip()
        rumbo_texto_entrada=(request.form.get('rumbo_texto') or '').strip()
        if rumbo_texto_entrada:
            parsed=parsear_rumbo_texto(rumbo_texto_entrada)
            if not parsed:
                errores.append("No pude interpretar el rumbo desde el texto. Revisa el formato.")
            else:
                c1,g,m,s,c2=parsed
        else:
            c1=(request.form.get('card1') or '').upper()
            c2=(request.form.get('card2') or '').upper()
            try:
                g=int(request.form.get('grados') or '0')
                m=int(request.form.get('minutos') or '0')
                s=int(request.form.get('segundos') or '0')
            except ValueError:
                errores.append("Grados/minutos/segundos deben ser números enteros.")
                g=m=s=0
        if not est_ini_raw or not est_fin_raw: errores.append("Ingresa ambas estaciones (inicio y fin).")
        if c1 not in {'N','S'} or c2 not in {'E','O'}: errores.append("Orientaciones inválidas. Usa N/S y E/O.")
        if not (0<=g<=359 and 0<=m<=59 and 0<=s<=59): errores.append("Rango inválido para grados (0–359), minutos/segundos (0–59).")
        if not errores:
            est_ini_txt=etiqueta_a_texto(est_ini_raw, convertir_numeros=convertir)
            est_fin_txt=etiqueta_a_texto(est_fin_raw, convertir_numeros=convertir)
            frase_estaciones=f"de la estación {est_ini_txt} a la estación {est_fin_txt}"
            texto_rumbo=rumbo_texto(c1,g,m,s,c2)
            compacto_solicitado=rumbo_compacto_solicitado(c1,g,m,s,c2)
            compacto_convencional=rumbo_compacto_convencional(c1,g,m,s,c2)
            redaccion=f"{frase_estaciones}, con rumbo {texto_rumbo}."
            if colindancia: redaccion+=f" {colindancia.strip()}"
            resultado={'redaccion':redaccion,'rumbo_texto':texto_rumbo,
                       'rumbo_compacto_solicitado':compacto_solicitado,
                       'rumbo_compacto_convencional':compacto_convencional,
                       'est_ini_txt':est_ini_txt,'est_fin_txt':est_fin_txt,'colindancia':colindancia}
    return render_template(
    'formulario.html',
    errores=errores,
    resultado=resultado,
    docx_ready=DOCX_AVAILABLE,
    app_version=APP_VERSION
)

@app.route('/descargar', methods=['POST'])
def descargar():
    redaccion=(request.form.get('redaccion') or '').strip()
    rumbo_texto_val=(request.form.get('rumbo_texto') or '').strip()
    rumbo_compacto_sol=(request.form.get('rumbo_compacto_solicitado') or '').strip()
    rumbo_compacto_conv=(request.form.get('rumbo_compacto_convencional') or '').strip()
    est_ini_txt=(request.form.get('est_ini_txt') or '').strip()
    est_fin_txt=(request.form.get('est_fin_txt') or '').strip()
    colindancia=(request.form.get('colindancia') or '').strip()
    if not redaccion: return "No hay redacción para exportar. Genera la redacción primero.", 400
    if not DOCX_AVAILABLE: return "La librería python-docx no está instalada en el servidor.", 500

    doc=Document()
    styles=doc.styles['Normal']; styles.font.name='Calibri'; styles.font.size=Pt(11)

    # Header with logo + text
    section=doc.sections[0]; header=section.header
    table=header.add_table(rows=1, cols=2); table.autofit=True
    # logo left
    try:
        if os.path.exists(LOGO_PATH):
            left_p=table.cell(0,0).paragraphs[0]
            left_p.alignment=WD_ALIGN_PARAGRAPH.LEFT
            left_p.add_run().add_picture(LOGO_PATH, width=Inches(1.8))
    except Exception:
        pass
    # text right
    right_p=table.cell(0,1).paragraphs[0]
    right_p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=right_p.add_run(HEADER_TEXT); r.bold=True; r.font.size=Pt(10)

    # Body
    doc.add_heading('Redacción topográfica', level=1)
    doc.add_paragraph(datetime.now().strftime("Generado el %Y-%m-%d %H:%M:%S"))
    doc.add_paragraph().add_run("Tramo: ").bold=True
    doc.add_paragraph(f"Estaciones: {est_ini_txt} → {est_fin_txt}")
    if colindancia: doc.add_paragraph(f"Colindancia: {colindancia}")
    doc.add_paragraph().add_run("Rumbos:").bold=True
    doc.add_paragraph(f"En texto: {rumbo_texto_val}")
    doc.add_paragraph(f"Compacto (solicitado): {rumbo_compacto_sol}")
    doc.add_paragraph(f"Compacto (convencional): {rumbo_compacto_conv}")
    doc.add_paragraph().add_run("Redacción completa:").bold=True
    doc.add_paragraph(redaccion)

    bio=BytesIO(); doc.save(bio); bio.seek(0)
    filename=f"Redaccion_{est_ini_txt.replace(' ','')}_{est_fin_txt.replace(' ','')}.docx" or "Redaccion.docx"
    return send_file(bio, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__=='__main__':
    app.run(debug=True)
