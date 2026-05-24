import os
import json
import base64
import requests
from datetime import datetime
from flask import Flask, request
from twilio.rest import Client
from twilio.twiml.messaging_response import MessagingResponse
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.cloud import vision

app = Flask(__name__)

# ── CONFIGURACIÓN ────────────────────────────────────────────────────────────
TWILIO_ACCOUNT_SID = os.environ["TWILIO_ACCOUNT_SID"]
TWILIO_AUTH_TOKEN  = os.environ["TWILIO_AUTH_TOKEN"]
TWILIO_NUMBER      = os.environ["TWILIO_NUMBER"]       # whatsapp:+14155238886
SHEET_ID           = os.environ["SHEET_ID"]
GOOGLE_CREDS_JSON  = os.environ["GOOGLE_CREDS_JSON"]   # contenido del JSON como string

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

CATEGORIAS = {
    # Supermercado
    "スーパー","イオン","マックスバリュ","ライフ","西友","ビッグ","コープ","業務スーパー",
    "ドン・キホーテ","ドンキ","コンビニ","セブン","ローソン","ファミマ","ミニストップ",
    # Restaurante
    "レストラン","食堂","居酒屋","ラーメン","寿司","すし","焼肉","定食","カフェ","cafe",
    "マクドナルド","モスバーガー","すき家","吉野家","松屋","ガスト","デニーズ","サイゼ",
    # Transporte
    "タクシー","電車","バス","suica","pasmo","駅","空港","JAL","ANA","新幹線",
    # Salud
    "薬局","ドラッグ","マツキヨ","ツルハ","ウエルシア","クリニック","病院","歯科",
    # Educación
    "書店","本屋","紀伊國屋","ジュンク堂","塾","スクール","学校",
    # Entretenimiento
    "映画","シネマ","カラオケ","ゲーム","ゲーセン","遊園地","netflix","spotify",
    # Ropa
    "ユニクロ","ZARA","H&M","GU","ジーユー","アパレル","洋服","靴",
    # Hogar/Servicios
    "電気","ガス","水道","ニトリ","イケア","IKEA","コジマ","ヤマダ","ビックカメラ",
    # Tecnología
    "アップル","apple","ヨドバシ","ソフトバンク","ドコモ","au","パソコン",
}

CATEGORIA_MAP = {
    "スーパー":"Supermercado","イオン":"Supermercado","マックスバリュ":"Supermercado",
    "ライフ":"Supermercado","西友":"Supermercado","コープ":"Supermercado",
    "業務スーパー":"Supermercado","ドン・キホーテ":"Supermercado","ドンキ":"Supermercado",
    "コンビニ":"Supermercado","セブン":"Supermercado","ローソン":"Supermercado",
    "ファミマ":"Supermercado","ミニストップ":"Supermercado","ビッグ":"Supermercado",
    "レストラン":"Restaurante / Delivery","食堂":"Restaurante / Delivery",
    "居酒屋":"Restaurante / Delivery","ラーメン":"Restaurante / Delivery",
    "寿司":"Restaurante / Delivery","すし":"Restaurante / Delivery",
    "焼肉":"Restaurante / Delivery","定食":"Restaurante / Delivery",
    "カフェ":"Restaurante / Delivery","cafe":"Restaurante / Delivery",
    "マクドナルド":"Restaurante / Delivery","モスバーガー":"Restaurante / Delivery",
    "すき家":"Restaurante / Delivery","吉野家":"Restaurante / Delivery",
    "松屋":"Restaurante / Delivery","ガスト":"Restaurante / Delivery",
    "デニーズ":"Restaurante / Delivery","サイゼ":"Restaurante / Delivery",
    "タクシー":"Transporte","電車":"Transporte","バス":"Transporte",
    "suica":"Transporte","pasmo":"Transporte","駅":"Transporte",
    "空港":"Transporte","JAL":"Transporte","ANA":"Transporte","新幹線":"Transporte",
    "薬局":"Salud / Farmacia","ドラッグ":"Salud / Farmacia","マツキヨ":"Salud / Farmacia",
    "ツルハ":"Salud / Farmacia","ウエルシア":"Salud / Farmacia",
    "クリニック":"Salud / Farmacia","病院":"Salud / Farmacia","歯科":"Salud / Farmacia",
    "書店":"Educación","本屋":"Educación","紀伊國屋":"Educación",
    "ジュンク堂":"Educación","塾":"Educación","スクール":"Educación","学校":"Educación",
    "映画":"Entretenimiento","シネマ":"Entretenimiento","カラオケ":"Entretenimiento",
    "ゲーム":"Entretenimiento","ゲーセン":"Entretenimiento","遊園地":"Entretenimiento",
    "netflix":"Entretenimiento","spotify":"Entretenimiento",
    "ユニクロ":"Ropa / Calzado","ZARA":"Ropa / Calzado","H&M":"Ropa / Calzado",
    "GU":"Ropa / Calzado","ジーユー":"Ropa / Calzado","アパレル":"Ropa / Calzado",
    "洋服":"Ropa / Calzado","靴":"Ropa / Calzado",
    "電気":"Hogar / Servicios","ガス":"Hogar / Servicios","水道":"Hogar / Servicios",
    "ニトリ":"Hogar / Servicios","イケア":"Hogar / Servicios","IKEA":"Hogar / Servicios",
    "コジマ":"Hogar / Servicios","ヤマダ":"Hogar / Servicios","ビックカメラ":"Hogar / Servicios",
    "アップル":"Tecnología","apple":"Tecnología","ヨドバシ":"Tecnología",
    "ソフトバンク":"Tecnología","ドコモ":"Tecnología","au":"Tecnología","パソコン":"Tecnología",
}

# Estado temporal por usuario (en memoria)
pending = {}

# ── GOOGLE SHEETS ────────────────────────────────────────────────────────────
def get_sheets_service():
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)

def append_row(fecha, descripcion, categoria, monto, notas, compartido, con_quien, mi_parte=None):
    service = get_sheets_service()
    mes = datetime.strptime(fecha, "%Y-%m-%d").strftime("%B")
    mi_parte = mi_parte if mi_parte is not None else (monto / 2 if compartido == "Sí" else monto)
    values = [[fecha, descripcion, categoria, monto, mes, notas, "✅", compartido, con_quien, mi_parte]]
    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range="Registro!A:J",
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()

# ── GOOGLE VISION ────────────────────────────────────────────────────────────
def analizar_ticket(image_url):
    # Descargar imagen desde Twilio (requiere auth)
    resp = requests.get(
        image_url,
        auth=(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
    )
    image_content = base64.b64encode(resp.content).decode("utf-8")

    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/cloud-platform"]
    )
    client = vision.ImageAnnotatorClient(credentials=creds)
    image  = vision.Image(content=base64.b64decode(image_content))
    result = client.text_detection(image=image)
    texto  = result.text_annotations[0].description if result.text_annotations else ""

    # Extraer monto (buscar 合計, 小計, お会計, o número grande con ¥)
    monto = extraer_monto(texto)
    categoria = detectar_categoria(texto)
    descripcion = extraer_descripcion(texto)

    return monto, categoria, descripcion, texto

def extraer_fecha(texto):
    import re
    # Patrones de fecha japonesa: 2026年05月20日, 2026/05/20, 2026-05-20
    patrones = [
        r"(\d{4})年(\d{1,2})月(\d{1,2})日",
        r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})",
    ]
    for pat in patrones:
        m = re.search(pat, texto)
        if m:
            anio, mes, dia = m.group(1), m.group(2).zfill(2), m.group(3).zfill(2)
            return f"{anio}-{mes}-{dia}"
    return None

def extraer_monto(texto):
    import re
    lineas = texto.split("\n")

    # Palabras clave de TOTAL (orden de prioridad)
    # Excluimos líneas que contengan 税 (impuesto) para no confundir con 税合計
    keywords_total = ["合計", "お会計", "請求額", "お支払", "合　計", "合 計"]
    keywords_excluir = ["税合計", "税込", "内税", "外税", "消費税", "8％税", "10％税",
                        "８%税", "１０%税", "税対象", "税　合計"]

    for keyword in keywords_total:
        for linea in lineas:
            # Saltar líneas de impuestos
            if any(ex in linea for ex in keywords_excluir):
                continue
            if keyword in linea:
                numeros = re.findall(r"[¥￥]?\s*(\d[\d,]+)", linea)
                if numeros:
                    candidatos = [int(n.replace(",","")) for n in numeros]
                    # Tomar el mayor de la línea (evita agarrar cantidades)
                    mejor = max(candidatos)
                    if mejor > 10:  # ignorar números muy chicos
                        return mejor

    # Fallback: mayor número del ticket excluyendo líneas de impuesto y códigos largos
    mayor = 0
    for linea in lineas:
        if any(ex in linea for ex in keywords_excluir):
            continue
        # Ignorar líneas con números de transacción/código (>10 dígitos seguidos)
        if re.search(r"\d{10,}", linea):
            continue
        # Ignorar líneas con palabras clave de referencia
        skip_refs = ["取引","番号","No","ＮＯ","ID","承認","処理","カード","CL","T7"]
        if any(s in linea for s in skip_refs):
            continue
        numeros = re.findall(r"\d[\d,]+", linea)
        for n in numeros:
            val = int(n.replace(",",""))
            # Ignorar montos irreales (>500,000 o <10)
            if val > 500000 or val < 10:
                continue
            if val > mayor:
                mayor = val
    return mayor

def detectar_categoria(texto):
    texto_lower = texto.lower()
    for keyword, cat in CATEGORIA_MAP.items():
        if keyword.lower() in texto_lower:
            return cat
    return "Otros"

def extraer_descripcion(texto):
    import re
    lineas = texto.split("\n")
    skip_patterns = [
        r"^\d{2,4}[-/年]\d{1,2}",
        r"^\d{10,}",
        r"http",
        r"^T\d{13}",
        r"^\d+$",
        r"登録番号",
        r"ありがとう",
        r"お問合",
        r"帮助",
        r"Help",
        r"領収",
        r"レシート",
        r"ご利用",
        r"^[-=＝－]+$",
    ]
    for linea in lineas:
        linea = linea.strip()
        if len(linea) < 2:
            continue
        if any(re.search(p, linea, re.IGNORECASE) for p in skip_patterns):
            continue
        # Priorizar líneas con kanji/katakana (nombre de comercio japonés)
        if re.search(r"[\u30A0-\u30FF\u4E00-\u9FFF]", linea):
            return linea[:50]
    # Fallback: primera línea no vacía
    for linea in lineas:
        linea = linea.strip()
        if len(linea) > 2:
            return linea[:50]
    return "Ticket"

# ── WEBHOOK PRINCIPAL ────────────────────────────────────────────────────────
@app.route("/webhook", methods=["POST"])
def webhook():
    sender      = request.form.get("From", "")
    num_media   = int(request.form.get("NumMedia", 0))
    body        = request.form.get("Body", "").strip().lower()
    media_url   = request.form.get("MediaUrl0", "")
    resp        = MessagingResponse()
    msg         = resp.message()

    # ── El usuario manda una IMAGEN ─────────────────────────────────────────
    if num_media > 0 and media_url:
        try:
            monto, categoria, descripcion, texto_raw = analizar_ticket(media_url)
            fecha_ticket = extraer_fecha(texto_raw) or datetime.today().strftime("%Y-%m-%d")
            pending[sender] = {
                "monto": monto,
                "categoria": categoria,
                "descripcion": descripcion,
                "fecha": fecha_ticket,
                "compartido": "No",
                "con_quien": "",
                "step": "confirmar"
            }
            msg.body(
                f"🧾 *Ticket detectado*\n\n"
                f"📍 Comercio: {descripcion}\n"
                f"💴 Monto: ¥{monto:,}\n"
                f"🏷️ Categoría: {categoria}\n"
                f"📅 Fecha: {fecha_ticket}\n\n"
                f"¿Es correcto? Respondé:\n"
                f"• *ok* para confirmar\n"
                f"• *compartido* si lo dividís con alguien\n"
                f"• *categoría: [nombre]* para corregir\n"
                f"• *monto: [número]* para corregir\n"
                f"• *fecha: [YYYY-MM-DD]* para corregir"
            )
        except Exception as e:
            msg.body(f"❌ No pude leer el ticket. Intentá con mejor luz o más cerca.\n({e})")
        return str(resp)

    # ── El usuario responde a un ticket pendiente ────────────────────────────
    if sender in pending:
        data = pending[sender]

        if body == "ok" or body == "si" or body == "sí":
            mi_parte = data.get("mi_parte", data["monto"] // 2 if data["compartido"] == "Sí" else data["monto"])
            append_row(data["fecha"], data["descripcion"], data["categoria"],
                       data["monto"], "", data["compartido"], data["con_quien"], mi_parte)
            del pending[sender]
            msg.body(
                f"✅ *Guardado*\n\n"
                f"📍 {data['descripcion']}\n"
                f"💴 ¥{data['monto']:,}"
                + (f"\n👥 Compartido con {data['con_quien']} → tu parte: ¥{mi_parte:,}" if data["compartido"] == "Sí" else "")
                + f"\n🏷️ {data['categoria']}"
            )

        elif body.startswith("compartido"):
            # Formatos soportados:
            # "compartido"               → mitad del total
            # "compartido con Juan"      → mitad del total, con Juan
            # "compartido 1500"          → monto específico compartido
            # "compartido 1500 con Juan" → monto específico, con Juan
            import re
            con_quien = ""
            monto_compartido = None

            match_con = re.search(r'con\s+([\w]+)', body)
            if match_con:
                con_quien = match_con.group(1).title()

            match_monto = re.search(r'compartido\s+(\d[\d,]*)', body)
            if match_monto:
                monto_compartido = int(match_monto.group(1).replace(",", ""))

            if monto_compartido:
                mi_parte = monto_compartido // 2
                detalle = (f"💴 Total ticket: ¥{data['monto']:,}\n"
                           f"💜 Parte compartida: ¥{monto_compartido:,}\n"
                           f"💰 Tu parte: ¥{mi_parte:,}")
            else:
                monto_compartido = data["monto"]
                mi_parte = monto_compartido // 2
                detalle = f"💴 Total: ¥{data['monto']:,}\n💰 Tu parte: ¥{mi_parte:,}"

            data["compartido"]       = "Sí"
            data["con_quien"]        = con_quien
            data["monto_compartido"] = monto_compartido
            data["mi_parte"]         = mi_parte
            data["step"]             = "confirmar"
            pending[sender]          = data

            msg.body(
                f"👥 *Gasto compartido*{' con ' + con_quien if con_quien else ''}\n\n"
                f"{detalle}\n\n"
                f"Respondé *ok* para confirmar."
            )

        elif body.startswith("fecha:"):
            import re
            fecha_input = body.split(":",1)[1].strip()
            # Aceptar formatos: 2026-05-20, 20/05/2026, 20-05-2026
            m = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", fecha_input)
            if not m:
                m = re.match(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", fecha_input)
                if m:
                    fecha_fmt = f"{m.group(3)}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
                else:
                    fecha_fmt = None
            else:
                fecha_fmt = f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
            if fecha_fmt:
                data["fecha"] = fecha_fmt
                pending[sender] = data
                msg.body(f"📅 Fecha cambiada a *{fecha_fmt}*. Respondé *ok* para confirmar.")
            else:
                msg.body("❌ No entendí la fecha. Formato: *fecha: 2026-05-20*")

        elif body.startswith("categoría:") or body.startswith("categoria:"):
            nueva = body.split(":",1)[1].strip().title()
            data["categoria"] = nueva
            pending[sender] = data
            msg.body(f"🏷️ Categoría cambiada a *{nueva}*. Respondé *ok* para confirmar.")

        elif body.startswith("monto:"):
            try:
                nuevo = int(body.split(":",1)[1].strip().replace(",","").replace("¥",""))
                data["monto"] = nuevo
                pending[sender] = data
                msg.body(f"💴 Monto cambiado a *¥{nuevo:,}*. Respondé *ok* para confirmar.")
            except:
                msg.body("❌ No entendí el monto. Ejemplo: *monto: 1500*")

        elif body == "cancelar":
            del pending[sender]
            msg.body("🗑️ Ticket descartado.")

        else:
            msg.body(
                "No entendí. Respondé:\n"
                "• *ok* para guardar\n"
                "• *compartido con [nombre]*\n"
                "• *categoría: [nombre]*\n"
                "• *monto: [número]*\n"
                "• *cancelar*"
            )
        return str(resp)

    # ── Mensaje sin contexto ─────────────────────────────────────────────────
    msg.body("👋 Mandame la foto de un ticket para registrarlo en tu planilla.")
    return str(resp)

if __name__ == "__main__":
    app.run(debug=False, port=5000)
