import os
import base64
import requests
from msal import ConfidentialClientApplication
from datetime import datetime

# =====================================================
# CONFIGURACION
# =====================================================

TENANT_ID = "tenant-id-correos"
CLIENT_ID = "client-id-correos"
CLIENT_SECRET = "token-secreto-graphl"

BUZON = "residentesgetxo@setex.es"

ASUNTO_BUSCADO = "Envio Correcto"

DESTINO = r"E:\FILER\setex\batchs\residentCard\20065"

# =====================================================

os.makedirs(DESTINO, exist_ok=True)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = ["https://graph.microsoft.com/.default"]

GRAPH_URL = "https://graph.microsoft.com/v1.0"

# =====================================================
# OBTENER TOKEN
# =====================================================

app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

token_result = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" not in token_result:
    print("ERROR obteniendo token")
    print(token_result)
    exit()

token = token_result["access_token"]

headers = {
    "Authorization": f"Bearer {token}"
}

# =====================================================
# FECHA DE HOY
# =====================================================

fecha_hoy = datetime.now().date()

# =====================================================
# LEER MENSAJES
# =====================================================

url = (
    f"{GRAPH_URL}/users/{BUZON}/messages"
    "?$top=500"
    "&$orderby=receivedDateTime desc"
)

resp = requests.get(url, headers=headers)

if resp.status_code != 200:
    print("ERROR leyendo mensajes")
    print(resp.text)
    exit()

mensajes = resp.json().get("value", [])

for mensaje in mensajes:

    asunto = mensaje.get("subject", "")

    if ASUNTO_BUSCADO.lower() not in asunto.lower():
        continue

    fecha_correo = mensaje.get("receivedDateTime")

    if not fecha_correo:
        continue

    fecha_correo = datetime.fromisoformat(
        fecha_correo.replace("Z", "+00:00")
    ).date()

    if fecha_correo != fecha_hoy:
        continue

    print(f"Correo encontrado: {asunto}")

    mensaje_id = mensaje["id"]

    # =================================================
    # ADJUNTOS
    # =================================================

    attach_url = (
        f"{GRAPH_URL}/users/{BUZON}/messages/"
        f"{mensaje_id}/attachments"
    )

    attach_resp = requests.get(
        attach_url,
        headers=headers
    )

    if attach_resp.status_code != 200:
        continue

    adjuntos = attach_resp.json().get("value", [])

    for adjunto in adjuntos:

        nombre = adjunto.get("name", "")

        if not nombre.lower().endswith(".txt"):
            continue

        contenido = adjunto.get("contentBytes")

        if not contenido:
            continue

        ruta = os.path.join(DESTINO, nombre)

        with open(ruta, "wb") as f:
            f.write(base64.b64decode(contenido))

        print(f"TXT descargado: {ruta}")

        # ==========================================
        # ENVIAR CORREO DE CONFIRMACION
        # ==========================================

        email_url = f"{GRAPH_URL}/users/{BUZON}/sendMail"

        email_body = {
            "message": {
                "subject": "Envio Correcto Realizado",
                "body": {
                    "contentType": "Text",
                    "content": (
                        f'El archivo "{nombre}" ha sido colocado '
                        f'en la carpeta satisfactoriamente.'
                    )
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": "informatica@setex.es"
                        }
                    }
                ]
            },
            "saveToSentItems": True
        }

        resp_mail = requests.post(
            email_url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            },
            json=email_body
        )

        if resp_mail.status_code == 202:
            print("Correo de confirmación enviado")
        else:
            print("Error enviando correo")
            print(resp_mail.text)

        # ==========================================
        # MARCAR COMO LEIDO
        # ==========================================

        requests.patch(
            f"{GRAPH_URL}/users/{BUZON}/messages/{mensaje_id}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            },
            json={
                "isRead": True
            }
        )

        print("Correo marcado como leído")
        print("Proceso finalizado correctamente")

        exit()

print(
    f"No existe ningún correo recibido hoy "
    f"con asunto '{ASUNTO_BUSCADO}'"
)

# ==========================================
# ENVIAR AVISO DE ERROR
# ==========================================

email_url = f"{GRAPH_URL}/users/{BUZON}/sendMail"

email_body = {
    "message": {
        "subject": "ERRROR - No se encontró archivo de residentes",
        "body": {
            "contentType": "Text",
            "content": (
                "No se encontró ningún correo con asunto "
                "'Envio Correcto' para el día de hoy.\n\n"
                "Por favor, revisar manualmente el proceso."
            )
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": "informatica@setex.es"
                }
            }
        ]
    },
    "saveToSentItems": True
}

resp_mail = requests.post(
    email_url,
    headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    },
    json=email_body
)

if resp_mail.status_code == 202:
    print("Correo de aviso enviado")
else:
    print("Error enviando correo de aviso")
    print(resp_mail.text)

exit()