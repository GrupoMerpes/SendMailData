from jinja2 import Template
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64
import os

class MonitorCampañas:

    def imagen_base64(self, ruta):
        if not os.path.exists(ruta):
            return ""

        with open(ruta, "rb") as img:
            return base64.b64encode(img.read()).decode()

    def generar_contenido_email(self, reporte):

        asunto = f"📊 Reporte Mensual - {reporte['cliente']}"

        imagenes = {
            "encabezado": self.imagen_base64("ENCABEZADO.png"),
            "footer": self.imagen_base64("FOOTER.png"),
            "titulo": self.imagen_base64("TITULO.png"),
            "icono1": self.imagen_base64("ICONO1.png"),
            "icono2": self.imagen_base64("ICONO2.png"),
        }

        with open("email_template.html", encoding="utf-8") as f:
            template = Template(f.read())

        html = template.render(
            cliente=reporte["cliente"],
            area=reporte["area_campaña"],
            merpes=reporte["merpes_consumidas"],
            merpes_meta=reporte["merpes_asignadas"],
            merpes_pct=reporte["porcentaje_merpes"],

            zoftinium=reporte["zoftinium_consumidas"],
            zoftinium_meta=reporte["zoftinium_asignadas"],
            zoftinium_pct=reporte["porcentaje_zoftinium"],

            diseño=reporte["diseño_consumidas"],
            diseño_meta=reporte["diseño_asignadas"],
            diseño_pct=reporte["porcentaje_diseño"],

            imagenes=imagenes
        )

        return asunto, html


    def enviar_email(self, correos, asunto, html):

        msg = MIMEMultipart("alternative")
        msg["Subject"] = asunto
        msg["From"] = self.config_email["email"]
        msg["To"] = ", ".join(correos)

        msg.attach(MIMEText(html, "html"))

        with smtplib.SMTP(
            self.config_email["smtp_server"],
            self.config_email["puerto"]
        ) as server:

            server.starttls()
            server.login(
                self.config_email["email"],
                self.config_email["password"]
            )

            server.send_message(msg)
