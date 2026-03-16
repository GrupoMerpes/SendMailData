import smtplib
from config import EMAIL_CONFIG

try:
    print("Conectando a Gmail...")
    server = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['puerto'])
    server.starttls()
    server.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
    print("✅ Conexión exitosa. El problema NO es la contraseña.")
    server.quit()
except Exception as e:
    print(f"❌ Error de conexión: {e}")