import pandas as pd
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Leer el archivo de Excel
df = pd.read_excel('contactos.xlsx')  # Cambia el nombre del archivo según sea necesario

# Inicializar el driver de Firefox
driver = webdriver.Firefox()
driver.get('https://web.whatsapp.com')
mensajes = {
    'Kiosko Simple': (
        "¡Hola {nombre}! Soy Daniel, creador de **Kiosko Simple**, un software que ha ayudado a muchos negocios "
        "como el tuyo a mejorar su gestión diaria y hacer el cierre de caja mucho más fácil y rápido. 🚀\n\n"
        "Me encantaría mostrarte cómo puedes ahorrar tiempo, reducir errores y llevar un control total de tus ventas "
        "en solo unos clics. 🎯\n\n"
        "Te ofrezco una **demostración gratuita** y sin compromiso. Estoy seguro de que en pocos minutos verás cómo "
        "**Kiosko Simple** puede ayudarte a optimizar tu negocio.\n\n"
        "Solo dime si te interesa y coordinamos una breve demo que se ajuste a tu horario. "
        " ¡No pierdes nada por probarlo! 😉\n\n"
        "Visita nuestra web para más información: www.kioskosimple.com. \n\n"
        "Espero tu respuesta. ¡Estoy aquí para ayudarte a llevar tu negocio al siguiente nivel!\n\n"
        "Saludos,\nDaniel"
    ),
    'Taller Expres': (
        "¡Hola {nombre}! Soy Daniel de **Taller Expres**. Nos especializamos en ofrecer soluciones rápidas y efectivas "
        "para tus necesidades automotrices. 🚗💨\n\n"
        "Te invito a que conozcas nuestros servicios y aproveches una **oferta especial** para nuevos clientes. "
        "Estoy seguro de que podemos ayudarte a mantener tu vehículo en óptimas condiciones.\n\n"
        "¿Te gustaría que te envíe más información? ¡Espero tu respuesta!\n\n"
        "Saludos,\nDaniel"
    ),
    'Rental Machine': (
        "¡Hola {nombre}! Soy Daniel de **Rental Machine**. Ofrecemos un servicio de alquiler de maquinaria eficiente "
        "y accesible para tus proyectos. 🛠️💼\n\n"
        "Quisiera compartir contigo nuestras opciones de alquiler y cómo podemos hacer que tu trabajo sea más fácil. "
        "¿Te gustaría recibir una **demostración gratuita** de nuestros servicios?\n\n"
        "Espero tu respuesta para coordinar una charla.\n\n"
        "Saludos,\nDaniel"
    )
}

try:
    WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH, '//h1[contains(text(), "Chats")]'))
    )
    print("¡WhatsApp Web está cargado y listo!")

    for index, row in df.iterrows():
        nombre = row['nombre']  # Asumiendo que hay una columna 'nombre'
        telefono = row['telefono']  # Asumiendo que hay una columna 'telefono'
        tipo_negocio = row['tipo_negocio']  # Asumiendo que hay una columna 'tipo_negocio'
        mensaje = mensajes.get(tipo_negocio, "¡Hola {nombre}! Espero que estés bien.").format(nombre=nombre)
        url = f'https://web.whatsapp.com/send?phone={telefono}&text={mensaje}'
        driver.get(url)
        try:
            message_box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
            message_box.send_keys(Keys.ENTER)
            print(f"Mensaje enviado a {nombre} con éxito.")
        except Exception as e:
            print(f"No se pudo enviar el mensaje a {nombre}: {e}")
        sleep(5)
except Exception as e:
    print("No se pudo cargar WhatsApp Web:", e)
    driver.quit()
    exit()
driver.quit()
