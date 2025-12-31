from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service # Ajouté
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas

# Chargement des données
excel_data = pandas.read_excel('Recipients data.xlsx', sheet_name='Recipients')

# Initialisation moderne du driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

driver.get('https://web.whatsapp.com')
input("Appuyez sur ENTRÉE après vous être connecté à WhatsApp Web et que vos messages soient visibles.")

# On parcourt les contacts et les messages
for contact, message in zip(excel_data['Contact'], excel_data['Message']):
    try:
        url = f'https://web.whatsapp.com/send?phone={contact}&text={message}'
        driver.get(url)
        
        try:
            # Note: Les noms de classes comme '_3XKXx' changent souvent sur WhatsApp.
            # Si le script bloque ici, il faudra inspecter le bouton "Envoyer" sur votre navigateur.
            click_btn = WebDriverWait(driver, 35).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            sleep(2)
            click_btn.click()
            sleep(5)
            print(f'Message envoyé à : {contact}')
        except Exception as e:
            print(f"Désolé, le message n'a pas pu être envoyé à {contact}")
            
    except Exception as e:
        print(f'Erreur avec le contact {contact} : {str(e)}')

driver.quit()
print("Le script s'est exécuté avec succès.")