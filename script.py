import os
import random
import pandas
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURATION ---
EXCEL_FILE = 'Recipients data.xlsx'
SESSION_DIR = "whatsapp_session"

def clean_number(phone):
    return "".join(filter(str.isdigit, str(phone)))

chrome_options = Options()
script_dir = os.path.dirname(os.path.abspath(__file__))
user_data_path = os.path.join(script_dir, SESSION_DIR)

chrome_options.add_argument(f"--user-data-dir={user_data_path}")
chrome_options.add_argument("--profile-directory=Default")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    excel_data = pandas.read_excel(EXCEL_FILE, sheet_name='Recipients')
    driver.get('https://web.whatsapp.com')
    
    print("\n" + "="*50)
    print("1. ATTENDEZ LE CHARGEMENT COMPLET DE WHATSAPP.")
    print("2. REVIENEZ ICI QUAND VOUS VOYEZ VOS CHATS.")
    print("="*50)
    input("\n>>> APPUYEZ SUR [ENTRÉE] POUR DÉMARRER...")

    for index, row in excel_data.iterrows():
        contact = clean_number(row['Contact'])
        message = row['Message']
        if not contact: continue

        print(f"\n[{index + 1}/{len(excel_data)}] Chargement de la page pour {contact}...")
        
        url = f'https://web.whatsapp.com/send?phone={contact}&text={message}'
        driver.get(url)

        # --- BOUCLE DE PATIENCE (Priorité 5 améliorée) ---
        sent = False
        max_attempts = 20 # 20 tentatives de 2 secondes = 40 secondes d'attente max
        
        for attempt in range(max_attempts):
            try:
                # On cherche si le message d'erreur "numéro invalide" apparaît
                invalid = driver.find_elements(By.XPATH, '//div[contains(text(), "invalide") or contains(text(), "invalid")]')
                if invalid:
                    print(f"❌ Le numéro {contact} est invalide.")
                    break
                
                # On cherche la zone de texte
                text_box = driver.find_element(By.XPATH, '//footer//div[@contenteditable="true"]')
                
                if text_box:
                    print("   Interface chargée. Envoi du message...")
                    sleep(2) # Sécurité pour laisser le texte se pré-remplir
                    text_box.send_keys(Keys.ENTER)
                    print(f"✅ Envoyé à {contact} !")
                    sent = True
                    break
            except:
                # Si l'élément n'est pas trouvé, on attend 2 sec et on recommence la boucle
                if attempt % 5 == 0 and attempt > 0:
                    print(f"   ...toujours en attente du chargement ({attempt*2}s écoulées)...")
                sleep(2)

        if not sent:
            print(f"⚠️ Échec : La page a mis trop de temps (>40s) ou le numéro {contact} pose problème.")

        # Pause entre deux contacts pour éviter d'être bloqué par WhatsApp
        wait_time = random.uniform(7, 12)
        print(f"Pause de sécurité : {wait_time:.1f}s")
        sleep(wait_time)

except Exception as e:
    print(f"\n❌ ERREUR : {e}")

finally:
    print("\n>>> Travail terminé.")
    driver.quit()