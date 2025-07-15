import sys
import os
os.environ['WDM_LOCAL'] = '1'  # Damit Webdriver im lokalen AppData gespeichert wird
import time
import pickle
import json
import tkinter as tk
from tkinter import messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
import zipfile
import cairosvg
from PIL import Image
import base64
from lxml import etree
import shutil

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# === Globale Variablen ===
COOKIE_FILE = "amazon_cookies.pkl"
SESSION_FILE = "amazon_session_info.json"
LOGIN_URL = "https://sellercentral.amazon.de"
SEARCH_FIELD_SELECTOR = 'input#sc-search-field.search-input.search-input-active'
SEARCH_BUTTON_SELECTOR = 'button.sc-search-button.search-icon-container'
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# Ermittle das Verzeichnis, in dem das Skript liegt
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(SCRIPT_DIR, "amazon_order_downloads")  # Download-Verzeichnis im Skript-Ordner

# === Verzeichnis erstellen ===
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# === Selenium Setup ===
def create_driver():
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.notifications": 2,
        "download.default_directory": DOWNLOAD_DIR,  # Set custom download directory
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument(f"--user-agent={USER_AGENT}")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

# === Session-Info speichern ===
def save_session_info(driver):
    session_info = {
        "url": driver.current_url,
        "timestamp": datetime.now().isoformat(),
        "user_agent": driver.execute_script("return navigator.userAgent;")
    }
    
    with open(SESSION_FILE, "w") as file:
        json.dump(session_info, file)

# === Session-Info laden ===
def load_session_info():
    if os.path.exists(SESSION_FILE):
        with open(SESSION_FILE, "r") as file:
            return json.load(file)
    return None

# === Cookies speichern ===
def save_cookies(driver):
    try:
        cookies = driver.get_cookies()
        # Filtere nur relevante Cookies für Amazon
        relevant_cookies = []
        for cookie in cookies:
            if any(domain in cookie.get('domain', '') for domain in ['amazon.de', 'amazon.com', 'sellercentral']):
                relevant_cookies.append(cookie)
        
        with open(COOKIE_FILE, "wb") as file:
            pickle.dump(relevant_cookies, file)
        
        save_session_info(driver)
        return True
    except Exception as e:
        print(f"Fehler beim Speichern der Cookies: {e}")
        return False

# === Cookies laden ===
def load_cookies(driver):
    if not os.path.exists(COOKIE_FILE):
        return False
    
    try:
        with open(COOKIE_FILE, "rb") as file:
            cookies = pickle.load(file)

        # Prüfe ob Cookies noch gültig sind
        session_info = load_session_info()
        if session_info:
            saved_time = datetime.fromisoformat(session_info["timestamp"])
            if datetime.now() - saved_time > timedelta(hours=12):  # 12 Stunden Gültigkeit
                print("Cookies sind abgelaufen")
                return False

        # Gehe erst zur Login-Seite, bevor Cookies gesetzt werden
        driver.get(LOGIN_URL)
        time.sleep(2)

        for cookie in cookies:
            # Bereinige Cookie-Daten
            cookie.pop('sameSite', None)
            cookie.pop('httpOnly', None)
            cookie.pop('secure', None)
            
            # Stelle sicher, dass die Domain korrekt ist
            if 'domain' not in cookie or not cookie['domain']:
                cookie['domain'] = '.amazon.de'
            
            try:
                driver.add_cookie(cookie)
            except Exception as e:
                print(f"Cookie konnte nicht gesetzt werden: {cookie.get('name')} - {e}")
        
        return True
    except Exception as e:
        print(f"Fehler beim Laden der Cookies: {e}")
        return False

# === Login-Status prüfen ===
def is_logged_in(driver):
    try:
        # Warte bis Seite geladen ist
        WebDriverWait(driver, 15).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        current_url = driver.current_url
        print(f"Aktuelle URL: {current_url}")
        
        # Prüfe ob wir auf einer Login-Seite sind
        if any(keyword in current_url.lower() for keyword in ["signin", "login", "auth"]):
            print("Auf Login-Seite erkannt")
            return False
        
        # Prüfe verschiedene Indikatoren für erfolgreichen Login
        login_indicators = [
            # Hauptsuchfeld
            'input#sc-search-field',
            'input.search-input',
            # Navigation/Header Elemente
            'nav[role="navigation"]',
            '#sc-navtab-reports',
            '#sc-navtab-orders',
            # Seller Central spezifische Elemente
            '.sc-dashboard',
            '[data-testid="sc-content"]',
            # Fallback: Jedes input field das nicht Login ist
            'input[type="text"]:not([name*="email"]):not([name*="password"])'
        ]
        
        for selector in login_indicators:
            try:
                element = driver.find_element(By.CSS_SELECTOR, selector)
                if element.is_displayed():
                    print(f"Login-Indikator gefunden: {selector}")
                    return True
            except:
                continue
        
        # Zusätzliche Prüfung: Wenn URL seller central enthält und nicht login
        if "sellercentral" in current_url and not any(keyword in current_url.lower() for keyword in ["signin", "login"]):
            print("Seller Central URL erkannt, Login wahrscheinlich erfolgreich")
            return True
            
        print("Keine Login-Indikatoren gefunden")
        return False
        
    except Exception as e:
        print(f"Login-Status konnte nicht geprüft werden: {e}")
        # Im Zweifelsfall True zurückgeben wenn keine klaren Fehler
        return True

# === Login durchführen (manuell) ===
def manual_login():
    driver = create_driver()
    
    try:
        driver.get(LOGIN_URL)
        
        messagebox.showinfo(
            "Login erforderlich", 
            "Bitte logge dich jetzt manuell ein (inkl. 2FA).\n\n"
            "Wichtig:\n"
            "- Stelle sicher, dass du komplett eingeloggt bist\n"
            "- Warte bis die Seller Central Startseite geladen ist\n"
            "- Drücke erst dann OK"
        )
        
        # Vereinfachte Login-Prüfung
        current_url = driver.current_url
        print(f"URL nach Login: {current_url}")
        
        # Wenn wir nicht mehr auf Login-Seite sind, als erfolgreich werten
        if not any(keyword in current_url.lower() for keyword in ["signin", "login", "auth"]):
            if save_cookies(driver):
                messagebox.showinfo("Erfolg", "Login-Cookies wurden erfolgreich gespeichert!")
            else:
                messagebox.showerror("Fehler", "Cookies konnten nicht gespeichert werden.")
        else:
            # Auch wenn Login-Seite erkannt wird, trotzdem versuchen Cookies zu speichern
            if save_cookies(driver):
                messagebox.showinfo("Cookies gespeichert", "Cookies wurden gespeichert. Teste die Suche um zu prüfen ob der Login funktioniert.")
            else:
                messagebox.showwarning("Warnung", "Login scheint nicht vollständig zu sein. Bitte versuche es erneut.")
            
    except Exception as e:
        messagebox.showerror("Fehler", f"Login fehlgeschlagen: {e}")
    finally:
        driver.quit()


# === Verarbeite heruntergeladene ZIP-Datei ===
def process_downloaded_zip(order_number):
    # Warte auf den Download
    time.sleep(5)  # Wartezeit für den Download
    
    # Finde die neueste ZIP-Datei im Download-Verzeichnis
    zip_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.zip')]
    if not zip_files:
        messagebox.showerror("Fehler", "Keine ZIP-Datei gefunden.")
        return
    
    # Nehme die neueste ZIP-Datei
    latest_zip = max(zip_files, key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)))
    zip_path = os.path.join(DOWNLOAD_DIR, latest_zip)
    
    # Erstelle einen Ordner für die entpackten Dateien (mit Bestellnummer als Name)
    extract_dir = os.path.join(DOWNLOAD_DIR, order_number)
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    # Entpacke die ZIP-Datei
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Entpacken der ZIP-Datei: {e}")
        return
    
    # Finde die SVG-Datei und die größte JPG-Datei
    svg_file = None
    jpg_files = []
    
    for root, dirs, files in os.walk(extract_dir):
        for file in files:
            if file.lower().endswith('.svg'):
                svg_file = os.path.join(root, file)
            elif file.lower().endswith('.jpg') or file.lower().endswith('.jpeg'):
                jpg_files.append(os.path.join(root, file))
    
    if not svg_file:
        messagebox.showerror("Fehler", "Keine SVG-Datei in der ZIP-Datei gefunden.")
        return
    
    if not jpg_files:
        messagebox.showerror("Fehler", "Keine JPG-Dateien in der ZIP-Datei gefunden.")
        return
    
    # Wähle die größte JPG-Datei (wahrscheinlich die richtige)
    largest_jpg = max(jpg_files, key=lambda f: os.path.getsize(f))
    
    # Verarbeite die Dateien
    if process_files(svg_file, largest_jpg, extract_dir, order_number):
        # Lösche die ZIP-Datei nur wenn die Verarbeitung erfolgreich war
        try:
            os.remove(zip_path)
            print(f"ZIP-Datei gelöscht: {zip_path}")
        except Exception as e:
            print(f"Warnung: ZIP-Datei konnte nicht gelöscht werden: {e}")


# === Verarbeite SVG und JPG zu TIFF ===
def process_files(svg_path, image_path, output_dir, order_number):
    """Verarbeite die SVG und JPG Dateien zu TIFF"""
    try:
        # 1. SVG modifizieren mit dem eingebetteten Bild
        new_svg_path = embed_image_in_svg(image_path, svg_path)
        if not new_svg_path:
            return False
        
        # 2. Konvertiere zu TIFF (mit Bestellnummer als Dateiname)
        output_path = os.path.join(output_dir, f"{order_number}.tiff")
        if convert_svg_to_tiff(new_svg_path, output_path):
            return True
        return False
        
    except Exception as e:
        messagebox.showerror("Fehler", f"Verarbeitung fehlgeschlagen: {str(e)}")
        return False


def embed_image_in_svg(image_path, svg_path):
    """Ersetzt nur das Bild innerhalb des clipPath"""
    try:
        with open(svg_path, "r+b") as f:
            tree = etree.parse(f)
            root = tree.getroot()

            namespaces = {
                'svg': 'http://www.w3.org/2000/svg',
                'xlink': 'http://www.w3.org/1999/xlink'
            }

            # Finde das Bild-Element innerhalb des clipPath
            clip_path_images = root.xpath('//svg:g[@clip-path]//svg:image', namespaces=namespaces)
            
            if not clip_path_images:
                raise Exception("Kein Bild im clipPath gefunden")
                
            # Das große Bild im clipPath (3960x2640)
            target_image = clip_path_images[0]

            with open(image_path, "rb") as img_file:
                encoded_string = "data:image/png;base64," + base64.b64encode(img_file.read()).decode('utf-8')

            target_image.set("{http://www.w3.org/1999/xlink}href", encoded_string)

            new_svg_path = os.path.splitext(svg_path)[0] + "_temp.svg"  # Als temporäre Datei markiert
            with open(new_svg_path, "wb") as new_svg:
                new_svg.write(etree.tostring(tree, pretty_print=True, encoding="UTF-8"))

            return new_svg_path

    except Exception as e:
        messagebox.showerror("Fehler", f"Bildeinbettung fehlgeschlagen: {str(e)}")
        return None

def convert_svg_to_tiff(svg_path, output_path=None):
    """Konvertiere SVG direkt zu TIFF und lösche temporäre Dateien"""
    try:
        if output_path is None:
            output_path = os.path.splitext(svg_path)[0].replace("_temp", "") + ".tiff"

        # 1. Konvertiere zu PNG
        png_path = os.path.splitext(output_path)[0] + "_temp.png"
        cairosvg.svg2png(
            url=svg_path,
            write_to=png_path,
            background_color=None,
            scale=1.0  # Keine Skalierung
        )

        # 2. Konvertiere PNG zu TIFF
        img = Image.open(png_path)
        img.save(
            output_path,
            format="TIFF",
            compression="tiff_deflate"
        )
        
        # Lösche temporäre Dateien
        os.remove(png_path)
        os.remove(svg_path)  # Lösche die temporäre SVG-Datei

        messagebox.showinfo("Erfolg", f"TIFF erfolgreich erstellt:\n{output_path}")
        return True

    except Exception as e:
        messagebox.showerror("Fehler", f"Konvertierung fehlgeschlagen: {str(e)}")
        return False

# === Bestellung suchen und verarbeiten ===
def search_order(order_number):
    driver = create_driver()
    
    try:
        # Cookies laden und prüfen
        if not load_cookies(driver):
            messagebox.showerror("Fehler", "Keine gültigen Cookies gefunden. Bitte logge dich zuerst manuell ein.")
            driver.quit()
            return

        # Seite neu laden um Login zu aktivieren
        driver.refresh()
        time.sleep(3)
        
        # Flexiblere Login-Prüfung
        current_url = driver.current_url
        print(f"URL nach Cookie-Login: {current_url}")
        
        # Wenn wir auf Login-Seite sind, Session ist abgelaufen
        if any(keyword in current_url.lower() for keyword in ["signin", "login", "auth"]):
            messagebox.showerror("Session abgelaufen", "Deine Session ist abgelaufen. Bitte logge dich erneut ein.")
            driver.quit()
            return

        # Suchfeld schneller finden mit explizitem Wait
        search_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input#sc-search-field"))
        )
        print("Suchfeld gefunden")
        
        # Suche durchführen
        search_field.clear()
        search_field.send_keys(order_number)
        
        # Suchbutton finden und klicken
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.sc-search-button.search-icon-container"))
        )
        search_button.click()
        print("Suche durchgeführt")
        
        try:
            # Warte auf Suchergebnisse mit kürzerem Timeout für die "Nicht gefunden"-Prüfung
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.sc-no-results-message, span.a-expander-prompt"))
            )
            
            # Prüfe ob "Keine Ergebnisse" Meldung vorhanden ist
            no_results = driver.find_elements(By.CSS_SELECTOR, "div.sc-no-results-message")
            if no_results:
                messagebox.showwarning("Nicht gefunden", f"Bestellung {order_number} wurde nicht gefunden. Bitte überprüfen Sie die Bestellnummer.")
                driver.quit()
                return
                
            # Wenn keine "Nicht gefunden" Meldung, dann normal fortfahren
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span.a-expander-prompt"))
            ).click()
            print("Auf Expander geklickt")
            
            # Klicke auf den Link "Anpassungsinformationen"
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.a-link-normal[href*='fulfillment']"))
            ).click()
            print("Auf Anpassungsinformationen geklickt")
            
            # Warte bis der Download-Button sichtbar ist und klicke ihn
            download_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "kat-button.download-zip-file-button"))
            )
            download_button.click()
            print("Download-Button geklickt")
            
            # Verarbeite die heruntergeladene ZIP-Datei
            process_downloaded_zip(order_number)
            
        except Exception as e:
            # Falls Timeout beim Finden der Elemente
            messagebox.showwarning("Nicht gefunden", f"Bestellung {order_number} wurde nicht gefunden oder die Seite hat zu lange geladen. Bitte überprüfen Sie die Bestellnummer.")
            
    except Exception as e:
        messagebox.showerror("Fehler", f"Prozess fehlgeschlagen: {e}")
        print(f"Fehler aufgetreten: {str(e)}")
    finally:
        driver.quit()

# === Cookie-Status prüfen ===
def check_cookie_status():
    if not os.path.exists(COOKIE_FILE):
        messagebox.showinfo("Cookie-Status", "Keine Cookies gespeichert.")
        return
    
    session_info = load_session_info()
    if session_info:
        saved_time = datetime.fromisoformat(session_info["timestamp"])
        time_diff = datetime.now() - saved_time
        
        status = f"Cookies gespeichert am: {saved_time.strftime('%d.%m.%Y %H:%M:%S')}\n"
        status += f"Alter: {time_diff.days} Tage, {time_diff.seconds // 3600} Stunden\n"
        
        if time_diff > timedelta(hours=12):
            status += "Status: Abgelaufen"
        else:
            status += "Status: Gültig"
        
        messagebox.showinfo("Cookie-Status", status)
    else:
        messagebox.showinfo("Cookie-Status", "Cookie-Datei vorhanden, aber keine Session-Info.")

# === GUI ===
def start_gui():
    window = tk.Tk()
    window.title("Amazon Seller Central - Bestellungssuche & Verarbeitung")
    window.geometry("450x300")
    
    # Titel
    tk.Label(window, text="Amazon Seller Central", font=("Arial", 16, "bold")).pack(pady=10)
    
    # Bestellnummer eingeben
    tk.Label(window, text="Bestellnummer eingeben:", font=("Arial", 12)).pack(pady=5)
    order_entry = tk.Entry(window, font=("Arial", 12), width=35)
    order_entry.pack(pady=5)

    def handle_search():
        order_number = order_entry.get().strip()
        if order_number:
            search_order(order_number)
        else:
            messagebox.showwarning("Hinweis", "Bitte eine Bestellnummer eingeben.")

    # Buttons
    tk.Button(window, text="Bestellung suchen & verarbeiten", command=handle_search, 
              font=("Arial", 12), bg="#FF9900", fg="black", width=25).pack(pady=10)
    
    tk.Button(window, text="Manuell einloggen & Cookies speichern", command=manual_login,
              font=("Arial", 10), width=35).pack(pady=5)
    
    tk.Button(window, text="Cookie-Status prüfen", command=check_cookie_status,
              font=("Arial", 10), width=35).pack(pady=5)

    # Hinweis
    tk.Label(window, text="Hinweis: Cookies sind ca. 12 Stunden gültig", 
             font=("Arial", 9), fg="gray").pack(pady=(20, 5))

    window.mainloop()

# === Programmstart ===
if __name__ == "__main__":
    start_gui()
