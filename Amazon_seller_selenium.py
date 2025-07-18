import os
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
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument(f"--user-agent={USER_AGENT}")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    # Create drivers directory if it doesn't exist
    drivers_dir = os.path.join(SCRIPT_DIR, "drivers")
    if not os.path.exists(drivers_dir):
        os.makedirs(drivers_dir)

    # Try local driver first, then fall back to ChromeDriverManager
    local_driver_path = os.path.join(drivers_dir, "chromedriver.exe")
    if os.path.exists(local_driver_path):
        try:
            driver = webdriver.Chrome(service=Service(local_driver_path), options=chrome_options)
        except:
            # If local driver fails, use ChromeDriverManager
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    else:
        # Use ChromeDriverManager if local driver doesn't exist
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # Stealth: Remove "webdriver" from navigator (fixed JavaScript syntax)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

# === Session-Info speichern ===
def save_session_info(driver):
    try:
        session_info = {
            "url": driver.current_url,
            "timestamp": datetime.now().isoformat(),
            "user_agent": driver.execute_script("return navigator.userAgent;")
        }
        
        with open(SESSION_FILE, "w") as file:
            json.dump(session_info, file)
        print(f"Session-Info gespeichert: {SESSION_FILE}")
        return True
    except Exception as e:
        print(f"Fehler beim Speichern der Session-Info: {e}")
        return False

# === Session-Info laden ===
def load_session_info():
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, "r") as file:
                return json.load(file)
        except Exception as e:
            print(f"Fehler beim Laden der Session-Info: {e}")
    return None

# === Cookies speichern ===
def save_cookies(driver):
    try:
        print("\n=== DEBUG: Start Cookie-Speicherung ===")
        print(f"Aktuelle URL: {driver.current_url}")
        
        # Warte kurz um sicherzustellen, dass die Seite vollständig geladen ist
        time.sleep(2)
        
        # Hole alle Cookies
        cookies = driver.get_cookies()
        print(f"Gefunden: {len(cookies)} Cookies insgesamt")
        
        # Filtere relevante Cookies
        relevant_cookies = []
        for cookie in cookies:
            domain = cookie.get('domain', '')
            if any(keyword in domain for keyword in ['amazon.de', 'amazon.com', 'sellercentral']):
                relevant_cookies.append(cookie)
                print(f" - Relevantes Cookie: {cookie['name']} (Domain: {domain})")
        
        print(f"Relevante Cookies: {len(relevant_cookies)}")
        
        if not relevant_cookies:
            print("WARNUNG: Keine relevanten Cookies gefunden!")
            return False
        
        # Speichere Cookies
        with open(COOKIE_FILE, "wb") as file:
            pickle.dump(relevant_cookies, file)
        print(f"Cookies gespeichert in: {os.path.abspath(COOKIE_FILE)}")
        
        # Speichere Session-Info
        if save_session_info(driver):
            print("Session-Info erfolgreich gespeichert")
        
        return True
        
    except Exception as e:
        print(f"FEHLER beim Speichern der Cookies: {e}")
        return False

# === Cookies laden ===
def load_cookies(driver):
    if not os.path.exists(COOKIE_FILE):
        print("Cookie-Datei nicht gefunden")
        return False
    
    try:
        with open(COOKIE_FILE, "rb") as file:
            cookies = pickle.load(file)
        
        print(f"Cookies geladen: {len(cookies)} Stück")

        # Prüfe ob Cookies noch gültig sind
        session_info = load_session_info()
        if session_info:
            saved_time = datetime.fromisoformat(session_info["timestamp"])
            time_diff = datetime.now() - saved_time
            print(f"Cookie-Alter: {time_diff}")
            
            if time_diff > timedelta(hours=12):  # 12 Stunden Gültigkeit
                print("Cookies sind abgelaufen")
                return False

        # Gehe erst zur Login-Seite, bevor Cookies gesetzt werden
        driver.get(LOGIN_URL)
        time.sleep(3)

        # Setze Cookies
        cookies_set = 0
        for cookie in cookies:
            try:
                # Bereinige Cookie-Daten
                clean_cookie = {
                    'name': cookie.get('name'),
                    'value': cookie.get('value'),
                    'domain': cookie.get('domain', '.amazon.de'),
                    'path': cookie.get('path', '/'),
                }
                
                # Füge optionale Attribute hinzu, falls vorhanden
                if 'expiry' in cookie:
                    clean_cookie['expiry'] = cookie['expiry']
                
                driver.add_cookie(clean_cookie)
                cookies_set += 1
                
            except Exception as e:
                print(f"Cookie konnte nicht gesetzt werden: {cookie.get('name')} - {e}")
        
        print(f"Erfolgreich {cookies_set} Cookies gesetzt")
        return cookies_set > 0
        
    except Exception as e:
        print(f"Fehler beim Laden der Cookies: {e}")
        return False

# === Verbesserte manuelle Login-Funktion ===
def manual_login():
    driver = None
    
    try:
        print("Starte manuellen Login-Prozess...")
        
        # Driver initialisieren
        driver = create_driver()
        print("Browser gestartet")
        
        # Zur Login-Seite navigieren
        driver.get(LOGIN_URL)
        print("Zur Login-Seite navigiert")
        
        # Zeige Login-Anweisungen
        messagebox.showinfo(
            "Manueller Login erforderlich", 
            "Bitte führe den Login jetzt durch:\n\n"
            "1. Gib deine Amazon-Anmeldedaten ein\n"
            "2. Führe die 2-Faktor-Authentifizierung durch\n"
            "3. Warte bis du zur Seller Central Startseite gelangst\n"
            "4. Klicke dann auf 'Login abgeschlossen' in diesem Fenster\n\n"
            "Hinweis: Das Browser-Fenster bleibt offen bis du bestätigst!"
        )
        
        # Warte auf Benutzer-Bestätigung mit Dialog
        root = tk.Tk()
        root.withdraw()  # Verstecke das Hauptfenster
        
        login_confirmed = False
        
        def confirm_login():
            nonlocal login_confirmed
            login_confirmed = True
            confirmation_window.destroy()
        
        def cancel_login():
            nonlocal login_confirmed
            login_confirmed = False
            confirmation_window.destroy()
        
        # Erstelle Bestätigungsfenster
        confirmation_window = tk.Toplevel(root)
        confirmation_window.title("Login-Bestätigung")
        confirmation_window.geometry("400x200")
        confirmation_window.grab_set()  # Modal window
        
        tk.Label(confirmation_window, 
                text="Bist du erfolgreich eingeloggt?", 
                font=("Arial", 12)).pack(pady=20)
        
        tk.Label(confirmation_window, 
                text="Klicke nur auf 'Login abgeschlossen' wenn du\n"
                     "komplett eingeloggt bist und die Seller Central\n"
                     "Startseite siehst!", 
                font=("Arial", 10)).pack(pady=10)
        
        button_frame = tk.Frame(confirmation_window)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, 
                 text="Login abgeschlossen", 
                 command=confirm_login, 
                 bg="green", 
                 fg="white",
                 font=("Arial", 12)).pack(side=tk.LEFT, padx=10)
        
        tk.Button(button_frame, 
                 text="Abbrechen", 
                 command=cancel_login, 
                 bg="red", 
                 fg="white",
                 font=("Arial", 12)).pack(side=tk.LEFT, padx=10)
        
        # Warte auf Benutzer-Entscheidung
        root.wait_window(confirmation_window)
        
        if not login_confirmed:
            print("Login abgebrochen durch Benutzer")
            messagebox.showinfo("Abgebrochen", "Login-Prozess wurde abgebrochen.")
            return
        
        print("Login-Bestätigung erhalten, speichere Cookies...")
        
        # Speichere Cookies direkt nach Bestätigung
        if save_cookies(driver):
            messagebox.showinfo("Erfolg!", 
                "Login erfolgreich abgeschlossen!\n\n"
                "Cookies wurden gespeichert und sind ca. 12 Stunden gültig.\n"
                "Du kannst jetzt Bestellungen suchen.")
            print("Cookie-Speicherung erfolgreich")
        else:
            messagebox.showwarning("Teilweise erfolgreich", 
                "Login war erfolgreich, aber Cookies konnten nicht gespeichert werden.\n"
                "Du musst dich beim nächsten Mal erneut einloggen.")
            print("Cookie-Speicherung fehlgeschlagen")
        
    except Exception as e:
        error_msg = f"Kritischer Fehler beim Login: {str(e)}"
        print(error_msg)
        messagebox.showerror("Fehler", error_msg)
        
    finally:
        # Browser nur schließen wenn der Benutzer es bestätigt hat
        if driver:
            try:
                print("Schließe Browser...")
                driver.quit()
                print("Browser erfolgreich geschlossen")
            except Exception as e:
                print(f"Fehler beim Schließen des Browsers: {e}")

# === Cookie-Status prüfen ===
def check_cookie_status():
    if not os.path.exists(COOKIE_FILE):
        messagebox.showinfo("Cookie-Status", "❌ Keine Cookies gespeichert.\n\nBitte logge dich zuerst manuell ein.")
        return
    
    try:
        # Lade Cookie-Datei
        with open(COOKIE_FILE, "rb") as file:
            cookies = pickle.load(file)
        
        session_info = load_session_info()
        
        if session_info:
            saved_time = datetime.fromisoformat(session_info["timestamp"])
            time_diff = datetime.now() - saved_time
            
            hours_old = time_diff.total_seconds() / 3600
            
            status = f"✅ Cookies gefunden!\n\n"
            status += f"Anzahl: {len(cookies)} Cookies\n"
            status += f"Gespeichert: {saved_time.strftime('%d.%m.%Y um %H:%M:%S')}\n"
            status += f"Alter: {int(hours_old)} Stunden\n\n"
            
            if time_diff > timedelta(hours=12):
                status += "⚠️ Status: Abgelaufen\n\n(Cookies sind älter als 12 Stunden)"
            else:
                remaining_hours = 12 - int(hours_old)
                status += f"✅ Status: Gültig\n\n(Noch {remaining_hours} Stunden gültig)"
            
            messagebox.showinfo("Cookie-Status", status)
        else:
            messagebox.showinfo("Cookie-Status", 
                f"⚠️ Cookies gefunden ({len(cookies)} Stück)\n\n"
                "Aber keine Session-Info vorhanden.\n"
                "Eventuell solltest du dich neu einloggen.")
    
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Prüfen der Cookies:\n{str(e)}")

# === Verarbeite heruntergeladene ZIP-Datei ===
def process_downloaded_zip(order_number):
    # Warte auf den Download
    time.sleep(5)
    
    # Finde die neueste ZIP-Datei im DOWNLOAD_DIR (nicht im temporären Verzeichnis)
    zip_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.zip')]
    if not zip_files:
        messagebox.showerror("Fehler", f"Keine ZIP-Datei gefunden in: {DOWNLOAD_DIR}")
        return
    
    latest_zip = max(zip_files, key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)))
    zip_path = os.path.join(DOWNLOAD_DIR, latest_zip)
    
    # Erstelle einen Ordner für die entpackten Dateien im DOWNLOAD_DIR
    extract_dir = os.path.join(DOWNLOAD_DIR, order_number)
    
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        print(f"Dateien entpackt nach: {extract_dir}")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Entpacken der ZIP-Datei: {e}\nPfad: {zip_path}")
        return
    
    # Verarbeite die Dateien
    tiff_path = process_files(extract_dir, order_number)
    
    if tiff_path and os.path.exists(tiff_path):
        messagebox.showinfo("Erfolg", 
            f"TIFF-Datei erfolgreich erstellt:\n{tiff_path}\n\n"
            "Die Datei befindet sich im Unterordner 'amazon_order_downloads'.")
        os.startfile(os.path.dirname(tiff_path))
    else:
        messagebox.showerror("Fehler", "TIFF-Datei konnte nicht erstellt werden.")
    
    # Lösche die ZIP-Datei nach erfolgreicher Verarbeitung
    try:
        os.remove(zip_path)
        print(f"ZIP-Datei gelöscht: {zip_path}")
    except Exception as e:
        print(f"Warnung: ZIP-Datei konnte nicht gelöscht werden: {e}")

def process_files(extract_dir, order_number):
    """Verarbeite die SVG und JPG Dateien zu TIFF"""
    try:
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
            return None
            
        if not jpg_files:
            messagebox.showerror("Fehler", "Keine JPG-Dateien in der ZIP-Datei gefunden.")
            return None
        
        largest_jpg = max(jpg_files, key=lambda f: os.path.getsize(f))
        
        # 1. SVG modifizieren mit dem eingebetteten Bild
        new_svg_path = embed_image_in_svg(largest_jpg, svg_file)
        if not new_svg_path:
            return None
        
        # 2. Konvertiere zu TIFF (im gleichen Verzeichnis)
        output_path = os.path.join(extract_dir, f"{order_number}.tiff")
        if convert_svg_to_tiff(new_svg_path, output_path):
            return output_path
        return None
        
    except Exception as e:
        messagebox.showerror("Fehler", f"Verarbeitung fehlgeschlagen: {str(e)}")
        return None

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

        # 2. Öffne das PNG und entferne Transparenz durch Cropping
        img = Image.open(png_path)
        
        if img.mode in ('RGBA', 'LA'):
            # Finde die Bounding Box des nicht-transparenten Bereichs
            bbox = img.getbbox()
            
            if bbox:
                # Schneide das Bild auf den nicht-transparenten Bereich zu
                img = img.crop(bbox)
            else:
                messagebox.showerror("Fehler", "Das Bild enthält nur transparente Pixel")
                return False

        # 3. Konvertiere PNG zu TIFF
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

        # Prüfen ob Account-Auswahlfenster erscheint (Deutschland auswählen)
        try:
            # Warte auf das Account-Auswahlfenster
            germany_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'button.full-page-account-switcher-account-details span.full-page-account-switcher-account-label'))
            )

            # Alternative Suche falls das erste nicht funktioniert
            if "Deutschland" not in germany_button.text:
                buttons = driver.find_elements(By.CSS_SELECTOR, 'button.full-page-account-switcher-account-details')
                for btn in buttons:
                    if "Deutschland" in btn.text:
                        germany_button = btn
                        break
    
            # Prüfe ob der Button "Deutschland" enthält
            if "Deutschland" in germany_button.text:
                germany_button.click()
                print("Deutschland Account ausgewählt")
                time.sleep(2)  # Warte bis die Auswahl wirksam wird

                # Warte auf und klicke den Bestätigungsbutton
            try:
                confirm_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.kat-button--primary.full-page-account-switcher-buttons'))
                )
                confirm_button.click()
                print("Konto-Auswahl bestätigt")
                time.sleep(2)  # Warte bis die Auswahl wirksam wird
            except Exception as e:
                print(f"Bestätigungsbutton nicht gefunden: {str(e)}")
        except Exception as e:
            # Falls das Fenster nicht erscheint, einfach fortfahren
            print(f"Kein Account-Auswahlfenster gefunden oder Fehler: {str(e)}")
            pass
        
        #Suchfeld schneller finden mit explizitem Wait
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
