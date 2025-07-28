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
import sys

# === Globale Variablen ===

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
COOKIE_FILE = os.path.join(BASE_DIR, "amazon_cookies.pkl")
SESSION_FILE = os.path.join(BASE_DIR, "amazon_session_info.json")
LOGIN_URL = "https://sellercentral.amazon.de"
SEARCH_FIELD_SELECTOR = 'input#sc-search-field.search-input.search-input-active'
SEARCH_BUTTON_SELECTOR = 'button.sc-search-button.search-icon-container'
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# Ermittle das Verzeichnis, in dem das Skript liegt
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "amazon_order_downloads")

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
            "Cookies wurden gespeichert und bleiben gültig, bis sie vom Server abgelehnt werden.\n"
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
            status += "✅ Status: Gültig (kein Ablaufdatum)"
            
            messagebox.showinfo("Cookie-Status", status)
        else:
            messagebox.showinfo("Cookie-Status", 
                f"⚠️ Cookies gefunden ({len(cookies)} Stück)\n\n"
                "Aber keine Session-Info vorhanden.\n"
                "Die Cookies sollten trotzdem funktionieren.")
    
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Prüfen der Cookies:\n{str(e)}")

# === Warte auf Download-Vollendung ===
def wait_for_download_completion(download_dir, timeout=30):
    """Warte bis der Download vollständig ist"""
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        # Prüfe ob .crdownload Dateien vorhanden sind (unvollständige Downloads)
        partial_files = [f for f in os.listdir(download_dir) if f.endswith('.crdownload')]
        if not partial_files:
            # Prüfe ob mindestens eine ZIP-Datei vorhanden ist
            zip_files = [f for f in os.listdir(download_dir) if f.endswith('.zip')]
            if zip_files:
                # Zusätzliche Wartezeit für Datei-Stabilität
                time.sleep(2)
                return True
        
        time.sleep(1)
    
    return False

# === Verbesserte Funktion zur Verarbeitung der ZIP-Datei ===
def process_downloaded_zip(order_number):
    """Verarbeite die heruntergeladene ZIP-Datei"""
    print(f"=== Starte Verarbeitung für Bestellung: {order_number} ===")
    
    # Warte auf Download-Vollendung
    if not wait_for_download_completion(DOWNLOAD_DIR, timeout=30):
        messagebox.showerror("Fehler", "Download wurde nicht rechtzeitig abgeschlossen.")
        return
    
    # Finde die neueste ZIP-Datei
    zip_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.zip')]
    if not zip_files:
        messagebox.showerror("Fehler", f"Keine ZIP-Datei gefunden in: {DOWNLOAD_DIR}")
        return
    
    latest_zip = max(zip_files, key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)))
    zip_path = os.path.join(DOWNLOAD_DIR, latest_zip)
    
    print(f"Gefundene ZIP-Datei: {zip_path}")
    
    # Erstelle Ordner für entpackte Dateien
    extract_dir = os.path.join(DOWNLOAD_DIR, order_number)
    
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    try:
        # Entpacke ZIP-Datei
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        print(f"Dateien entpackt nach: {extract_dir}")
        
        # Liste alle entpackten Dateien
        print("Entpackte Dateien:")
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                print(f"  - {file_path}")
        
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Entpacken der ZIP-Datei: {e}")
        return
    
    # Verarbeite die Dateien zu TIFF
    tiff_path = process_files_to_tiff(extract_dir, order_number)
    
    if tiff_path and os.path.exists(tiff_path):
        messagebox.showinfo("Erfolg", 
            f"TIFF-Datei erfolgreich erstellt:\n{tiff_path}\n\n"
            "Die Datei befindet sich im 'amazon_order_downloads' Ordner.")
        
        # Öffne den Ordner mit der TIFF-Datei
        try:
            os.startfile(os.path.dirname(tiff_path))
        except:
            pass  # Falls startfile nicht verfügbar ist
    else:
        messagebox.showerror("Fehler", "TIFF-Datei konnte nicht erstellt werden.")
    
    # Lösche die ZIP-Datei nach erfolgreicher Verarbeitung
    try:
        os.remove(zip_path)
        print(f"ZIP-Datei gelöscht: {zip_path}")
    except Exception as e:
        print(f"Warnung: ZIP-Datei konnte nicht gelöscht werden: {e}")


def extract_dimensions_and_check_text(extract_dir):
    """Sucht nach JSON-Datei, extrahiert Druckdimensionen und prüft Verkäufertext"""
    try:
        # Suche nach JSON-Dateien
        json_files = [f for f in os.listdir(extract_dir) if f.lower().endswith('.json')]
        if not json_files:
            messagebox.showerror("Fehler", "Keine JSON-Datei im Download gefunden")
            return None
        
        json_path = os.path.join(extract_dir, json_files[0])
        print(f"Verarbeite JSON-Datei: {json_path}")

        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 1. Prüfe Verkäufertext zuerst
        seller_message = None
        if 'customizationInfo' in data:
            for surface in data['customizationInfo'].get('version3.0', {}).get('surfaces', []):
                for area in surface.get('areas', []):
                    if area.get('customizationType') == "TextPrinting" and area.get('label') == "Verkäufer nachricht":
                        if area.get('text', '').strip():
                            seller_message = area['text'].strip()
                            messagebox.showinfo(
                                "Verkäuferhinweis", 
                                f"Nachricht vom Verkäufer:\n\n{seller_message}"
                            )
        
        # 2. Suche Druckdimensionen (kritischer Teil)
        required_dimensions = None
        
        # Zuerst in customizationData suchen
        if 'customizationData' in data:
            for child in data['customizationData'].get('children', []):
                for subchild in child.get('children', []):
                    for item in subchild.get('children', []):
                        if item.get('type') == "PlacementContainerCustomization":
                            dims = item.get('dimension', {})
                            if 'width' in dims and 'height' in dims:
                                required_dimensions = {
                                    'width': dims['width'],
                                    'height': dims['height'],
                                    'ratio': dims['width'] / dims['height']
                                }
                                print(f"Gefundene Druckdimensionen: {dims['width']}x{dims['height']}")
                                return required_dimensions
        
        # Falls nicht gefunden, in customizationInfo suchen
        if not required_dimensions and 'customizationInfo' in data:
            for surface in data['customizationInfo'].get('version3.0', {}).get('surfaces', []):
                for area in surface.get('areas', []):
                    if area.get('customizationType') == "ImagePrinting":
                        dims = area.get('Dimensions', {})
                        if 'width' in dims and 'height' in dims:
                            required_dimensions = {
                                'width': dims['width'],
                                'height': dims['height'],
                                'ratio': dims['width'] / dims['height']
                            }
                            print(f"Gefundene Druckdimensionen (ImagePrinting): {dims['width']}x{dims['height']}")
                            return required_dimensions
        
        if not required_dimensions:
            messagebox.showerror(
                "Fehler", 
                "Konnte Druckdimensionen nicht ermitteln.\n"
                "Die Verarbeitung wird abgebrochen."
            )
            return None
        
        return required_dimensions
        
    except Exception as e:
        messagebox.showerror("Fehler", f"JSON-Verarbeitung fehlgeschlagen: {str(e)}")
        return None

def check_and_correct_aspect_ratio(tiff_path, target_ratio, tolerance=0.01):
    """Überprüft und korrigiert das Bildverhältnis der TIFF-Datei"""
    try:
        img = Image.open(tiff_path)
        current_width, current_height = img.size
        current_ratio = current_width / current_height
        
        print(f"Aktuelle Dimensionen: {current_width}x{current_height}, Verhältnis: {current_ratio:.4f}")
        print(f"Ziel-Verhältnis: {target_ratio:.4f}")
        
        # Prüfe ob Korrektur notwendig ist
        ratio_diff = abs(current_ratio - target_ratio)
        if ratio_diff <= tolerance:
            print("Bildverhältnis ist bereits korrekt")
            return True
        
        print(f"Korrektur notwendig (Abweichung: {ratio_diff:.4f})")
        
        # Berechne neue Dimensionen
        if current_ratio > target_ratio:  # Zu breit
            new_width = int(current_height * target_ratio)
            new_height = current_height
        else:  # Zu hoch
            new_width = current_width
            new_height = int(current_width / target_ratio)
        
        print(f"Neue Dimensionen: {new_width}x{new_height}")
        
        # Führe die Korrektur durch
        corrected_img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Backup der Originaldatei
        backup_path = tiff_path.replace('.tiff', '_original.tiff')
        shutil.copy2(tiff_path, backup_path)
        
        # Speichere korrigierte Version
        corrected_img.save(tiff_path, format="TIFF", compression="tiff_deflate")
        
        # Verifiziere das Ergebnis
        verify_img = Image.open(tiff_path)
        verify_ratio = verify_img.width / verify_img.height
        final_diff = abs(verify_ratio - target_ratio)
        
        if final_diff <= tolerance:
            print("Erfolgreich korrigiert!")
            return True
        else:
            print(f"Warnung: Restabweichung {final_diff:.4f}")
            return False
            
    except Exception as e:
        print(f"Fehler bei der Bildkorrektur: {e}")
        return False

        

# Aktualisierte process_files_to_tiff Funktion
def process_files_to_tiff(extract_dir, order_number):
    """Verarbeite SVG und JPG Dateien zu TIFF mit Verhältniskontrolle"""
    try:
        print(f"=== Starte Dateiverarbeitung für {order_number} ===")
        
        # 1. Finde Dateien
        svg_file = None
        jpg_files = []
        
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                if file.lower().endswith('.svg'):
                    svg_file = file_path
                elif file.lower().endswith(('.jpg', '.jpeg')):
                    jpg_files.append(file_path)
        
        if not svg_file or not jpg_files:
            messagebox.showerror("Fehler", "SVG oder JPG Dateien fehlen")
            return None
        
        # 2. Extrahiere Dimensionen
        dimensions = extract_dimensions_and_check_text(extract_dir)
        if not dimensions:  # Wenn keine Dimensionen gefunden wurden
            return None  # Verarbeitung abbrechen
        
        # 3. Verarbeite Bild
        largest_jpg = max(jpg_files, key=lambda f: os.path.getsize(f))
        modified_svg = embed_image_in_svg(largest_jpg, svg_file)
        if not modified_svg:
            return None
        
        # 4. Konvertiere zu TIFF
        output_path = os.path.join(extract_dir, f"{order_number}.tiff")
        if not convert_svg_to_tiff(modified_svg, output_path):
            return None
        
        # 5. Verhältniskontrolle
        if dimensions and 'ratio' in dimensions:
            print("Führe Verhältniskontrolle durch...")
            if not check_and_correct_aspect_ratio(output_path, dimensions['ratio']):
                messagebox.showwarning("Warnung", "Bildverhältnis konnte nicht perfekt korrigiert werden")
        
        return output_path
        
    except Exception as e:
        messagebox.showerror("Fehler", f"Verarbeitung fehlgeschlagen: {str(e)}")
        return None

def embed_image_in_svg(image_path, svg_path):
    """Ersetzt das Bild innerhalb des clipPath"""
    try:
        print(f"=== Bette Bild ein: {image_path} in {svg_path} ===")
        
        # Lese SVG-Datei
        with open(svg_path, "rb") as f:
            tree = etree.parse(f)
            root = tree.getroot()

        namespaces = {
            'svg': 'http://www.w3.org/2000/svg',
            'xlink': 'http://www.w3.org/1999/xlink'
        }

        # Finde das Bild-Element innerhalb des clipPath
        clip_path_images = root.xpath('//svg:g[@clip-path]//svg:image', namespaces=namespaces)
        
        if not clip_path_images:
            print("Kein Bild im clipPath gefunden, suche nach allen Bildern...")
            # Fallback: Suche nach allen Bildern
            all_images = root.xpath('//svg:image', namespaces=namespaces)
            if all_images:
                clip_path_images = all_images[:1]  # Nimm das erste Bild
            else:
                raise Exception("Keine Bild-Elemente in der SVG gefunden")
        
        # Das Ziel-Bild (normalerweise das große Bild)
        target_image = clip_path_images[0]
        print(f"Ziel-Bild-Element gefunden: {target_image.get('width')}x{target_image.get('height')}")

        # Lade und kodiere das Bild
        with open(image_path, "rb") as img_file:
            encoded_string = "data:image/jpeg;base64," + base64.b64encode(img_file.read()).decode('utf-8')

        # Setze das neue Bild
        target_image.set("{http://www.w3.org/1999/xlink}href", encoded_string)
        print("Bild erfolgreich eingebettet")

        # Speichere die modifizierte SVG
        new_svg_path = os.path.splitext(svg_path)[0] + "_modified.svg"
        with open(new_svg_path, "wb") as new_svg:
            new_svg.write(etree.tostring(tree, pretty_print=True, encoding="UTF-8"))

        print(f"Modifizierte SVG gespeichert: {new_svg_path}")
        return new_svg_path

    except Exception as e:
        print(f"Fehler bei der Bildeinbettung: {e}")
        messagebox.showerror("Fehler", f"Bildeinbettung fehlgeschlagen: {str(e)}")
        return None

def convert_svg_to_tiff(svg_path, output_path):
    """Konvertiere SVG zu TIFF"""
    try:
        print(f"=== Konvertiere SVG zu TIFF: {svg_path} -> {output_path} ===")
        
        # Temporärer PNG-Pfad
        png_path = os.path.splitext(output_path)[0] + "_temp.png"
        
        # Konvertiere SVG zu PNG
        cairosvg.svg2png(
            url=svg_path,
            write_to=png_path,
            background_color=None,
            scale=1.0
        )
        print("SVG zu PNG konvertiert")

        # Öffne PNG und verarbeite es
        img = Image.open(png_path)
        print(f"PNG geöffnet: {img.size}, Modus: {img.mode}")
        
        # Entferne Transparenz durch Cropping
        if img.mode in ('RGBA', 'LA'):
            bbox = img.getbbox()
            if bbox:
                img = img.crop(bbox)
                print(f"Bild beschnitten auf: {img.size}")
            else:
                print("Warnung: Bild enthält nur transparente Pixel")

        # Speichere als TIFF
        img.save(output_path, format="TIFF", compression="tiff_deflate")
        print(f"TIFF gespeichert: {output_path}")
        
        # Lösche temporäre Dateien
        try:
            os.remove(png_path)
            os.remove(svg_path)  # Lösche die modifizierte SVG
            print("Temporäre Dateien gelöscht")
        except Exception as e:
            print(f"Warnung: Temporäre Dateien konnten nicht gelöscht werden: {e}")

        return True

    except Exception as e:
        print(f"Fehler bei der TIFF-Konvertierung: {e}")
        messagebox.showerror("Fehler", f"Konvertierung fehlgeschlagen: {str(e)}")
        return False

# === Bestellung suchen und verarbeiten ===
def search_order(order_number):
    driver = create_driver()
    
    try:
        print(f"=== Starte Suche nach Bestellung: {order_number} ===")
        
        # Cookies laden und prüfen
        if not load_cookies(driver):
            messagebox.showerror("Fehler", "Keine gültigen Cookies gefunden. Bitte logge dich zuerst manuell ein.")
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
            return

        # Prüfen auf Account-Auswahlfenster (Deutschland auswählen)
        try:
            print("Prüfe auf Account-Auswahlfenster...")
            # Warte auf Account-Auswahlfenster mit verschiedenen Selektoren
            germany_selectors = [
                'button.full-page-account-switcher-account-details span.full-page-account-switcher-account-label',
                'button.full-page-account-switcher-account-details',
                '[data-testid="account-switcher-account-details"]',
                'button[class*="account-switcher"]'
            ]
            
            germany_button = None
            for selector in germany_selectors:
                try:
                    elements = WebDriverWait(driver, 5).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector))
                    )
                    
                    for element in elements:
                        if "Deutschland" in element.text or "Germany" in element.text:
                            germany_button = element
                            break
                    
                    if germany_button:
                        break
                        
                except:
                    continue
            
            if germany_button:
                print("Deutschland Account gefunden, klicke darauf...")
                germany_button.click()
                time.sleep(2)
                
                
            # Suche nach Bestätigungsbutton mit verschiedenen Selektoren (inkl. kat-button)
            confirm_selectors = [
                'kat-button[data-test="confirm-selection"]',
                'kat-button.full-page-account-switcher-button',
                'button.kat-button.full-page-account-switcher-button',
                'button[class*="full-page-account-switcher-button"]',
                'button[class*="account-switcher-button"]',
                'button.kat-button'
            ]

            confirm_button = None
            for selector in confirm_selectors:
                try:
                    confirm_button = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    if confirm_button:
                        print(f"Bestätigungsbutton gefunden mit Selektor: {selector}")
                        # Versuche JavaScript-Klick (robuster bei Custom Elements)
                        driver.execute_script("arguments[0].click();", confirm_button)
                        time.sleep(2)
                        break
                except Exception as e:
                    print(f"Fehler mit Selektor {selector}: {e}")
                    continue

            if not confirm_button:
                print("Bestätigungsbutton nicht gefunden, versuche trotzdem fortzufahren...")
            else:
                print("Bestätigungsbutton wurde erfolgreich geklickt.")

        except Exception as e:
            print(f"Account-Auswahl übersprungen: {str(e)}")
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
    tk.Label(window, text="Hinweis: Cookies bleiben gültig bis sie ablaufen", font=("Arial", 9), fg="gray").pack(pady=(20, 5))
    
    # Bestellnummer eingeben
    tk.Label(window, text="Bestellnummer eingeben:", font=("Arial", 12)).pack(pady=5)
    order_entry = tk.Entry(window, font=("Arial", 12), width=35)
    order_entry.pack(pady=5)
    
    # Fokus auf das Eingabefeld setzen
    order_entry.focus_set()

    def handle_search(order_number=None):
        if order_number is None:
            order_number = order_entry.get().strip()
        if order_number:
            search_order(order_number)
        else:
            messagebox.showwarning("Hinweis", "Bitte eine Bestellnummer eingeben.")

    def on_barcode_input(event):
        # Der Barcode-Scanner sendet die Daten + Enter
        # Wir nehmen den aktuellen Inhalt des Feldes (ohne den letzten Zeilenumbruch)
        barcode = order_entry.get().strip()
        if barcode:
            handle_search(barcode)

    # Enter-Taste und Barcode-Eingabe binden
    order_entry.bind('<Return>', on_barcode_input)
    
    # Buttons
    tk.Button(window, text="Bestellung suchen & verarbeiten", 
              command=lambda: handle_search(), 
              font=("Arial", 12), bg="#FF9900", fg="black", width=25).pack(pady=10)
    
    tk.Button(window, text="Manuell einloggen & Cookies speichern", command=manual_login,
              font=("Arial", 10), width=35).pack(pady=5)
    
    tk.Button(window, text="Cookie-Status prüfen", command=check_cookie_status,
              font=("Arial", 10), width=35).pack(pady=5)

    # Hinweis
    tk.Label(window, text="Hinweis: Cookies sind ca. 12 Stunden gültig", 
             font=("Arial", 9), fg="gray").pack(pady=(20, 5))
    
    # Automatisch nach Barcode-Eingabe suchen
    window.mainloop()

# === Programmstart ===
if __name__ == "__main__":
    start_gui()
