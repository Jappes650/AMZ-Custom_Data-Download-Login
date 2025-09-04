import os
import time
import pickle
import json
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
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
import threading

# Thread-sichere Messagebox Wrapper
def safe_messagebox(func, *args, **kwargs):
    import tkinter as tk
    root = tk._default_root
    if root:
        root.after(0, lambda: func(*args, **kwargs))
    else:
        func(*args, **kwargs)

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

# === NEUE FUNKTIONEN: Heizungstyp-Erkennung ===

def create_default_config():
    """Erstellt Standard-Konfigurationsdatei falls sie nicht existiert"""
    config = {
        "heating_panels": {
            "130W Standard": {
                "width": 500,
                "height": 380,
                "tolerance": 0.01,
                "watt": 130,
                "description": "Kleine Infrarotheizung"
            },
            "300W Standard": {
                "width": 600,
                "height": 500,
                "tolerance": 0.01,
                "watt": 300,
                "description": "Mittlere Infrarotheizung"
            },
            "450W Standard": {
                "width": 900,
                "height": 500,
                "tolerance": 0.01,
                "watt": 450,
                "description": "Mittlere Infrarotheizung"
            },
            "600W / 800W Standard": {
                "width": 1000,
                "height": 600,
                "tolerance": 0.01,
                "watt": "600 - 800",
                "description": "Große Infrarotheizung"
            },
            "1000W / 1200W Standard": {
                "width": 1200,
                "height": 600,
                "tolerance": 0.01,
                "watt": "1000 - 1200",
                "description": "Große Infrarotheizung"
            }
        },
        "quality_settings": {
            "min_dpi": 150,
            "tiff_compression": "tiff_lzw",
            "max_ratio_deviation": 0.05
        }
    }
    return config

def load_config():
    """Lädt Konfiguration aus Datei oder erstellt Standard-Config"""
    config_path = os.path.join(BASE_DIR, "heating_config.json")
    
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Fehler beim Laden der Config: {e}")
            safe_messagebox(messagebox.showwarning, "Config-Fehler", 
                "Fehler beim Laden der Konfiguration. Verwende Standard-Einstellungen.")
    
    # Erstelle Standard-Config
    config = create_default_config()
    save_config(config)
    return config

def save_config(config):
    """Speichert Konfiguration in Datei"""
    config_path = os.path.join(BASE_DIR, "heating_config.json")
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        print(f"Konfiguration gespeichert: {config_path}")
    except Exception as e:
        print(f"Fehler beim Speichern der Config: {e}")

def ask_yes_no_safe(title, message):
    """Thread-sicheres askyesno mit Rückgabe"""
    result = {"value": False}
    done = tk.BooleanVar()  # Synchronisations-Flag

    def _ask():
        try:
            res = messagebox.askyesno(title, message)
            result["value"] = res
        finally:
            done.set(True)  # Signal: Dialog geschlossen

    root = tk._default_root
    if root:
        root.after(0, _ask)
        root.wait_variable(done)  # Warten, bis _ask fertig ist
    else:
        result["value"] = messagebox.askyesno(title, message)

    return result["value"]


def detect_heating_type(dimensions):
    """
    Erkennt Heizungstyp anhand der Dimensionen
    
    Args:
        dimensions (dict): Dictionary mit 'width', 'height', 'ratio'
    
    Returns:
        tuple: (heating_type_name, heating_specs) oder ("Unbekannt", None)
    """
    try:
        config = load_config()
        
        print(f"=== Heizungstyp-Erkennung ===")
        print(f"Bild-Verhältnis: {dimensions['ratio']:.4f}")
        
        best_match = None
        best_deviation = float('inf')
        
        # Durchsuche alle konfigurierten Heizungstypen
        for heating_type, specs in config["heating_panels"].items():
            target_ratio = specs["width"] / specs["height"]
            deviation = abs(dimensions["ratio"] - target_ratio)
            
            print(f"Prüfe {heating_type}:")
            print(f"  - Ziel-Verhältnis: {target_ratio:.4f}")
            print(f"  - Abweichung: {deviation:.4f}")
            print(f"  - Toleranz: {specs['tolerance']:.4f}")
            
            # Prüfe ob innerhalb der Toleranz und besser als bisheriger Match
            if deviation <= specs["tolerance"] and deviation < best_deviation:
                best_match = (heating_type, specs)
                best_deviation = deviation
                print(f"  ✓ Neuer bester Match!")
            else:
                print(f"  ✗ Außerhalb Toleranz oder schlechter Match")
        
        if best_match:
            heating_type, specs = best_match
            print(f"\n🎯 ERKANNT: {heating_type}")
            print(f"   Größe: {specs['width']}x{specs['height']}mm")
            print(f"   Leistung: {specs['watt']}W")
            print(f"   Beschreibung: {specs['description']}")
            return best_match
        else:
            print(f"\n❌ KEIN MATCH: Keine passende Heizung gefunden")
            return "Unbekannt", None
            
    except Exception as e:
        print(f"Fehler bei Heizungstyp-Erkennung: {e}")
        return "Fehler", None

def validate_heating_match(heating_type, specs, dimensions, show_dialog=True):
    """
    Validiert die Heizungstyp-Erkennung und zeigt Bestätigung
    
    Args:
        heating_type (str): Name des erkannten Heizungstyps
        specs (dict): Spezifikationen des Heizungstyps
        dimensions (dict): Bild-Dimensionen
        show_dialog (bool): Ob Bestätigungsdialog gezeigt werden soll
    
    Returns:
        bool: True wenn Benutzer bestätigt oder kein Dialog
    """
    if heating_type == "Unbekannt":
        if show_dialog:
            safe_messagebox(messagebox.showwarning, 
                "Heizungstyp unbekannt",
                f"Bildverhältnis: {dimensions['ratio']:.4f}\n\n"
                "Kein passender Heizungstyp gefunden!\n"
                "Bitte prüfe die Maske bei Amazon oder erweitere die Konfiguration."
            )
        return False
    
    if heating_type == "Fehler":
        if show_dialog:
            safe_messagebox(messagebox.showerror, "Fehler", "Fehler bei der Heizungstyp-Erkennung!")
        return False
    
    # Zeige Bestätigung
    if show_dialog:
        target_ratio = specs["width"] / specs["height"]
        deviation = abs(dimensions["ratio"] - target_ratio)
        
        message = f"🎯 ERKANNTER HEIZUNGSTYP:\n\n"
        message += f"Typ: {heating_type}\n"
        message += f"Größe: {specs['width']} x {specs['height']} mm\n"
        message += f"Leistung: {specs['watt']} Watt\n"
        message += f"Beschreibung: {specs['description']}\n\n"
        message += f"TECHNISCHE DETAILS:\n"
        message += f"Bild-Verhältnis: {dimensions['ratio']:.4f}\n"
        message += f"Ziel-Verhältnis: {target_ratio:.4f}\n"
        message += f"Abweichung: {deviation:.4f}\n"
        message += f"Toleranz: {specs['tolerance']:.4f}\n\n"
        message += "Soll die Verarbeitung fortgesetzt werden?"
        
        result = ask_yes_no_safe("Heizungstyp bestätigen", message)
        return result
    
    return True

def get_heating_recommendations(dimensions):
    """
    Gibt Empfehlungen für ähnliche Heizungstypen wenn kein exakter Match
    
    Args:
        dimensions (dict): Bild-Dimensionen
    
    Returns:
        list: Liste von (heating_type, deviation) Tupeln, sortiert nach Abweichung
    """
    try:
        config = load_config()
        recommendations = []
        
        for heating_type, specs in config["heating_panels"].items():
            target_ratio = specs["width"] / specs["height"]
            deviation = abs(dimensions["ratio"] - target_ratio)
            recommendations.append((heating_type, specs, deviation))
        
        # Sortiere nach Abweichung (beste zuerst)
        recommendations.sort(key=lambda x: x[2])
        
        return recommendations[:3]  # Top 3 Empfehlungen
        
    except Exception as e:
        print(f"Fehler bei Empfehlungen: {e}")
        return []

def edit_heating_config():
    """Öffnet Fenster zum Bearbeiten der Heizungstypen"""
    config_window = tk.Toplevel()
    config_window.title("Heizungstypen konfigurieren")
    config_window.geometry("600x400")
    
    # Zeige aktuelle Konfiguration
    config = load_config()
    
    tk.Label(config_window, 
             text="AKTUELLE HEIZUNGSTYPEN:", 
             font=("Arial", 12, "bold")).pack(pady=(20, 10))
    
    # Scrollbare Liste der Heizungstypen
    frame = tk.Frame(config_window)
    frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
    
    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, font=("Courier", 9))
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)
    
    # Fülle Liste mit Heizungstypen
    for heating_type, specs in config["heating_panels"].items():
        ratio = specs["width"] / specs["height"]
        text = f"{heating_type:<20} {specs['width']}x{specs['height']}mm {specs['watt']}W (Ratio: {ratio:.3f})"
        listbox.insert(tk.END, text)
    
    # Config-Datei Pfad
    config_path = os.path.join(BASE_DIR, "heating_config.json")
    
    tk.Label(config_window, 
             text=f"Konfigurationsdatei: {config_path}", 
             font=("Arial", 8), fg="gray").pack(pady=5)
    
    def open_config_folder():
        try:
            os.startfile(BASE_DIR)
        except:
            ("Pfad", f"Öffne diesen Ordner:\n{BASE_DIR}")
    
    tk.Button(config_window, 
              text="Ordner öffnen", 
              command=open_config_folder,
              font=("Arial", 10)).pack(pady=10)
    
    tk.Label(config_window, 
             text="Bearbeite die 'heating_config.json' Datei mit einem Texteditor\n"
                  "und starte das Programm neu um Änderungen zu übernehmen.", 
             font=("Arial", 9)).pack(pady=10)

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
        safe_messagebox(messagebox.showinfo, 
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
            safe_messagebox(messagebox.showinfo, "Abgebrochen", "Login-Prozess wurde abgebrochen.")
            return
        
        print("Login-Bestätigung erhalten, speichere Cookies...")
        
        # Speichere Cookies direkt nach Bestätigung
        if save_cookies(driver):
            safe_messagebox(messagebox.showinfo, "Erfolg!", 
            "Login erfolgreich abgeschlossen!\n\n"
            "Cookies wurden gespeichert und bleiben gültig, bis sie vom Server abgelehnt werden.\n"
            "Du kannst jetzt Bestellungen suchen.")
            print("Cookie-Speicherung erfolgreich")
        else:
            safe_messagebox(messagebox.showwarning, "Teilweise erfolgreich", 
                "Login war erfolgreich, aber Cookies konnten nicht gespeichert werden.\n"
                "Du musst dich beim nächsten Mal erneut einloggen.")
            print("Cookie-Speicherung fehlgeschlagen")
        
    except Exception as e:
        error_msg = f"Kritischer Fehler beim Login: {str(e)}"
        print(error_msg)
        safe_messagebox(messagebox.showerror, "Fehler", error_msg)
        
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
        safe_messagebox(messagebox.showinfo, "Cookie-Status", "❌ Keine Cookies gespeichert.\n\nBitte logge dich zuerst manuell ein.")
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
            
            safe_messagebox(messagebox.showinfo, "Cookie-Status", status)
        else:
            safe_messagebox(messagebox.showinfo, "Cookie-Status", 
                f"⚠️ Cookies gefunden ({len(cookies)} Stück)\n\n"
                "Aber keine Session-Info vorhanden.\n"
                "Die Cookies sollten trotzdem funktionieren.")
    
    except Exception as e:
        safe_messagebox(messagebox.showerror, "Fehler", f"Fehler beim Prüfen der Cookies:\n{str(e)}")

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

# === Multi-Position Verarbeitung ===

def find_order_positions(driver):
    """
    Erkennt alle Positionen einer Bestellung und prüft welche Anpassungsinformationen haben.
    Optimiert: weniger sleep-Zeiten, robustes Verhalten bleibt.
    
    Returns:
        list: [{'position': 1, 'has_customization': True, 'element': WebElement, 'expander': WebElement}, ...]
    """
    try:
        print("=== Optimierte Suche nach Bestellpositionen (stabil & schnell) ===")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span.a-expander-prompt, .a-expander-prompt, [class*='expander-prompt']"))
        )

        positions = []

        # Suche nach Expandern
        expander_selectors = [
            "span.a-expander-prompt",
            ".a-expander-prompt",
            "[class*='expander-prompt']"
        ]
        
        expanders = []
        for selector in expander_selectors:
            found = driver.find_elements(By.CSS_SELECTOR, selector)
            if found:
                expanders = found
                print(f"Expanders gefunden mit Selektor: {selector}")
                break
        
        if not expanders:
            print("Keine Expander gefunden – möglicherweise nur eine Position.")
            # Einzelposition prüfen
            customization_links = driver.find_elements(By.CSS_SELECTOR, "a.a-link-normal[href*='fulfillment']")
            if customization_links:
                positions.append({
                    'position': 1,
                    'has_customization': True,
                    'element': customization_links[0],
                    'expander': None
                })
                print("Einzelne Position mit Anpassungsinformationen gefunden.")
            else:
                print("Keine Anpassungsinformationen gefunden.")
            return positions

        print(f"Gefundene Expander: {len(expanders)}")

        for i, expander in enumerate(expanders, 1):
            try:
                print(f"\n--- Prüfe Position {i} ---")
                driver.execute_script("arguments[0].scrollIntoView(true);", expander)
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable(expander)).click()
                
                # Suche nach Link innerhalb übergeordnetem Container
                parent = expander
                for _ in range(5):
                    parent = parent.find_element(By.XPATH, "./..")
                    customization_links = parent.find_elements(By.CSS_SELECTOR, "a.a-link-normal[href*='fulfillment']")
                    if customization_links:
                        positions.append({
                            'position': i,
                            'has_customization': True,
                            'element': customization_links[0],
                            'expander': expander
                        })
                        print(f"✅ Position {i}: Hat Anpassungsinformationen")
                        break
                else:
                    positions.append({
                        'position': i,
                        'has_customization': False,
                        'element': None,
                        'expander': expander
                    })
                    print(f"⚪ Position {i}: Keine Anpassungsinformationen")

                # Expander wieder schließen (optional, aber kann DOM sauber halten)
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable(expander)).click()
                except:
                    pass

            except Exception as e:
                print(f"⚠️ Fehler bei Position {i}: {e}")
                positions.append({
                    'position': i,
                    'has_customization': False,
                    'element': None,
                    'expander': expander
                })

        print(f"\n=== ZUSAMMENFASSUNG ===")
        print(f"Gesamtpositionen: {len(positions)}")
        print(f"Mit Anpassungsinformationen: {len([p for p in positions if p['has_customization']])}")

        return positions

    except Exception as e:
        print(f"❌ Fehler bei der Positionssuche: {e}")
        return []

def process_single_position(driver, position_info, order_number):
    """
    Verarbeitet eine einzelne Position
    
    Args:
        driver: WebDriver-Instanz
        position_info: Dictionary mit Position-Informationen
        order_number: Bestellnummer
    
    Returns:
        bool: True wenn erfolgreich verarbeitet
    """
    try:
        position_num = position_info['position']
        print(f"\n=== Verarbeite Position {position_num} ===")
        
        if not position_info['has_customization']:
            print(f"Position {position_num}: Keine Anpassungsinformationen - übersprungen")
            return True
        
        # Öffne die Position falls sie geschlossen ist
        if position_info['expander']:
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", position_info['expander'])
                time.sleep(1)
                position_info['expander'].click()
                time.sleep(2)
            except:
                print("Position bereits geöffnet oder Fehler beim Öffnen")
        
        # Klicke auf Anpassungsinformationen-Link
        customization_link = position_info['element']
        driver.execute_script("arguments[0].scrollIntoView(true);", customization_link)
        time.sleep(1)
        customization_link.click()
        print(f"Anpassungsinformationen für Position {position_num} geöffnet")
        time.sleep(3)
        
        # Warte und klicke Download-Button
        download_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "kat-button.download-zip-file-button"))
        )
        download_button.click()
        print(f"Download für Position {position_num} gestartet")
        
        # Verarbeite die ZIP-Datei
        position_order_number = f"{order_number}_pos{position_num}"
        success = process_downloaded_zip(position_order_number)
        
        if success:
            print(f"✅ Position {position_num} erfolgreich verarbeitet")
        else:
            print(f"❌ Position {position_num} Verarbeitung fehlgeschlagen")
        
        # NEU: Direkt zur Bestellübersicht zurück über Breadcrumb
        try:
            breadcrumb = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "kat-breadcrumb-item[label='Bestelldetails']"))
            )
            breadcrumb.click()
            print("Zurück zur Bestellübersicht über Breadcrumb")
            time.sleep(3)
            
            # Warte bis die Bestellübersicht wieder geladen ist
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "span.a-expander-prompt"))
            )
            
        except Exception as e:
            print(f"Warnung: Konnte nicht über Breadcrumb zurück: {e}")
            # Fallback: Browser zurück
            driver.back()
            time.sleep(3)
        
        return success
        
    except Exception as e:
        print(f"Fehler bei Position {position_num}: {e}")
        # Versuche zurück zur Bestellübersicht zu gehen
        try:
            driver.back()
            time.sleep(2)
        except:
            pass
        return False

def search_order_multi_position(order_number):
    """
    Erweiterte Bestellungssuche mit Multi-Position-Unterstützung
    """
    driver = create_driver()
    
    try:
        print(f"=== Starte Multi-Position-Suche für Bestellung: {order_number} ===")
        
        # Standard Login-Prozess
        if not load_cookies(driver):
            safe_messagebox(messagebox.showerror, "Fehler", "Keine gültigen Cookies gefunden. Bitte logge dich zuerst manuell ein.")
            return

        driver.refresh()
        time.sleep(3)
        
        current_url = driver.current_url
        print(f"URL nach Cookie-Login: {current_url}")
        
        if any(keyword in current_url.lower() for keyword in ["signin", "login", "auth"]):
            safe_messagebox(messagebox.showerror, "Session abgelaufen", "Deine Session ist abgelaufen. Bitte logge dich erneut ein.")
            return

        # Account-Auswahl handling (Deutschland)
        try:
            print("Prüfe auf Account-Auswahlfenster...")
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
                print("Deutschland Account gefunden")
                germany_button.click()
                time.sleep(2)
                
                # Bestätigungsbutton
                confirm_selectors = [
                    'kat-button[data-test="confirm-selection"]',
                    'kat-button.full-page-account-switcher-button',
                    'button.kat-button.full-page-account-switcher-button',
                    'button[class*="full-page-account-switcher-button"]',
                    'button[class*="account-switcher-button"]',
                    'button.kat-button'
                ]

                for selector in confirm_selectors:
                    try:
                        confirm_button = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                        if confirm_button:
                            driver.execute_script("arguments[0].click();", confirm_button)
                            time.sleep(2)
                            break
                    except:
                        continue

        except Exception as e:
            print(f"Account-Auswahl übersprungen: {str(e)}")
        
        # Suche nach der Bestellung
        search_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input#sc-search-field"))
        )
        print("Suchfeld gefunden")
        
        search_field.clear()
        search_field.send_keys(order_number)
        
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.sc-search-button.search-icon-container"))
        )
        search_button.click()
        print("Suche durchgeführt")
        
        # Prüfe auf "Nicht gefunden"
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.sc-no-results-message, span.a-expander-prompt"))
            )
            
            no_results = driver.find_elements(By.CSS_SELECTOR, "div.sc-no-results-message")
            if no_results:
                safe_messagebox(messagebox.showwarning, "Nicht gefunden", f"Bestellung {order_number} wurde nicht gefunden.")
                return
            
            # Erkenne alle Positionen
            positions = find_order_positions(driver)
            
            if not positions:
                safe_messagebox(messagebox.showwarning, "Keine Positionen", "Keine Bestellpositionen gefunden.")
                return
            
            # Zeige Übersicht der gefundenen Positionen
            customizable_positions = [p for p in positions if p['has_customization']]
            
            if not customizable_positions:
                safe_messagebox(messagebox.showinfo, "Keine Anpassungen", 
                    f"Bestellung {order_number} hat {len(positions)} Position(en), "
                    "aber keine davon hat Anpassungsinformationen.")
                return
            
            # Bestätigung anzeigen
            message = f"BESTELLUNG: {order_number}\n\n"
            message += f"Gefundene Positionen: {len(positions)}\n"
            message += f"Mit Anpassungsinformationen: {len(customizable_positions)}\n\n"
            message += "Positionen mit Anpassungen:\n"
            for pos in customizable_positions:
                message += f"• Position {pos['position']}\n"
            message += "\nSollen alle Positionen verarbeitet werden?"
            
            result = ask_yes_no_safe("Multi-Position Verarbeitung", message)
            if not result:
                print("Verarbeitung vom Benutzer abgebrochen")
                return
            
            # Verarbeite alle Positionen mit Anpassungsinformationen
            processed_count = 0
            failed_count = 0
            
            for position_info in customizable_positions:
                try:
                    print(f"\n{'='*50}")
                    print(f"VERARBEITE POSITION {position_info['position']} VON {len(customizable_positions)}")
                    print(f"{'='*50}")
                    
                    # NEU: Nach jeder Navigation die Positionen neu laden
                    current_positions = find_order_positions(driver)
                    current_position_info = next(
                        (p for p in current_positions if p['position'] == position_info['position']), 
                        None
                    )
                    
                    if not current_position_info:
                        print(f"Position {position_info['position']} nicht mehr gefunden")
                        failed_count += 1
                        continue
                    
                    success = process_single_position(driver, current_position_info, order_number)
                    
                    if success:
                        processed_count += 1
                    else:
                        failed_count += 1
                    
                    # Kurze Pause zwischen Positionen
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"Fehler bei Position {position_info['position']}: {e}")
                    failed_count += 1
                    continue
            
            # Abschlussmeldung
            final_message = f"VERARBEITUNG ABGESCHLOSSEN!\n\n"
            final_message += f"Bestellung: {order_number}\n"
            final_message += f"Erfolgreich verarbeitet: {processed_count}\n"
            final_message += f"Fehlgeschlagen: {failed_count}\n"
            final_message += f"Gesamtpositionen: {len(positions)}\n\n"
            final_message += "Die TIFF-Dateien befinden sich im 'amazon_order_downloads' Ordner."
            
            safe_messagebox(messagebox.showinfo, "Verarbeitung abgeschlossen", final_message)
            
            # Öffne den Download-Ordner
            try:
                os.startfile(DOWNLOAD_DIR)
            except:
                pass
                
        except Exception as e:
            safe_messagebox(messagebox.showwarning, "Nicht gefunden", 
                f"Bestellung {order_number} wurde nicht gefunden oder die Seite hat zu lange geladen.")
            
    except Exception as e:
        safe_messagebox(messagebox.showerror, "Fehler", f"Multi-Position-Prozess fehlgeschlagen: {e}")
        print(f"Kritischer Fehler: {str(e)}")
        
    finally:
        # Browser erst am Ende schließen
        driver.quit()

# Aktualisierte process_downloaded_zip für Multi-Position
def process_downloaded_zip(order_number):
    """Verarbeite die heruntergeladene ZIP-Datei - Multi-Position Version"""
    print(f"=== Starte Verarbeitung für: {order_number} ===")
    
    # Warte auf Download-Vollendung
    if not wait_for_download_completion(DOWNLOAD_DIR, timeout=30):
        print("❌ Download nicht rechtzeitig abgeschlossen")
        return False
    
    # Finde die neueste ZIP-Datei
    zip_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.zip')]
    if not zip_files:
        print(f"❌ Keine ZIP-Datei gefunden in: {DOWNLOAD_DIR}")
        return False
    
    latest_zip = max(zip_files, key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)))
    zip_path = os.path.join(DOWNLOAD_DIR, latest_zip)
    
    print(f"Verarbeite ZIP-Datei: {zip_path}")
    
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
        
        # Liste entpackte Dateien
        print("Entpackte Dateien:")
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                print(f"  - {file_path}")
        
    except Exception as e:
        print(f"❌ Fehler beim Entpacken: {e}")
        return False
    
    # Verarbeite die Dateien zu TIFF
    tiff_path = process_files_to_tiff(extract_dir, order_number)
    
    if tiff_path and os.path.exists(tiff_path):
        print(f"✅ TIFF-Datei erfolgreich erstellt: {tiff_path}")
        
        # Lösche die ZIP-Datei nach erfolgreicher Verarbeitung
        try:
            os.remove(zip_path)
            print(f"ZIP-Datei gelöscht: {zip_path}")
        except Exception as e:
            print(f"Warnung: ZIP-Datei konnte nicht gelöscht werden: {e}")
        
        return True
    else:
        print("❌ TIFF-Datei konnte nicht erstellt werden")
        return False


def extract_image_filename_from_json(extract_dir):
    """
    Extrahiert den korrekten Bildnamen aus der JSON-Datei
    
    Returns:
        str: Der korrekte Bildname oder None wenn nicht gefunden
    """
    try:
        # Suche nach JSON-Dateien
        json_files = [f for f in os.listdir(extract_dir) if f.lower().endswith('.json')]
        if not json_files:
            print("Keine JSON-Datei gefunden")
            return None
        
        json_path = os.path.join(extract_dir, json_files[0])
        print(f"Durchsuche JSON-Datei nach Bildname: {json_path}")

        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Durchsuche customizationData nach ImageCustomization
        if 'customizationData' in data:
            print("Durchsuche customizationData...")
            image_name = search_for_image_in_data(data['customizationData'])
            if image_name:
                print(f"✅ Bildname in customizationData gefunden: {image_name}")
                return image_name
        
        # Fallback: Durchsuche customizationInfo
        if 'customizationInfo' in data:
            print("Durchsuche customizationInfo als Fallback...")
            # Hier könnten weitere Suchlogiken implementiert werden
            # Da in deinem Beispiel der Bildname nicht in customizationInfo steht
        
        print("❌ Kein Bildname in JSON gefunden")
        return None
        
    except Exception as e:
        print(f"Fehler beim Extrahieren des Bildnamens aus JSON: {e}")
        return None

def search_for_image_in_data(data):
    """
    Rekursive Suche nach ImageCustomization in den Datenstrukturen
    
    Args:
        data: Dictionary oder Liste zum Durchsuchen
    
    Returns:
        str: Bildname oder None
    """
    if isinstance(data, dict):
        # Prüfe ob dies eine ImageCustomization ist
        if data.get('type') == 'ImageCustomization' and 'image' in data:
            image_info = data['image']
            if 'imageName' in image_info:
                return image_info['imageName']
        
        # Durchsuche alle Werte im Dictionary
        for key, value in data.items():
            result = search_for_image_in_data(value)
            if result:
                return result
                
    elif isinstance(data, list):
        # Durchsuche alle Elemente in der Liste
        for item in data:
            result = search_for_image_in_data(item)
            if result:
                return result
    
    return None

def find_correct_image_file(extract_dir, target_filename):
    """
    Findet die korrekte Bilddatei basierend auf dem Ziel-Dateinamen
    
    Args:
        extract_dir (str): Verzeichnis mit den extrahierten Dateien
        target_filename (str): Ziel-Dateiname aus der JSON
    
    Returns:
        str: Vollständiger Pfad zur korrekten Bilddatei oder None
    """
    try:
        # Erstelle eine Liste aller Bilddateien
        image_files = []
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                    image_files.append(os.path.join(root, file))
        
        print(f"Gefundene Bilddateien: {[os.path.basename(f) for f in image_files]}")
        print(f"Suche nach: {target_filename}")
        
        # Exakte Übereinstimmung
        for image_file in image_files:
            if os.path.basename(image_file) == target_filename:
                print(f"✅ Exakte Übereinstimmung gefunden: {image_file}")
                return image_file
        
        # Suche ohne Dateiendung (falls sich die Endung geändert hat)
        target_base = os.path.splitext(target_filename)[0]
        for image_file in image_files:
            file_base = os.path.splitext(os.path.basename(image_file))[0]
            if file_base == target_base:
                print(f"✅ Übereinstimmung ohne Dateiendung gefunden: {image_file}")
                return image_file
        
        print(f"❌ Keine Übereinstimmung für {target_filename} gefunden")
        return None
        
    except Exception as e:
        print(f"Fehler beim Suchen der korrekten Bilddatei: {e}")
        return None  

# === ERWEITERTE VERSION mit Heizungstyp-Erkennung ===
def extract_dimensions_and_check_text(extract_dir):
    """
    Erweiterte Version mit Heizungstyp-Erkennung
    """
    try:
        # Suche nach JSON-Dateien
        json_files = [f for f in os.listdir(extract_dir) if f.lower().endswith('.json')]
        if not json_files:
            safe_messagebox(messagebox.showerror, "Fehler", "Keine JSON-Datei im Download gefunden")
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
                            safe_messagebox(messagebox.showinfo, 
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
                                break
                    if required_dimensions:
                        break
                if required_dimensions:
                    break
        
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
                            break
                if required_dimensions:
                    break
        
        if not required_dimensions:
            safe_messagebox(messagebox.showerror, 
                "Fehler", 
                "Konnte Druckdimensionen nicht ermitteln.\n"
                "Die Verarbeitung wird abgebrochen."
            )
            return None
        
        # 3. NEU: Heizungstyp-Erkennung
        heating_type, heating_specs = detect_heating_type(required_dimensions)
        
        # 4. NEU: Validierung mit Benutzer-Bestätigung
        if not validate_heating_match(heating_type, heating_specs, required_dimensions):
            # Falls Benutzer ablehnt oder kein Match, zeige Empfehlungen
            if heating_type == "Unbekannt":
                recommendations = get_heating_recommendations(required_dimensions)
                if recommendations:
                    rec_text = "ÄHNLICHE HEIZUNGSTYPEN:\n\n"
                    for i, (rec_type, rec_specs, deviation) in enumerate(recommendations, 1):
                        rec_text += f"{i}. {rec_type}\n"
                        rec_text += f"   Verhältnis: {rec_specs['width']/rec_specs['height']:.4f}\n"
                        rec_text += f"   Abweichung: {deviation:.4f}\n\n"
                    
                    safe_messagebox(messagebox.showinfo, "Empfehlungen", rec_text)
            
            return None  # Abbruch der Verarbeitung
        
        # 5. Erweitere Dimensions um Heizungsinfo
        required_dimensions['heating_type'] = heating_type
        required_dimensions['heating_specs'] = heating_specs
        
        return required_dimensions
        
    except Exception as e:
        safe_messagebox(messagebox.showerror, "Fehler", f"JSON-Verarbeitung fehlgeschlagen: {str(e)}")
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
    """Verarbeite SVG und Bilddateien zu TIFF mit korrekter Bilderkennung"""
    try:
        print(f"=== Starte Dateiverarbeitung für {order_number} ===")
        
        # 1. Finde SVG-Datei
        svg_file = None
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                if file.lower().endswith('.svg'):
                    svg_file = os.path.join(root, file)
                    break
            if svg_file:
                break
        
        if not svg_file:
            safe_messagebox(messagebox.showerror, "Fehler", "Keine SVG-Datei gefunden")
            return None
        
        # 2. NEUE LOGIK: Extrahiere korrekten Bildnamen aus JSON
        target_image_name = extract_image_filename_from_json(extract_dir)
        target_image_file = None
        
        if target_image_name:
            # Suche die korrekte Bilddatei
            target_image_file = find_correct_image_file(extract_dir, target_image_name)
            
        if not target_image_file:
            print("⚠️ Fallback: Verwende größte Bilddatei")
            # Fallback zur alten Methode: größte Bilddatei
            image_files = []
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                        image_files.append(os.path.join(root, file))
            
            if not image_files:
                safe_messagebox(messagebox.showerror, "Fehler", "Keine Bilddateien gefunden")
                return None
            
            target_image_file = max(image_files, key=lambda f: os.path.getsize(f))
            print(f"Fallback: Verwende größte Datei: {os.path.basename(target_image_file)}")
        
        print(f"🎯 Verwende Bilddatei: {os.path.basename(target_image_file)}")
        
        # 3. Extrahiere Dimensionen und Heizungstyp
        dimensions = extract_dimensions_and_check_text(extract_dir)
        if not dimensions:  # Wenn keine Dimensionen gefunden wurden oder Benutzer abgebrochen
            return None  # Verarbeitung abbrechen
        
        # 4. Verarbeite Bild
        modified_svg = embed_image_in_svg(target_image_file, svg_file)
        if not modified_svg:
            return None
        
        # 5. Konvertiere zu TIFF
        output_path = os.path.join(extract_dir, f"{order_number}.tiff")
        if not convert_svg_to_tiff(modified_svg, output_path):
            return None
        
        # 6. Verhältniskontrolle (mit Heizungstyp-Info)
        if dimensions and 'ratio' in dimensions:
            print("Führe Verhältniskontrolle durch...")
            if not check_and_correct_aspect_ratio(output_path, dimensions['ratio']):
                safe_messagebox(messagebox.showwarning, "Warnung", "Bildverhältnis konnte nicht perfekt korrigiert werden")
        
        return output_path
        
    except Exception as e:
        safe_messagebox(messagebox.showerror, "Fehler", f"Verarbeitung fehlgeschlagen: {str(e)}")
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
        safe_messagebox(messagebox.showerror, "Fehler", f"Bildeinbettung fehlgeschlagen: {str(e)}")
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
        safe_messagebox(messagebox.showerror, "Fehler", f"Konvertierung fehlgeschlagen: {str(e)}")
        return False

# === Bestellung suchen und verarbeiten ===
def search_order(order_number):
    driver = create_driver()
    
    try:
        print(f"=== Starte Suche nach Bestellung: {order_number} ===")
        
        # Cookies laden und prüfen
        if not load_cookies(driver):
            safe_messagebox(messagebox.showerror, "Fehler", "Keine gültigen Cookies gefunden. Bitte logge dich zuerst manuell ein.")
            return

        # Seite neu laden um Login zu aktivieren
        driver.refresh()
        time.sleep(3)
        
        # Flexiblere Login-Prüfung
        current_url = driver.current_url
        print(f"URL nach Cookie-Login: {current_url}")
        
        # Wenn wir auf Login-Seite sind, Session ist abgelaufen
        if any(keyword in current_url.lower() for keyword in ["signin", "login", "auth"]):
            safe_messagebox(messagebox.showerror, "Session abgelaufen", "Deine Session ist abgelaufen. Bitte logge dich erneut ein.")
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
                safe_messagebox(messagebox.showwarning, "Nicht gefunden", f"Bestellung {order_number} wurde nicht gefunden. Bitte überprüfen Sie die Bestellnummer.")
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
            safe_messagebox(messagebox.showwarning, "Nicht gefunden", f"Bestellung {order_number} wurde nicht gefunden oder die Seite hat zu lange geladen. Bitte überprüfen Sie die Bestellnummer.")
            
    except Exception as e:
        safe_messagebox(messagebox.showerror, "Fehler", f"Prozess fehlgeschlagen: {e}")
        print(f"Fehler aufgetreten: {str(e)}")
    finally:
        driver.quit()

# === GUI ===
def start_gui():
    window = tk.Tk()
    window.title("Amazon Seller Central - Bestellungssuche & Verarbeitung mit Heizungstyp-Erkennung")
    window.geometry("500x400")
    
    # Titel
    tk.Label(window, text="INFRAROTHEIZUNG DRUCKDATEI-GENERATOR", 
             font=("Arial", 14, "bold"), fg="red").pack(pady=(20, 5))
    
    tk.Label(window, text="Hinweis: Cookies bleiben gültig bis sie ablaufen", 
             font=("Arial", 9), fg="gray").pack(pady=(5, 15))
    
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
            threading.Thread(target=search_order_multi_position, args=(order_number,), daemon=True).start()
        else:
            safe_messagebox(messagebox.showwarning, "Hinweis", "Bitte eine Bestellnummer eingeben.")

    def on_barcode_input(event):
        # Der Barcode-Scanner sendet die Daten + Enter
        # Wir nehmen den aktuellen Inhalt des Feldes (ohne den letzten Zeilenumbruch)
        barcode = order_entry.get().strip()
        if barcode:
            handle_search(barcode)

    # Enter-Taste und Barcode-Eingabe binden
    order_entry.bind('<Return>', on_barcode_input)
    
    # Buttons
    tk.Button(window, text="🔍 Bestellung suchen & verarbeiten", 
              command=lambda: handle_search(), 
              font=("Arial", 12), bg="#FF9900", fg="black", width=30).pack(pady=10)
    
    # Separator
    separator = tk.Frame(window, height=2, bd=1, relief=tk.SUNKEN)
    separator.pack(fill=tk.X, padx=20, pady=10)
    
    # Management Buttons
    tk.Label(window, text="VERWALTUNG:", font=("Arial", 10, "bold")).pack(pady=(5, 5))
    
    tk.Button(window, text="🔐 Manuell einloggen & Cookies speichern", 
              command=manual_login,
              font=("Arial", 10), width=40).pack(pady=2)
    
    tk.Button(window, text="📊 Cookie-Status prüfen", 
              command=check_cookie_status,
              font=("Arial", 10), width=40).pack(pady=2)
    
    tk.Button(window, text="⚙️ Heizungstypen konfigurieren", 
              command=edit_heating_config,
              font=("Arial", 10), width=40).pack(pady=2)

    # Info-Bereich
    info_frame = tk.Frame(window, bg="#f0f0f0", relief=tk.RIDGE, bd=1)
    info_frame.pack(fill=tk.X, padx=20, pady=15)
    
    tk.Label(info_frame, text="ℹ️ NEUE FEATURES:", 
             font=("Arial", 9, "bold"), bg="#f0f0f0").pack(pady=(5, 2))
    
    tk.Label(info_frame, text="• Automatische Heizungstyp-Erkennung", 
             font=("Arial", 8), bg="#f0f0f0").pack(anchor="w", padx=10)
    
    tk.Label(info_frame, text="• Bildverhältnis-Kontrolle", 
             font=("Arial", 8), bg="#f0f0f0").pack(anchor="w", padx=10)
    
    tk.Label(info_frame, text="• Konfigurierbare Heizungsgrößen", 
             font=("Arial", 8), bg="#f0f0f0").pack(anchor="w", padx=10, pady=(0, 5))

    # Hinweis
    tk.Label(window, text="Hinweis: Cookies sind ca. 12 Stunden gültig | Barcode-Scanner unterstützt", 
             font=("Arial", 8), fg="gray").pack(pady=(10, 5))
    
    # Automatisch nach Barcode-Eingabe suchen
    window.mainloop()

# === Programmstart ===
if __name__ == "__main__":
    # Erstelle Standard-Config beim ersten Start
    load_config()
    start_gui()
