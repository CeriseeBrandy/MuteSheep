import customtkinter as ctk
from pynput import keyboard
import threading
import os
import sys
import winreg
import winsound
from PIL import Image
import pystray

# Bibliothèque pour le contrôle sonore système
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

ctk.set_appearance_mode("dark")

class MuteSheep(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Mute Sheep")
        self.geometry("500x650")
        self.configure(fg_color="#333333")
        
        # Définir l'icône si elle existe
        self.set_icon()

        self.shortcut = "f8"
        self.is_muted = False
        self.waiting_for_key = False
        
        # Initialiser le volume du microphone
        self.volume = None
        self.init_microphone()
        
        # System tray
        self.tray_icon = None
        self.is_minimized_to_tray = False

        self.setup_ui()
        
        # Lancement de l'écoute clavier
        self.listener = keyboard.Listener(on_press=self.on_press)
        self.listener.start()

    def set_icon(self):
        """Définit l'icône de l'application si MuteSheep.ico existe"""
        # Chercher l'icône dans plusieurs emplacements possibles
        possible_paths = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "MuteSheep.ico"),
            os.path.join(os.path.dirname(sys.executable), "MuteSheep.ico"),
            os.path.join(os.getcwd(), "MuteSheep.ico"),
            "MuteSheep.ico"
        ]
        
        # Si l'application est packagée avec PyInstaller
        if getattr(sys, 'frozen', False):
            # Chemin pour les applications PyInstaller
            base_path = sys._MEIPASS
            possible_paths.insert(0, os.path.join(base_path, "MuteSheep.ico"))
        
        for icon_path in possible_paths:
            if os.path.exists(icon_path):
                try:
                    self.iconbitmap(icon_path)
                    # Pour Windows : définir aussi l'icône de la barre des tâches
                    self.iconbitmap(default=icon_path)
                    print(f"Icône chargée : {icon_path}")
                    return
                except Exception as e:
                    print(f"Erreur lors du chargement de l'icône {icon_path} : {e}")
        
        print("Icône non trouvée dans aucun emplacement")

    def init_microphone(self):
        """Initialise la connexion au microphone au démarrage"""
        try:
            devices = AudioUtilities.GetMicrophone()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume = cast(interface, POINTER(IAudioEndpointVolume))
            print("Microphone initialisé avec succès")
        except Exception as e:
            print(f"Erreur d'initialisation du microphone : {e}")
            self.volume = None

    def setup_ui(self):
        ctk.CTkLabel(self, text="🐑", font=("Arial", 60)).pack(pady=(40, 5))
        ctk.CTkLabel(self, text="Mute Sheep", font=("Arial", 36, "bold")).pack()
        ctk.CTkLabel(self, text="Appuyez sur la touche choisie pour couper le micro", font=("Arial", 13), text_color="#AAAAAA").pack(pady=20)

        self.btn_change = ctk.CTkButton(self, text="Changer Sheep ⌨", command=self.start_key_assignment,
                                       fg_color="#3498DB", hover_color="#2980B9", font=("Arial", 20, "bold"), height=60, width=300)
        self.btn_change.pack(pady=10)

        self.cur_key_lbl = ctk.CTkLabel(self, text=f"Touche actuelle : {self.shortcut.upper()}")
        self.cur_key_lbl.pack()

        ctk.CTkFrame(self, height=1, width=420, fg_color="#555555").pack(pady=30)

        self.status_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.status_frame.pack(padx=40, fill="x")
        ctk.CTkLabel(self.status_frame, text="Statut le raccourci : ", font=("Arial", 15)).pack(side="left")
        
        self.status_val = ctk.CTkLabel(self.status_frame, text="ACTIF", text_color="#2ECC71", font=("Arial", 15, "bold"))
        self.status_val.pack(side="left")
        
        self.mute_val = ctk.CTkLabel(self.status_frame, text="  MUET", text_color="#555555", font=("Arial", 15, "bold"))
        self.mute_val.pack(side="left")
        
        # Ajout du bouton auto-démarrage
        ctk.CTkFrame(self, height=1, width=420, fg_color="#555555").pack(pady=20)
        
        self.autostart_btn = ctk.CTkButton(self, text="⚙ Auto-démarrage Windows", 
                                          command=self.toggle_autostart,
                                          fg_color="#9B59B6", hover_color="#8E44AD", 
                                          font=("Arial", 16, "bold"), height=50, width=300)
        self.autostart_btn.pack(pady=10)
        
        self.autostart_status = ctk.CTkLabel(self, text="", font=("Arial", 12), text_color="#AAAAAA")
        self.autostart_status.pack()
        
        # Bouton pour minimiser dans la barre des tâches
        self.minimize_btn = ctk.CTkButton(self, text="🗕 Minimiser dans la barre des tâches", 
                                         command=self.minimize_to_tray,
                                         fg_color="#34495E", hover_color="#2C3E50", 
                                         font=("Arial", 14), height=45, width=300)
        self.minimize_btn.pack(pady=10)
        
        # Vérifier le statut de l'auto-démarrage
        self.update_autostart_status()

    def play_sound(self, sound_type):
        """Joue un son système Windows"""
        def _play():
            try:
                if sound_type == "mute":
                    # Son pour le mute (son d'erreur ou critique)
                    winsound.MessageBeep(winsound.MB_ICONHAND)
                else:
                    # Son pour l'activation (son par défaut ou astérisque)
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except Exception as e:
                print(f"Erreur lors de la lecture du son : {e}")
        
        # Jouer le son dans un thread séparé pour ne pas bloquer l'UI
        threading.Thread(target=_play, daemon=True).start()

    def toggle_mute(self):
        """Bascule l'état muet du microphone"""
        self.is_muted = not self.is_muted
        
        # Mise à jour de l'interface
        if self.is_muted:
            self.status_val.configure(text_color="#555555")
            self.mute_val.configure(text_color="#E74C3C")
        else:
            self.status_val.configure(text_color="#2ECC71")
            self.mute_val.configure(text_color="#555555")
        
        # Jouer le son approprié
        self.play_sound("mute" if self.is_muted else "unmute")
        
        # Exécution de la commande système
        threading.Thread(target=self.apply_mute_windows, daemon=True).start()

    def apply_mute_windows(self):
        """Applique l'état muet au microphone Windows"""
        try:
            if self.volume is None:
                self.init_microphone()
            
            if self.volume is not None:
                # Application du mute
                self.volume.SetMute(1 if self.is_muted else 0, None)
                print(f"Action Windows : {'MUET' if self.is_muted else 'ACTIF'}")
            else:
                print("Erreur : Impossible d'accéder au microphone")
        except Exception as e:
            print(f"Erreur lors du mute : {e}")
            # Réinitialiser la connexion au microphone
            self.volume = None
            self.init_microphone()

    def start_key_assignment(self):
        """Démarre l'assignation d'une nouvelle touche"""
        self.waiting_for_key = True
        self.btn_change.configure(text="Appuyez sur une touche...", fg_color="#E67E22")

    def on_press(self, key):
        """Gère les événements de pression de touche"""
        if self.waiting_for_key:
            try:
                self.shortcut = key.char.lower()
            except:
                self.shortcut = str(key).replace("Key.", "")
            self.after(0, self.update_key_ui)
            self.waiting_for_key = False
            return

        try:
            k = key.char.lower()
        except:
            k = str(key).replace("Key.", "")
        
        if k == self.shortcut:
            self.after(0, self.toggle_mute)

    def update_key_ui(self):
        """Met à jour l'affichage de la touche actuelle"""
        self.cur_key_lbl.configure(text=f"Touche actuelle : {self.shortcut.upper()}")
        self.btn_change.configure(text="Changer Sheep ⌨", fg_color="#3498DB")

    def toggle_autostart(self):
        """Active ou désactive l'auto-démarrage de l'application"""
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "MuteSheep"
        
        try:
            # Ouvrir la clé de registre
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            
            # Vérifier si déjà configuré
            try:
                winreg.QueryValueEx(key, app_name)
                # Si présent, le supprimer
                winreg.DeleteValue(key, app_name)
                winreg.CloseKey(key)
                print("Auto-démarrage désactivé")
            except FileNotFoundError:
                # Si absent, l'ajouter
                script_path = os.path.abspath(sys.argv[0])
                
                # Si c'est un script .py, utiliser pythonw.exe pour éviter la console
                if script_path.endswith('.py'):
                    python_path = sys.executable.replace('python.exe', 'pythonw.exe')
                    if not os.path.exists(python_path):
                        python_path = sys.executable
                    value = f'"{python_path}" "{script_path}"'
                else:
                    value = f'"{script_path}"'
                
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, value)
                winreg.CloseKey(key)
                print("Auto-démarrage activé")
            
            # Mettre à jour le statut
            self.update_autostart_status()
            
        except Exception as e:
            print(f"Erreur lors de la configuration de l'auto-démarrage : {e}")

    def update_autostart_status(self):
        """Met à jour l'affichage du statut de l'auto-démarrage"""
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "MuteSheep"
        
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
            try:
                winreg.QueryValueEx(key, app_name)
                self.autostart_status.configure(text="✓ Activé au démarrage", text_color="#2ECC71")
                winreg.CloseKey(key)
            except FileNotFoundError:
                self.autostart_status.configure(text="✗ Non activé au démarrage", text_color="#E74C3C")
                winreg.CloseKey(key)
        except Exception as e:
            self.autostart_status.configure(text="Erreur de vérification", text_color="#E67E22")
            print(f"Erreur lors de la vérification : {e}")

    def on_closing(self):
        """Gère la fermeture propre de l'application"""
        if self.tray_icon:
            self.tray_icon.stop()
        if self.listener:
            self.listener.stop()
        self.destroy()
    
    def create_tray_icon(self):
        """Crée l'icône dans la barre des tâches"""
        # Créer une image pour l'icône (mouton emoji en image)
        icon_path = self.find_icon_path()
        
        if icon_path and os.path.exists(icon_path):
            try:
                image = Image.open(icon_path)
            except:
                image = self.create_default_icon()
        else:
            image = self.create_default_icon()
        
        # Menu du system tray
        menu = pystray.Menu(
            pystray.MenuItem("🐑 Mute Sheep", self.show_window, default=True),
            pystray.MenuItem(f"Touche : {self.shortcut.upper()}", lambda: None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Afficher", self.show_window),
            pystray.MenuItem("Quitter", self.quit_app)
        )
        
        self.tray_icon = pystray.Icon("MuteSheep", image, "Mute Sheep", menu)
    
    def create_default_icon(self):
        """Crée une icône par défaut si aucune n'est trouvée"""
        # Créer une image simple 64x64 avec un fond
        width = 64
        height = 64
        color1 = (52, 152, 219)  # Bleu
        color2 = (255, 255, 255)  # Blanc
        
        image = Image.new('RGB', (width, height), color1)
        pixels = image.load()
        
        # Dessiner un cercle simple
        for x in range(width):
            for y in range(height):
                dist = ((x - width/2)**2 + (y - height/2)**2)**0.5
                if dist < width/3:
                    pixels[x, y] = color2
        
        return image
    
    def find_icon_path(self):
        """Trouve le chemin de l'icône"""
        possible_paths = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "MuteSheep.ico"),
            os.path.join(os.path.dirname(sys.executable), "MuteSheep.ico"),
            os.path.join(os.getcwd(), "MuteSheep.ico"),
            "MuteSheep.ico"
        ]
        
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
            possible_paths.insert(0, os.path.join(base_path, "MuteSheep.ico"))
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None
    
    def minimize_to_tray(self):
        """Minimise l'application dans la barre des tâches"""
        if not self.tray_icon:
            self.create_tray_icon()
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
        
        self.withdraw()  # Cache la fenêtre
        self.is_minimized_to_tray = True
    
    def show_window(self, icon=None, item=None):
        """Affiche la fenêtre depuis la barre des tâches"""
        self.after(0, self._show_window)
    
    def _show_window(self):
        """Affiche réellement la fenêtre (doit être appelé dans le thread principal)"""
        self.deiconify()  # Affiche la fenêtre
        self.lift()  # Met la fenêtre au premier plan
        self.focus_force()  # Donne le focus
        self.is_minimized_to_tray = False
    
    def quit_app(self, icon=None, item=None):
        """Quitte complètement l'application"""
        if self.tray_icon:
            self.tray_icon.stop()
        self.after(0, self.on_closing)

if __name__ == "__main__":
    app = MuteSheep()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()