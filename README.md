# 🐑 Mute Sheep

Une application Windows simple et élégante pour contrôler votre microphone d'un seul clic ou raccourci clavier.

## ✨ Fonctionnalités

- 🎤 **Activation/Désactivation rapide** : Contrôlez votre microphone avec un raccourci clavier personnalisable
- ⌨️ **Raccourci personnalisable** : Choisissez votre propre touche (par défaut : F8)
- 🔊 **Notifications sonores** : Sons Windows lors de l'activation/désactivation
- 🗕 **Mode système tray** : Minimisez l'application dans la barre des tâches
- 🚀 **Auto-démarrage Windows** : Lancez automatiquement au démarrage de Windows
- 🎨 **Interface moderne** : Design sombre et épuré avec CustomTkinter

## 🔧 Installation

### Utiliser l'exécutable

1. Téléchargez la dernière version depuis [Releases](https://github.com/CeriseeBrandy/MuteSheep/releases)
2. Lancez `MuteSheep.exe`
3. C'est tout ! 🎉

### Depuis le code source
```bash
git clone https://github.com/CeriseeBrandy/MuteSheep.git
cd MuteSheep
pip install -r requirements.txt
python mute_sheep.py
```

## 🎮 Utilisation

1. **Changer le raccourci** : Cliquez sur "Changer Sheep ⌨" puis appuyez sur la touche de votre choix
2. **Activer/Désactiver le micro** : Appuyez sur votre raccourci (F8 par défaut)
3. **Minimiser** : Cliquez sur "🗕 Minimiser dans la barre des tâches"
4. **Auto-démarrage** : Cliquez sur "⚙ Auto-démarrage Windows"

## 🛠️ Technologies utilisées

- Python 3.8+
- CustomTkinter
- pynput
- pycaw
- pystray
- Pillow

## 📝 License

Ce projet est sous licence MIT.