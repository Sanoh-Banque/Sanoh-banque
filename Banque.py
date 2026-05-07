"""
=============================================================
  BanquePy v2 – Application Bancaire avec Interface Graphique
  -----------------------------------------------------------
  Ce programme simule une banque avec une interface graphique.
  Il utilise plusieurs bibliothèques Python :
    • tkinter    → créer les fenêtres et boutons (interface)
    • sqlite3    → sauvegarder les données dans un fichier .db
    • matplotlib → dessiner des graphiques
    • csv        → exporter l'historique dans un tableur
    • datetime   → obtenir la date et l'heure actuelles
    • os         → accéder aux chemins de fichiers sur l'ordinateur

  Fonctionnalités :
    ✔ Créer des comptes bancaires
    ✔ Déposer et retirer de l'argent
    ✔ Faire des virements entre comptes
    ✔ Consulter le solde et l'historique
    ✔ Voir des graphiques (évolution, répartition, activité)
    ✔ Exporter l'historique en fichier CSV (Excel)
    ✔ Rechercher et supprimer des comptes
    ✔ Les données sont SAUVEGARDÉES même après fermeture
=============================================================
"""

# ─────────────────────────────────────────────────────────────
#  IMPORTATION DES BIBLIOTHÈQUES
# ─────────────────────────────────────────────────────────────

# tkinter : bibliothèque standard Python pour créer des interfaces graphiques
import tkinter as tk
# ttk : widgets plus modernes (tableaux, barres de défilement...)
# messagebox : boîtes de dialogue (erreur, succès, confirmation)
# filedialog : fenêtre "Enregistrer sous..." pour choisir un fichier
from tkinter import ttk, messagebox, filedialog

# datetime : pour obtenir la date et l'heure du moment présent
import datetime

# sqlite3 : base de données légère intégrée à Python (pas besoin d'installation)
# Les données sont stockées dans un fichier .db sur l'ordinateur
import sqlite3

# csv : permet de lire et écrire des fichiers CSV (lisibles par Excel)
import csv

# os : permet de manipuler les chemins de fichiers (dossiers, noms de fichiers)
import os


# ─────────────────────────────────────────────────────────────
#  CHARGEMENT DE MATPLOTLIB (OPTIONNEL)
# ─────────────────────────────────────────────────────────────

# On essaie d'importer matplotlib pour faire des graphiques.
# Si la bibliothèque n'est pas installée, le programme continue quand même
# mais sans les graphiques. Pour l'installer : pip install matplotlib
try:
    import matplotlib
    # "TkAgg" = mode d'affichage compatible avec Tkinter
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    # FigureCanvasTkAgg = pont entre Matplotlib et Tkinter (affiche le graphique dans la fenêtre)
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    MATPLOTLIB_OK = True   # On note que Matplotlib est disponible
except ImportError:
    MATPLOTLIB_OK = False  # Matplotlib absent : les graphiques seront désactivés


# ─────────────────────────────────────────────────────────────
#  CHARGEMENT DE PANDAS (OPTIONNEL)
# ─────────────────────────────────────────────────────────────

# Pandas est une bibliothèque pour manipuler des tableaux de données.
# Elle n'est pas indispensable ici, mais elle est prête à être utilisée.
# Pour l'installer : pip install pandas
try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False


# ─────────────────────────────────────────────────────────────
#  COULEURS ET POLICES (THÈME VISUEL)
# ─────────────────────────────────────────────────────────────

# Dictionnaire C : toutes les couleurs de l'application en un seul endroit.
# Chaque couleur est un code hexadécimal (ex: "#0A1628" = bleu très foncé).
# Avantage : si on veut changer une couleur, on ne la modifie qu'ici.
C = {
    "navy":        "#0A1628",   # Fond principal (bleu marine très foncé)
    "navy_mid":    "#0F2044",   # Fond des panneaux (bleu marine moyen)
    "navy_light":  "#162A55",   # Fond des cartes (bleu marine clair)
    "navy_hover":  "#1C3366",   # Couleur au survol des boutons
    "gold":        "#C9A84C",   # Or → titres, bouton actif dans le menu
    "gold_light":  "#E8C96A",   # Or clair → survol des boutons principaux
    "white":       "#FFFFFF",   # Blanc pur
    "off_white":   "#E8EDF5",   # Blanc cassé → texte normal
    "muted":       "#7A8BAD",   # Gris-bleu → texte secondaire (labels)
    "border":      "#1E3560",   # Couleur des bordures et séparateurs
    "green":       "#2ECC8D",   # Vert → succès, solde positif, dépôt
    "red":         "#E05555",   # Rouge → erreur, retrait
    "amber":       "#F0A030",   # Orange → avertissement, retrait effectué
    "teal":        "#22A8B8",   # Bleu-vert → statistiques, badges
    "sidebar_w":   230,         # Largeur de la barre latérale (en pixels)
    "header_h":    64,          # Hauteur de l'en-tête (en pixels)
}

# Polices de caractères utilisées dans l'application.
# Format : ("nom de la police", taille, style)
FONT_TITLE   = ("Georgia", 22, "bold")    # Grands titres de page
FONT_HEADING = ("Georgia", 14, "bold")    # Sous-titres de section
FONT_LABEL   = ("Helvetica", 10)          # Texte normal dans les formulaires
FONT_VALUE   = ("Helvetica", 13, "bold")  # Valeurs importantes (solde...)
FONT_SMALL   = ("Helvetica", 9)           # Petits textes (hints, dates...)
FONT_MONO    = ("Courier", 10)            # Police à chasse fixe (non utilisée ici)
FONT_BTN     = ("Helvetica", 10, "bold")  # Texte dans les boutons
FONT_NAV     = ("Helvetica", 11)          # Texte dans le menu de navigation

# Chemin vers le fichier de base de données SQLite.
# os.path.abspath(__file__) donne le dossier où se trouve ce script Python.
# Le fichier "banquepy.db" sera créé dans le même dossier que ce script.
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "banquepy.db")


# ─────────────────────────────────────────────────────────────
#  CLASSE DATABASE – GESTION DE LA BASE DE DONNÉES SQLite
# ─────────────────────────────────────────────────────────────

class Database:
    """
    Cette classe s'occupe de TOUT ce qui touche à la base de données.
    Elle fait le lien entre le programme et le fichier banquepy.db.

    Une base de données SQLite, c'est comme un fichier Excel intelligent :
    - Il contient des "tables" (comme des feuilles Excel)
    - Chaque table a des colonnes et des lignes
    - On peut ajouter, modifier, supprimer et rechercher des données
    """

    def __init__(self, path=DB_PATH):
        """
        Constructeur : appelé automatiquement quand on crée un objet Database.
        - path : chemin du fichier .db (par défaut : banquepy.db)
        """
        self.path = path

        # Connexion à la base de données (crée le fichier s'il n'existe pas)
        self.conn = sqlite3.connect(path, check_same_thread=False)

        # row_factory = sqlite3.Row : permet d'accéder aux colonnes par leur nom
        # Exemple : ligne["titulaire"] au lieu de ligne[1]
        self.conn.row_factory = sqlite3.Row

        # Création des tables si elles n'existent pas encore
        self._init_schema()

    def _init_schema(self):
        """
        Crée la structure de la base de données (les tables).
        "IF NOT EXISTS" = si la table existe déjà, on ne fait rien.

        Tables créées :
          • comptes      → stocke les informations de chaque compte
          • transactions → stocke chaque dépôt, retrait, virement
          • meta         → stocke des paramètres internes (ex: compteur)
        """
        cur = self.conn.cursor()  # Le curseur est l'outil pour envoyer des commandes SQL

        # executescript : exécute plusieurs commandes SQL d'un coup
        cur.executescript("""
            CREATE TABLE IF NOT EXISTS comptes (
                numero      TEXT PRIMARY KEY,   -- Identifiant unique (ex: ACC001)
                titulaire   TEXT NOT NULL,       -- Nom du propriétaire du compte
                solde       REAL NOT NULL DEFAULT 0.0,  -- Argent disponible
                cree_le     TEXT NOT NULL        -- Date de création du compte
            );

            CREATE TABLE IF NOT EXISTS transactions (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,  -- Numéro automatique
                compte_num  TEXT NOT NULL,       -- Numéro du compte concerné
                date_op     TEXT NOT NULL,       -- Date de l'opération
                type_op     TEXT NOT NULL,       -- "Dépôt", "Retrait", "Virement"...
                montant     REAL NOT NULL,       -- Montant (positif ou négatif)
                note        TEXT DEFAULT '',     -- Commentaire optionnel
                FOREIGN KEY (compte_num) REFERENCES comptes(numero)
                -- La clé étrangère garantit que le compte_num existe bien dans comptes
            );

            CREATE TABLE IF NOT EXISTS meta (
                cle         TEXT PRIMARY KEY,   -- Nom du paramètre (ex: "compteur")
                valeur      TEXT                -- Valeur du paramètre
            );
        """)

        # Initialise le compteur de numéros de compte à 1 (si pas déjà fait)
        # "INSERT OR IGNORE" = n'insère que si la clé n'existe pas encore
        cur.execute("INSERT OR IGNORE INTO meta VALUES ('compteur','1')")
        self.conn.commit()  # commit() = valide et sauvegarde les changements

    # ── GESTION DES COMPTES ────────────────────────────────────

    def prochain_numero(self):
        """
        Génère le prochain numéro de compte unique (ACC001, ACC002, etc.)
        À chaque appel, le compteur augmente de 1 dans la base de données.
        """
        cur = self.conn.cursor()

        # Lit la valeur actuelle du compteur
        cur.execute("SELECT valeur FROM meta WHERE cle='compteur'")
        n = int(cur.fetchone()[0])  # fetchone() récupère la première ligne

        # Incrémente le compteur pour la prochaine fois
        cur.execute("UPDATE meta SET valeur=? WHERE cle='compteur'", (n + 1,))
        self.conn.commit()

        # Retourne le numéro formaté : 1 → "ACC001", 12 → "ACC012"
        return f"ACC{n:03d}"

    def creer_compte(self, titulaire, solde_initial=0.0):
        """
        Crée un nouveau compte bancaire dans la base de données.
        - titulaire    : nom du propriétaire
        - solde_initial : argent de départ (doit être ≥ 0)
        Retourne le numéro du compte créé (ex: "ACC001")
        """
        # Vérification : le solde ne peut pas être négatif
        if solde_initial < 0:
            raise ValueError("Le solde initial ne peut pas être négatif.")

        # Génère un numéro unique pour ce compte
        numero = self.prochain_numero()

        # Récupère la date et l'heure actuelles, formatées lisiblement
        now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        cur = self.conn.cursor()

        # Insère le nouveau compte dans la table "comptes"
        cur.execute(
            "INSERT INTO comptes VALUES (?,?,?,?)",
            (numero, titulaire, solde_initial, now)
            # Les "?" sont des paramètres : cela protège contre les injections SQL
        )

        # Si un solde initial est fourni, on enregistre aussi une transaction
        if solde_initial > 0:
            cur.execute(
                "INSERT INTO transactions(compte_num,date_op,type_op,montant,note) VALUES (?,?,?,?,?)",
                (numero, now, "Dépôt initial", solde_initial, "")
            )

        self.conn.commit()
        return numero  # On retourne le numéro pour l'afficher à l'utilisateur

    def trouver_compte(self, numero):
        """
        Cherche un compte par son numéro (ex: "ACC001").
        Retourne les données du compte, ou None s'il n'existe pas.
        .upper() = convertit en majuscules pour éviter les erreurs de casse
        """
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM comptes WHERE numero=?", (numero.upper(),))
        return cur.fetchone()  # fetchone() = une seule ligne (ou None)

    def tous_les_comptes(self):
        """
        Retourne la liste de tous les comptes, triés par date de création.
        fetchall() = toutes les lignes (une liste)
        """
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM comptes ORDER BY cree_le")
        return cur.fetchall()

    def supprimer_compte(self, numero):
        """
        Supprime définitivement un compte ET tout son historique.
        Ordre important : supprimer d'abord les transactions,
        ensuite le compte (sinon la contrainte de clé étrangère bloque).
        """
        cur = self.conn.cursor()
        cur.execute("DELETE FROM transactions WHERE compte_num=?", (numero,))
        cur.execute("DELETE FROM comptes WHERE numero=?", (numero,))
        self.conn.commit()

    def rechercher_comptes(self, terme):
        """
        Recherche des comptes dont le numéro OU le titulaire contient le terme.
        LIKE "%terme%" = cherche partout dans la chaîne (début, milieu, fin)
        """
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM comptes WHERE numero LIKE ? OR titulaire LIKE ?",
            (f"%{terme}%", f"%{terme}%")
        )
        return cur.fetchall()

    # ── OPÉRATIONS BANCAIRES ───────────────────────────────────

    def deposer(self, numero, montant, note=""):
        """
        Crédite un compte (ajoute de l'argent).
        - numero  : numéro du compte à créditer
        - montant : somme à déposer (doit être > 0)
        - note    : commentaire optionnel sur l'opération
        """
        # Vérifications avant d'effectuer l'opération
        if montant <= 0:
            raise ValueError("Le montant doit être positif.")
        c = self.trouver_compte(numero)
        if not c:
            raise ValueError("Compte introuvable.")

        now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        cur = self.conn.cursor()

        # Met à jour le solde : solde = solde + montant
        cur.execute("UPDATE comptes SET solde=solde+? WHERE numero=?", (montant, numero))

        # Enregistre la transaction dans l'historique
        cur.execute(
            "INSERT INTO transactions(compte_num,date_op,type_op,montant,note) VALUES (?,?,?,?,?)",
            (numero, now, "Dépôt", montant, note)
        )
        self.conn.commit()

    def retirer(self, numero, montant, note=""):
        """
        Débite un compte (retire de l'argent).
        - numero  : numéro du compte à débiter
        - montant : somme à retirer (doit être > 0 et ≤ solde)
        - note    : commentaire optionnel
        """
        if montant <= 0:
            raise ValueError("Le montant doit être positif.")
        c = self.trouver_compte(numero)
        if not c:
            raise ValueError("Compte introuvable.")

        # Vérifie qu'il y a assez d'argent sur le compte
        if montant > c["solde"]:
            raise ValueError("Solde insuffisant.")

        now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        cur = self.conn.cursor()

        # Diminue le solde : solde = solde - montant
        cur.execute("UPDATE comptes SET solde=solde-? WHERE numero=?", (montant, numero))

        # Le montant est enregistré en NÉGATIF dans l'historique (c'est un débit)
        cur.execute(
            "INSERT INTO transactions(compte_num,date_op,type_op,montant,note) VALUES (?,?,?,?,?)",
            (numero, now, "Retrait", -montant, note)
        )
        self.conn.commit()

    def virer(self, num_src, num_dst, montant, note=""):
        """
        Effectue un virement d'un compte vers un autre.
        C'est une opération "atomique" : soit tout réussit, soit rien ne change.
        - num_src : numéro du compte qui envoie l'argent (source)
        - num_dst : numéro du compte qui reçoit l'argent (destination)
        - montant : somme à transférer
        - note    : libellé du virement (optionnel)
        """
        # Vérifications préalables
        if num_src == num_dst:
            raise ValueError("Les deux comptes doivent être différents.")

        src = self.trouver_compte(num_src)
        dst = self.trouver_compte(num_dst)

        if not src:
            raise ValueError("Compte source introuvable.")
        if not dst:
            raise ValueError("Compte destinataire introuvable.")
        if montant <= 0:
            raise ValueError("Le montant doit être positif.")
        if montant > src["solde"]:
            raise ValueError("Solde insuffisant sur le compte source.")

        now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        cur = self.conn.cursor()

        # Débite le compte source
        cur.execute("UPDATE comptes SET solde=solde-? WHERE numero=?", (montant, num_src))

        # Crédite le compte destinataire
        cur.execute("UPDATE comptes SET solde=solde+? WHERE numero=?", (montant, num_dst))

        # Prépare les libellés des transactions
        lib_src = f"Virement vers {num_dst}" + (f" – {note}" if note else "")
        lib_dst = f"Virement de {num_src}"   + (f" – {note}" if note else "")

        # Enregistre la transaction côté SOURCE (montant négatif)
        cur.execute(
            "INSERT INTO transactions(compte_num,date_op,type_op,montant,note) VALUES (?,?,?,?,?)",
            (num_src, now, "Virement émis", -montant, lib_src)
        )

        # Enregistre la transaction côté DESTINATION (montant positif)
        cur.execute(
            "INSERT INTO transactions(compte_num,date_op,type_op,montant,note) VALUES (?,?,?,?,?)",
            (num_dst, now, "Virement reçu", montant, lib_dst)
        )

        # Un seul commit pour les deux opérations : si l'une échoue, rien n'est sauvegardé
        self.conn.commit()

    # ── CONSULTATION DE L'HISTORIQUE ───────────────────────────

    def historique(self, numero, limit=500):
        """
        Retourne toutes les transactions d'un compte donné.
        - numero : numéro du compte
        - limit  : nombre maximum de transactions à retourner
        Les transactions sont triées par ordre chronologique (les plus anciennes en premier).
        """
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM transactions WHERE compte_num=? ORDER BY id",
            (numero,)
        )
        return cur.fetchall()

    def toutes_transactions(self, limit=50):
        """
        Retourne les dernières transactions de TOUS les comptes.
        Utilisé pour le tableau de bord (page d'accueil).
        La jointure (JOIN) permet de récupérer aussi le nom du titulaire.
        """
        cur = self.conn.cursor()
        cur.execute(
            """SELECT t.*, c.titulaire
               FROM transactions t
               JOIN comptes c ON t.compte_num = c.numero
               ORDER BY t.id DESC LIMIT ?""",
            (limit,)
        )
        return cur.fetchall()

    # ── EXPORT CSV ─────────────────────────────────────────────

    def exporter_historique_csv(self, numero, filepath):
        """
        Exporte l'historique d'un compte dans un fichier CSV.
        - numero   : numéro du compte
        - filepath : chemin où enregistrer le fichier (ex: "C:/Users/.../historique.csv")
        Le fichier est lisible par Excel ou LibreOffice Calc.
        encoding="utf-8-sig" : ajoute un marqueur BOM pour que les accents s'affichent bien dans Excel
        """
        rows = self.historique(numero)

        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")  # ";" comme séparateur (standard français)

            # Écriture de l'en-tête du tableau
            w.writerow(["#", "Date", "Type", "Montant (DA)", "Note"])

            # Écriture de chaque ligne de transaction
            for i, r in enumerate(rows, 1):
                w.writerow([i, r["date_op"], r["type_op"],
                            f"{r['montant']:.2f}", r["note"]])
    # ── DONNÉES POUR LES GRAPHIQUES ────────────────────────────

    def solde_evolution(self, numero):
        """
        Calcule l'évolution du solde dans le temps pour un compte.
        Retourne une liste de tuples (date, solde_cumulé).
        Le solde cumulé = somme de tous les montants jusqu'à ce point.
        Utilisé par le graphique "Évolution du solde" dans Matplotlib.
        """
        rows = self.historique(numero)
        solde = 0.0
        pts = []  # pts = liste de points (date, solde)

        for r in rows:
            solde += r["montant"]  # On cumule les montants (+ pour dépôt, - pour retrait)
            pts.append((r["date_op"], solde))

        return pts

    def close(self):
        """Ferme proprement la connexion à la base de données."""
        self.conn.close()


# ─────────────────────────────────────────────────────────────
#  FONCTIONS UTILITAIRES (WIDGETS RÉUTILISABLES)
# ─────────────────────────────────────────────────────────────

def make_entry(parent, textvariable=None, show="", width=28):
    """
    Crée un champ de saisie (zone de texte) stylisé.
    - parent       : le conteneur parent (frame, fenêtre...)
    - textvariable : variable tkinter liée au contenu du champ
    - show         : si show="*", affiche des étoiles (pour les mots de passe)
    - width        : largeur du champ en nombre de caractères
    """
    return tk.Entry(parent,
                    textvariable=textvariable,
                    show=show,
                    width=width,
                    bg=C["navy_light"],        # Fond bleu foncé
                    fg=C["off_white"],          # Texte blanc cassé
                    insertbackground=C["gold"], # Curseur doré
                    relief="flat",              # Pas de relief 3D
                    font=FONT_LABEL,
                    highlightthickness=1,       # Bordure fine
                    highlightbackground=C["border"],  # Bordure normale
                    highlightcolor=C["gold"])   # Bordure dorée quand sélectionné


def make_btn(parent, text, command, primary=True, width=20, icon=""):
    """
    Crée un bouton stylisé avec effet de survol (hover).
    - parent  : conteneur parent
    - text    : texte affiché sur le bouton
    - command : fonction appelée quand on clique
    - primary : True = bouton doré (action principale), False = bouton discret
    - width   : largeur du bouton
    - icon    : icône optionnelle affichée avant le texte
    """
    # Ajoute l'icône si fournie
    label = f"{icon}  {text}" if icon else text

    # Choisit les couleurs selon le type de bouton
    if primary:
        bg, fg, hbg = C["gold"], C["navy"], C["gold_light"]    # Doré
    else:
        bg, fg, hbg = C["navy_light"], C["off_white"], C["navy_hover"]  # Discret

    btn = tk.Button(parent, text=label, command=command,
                    bg=bg, fg=fg,
                    activebackground=hbg, activeforeground=fg,
                    font=FONT_BTN, relief="flat",
                    cursor="hand2",   # Le curseur devient une main au survol
                    width=width, pady=7, bd=0)

    # Effet hover : change la couleur quand la souris entre/sort
    btn.bind("<Enter>", lambda e: btn.config(bg=hbg))  # Souris entre → couleur hover
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))   # Souris sort → couleur normale

    return btn


def separator(parent, bg=None):
    """
    Crée une ligne horizontale de séparation (1 pixel de hauteur).
    Utilisée pour séparer visuellement les sections.
    """
    return tk.Frame(parent, bg=bg or C["border"], height=1)


def tooltip(widget, text):
    """
    Affiche une petite bulle d'aide (tooltip) quand on survole un widget.
    - widget : l'élément sur lequel afficher le tooltip
    - text   : le texte de la bulle d'aide
    """
    tip = None  # Référence à la fenêtre tooltip (None = pas visible)

    def show(e):
        """Crée et affiche la bulle d'aide."""
        nonlocal tip
        tip = tk.Toplevel(widget)           # Fenêtre flottante
        tip.wm_overrideredirect(True)       # Pas de barre de titre
        tip.wm_geometry(f"+{e.x_root+10}+{e.y_root+10}")  # Position près du curseur
        tk.Label(tip, text=text, bg=C["gold"], fg=C["navy"],
                 font=FONT_SMALL, padx=6, pady=4).pack()

    def hide(e):
        """Détruit la bulle d'aide quand la souris part."""
        nonlocal tip
        if tip:
            tip.destroy()
            tip = None

    widget.bind("<Enter>", show)  # Souris entre → afficher
    widget.bind("<Leave>", hide)  # Souris sort → cacher


# ─────────────────────────────────────────────────────────────
#  CLASSE PRINCIPALE – FENÊTRE DE L'APPLICATION
# ─────────────────────────────────────────────────────────────

class BanquePyApp(tk.Tk):
    """
    Classe principale de l'application. Elle hérite de tk.Tk,
    ce qui signifie qu'ELLE EST la fenêtre principale.

    Elle contient :
    - La base de données (self.db)
    - L'en-tête (logo, titre, horloge)
    - La barre latérale de navigation (sidebar)
    - La zone de contenu où s'affichent les différentes pages
    - La barre de statut en bas
    """

    def __init__(self):
        """
        Initialisation de l'application :
        1. Crée la fenêtre principale
        2. Ouvre la connexion à la base de données
        3. Construit toute l'interface graphique
        4. Affiche la page d'accueil
        """
        super().__init__()  # Initialise tk.Tk (la fenêtre de base)

        # Connexion à la base de données SQLite
        self.db = Database()

        # Numéro du compte actuellement sélectionné (None = aucun)
        self.compte_actif_num = None

        # Configuration de la fenêtre principale
        self.title("BanquePy v2 – Espace Client")
        self.geometry("1100x720")       # Taille initiale : 1100 × 720 pixels
        self.minsize(950, 620)          # Taille minimale autorisée
        self.configure(bg=C["navy"])    # Couleur de fond
        self.resizable(True, True)      # Redimensionnable en largeur et hauteur

        # Quand on ferme la fenêtre, appelle _on_quit au lieu de fermer brutalement
        self.protocol("WM_DELETE_WINDOW", self._on_quit)

        # Construction de l'interface
        self._build_ui()

        # Affiche la page d'accueil au démarrage
        self._show_page("accueil")

    # ── CONSTRUCTION DE L'INTERFACE ───────────────────────────

    def _build_ui(self):
        """
        Construit toute l'interface : en-tête, corps, pages, barre de statut.
        Cette méthode appelle d'autres méthodes spécialisées.
        """
        self._build_header()  # En-tête en haut

        # Corps principal : contient la sidebar + la zone de contenu
        body = tk.Frame(self, bg=C["navy"])
        body.pack(fill="both", expand=True)  # S'étend dans toute la place disponible

        self._build_sidebar(body)   # Menu de gauche

        # Zone de contenu à droite (les pages s'affichent ici)
        self.content_frame = tk.Frame(body, bg=C["navy_mid"])
        self.content_frame.pack(side="left", fill="both", expand=True)

        # Création de toutes les pages (elles se superposent, on n'en voit qu'une à la fois)
        self.pages = {}
        for PageClass in (AccueilPage, CreerComptePage, DepotPage,
                          RetraitPage, VirementPage, SoldePage,
                          HistoriquePage, ListePage, GraphiquesPage):
            page = PageClass(self.content_frame, self)
            self.pages[page.PAGE_ID] = page
            # place() avec relwidth=1, relheight=1 = occupe toute la zone
            page.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._build_statusbar()  # Barre de statut en bas

    def _build_header(self):
        """
        Construit l'en-tête de l'application :
        - Logo "BanquePy" à gauche
        - Titre de la page courante au centre
        - Badge SQLite et horloge à droite
        """
        # Cadre de l'en-tête (hauteur fixe)
        hdr = tk.Frame(self, bg=C["navy"], height=C["header_h"])
        hdr.pack(fill="x")
        hdr.pack_propagate(False)  # Empêche le cadre de rétrécir selon son contenu

        # Zone du logo (même largeur que la sidebar)
        logo_frame = tk.Frame(hdr, bg=C["navy"], width=C["sidebar_w"])
        logo_frame.pack(side="left", fill="y")
        logo_frame.pack_propagate(False)
        tk.Label(logo_frame, text="◈  BanquePy",
                 bg=C["navy"], fg=C["gold"],
                 font=("Georgia", 16, "bold")).pack(side="left", padx=20, pady=15)

        # Séparateur vertical entre le logo et le titre
        tk.Frame(hdr, bg=C["border"], width=1).pack(side="left", fill="y")

        # Titre de la page affichée (mis à jour par _show_page)
        self.header_title = tk.Label(hdr, text="Tableau de bord",
                                     bg=C["navy"], fg=C["off_white"],
                                     font=("Helvetica", 13))
        self.header_title.pack(side="left", padx=24)

        # Badge indiquant que SQLite est utilisé (à droite)
        badge = tk.Label(hdr, text="💾 SQLite",
                         bg=C["navy_light"], fg=C["teal"],
                         font=FONT_SMALL, padx=8, pady=3)
        badge.pack(side="right", padx=8)
        tooltip(badge, f"Base de données : {DB_PATH}")  # Affiche le chemin du fichier

        # Horloge mise à jour chaque seconde
        self.clock_lbl = tk.Label(hdr, bg=C["navy"], fg=C["muted"],
                                  font=FONT_SMALL)
        self.clock_lbl.pack(side="right", padx=16)
        self._tick()  # Démarre l'horloge

        # Ligne de séparation sous l'en-tête
        separator(self, bg=C["border"]).pack(fill="x")

    def _build_sidebar(self, parent):
        """
        Construit la barre de navigation latérale (à gauche).
        Contient :
        - Une mini-carte du compte actif
        - Les boutons de navigation vers chaque page
        """
        sb = tk.Frame(parent, bg=C["navy"], width=C["sidebar_w"])
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)

        # Mini-carte du compte actuellement sélectionné
        self.compte_card = tk.Frame(sb, bg=C["navy_light"], relief="flat")
        self.compte_card.pack(fill="x", padx=14, pady=(18, 6))
        self.compte_lbl = tk.Label(self.compte_card,
                                   text="Aucun compte sélectionné",
                                   bg=C["navy_light"], fg=C["muted"],
                                   font=FONT_SMALL, wraplength=180, justify="left")
        self.compte_lbl.pack(padx=10, pady=10, anchor="w")

        separator(sb).pack(fill="x", padx=14, pady=6)

        # Liste des items de navigation : (identifiant_page, icône, libellé)
        nav_items = [
            ("accueil",     "⌂",  "Tableau de bord"),
            ("creer",       "＋", "Nouveau compte"),
            ("depot",       "↑",  "Dépôt"),
            ("retrait",     "↓",  "Retrait"),
            ("virement",    "⇄",  "Virement"),
            ("solde",       "◎",  "Consulter le solde"),
            ("historique",  "≡",  "Historique"),
            ("graphiques",  "📊", "Graphiques"),
            ("liste",       "▤",  "Tous les comptes"),
        ]

        # Dictionnaire pour mémoriser les boutons (pour changer leur couleur)
        self.nav_buttons = {}

        for page_id, icon, label in nav_items:
            text = f"  {icon}  {label}"  # Texte du bouton avec indentation

            # Crée le bouton de navigation
            btn = tk.Button(sb, text=text, anchor="w",
                            bg=C["navy"], fg=C["off_white"],
                            activebackground=C["navy_hover"],
                            activeforeground=C["gold"],
                            font=FONT_NAV, relief="flat", bd=0,
                            cursor="hand2", padx=14, pady=9,
                            # lambda avec pid=page_id : capture la valeur de page_id à cet instant
                            command=lambda pid=page_id: self._show_page(pid))
            btn.pack(fill="x")

            # Effets visuels au survol
            btn.bind("<Enter>", lambda e, b=btn: b.config(bg=C["navy_hover"]))
            btn.bind("<Leave>", lambda e, b=btn, pid=page_id:
                     # Reste doré si c'est la page active, sinon revient à la couleur normale
                     b.config(bg=C["gold"] if getattr(self, "_current", "") == pid
                              else C["navy"]))

            self.nav_buttons[page_id] = btn

        # Séparateur en bas de la sidebar
        separator(sb).pack(fill="x", padx=14, pady=10, side="bottom")

    def _build_statusbar(self):
        """
        Construit la barre de statut en bas de la fenêtre.
        Elle affiche des messages temporaires après chaque action
        (ex: "✔ Dépôt de 500 DA effectué").
        """
        sb = tk.Frame(self, bg=C["navy"], height=24)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)

        separator(sb, C["border"]).pack(fill="x")

        # Message de statut (mis à jour par set_status)
        self.status_lbl = tk.Label(sb, text="Prêt.",
                                   bg=C["navy"], fg=C["muted"],
                                   font=FONT_SMALL, anchor="w")
        self.status_lbl.pack(side="left", padx=16)

        # Chemin du fichier DB affiché à droite (en très discret)
        tk.Label(sb, text=f"DB : {DB_PATH}",
                 bg=C["navy"], fg=C["border"],
                 font=FONT_SMALL).pack(side="right", padx=16)

    # ── NAVIGATION ENTRE LES PAGES ────────────────────────────

    def _show_page(self, page_id):
        """
        Affiche la page demandée et met à jour le menu de navigation.
        - page_id : identifiant de la page (ex: "accueil", "depot"...)
        """
        self._current = page_id  # Mémorise la page active

        # Correspondance entre les identifiants et les titres affichés
        titles = {
            "accueil":    "Tableau de bord",
            "creer":      "Nouveau compte",
            "depot":      "Dépôt",
            "retrait":    "Retrait",
            "virement":   "Virement",
            "solde":      "Consulter le solde",
            "historique": "Historique des transactions",
            "graphiques": "Graphiques & Statistiques",
            "liste":      "Tous les comptes",
        }

        # Met à jour la couleur des boutons de navigation
        for pid, btn in self.nav_buttons.items():
            if pid == page_id:
                btn.config(bg=C["gold"], fg=C["navy"])     # Bouton actif : doré
            else:
                btn.config(bg=C["navy"], fg=C["off_white"])  # Autres : normal

        # Met à jour le titre dans l'en-tête
        self.header_title.config(text=titles.get(page_id, ""))

        # Amène la page au premier plan (.lift() = "soulève" le widget)
        self.pages[page_id].lift()

        # Appelle on_show() pour que la page se rafraîchisse si besoin
        self.pages[page_id].on_show()

    # ── GESTION DU COMPTE ACTIF ───────────────────────────────

    def set_compte_actif(self, numero):
        """
        Définit le compte actuellement sélectionné et rafraîchit la mini-carte.
        - numero : numéro du compte à définir comme actif
        """
        self.compte_actif_num = numero
        self.refresh_compte_card()

    def refresh_compte_card(self):
        """
        Met à jour la mini-carte du compte actif dans la sidebar.
        Si aucun compte n'est sélectionné, affiche un message par défaut.
        """
        if self.compte_actif_num:
            c = self.db.trouver_compte(self.compte_actif_num)
            if c:
                # Affiche le numéro, le nom et le solde du compte
                self.compte_lbl.config(
                    text=f"{c['numero']}\n{c['titulaire']}\n{c['solde']:,.2f} DA",
                    fg=C["gold"])
                return

        # Aucun compte sélectionné
        self.compte_lbl.config(text="Aucun compte sélectionné", fg=C["muted"])

    # ── NOTIFICATIONS ET STATUT ───────────────────────────────

    def notify(self, msg, kind="ok"):
        """
        Affiche une boîte de dialogue ET met à jour la barre de statut.
        - msg  : message à afficher
        - kind : "ok" (succès), "err" (erreur) ou "warn" (avertissement)
        """
        icon = {"ok": "✔", "err": "✘", "warn": "⚠"}
        self.set_status(f"{icon.get(kind,'·')} {msg}")  # Met à jour la barre de statut

        # Affiche la boîte de dialogue appropriée
        if kind == "ok":
            messagebox.showinfo("Succès", f"{icon['ok']}  {msg}")
        elif kind == "err":
            messagebox.showerror("Erreur", f"{icon['err']}  {msg}")
        else:
            messagebox.showwarning("Attention", f"{icon['warn']}  {msg}")

    def set_status(self, msg):
        """
        Affiche un message dans la barre de statut en bas.
        Le message disparaît automatiquement après 6 secondes.
        """
        self.status_lbl.config(text=msg)
        # after(6000, ...) = appelle la fonction après 6000 millisecondes (6 secondes)
        self.after(6000, lambda: self.status_lbl.config(text="Prêt."))

    def _tick(self):
        """
        Met à jour l'horloge dans l'en-tête chaque seconde.
        S'appelle elle-même en boucle grâce à after(1000, ...).
        """
        now = datetime.datetime.now().strftime("%A %d/%m/%Y  %H:%M:%S")
        self.clock_lbl.config(text=now)
        self.after(1000, self._tick)  # Se rappelle dans 1 seconde

    def _on_quit(self):
        """
        Appelée quand l'utilisateur ferme la fenêtre.
        Ferme proprement la connexion SQLite avant de quitter.
        """
        self.db.close()   # Fermeture propre de la base de données
        self.destroy()    # Ferme la fenêtre Tkinter
# ─────────────────────────────────────────────────────────────
#  CLASSE DE BASE DES PAGES
# ─────────────────────────────────────────────────────────────

class BasePage(tk.Frame):
    """
    Classe parente de toutes les pages de l'application.
    Elle hérite de tk.Frame : chaque page est un cadre Tkinter.

    Elle fournit des méthodes utiles communes à toutes les pages :
    - _page_title() : affiche le titre de la page
    - _card()       : crée une "carte" (zone encadrée)
    - _field()      : crée un champ de formulaire avec son label
    - _lookup_compte() : cherche un compte et affiche son info
    - _make_tree()  : crée un tableau avec colonnes et scrollbar

    PAGE_ID : identifiant unique de la page (à redéfinir dans chaque sous-classe)
    """
    PAGE_ID = ""

    def __init__(self, parent, app):
        """
        - parent : le conteneur dans lequel cette page est placée
        - app    : référence à l'application principale (pour accéder à self.db, etc.)
        """
        super().__init__(parent, bg=C["navy_mid"])
        self.app = app   # Référence à l'application principale
        self.db  = app.db  # Raccourci vers la base de données
        self._build()    # Construction du contenu de la page

    def _build(self):
        """À redéfinir dans chaque page pour construire son contenu."""
        pass

    def on_show(self):
        """Appelée chaque fois que la page devient visible. À redéfinir si besoin."""
        pass

    def _page_title(self, text, subtitle=""):
        """
        Affiche le titre principal et un sous-titre optionnel de la page,
        suivis d'une ligne de séparation.
        """
        f = tk.Frame(self, bg=C["navy_mid"])
        f.pack(fill="x", padx=30, pady=(28, 0))

        tk.Label(f, text=text, bg=C["navy_mid"], fg=C["gold"],
                 font=FONT_TITLE).pack(anchor="w")  # anchor="w" = aligné à gauche

        if subtitle:
            tk.Label(f, text=subtitle, bg=C["navy_mid"],
                     fg=C["muted"], font=FONT_SMALL).pack(anchor="w", pady=(2, 0))

        separator(self, C["border"]).pack(fill="x", padx=30, pady=14)

    def _card(self, parent, **kw):
        """
        Crée une "carte" : un cadre avec le fond navy_light.
        Utilisée pour grouper visuellement des éléments (formulaires, tableaux...).
        **kw : arguments supplémentaires passés à tk.Frame
        """
        return tk.Frame(parent, bg=C["navy_light"], relief="flat", bd=0, **kw)

    def _field(self, parent, label, var, show="", hint=""):
        """
        Crée un champ de formulaire complet :
        - Un label au-dessus (ex: "Numéro de compte")
        - Un champ de saisie lié à la variable var
        - Un hint optionnel en dessous (ex: "Nombre ≥ 0")

        Retourne le widget Entry pour pouvoir y ajouter des événements.
        """
        row = tk.Frame(parent, bg=C["navy_light"])
        row.pack(fill="x", pady=(8, 0))

        tk.Label(row, text=label, bg=C["navy_light"],
                 fg=C["muted"], font=FONT_SMALL).pack(anchor="w")

        e = make_entry(row, textvariable=var, show=show)
        e.pack(fill="x", ipady=6)  # ipady : espace intérieur vertical

        if hint:
            tk.Label(row, text=hint, bg=C["navy_light"],
                     fg=C["border"], font=FONT_SMALL).pack(anchor="w")
        return e

    def _lookup_compte(self, num_var, info_lbl):
        """
        Recherche le compte dont le numéro est dans num_var
        et met à jour info_lbl avec le résultat (nom + solde, ou erreur).
        Utilisé dans les pages Dépôt, Retrait et Virement.
        """
        c = self.db.trouver_compte(num_var.get())
        if c:
            info_lbl.config(
                text=f"  ✔  {c['titulaire']}  ·  Solde : {c['solde']:,.2f} DA",
                fg=C["green"])
        else:
            info_lbl.config(text="  ✘  Compte introuvable.", fg=C["red"])

    def _make_tree(self, parent, cols, widths, style_name="Bank", height=10):
        """
        Crée un tableau (Treeview) avec des colonnes définies.
        - cols       : noms des colonnes (ex: ("Numéro", "Titulaire", "Solde"))
        - widths     : largeurs des colonnes en pixels
        - style_name : nom du style ttk (doit être unique par tableau)
        - height     : nombre de lignes visibles sans défilement
        """
        # Configuration du style visuel du tableau
        style = ttk.Style()
        style.configure(f"{style_name}.Treeview",
                        background=C["navy_light"],
                        foreground=C["off_white"],
                        fieldbackground=C["navy_light"],
                        rowheight=28,
                        font=FONT_LABEL,
                        borderwidth=0)
        style.configure(f"{style_name}.Treeview.Heading",
                        background=C["navy"],
                        foreground=C["gold"],
                        font=FONT_SMALL,
                        relief="flat")
        style.map(f"{style_name}.Treeview",
                  background=[("selected", C["navy_hover"])],
                  foreground=[("selected", C["gold"])])

        # Création du tableau
        tree = ttk.Treeview(parent, columns=cols, show="headings",
                            style=f"{style_name}.Treeview", height=height)

        # Configuration de chaque colonne
        for col, w in zip(cols, widths):
            tree.heading(col, text=col)              # En-tête de colonne
            tree.column(col, width=w, anchor="center")  # Largeur et alignement

        # Barre de défilement verticale
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=2, pady=2)

        return tree


# ─────────────────────────────────────────────────────────────
#  PAGE ACCUEIL – TABLEAU DE BORD
# ─────────────────────────────────────────────────────────────

class AccueilPage(BasePage):
    """
    Page d'accueil : affiche un résumé statistique de la banque
    et les dernières transactions effectuées.
    """
    PAGE_ID = "accueil"

    def _build(self):
        """Construction de la page d'accueil."""
        self._page_title("Bienvenue", "Votre espace bancaire personnel · BanquePy v2 + SQLite")

        body = tk.Frame(self, bg=C["navy_mid"])
        body.pack(fill="both", expand=True, padx=30, pady=10)

        # Ligne de métriques (3 cartes côte à côte)
        metrics = tk.Frame(body, bg=C["navy_mid"])
        metrics.pack(fill="x", pady=(0, 16))

        # Crée 3 cartes de statistiques
        self.m_comptes = self._metric(metrics, "Comptes",              "0",         C["gold"])
        self.m_total   = self._metric(metrics, "Total des fonds",      "0,00 DA",   C["green"])
        self.m_ops     = self._metric(metrics, "Opérations aujourd'hui", "0",       C["teal"])

        # Tableau des dernières transactions
        card = self._card(body)
        card.pack(fill="both", expand=True)

        tk.Label(card, text="Dernières opérations",
                 bg=C["navy_light"], fg=C["gold"],
                 font=FONT_HEADING).pack(anchor="w", padx=20, pady=(14, 4))
        separator(card, C["border"]).pack(fill="x", padx=20)

        cols   = ("Compte", "Titulaire", "Type", "Montant", "Date")
        widths = [80, 160, 130, 120, 150]
        self.tree = self._make_tree(card, cols, widths, "Acc", 10)

    def _metric(self, parent, label, value, color):
        """
        Crée une carte de statistique (une "métrique") :
        - Un petit label gris en haut (ex: "Comptes")
        - Une grande valeur colorée en bas (ex: "5")
        """
        f = self._card(parent)
        f.pack(side="left", fill="both", expand=True, padx=(0, 12))

        tk.Label(f, text=label, bg=C["navy_light"], fg=C["muted"],
                 font=FONT_SMALL).pack(anchor="w", padx=16, pady=(14, 2))

        lbl = tk.Label(f, text=value, bg=C["navy_light"], fg=color,
                       font=FONT_VALUE)
        lbl.pack(anchor="w", padx=16, pady=(0, 14))

        return lbl  # On retourne le label pour pouvoir le mettre à jour dans on_show()

    def on_show(self):
        """
        Rafraîchit les données à chaque fois que la page est affichée.
        Calcule les statistiques et remplit le tableau des transactions.
        """
        comptes = self.db.tous_les_comptes()
        total   = sum(c["solde"] for c in comptes)  # Somme de tous les soldes

        # Compte les opérations d'aujourd'hui
        today = datetime.date.today().strftime("%d/%m/%Y")
        rows  = self.db.toutes_transactions(50)
        ops_today = sum(1 for r in rows if r["date_op"].startswith(today))

        # Met à jour les 3 métriques
        self.m_comptes.config(text=str(len(comptes)))
        self.m_total.config(text=f"{total:,.2f} DA")
        self.m_ops.config(text=str(ops_today))

        # Vide le tableau puis le remplit avec les 20 dernières transactions
        for row in self.tree.get_children():
            self.tree.delete(row)

        for r in rows[:20]:
            signe = "+" if r["montant"] >= 0 else ""  # Ajoute "+" pour les montants positifs
            self.tree.insert("", "end", values=(
                r["compte_num"], r["titulaire"], r["type_op"],
                f"{signe}{r['montant']:.2f} DA", r["date_op"]
            ))


# ─────────────────────────────────────────────────────────────
#  PAGE CRÉER COMPTE
# ─────────────────────────────────────────────────────────────

class CreerComptePage(BasePage):
    """Page permettant d'ouvrir un nouveau compte bancaire."""
    PAGE_ID = "creer"

    def _build(self):
        self._page_title("Nouveau compte", "Ouvrir un compte bancaire")

        # Centre le formulaire dans la page
        outer = tk.Frame(self, bg=C["navy_mid"])
        outer.pack(expand=True)

        card = self._card(outer)
        card.pack(padx=30, pady=10, ipadx=20, ipady=10)

        tk.Label(card, text="Informations du titulaire",
                 bg=C["navy_light"], fg=C["gold"],
                 font=FONT_HEADING).pack(anchor="w", padx=20, pady=(18, 0))
        separator(card, C["border"]).pack(fill="x", padx=20, pady=10)

        form = tk.Frame(card, bg=C["navy_light"])
        form.pack(padx=20, fill="x")

        # Variables tkinter liées aux champs de saisie
        # StringVar() = variable qui peut contenir du texte
        self.v_nom   = tk.StringVar()
        self.v_solde = tk.StringVar(value="0")  # Valeur par défaut : 0

        self._field(form, "Nom complet du titulaire", self.v_nom)
        self._field(form, "Solde initial (DA)", self.v_solde, hint="Nombre ≥ 0")

        tk.Frame(card, bg=C["navy_light"], height=16).pack()
        make_btn(card, "Ouvrir le compte", self._creer, width=30, icon="＋").pack(pady=(0, 20))

    def _creer(self):
        """
        Appelée quand on clique sur "Ouvrir le compte".
        Valide les données saisies et crée le compte dans la base de données.
        """
        nom   = self.v_nom.get().strip()    # .strip() supprime les espaces en début/fin
        solde = self.v_solde.get().strip()

        # Validation : le nom ne peut pas être vide
        if not nom:
            self.app.notify("Le nom du titulaire est requis.", "err")
            return  # On arrête ici si invalide

        # Validation : le solde doit être un nombre ≥ 0
        try:
            solde_f = float(solde)  # Convertit le texte en nombre décimal
            if solde_f < 0:
                raise ValueError()  # Déclenche l'erreur si négatif
        except ValueError:
            self.app.notify("Solde initial invalide (nombre ≥ 0).", "err")
            return

        # Tout est valide : création du compte
        try:
            numero = self.db.creer_compte(nom, solde_f)
            self.app.set_compte_actif(numero)  # Active ce compte dans la sidebar

            # Remet les champs à zéro pour la prochaine saisie
            self.v_nom.set("")
            self.v_solde.set("0")

            c = self.db.trouver_compte(numero)
            self.app.notify(
                f"Compte {numero} créé pour {nom}.\n"
                f"Solde initial : {c['solde']:,.2f} DA")
        except Exception as ex:
            self.app.notify(str(ex), "err")


# ─────────────────────────────────────────────────────────────
#  PAGE DÉPÔT
# ─────────────────────────────────────────────────────────────

class DepotPage(BasePage):
    """Page permettant de déposer de l'argent sur un compte."""
    PAGE_ID = "depot"

    def _build(self):
        self._page_title("Dépôt", "Créditer un compte")

        outer = tk.Frame(self, bg=C["navy_mid"])
        outer.pack(expand=True)

        card = self._card(outer)
        card.pack(padx=30, pady=10, ipadx=20, ipady=10)

        tk.Label(card, text="Effectuer un dépôt",
                 bg=C["navy_light"], fg=C["gold"],
                 font=FONT_HEADING).pack(anchor="w", padx=20, pady=(18, 0))
        separator(card, C["border"]).pack(fill="x", padx=20, pady=10)

        form = tk.Frame(card, bg=C["navy_light"])
        form.pack(padx=20, fill="x")

        # Variables liées aux champs
        self.v_num     = tk.StringVar()  # Numéro de compte
        self.v_montant = tk.StringVar()  # Montant à déposer
        self.v_note    = tk.StringVar()  # Note optionnelle

        # Label d'information (affiche le nom et solde du compte trouvé)
        self.info_lbl = tk.Label(card, text="", bg=C["navy_light"],
                                 fg=C["muted"], font=FONT_SMALL)
        self.info_lbl.pack(anchor="w", padx=20)

        # Champ numéro de compte : quand on quitte ou appuie Entrée → recherche auto
        e_num = self._field(form, "Numéro de compte (ex. ACC001)", self.v_num)
        e_num.bind("<FocusOut>", lambda _: self._lookup_compte(self.v_num, self.info_lbl))
        e_num.bind("<Return>",   lambda _: self._lookup_compte(self.v_num, self.info_lbl))

        self._field(form, "Montant à déposer (DA)", self.v_montant)
        self._field(form, "Note (optionnel)", self.v_note)

        tk.Frame(card, bg=C["navy_light"], height=16).pack()
        make_btn(card, "Valider le dépôt", self._deposer, width=30, icon="↑").pack(pady=(0, 20))

    def _deposer(self):
        """Exécute le dépôt après validation des données."""
        num = self.v_num.get().strip().upper()  # Convertit en majuscules
        try:
            montant = float(self.v_montant.get())

            # Appel à la base de données pour effectuer le dépôt
            self.db.deposer(num, montant, self.v_note.get().strip())

            # Met à jour la sidebar avec le nouveau solde
            self.app.set_compte_actif(num)

            c = self.db.trouver_compte(num)
            self.info_lbl.config(
                text=f"  ✔  {c['titulaire']}  ·  Nouveau solde : {c['solde']:,.2f} DA",
                fg=C["green"])

            # Remet les champs à zéro
            self.v_montant.set("")
            self.v_note.set("")

            self.app.notify(f"Dépôt de {montant:,.2f} DA effectué sur {num}.")
        except ValueError as ex:
            self.app.notify(str(ex), "err")

    def on_show(self):
        """Pré-remplit le numéro de compte si un compte est déjà actif."""
        if self.app.compte_actif_num:
            self.v_num.set(self.app.compte_actif_num)
            self._lookup_compte(self.v_num, self.info_lbl)


# ─────────────────────────────────────────────────────────────
#  PAGE RETRAIT
# ─────────────────────────────────────────────────────────────

class RetraitPage(BasePage):
    """Page permettant de retirer de l'argent d'un compte."""
    PAGE_ID = "retrait"

    def _build(self):
        self._page_title("Retrait", "Débiter un compte")

        outer = tk.Frame(self, bg=C["navy_mid"])
        outer.pack(expand=True)

        card = self._card(outer)
        card.pack(padx=30, pady=10, ipadx=20, ipady=10)

        tk.Label(card, text="Effectuer un retrait",
                 bg=C["navy_light"], fg=C["gold"],
                 font=FONT_HEADING).pack(anchor="w", padx=20, pady=(18, 0))
        separator(card, C["border"]).pack(fill="x", padx=20, pady=10)

        form = tk.Frame(card, bg=C["navy_light"])
        form.pack(padx=20, fill="x")

        self.v_num     = tk.StringVar()
        self.v_montant = tk.StringVar()
        self.v_note    = tk.StringVar()

        self.info_lbl = tk.Label(card, text="", bg=C["navy_light"],
                                 fg=C["muted"], font=FONT_SMALL)
        self.info_lbl.pack(anchor="w", padx=20)

        e_num = self._field(form, "Numéro de compte", self.v_num)
        e_num.bind("<FocusOut>", lambda _: self._lookup_compte(self.v_num, self.info_lbl))
        e_num.bind("<Return>",   lambda _: self._lookup_compte(self.v_num, self.info_lbl))

        self._field(form, "Montant à retirer (DA)", self.v_montant)
        self._field(form, "Note (optionnel)", self.v_note)

        tk.Frame(card, bg=C["navy_light"], height=16).pack()
        make_btn(card, "Valider le retrait", self._retirer, width=30, icon="↓").pack(pady=(0, 20))

    def _retirer(self):
        """Exécute le retrait après validation des données."""
        num = self.v_num.get().strip().upper()
        try:
            montant = float(self.v_montant.get())

            # La méthode db.retirer() vérifie que le solde est suffisant
            self.db.retirer(num, montant, self.v_note.get().strip())

            self.app.set_compte_actif(num)

            c = self.db.trouver_compte(num)
            self.info_lbl.config(
                text=f"  ✔  {c['titulaire']}  ·  Nouveau solde : {c['solde']:,.2f} DA",
                fg=C["amber"])  # Orange pour le retrait (attention)

            self.v_montant.set("")
            self.v_note.set("")

            self.app.notify(f"Retrait de {montant:,.2f} DA effectué depuis {num}.")
        except ValueError as ex:
            self.app.notify(str(ex), "err")

    def on_show(self):
        """Pré-remplit le numéro de compte si un compte est déjà actif."""
        if self.app.compte_actif_num:
            self.v_num.set(self.app.compte_actif_num)
            self._lookup_compte(self.v_num, self.info_lbl)
# ─────────────────────────────────────────────────────────────
#  PAGE VIREMENT
# ─────────────────────────────────────────────────────────────

class VirementPage(BasePage):
    """
    Page permettant de faire un virement d'un compte vers un autre.
    Un virement = débit du compte source + crédit du compte destinataire.
    """
    PAGE_ID = "virement"

    def _build(self):
        self._page_title("Virement", "Transfert entre comptes")

        outer = tk.Frame(self, bg=C["navy_mid"])
        outer.pack(expand=True)

        card = self._card(outer)
        card.pack(padx=30, pady=10, ipadx=24, ipady=10)

        tk.Label(card, text="Effectuer un virement",
                 bg=C["navy_light"], fg=C["gold"],
                 font=FONT_HEADING).pack(anchor="w", padx=20, pady=(18, 0))
        separator(card, C["border"]).pack(fill="x", padx=20, pady=10)

        form = tk.Frame(card, bg=C["navy_light"])
        form.pack(padx=20, fill="x")

        self.v_src     = tk.StringVar()  # Numéro du compte qui envoie
        self.v_dst     = tk.StringVar()  # Numéro du compte qui reçoit
        self.v_montant = tk.StringVar()
        self.v_note    = tk.StringVar()

        # Labels d'information pour les deux comptes
        self.info_src = tk.Label(card, text="", bg=C["navy_light"],
                                 fg=C["muted"], font=FONT_SMALL)
        self.info_src.pack(anchor="w", padx=20)

        self.info_dst = tk.Label(card, text="", bg=C["navy_light"],
                                 fg=C["muted"], font=FONT_SMALL)
        self.info_dst.pack(anchor="w", padx=20)

        # Champ source avec vérification automatique
        e_src = self._field(form, "Compte source (qui envoie)", self.v_src)
        e_src.bind("<FocusOut>", lambda _: self._lookup_compte(self.v_src, self.info_src))
        e_src.bind("<Return>",   lambda _: self._lookup_compte(self.v_src, self.info_src))

        # Champ destination avec vérification automatique
        e_dst = self._field(form, "Compte destinataire (qui reçoit)", self.v_dst)
        e_dst.bind("<FocusOut>", lambda _: self._lookup_compte(self.v_dst, self.info_dst))
        e_dst.bind("<Return>",   lambda _: self._lookup_compte(self.v_dst, self.info_dst))

        self._field(form, "Montant (DA)", self.v_montant)
        self._field(form, "Libellé (optionnel)", self.v_note)

        tk.Frame(card, bg=C["navy_light"], height=16).pack()
        make_btn(card, "Effectuer le virement", self._virer, width=30, icon="⇄").pack(pady=(0, 20))

    def _virer(self):
        """
        Exécute le virement.
        La méthode db.virer() est atomique : si une erreur survient,
        aucune modification n'est sauvegardée.
        """
        src = self.v_src.get().strip().upper()
        dst = self.v_dst.get().strip().upper()
        try:
            montant = float(self.v_montant.get())
            self.db.virer(src, dst, montant, self.v_note.get().strip())

            # Met à jour la sidebar avec le compte source
            self.app.set_compte_actif(src)

            c_src = self.db.trouver_compte(src)
            self.info_src.config(
                text=f"  ✔  Nouveau solde {src} : {c_src['solde']:,.2f} DA",
                fg=C["amber"])

            self.v_montant.set("")
            self.v_note.set("")
            self.app.notify(f"Virement de {montant:,.2f} DA : {src} → {dst} effectué.")
        except ValueError as ex:
            self.app.notify(str(ex), "err")

    def on_show(self):
        """Pré-remplit le compte source si un compte est actif."""
        if self.app.compte_actif_num:
            self.v_src.set(self.app.compte_actif_num)
            self._lookup_compte(self.v_src, self.info_src)


# ─────────────────────────────────────────────────────────────
#  PAGE SOLDE
# ─────────────────────────────────────────────────────────────

class SoldePage(BasePage):
    """Page permettant de consulter le solde et les informations d'un compte."""
    PAGE_ID = "solde"

    def _build(self):
        self._page_title("Solde", "Consulter le solde d'un compte")

        outer = tk.Frame(self, bg=C["navy_mid"])
        outer.pack(expand=True)

        card = self._card(outer)
        card.pack(padx=30, pady=10, ipadx=30, ipady=20)

        tk.Label(card, text="Consulter le solde",
                 bg=C["navy_light"], fg=C["gold"],
                 font=FONT_HEADING).pack(anchor="w", padx=20, pady=(18, 0))
        separator(card, C["border"]).pack(fill="x", padx=20, pady=10)

        form = tk.Frame(card, bg=C["navy_light"])
        form.pack(padx=20, fill="x")

        self.v_num = tk.StringVar()
        self._field(form, "Numéro de compte", self.v_num)

        tk.Frame(card, bg=C["navy_light"], height=10).pack()
        make_btn(card, "Consulter", self._afficher, width=28, icon="◎").pack()
        tk.Frame(card, bg=C["navy_light"], height=20).pack()

        # Zone de résultat dynamique : son contenu est recréé à chaque recherche
        self.result_frame = tk.Frame(card, bg=C["navy_light"])
        self.result_frame.pack(fill="x", padx=20, pady=(0, 20))

    def _afficher(self):
        """
        Cherche le compte et affiche ses informations.
        winfo_children() = liste des widgets enfants du frame.
        On les détruit tous avant d'en créer de nouveaux.
        """
        # Vide la zone de résultat
        for w in self.result_frame.winfo_children():
            w.destroy()

        c = self.db.trouver_compte(self.v_num.get())
        if not c:
            tk.Label(self.result_frame, text="Compte introuvable.",
                     bg=C["navy_light"], fg=C["red"],
                     font=FONT_LABEL).pack()
            return

        separator(self.result_frame, C["border"]).pack(fill="x", pady=(0, 14))

        # Données à afficher : (label, valeur, couleur)
        rows = [
            ("Numéro de compte", c["numero"],    C["off_white"]),
            ("Titulaire",        c["titulaire"], C["off_white"]),
            ("Solde disponible", f"{c['solde']:,.2f} DA", C["green"]),
            ("Compte ouvert le", c["cree_le"],   C["muted"]),
        ]

        # Crée une ligne pour chaque information
        for label, val, color in rows:
            row = tk.Frame(self.result_frame, bg=C["navy_light"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, width=20, anchor="w",
                     bg=C["navy_light"], fg=C["muted"],
                     font=FONT_SMALL).pack(side="left")
            tk.Label(row, text=val, anchor="w",
                     bg=C["navy_light"], fg=color,
                     font=FONT_VALUE).pack(side="left", padx=8)

        # Active ce compte dans la sidebar
        self.app.set_compte_actif(c["numero"])

    def on_show(self):
        """Pré-remplit le champ si un compte est actif."""
        if self.app.compte_actif_num:
            self.v_num.set(self.app.compte_actif_num)


# ─────────────────────────────────────────────────────────────
#  PAGE HISTORIQUE
# ─────────────────────────────────────────────────────────────

class HistoriquePage(BasePage):
    """
    Page affichant le détail de toutes les transactions d'un compte.
    Permet aussi d'exporter l'historique en fichier CSV.
    """
    PAGE_ID = "historique"

    def _build(self):
        self._page_title("Historique", "Détail de toutes les transactions")

        # Barre de recherche et boutons en haut
        top = tk.Frame(self, bg=C["navy_mid"])
        top.pack(fill="x", padx=30, pady=(0, 8))

        self.v_num = tk.StringVar()
        tk.Label(top, text="Numéro de compte :", bg=C["navy_mid"],
                 fg=C["muted"], font=FONT_SMALL).pack(side="left")
        make_entry(top, textvariable=self.v_num, width=14).pack(side="left", padx=8)
        make_btn(top, "Afficher", self._afficher, primary=True, width=12, icon="≡").pack(side="left")

        # Bouton d'export CSV
        btn_csv = make_btn(top, "Export CSV", self._export_csv, primary=False, width=12, icon="📁")
        btn_csv.pack(side="left", padx=8)
        tooltip(btn_csv, "Exporter l'historique en fichier CSV (lisible par Excel)")

        # Tableau de l'historique
        card = self._card(self)
        card.pack(fill="both", expand=True, padx=30, pady=(0, 8))

        cols   = ("#", "Date", "Type", "Montant", "Note")
        widths = [40, 160, 160, 120, 200]
        self.tree = self._make_tree(card, cols, widths, "Hist", 12)

        # Label de résumé en bas du tableau
        self.total_lbl = tk.Label(card, text="", bg=C["navy_light"],
                                  fg=C["gold"], font=FONT_LABEL, anchor="e")
        self.total_lbl.pack(fill="x", padx=16, pady=8)

    def _afficher(self):
        """Remplit le tableau avec l'historique du compte saisi."""
        # Vide le tableau
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.total_lbl.config(text="")

        num = self.v_num.get().strip().upper()
        c   = self.db.trouver_compte(num)
        if not c:
            self.app.notify("Compte introuvable.", "err")
            return

        # Récupère et affiche toutes les transactions
        rows = self.db.historique(num)
        for i, h in enumerate(rows, 1):
            signe = "+" if h["montant"] >= 0 else ""  # Signe "+" pour les crédits
            self.tree.insert("", "end", values=(
                i, h["date_op"], h["type_op"],
                f"{signe}{h['montant']:,.2f} DA",
                h["note"] or ""  # Affiche "" si pas de note
            ))

        # Met à jour le résumé en bas
        self.total_lbl.config(
            text=f"Solde actuel : {c['solde']:,.2f} DA   |   {len(rows)} transaction(s)")
        self.app.set_compte_actif(num)

    def _export_csv(self):
        """
        Exporte l'historique du compte dans un fichier CSV.
        Ouvre une boîte de dialogue "Enregistrer sous..." pour choisir l'emplacement.
        """
        num = self.v_num.get().strip().upper()
        if not self.db.trouver_compte(num):
            self.app.notify("Entrez un numéro de compte valide avant l'export.", "warn")
            return

        # Boîte de dialogue "Enregistrer sous"
        path = filedialog.asksaveasfilename(
            title="Enregistrer l'historique CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"historique_{num}.csv")  # Nom de fichier proposé par défaut

        if path:  # Si l'utilisateur n'a pas annulé
            try:
                self.db.exporter_historique_csv(num, path)
                self.app.notify(f"Fichier exporté :\n{path}")
            except Exception as ex:
                self.app.notify(str(ex), "err")

    def on_show(self):
        """Si un compte est actif, l'affiche directement."""
        if self.app.compte_actif_num:
            self.v_num.set(self.app.compte_actif_num)
            self._afficher()
# ─────────────────────────────────────────────────────────────
#  PAGE GRAPHIQUES (Matplotlib)
# ─────────────────────────────────────────────────────────────

class GraphiquesPage(BasePage):
    """
    Page de visualisation graphique avec Matplotlib.
    Propose 3 types de graphiques :
    1. Évolution du solde d'un compte dans le temps (courbe)
    2. Répartition des fonds entre tous les comptes (camembert)
    3. Activité mensuelle de la banque (barres)
    """
    PAGE_ID = "graphiques"

    def _build(self):
        self._page_title("Graphiques", "Visualisation statistique des comptes")

        # Si Matplotlib n'est pas installé, on affiche un message
        if not MATPLOTLIB_OK:
            tk.Label(self,
                     text="⚠ Matplotlib non installé.\n\nPour l'activer :\npip install matplotlib",
                     bg=C["navy_mid"], fg=C["amber"],
                     font=FONT_HEADING).pack(expand=True)
            return

        # Barre de contrôles en haut
        top = tk.Frame(self, bg=C["navy_mid"])
        top.pack(fill="x", padx=30, pady=(0, 10))

        self.v_num = tk.StringVar()
        tk.Label(top, text="Compte (pour évolution solde) :",
                 bg=C["navy_mid"], fg=C["muted"],
                 font=FONT_SMALL).pack(side="left")
        make_entry(top, textvariable=self.v_num, width=12).pack(side="left", padx=8)

        # 3 boutons pour les 3 types de graphiques
        make_btn(top, "Évolution du solde",   self._graph_evolution,   primary=True,  width=18).pack(side="left")
        make_btn(top, "Répartition comptes",  self._graph_repartition, primary=False, width=18).pack(side="left", padx=8)
        make_btn(top, "Activité mensuelle",   self._graph_activite,    primary=False, width=18).pack(side="left")

        # Zone d'affichage du graphique (canvas Matplotlib)
        self.canvas_frame = tk.Frame(self, bg=C["navy_mid"])
        self.canvas_frame.pack(fill="both", expand=True, padx=30, pady=(0, 16))
        self._canvas = None  # Le canvas sera créé quand on génère un graphique

    def _clear_canvas(self):
        """Supprime le graphique précédent avant d'en afficher un nouveau."""
        if self._canvas:
            self._canvas.get_tk_widget().destroy()
            self._canvas = None
        # Supprime aussi tous les autres widgets éventuels
        for w in self.canvas_frame.winfo_children():
            w.destroy()

    def _embed_figure(self, fig):
        """
        Intègre une figure Matplotlib dans la fenêtre Tkinter.
        FigureCanvasTkAgg = adaptateur entre Matplotlib et Tkinter.
        """
        self._clear_canvas()
        self._canvas = FigureCanvasTkAgg(fig, master=self.canvas_frame)
        self._canvas.draw()  # Dessine le graphique
        self._canvas.get_tk_widget().pack(fill="both", expand=True)

    def _mpl_style(self, fig, ax_list):
        """
        Applique le thème de couleurs BanquePy aux graphiques Matplotlib.
        - fig     : la figure (le cadre global du graphique)
        - ax_list : les axes (la zone de dessin du graphique)
        """
        fig.patch.set_facecolor(C["navy_light"])  # Fond de la figure

        # Si un seul axe est passé, on le met dans une liste pour uniformiser
        axes = ax_list if isinstance(ax_list, (list, tuple)) else [ax_list]

        for ax in axes:
            ax.set_facecolor(C["navy_mid"])                    # Fond du graphique
            ax.tick_params(colors=C["muted"])                  # Couleur des graduations
            ax.xaxis.label.set_color(C["muted"])               # Label axe X
            ax.yaxis.label.set_color(C["muted"])               # Label axe Y
            ax.title.set_color(C["gold"])                      # Titre du graphique
            for spine in ax.spines.values():
                spine.set_edgecolor(C["border"])               # Bordures du graphique

    def _graph_evolution(self):
        """
        Trace l'évolution du solde d'un compte dans le temps.
        Type de graphique : courbe avec zone colorée en dessous (area chart).
        """
        if not MATPLOTLIB_OK: return

        num = self.v_num.get().strip().upper()
        pts = self.db.solde_evolution(num)  # Liste de (date, solde_cumulé)

        if not pts:
            self.app.notify("Aucune transaction pour ce compte.", "warn")
            return

        # Sépare les dates et les soldes dans deux listes distinctes
        dates  = [p[0] for p in pts]    # Ex: ["01/01/2025 09:00", "02/01/2025 14:30"...]
        soldes = [p[1] for p in pts]    # Ex: [1000.0, 1500.0, 1200.0...]
        x      = range(len(dates))      # Indices numériques pour l'axe X

        # Création du graphique
        fig, ax = plt.subplots(figsize=(9, 4))

        # Zone colorée sous la courbe (fill_between = remplissage entre 0 et la courbe)
        ax.fill_between(x, soldes, alpha=0.18, color=C["green"])

        # Courbe principale avec points
        ax.plot(x, soldes, color=C["green"], linewidth=2.5, marker="o", markersize=4)

        ax.set_title(f"Évolution du solde – {num}")
        ax.set_ylabel("Solde (DA)")
        ax.set_xticks(x)
        ax.set_xticklabels(dates, rotation=45, ha="right", fontsize=7)
        ax.grid(True, color=C["border"], linestyle="--", linewidth=0.5)

        self._mpl_style(fig, ax)
        fig.tight_layout()  # Ajuste automatiquement les marges
        self._embed_figure(fig)

    def _graph_repartition(self):
        """
        Trace un graphique en camembert (donut) montrant la répartition
        des fonds entre tous les comptes de la banque.
        """
        if not MATPLOTLIB_OK: return

        comptes = self.db.tous_les_comptes()
        if not comptes:
            self.app.notify("Aucun compte enregistré.", "warn")
            return

        # Noms et soldes des comptes (max(solde, 0) pour éviter les négatifs)
        labels = [c["titulaire"] for c in comptes]
        vals   = [max(c["solde"], 0) for c in comptes]

        if sum(vals) == 0:
            self.app.notify("Tous les soldes sont nuls.", "warn")
            return

        # Palette de couleurs pour les tranches du camembert
        palette = ["#C9A84C", "#2ECC8D", "#22A8B8", "#E05555", "#F0A030",
                   "#7A8BAD", "#E8C96A", "#1C3366"]

        fig, ax = plt.subplots(figsize=(7, 5))

        # pie() = graphique camembert
        # autopct = affiche le pourcentage sur chaque tranche
        # wedgeprops = style des tranches (width=0.6 → donut)
        wedges, texts, autotexts = ax.pie(
            vals, labels=labels, autopct="%1.1f%%",
            colors=palette[:len(vals)],
            wedgeprops=dict(width=0.6, edgecolor=C["navy_light"]),
            textprops=dict(color=C["off_white"], fontsize=9))

        # Rend les pourcentages en gras et de couleur marine
        for at in autotexts:
            at.set_color(C["navy"])
            at.set_fontweight("bold")

        ax.set_title("Répartition des fonds par compte")
        self._mpl_style(fig, ax)
        fig.tight_layout()
        self._embed_figure(fig)

    def _graph_activite(self):
        """
        Trace un histogramme montrant le nombre d'opérations par mois.
        Utile pour visualiser les périodes d'activité bancaire.
        """
        if not MATPLOTLIB_OK: return

        rows = self.db.toutes_transactions(500)
        if not rows:
            self.app.notify("Aucune transaction enregistrée.", "warn")
            return

        # Compte le nombre de transactions par mois
        # mois_count : dictionnaire {"2025-01": 5, "2025-02": 3, ...}
        mois_count = {}
        for r in rows:
            try:
                # La date est au format "DD/MM/YYYY HH:MM"
                # On la convertit en "YYYY-MM" pour un tri chronologique correct
                parts = r["date_op"].split("/")
                if len(parts) >= 3:
                    mois = f"{parts[2][:4]}-{parts[1]}"  # Ex: "2025-01"
                else:
                    mois = "?"
            except Exception:
                mois = "?"

            # Incrémente le compteur du mois
            mois_count[mois] = mois_count.get(mois, 0) + 1

        # Trie les mois par ordre chronologique
        mois_sorted = sorted(mois_count.keys())
        counts = [mois_count[m] for m in mois_sorted]

        # Création du graphique en barres
        fig, ax = plt.subplots(figsize=(9, 4))
        bars = ax.bar(mois_sorted, counts, color=C["teal"], width=0.5,
                      edgecolor=C["border"])

        # Affiche le nombre au-dessus de chaque barre
        for bar, count in zip(bars, counts):
            ax.text(bar.get_x() + bar.get_width() / 2,    # Position X : centre de la barre
                    bar.get_height() + 0.1,                # Position Y : juste au-dessus
                    str(count),
                    ha="center", va="bottom",
                    color=C["gold"], fontsize=9)

        ax.set_title("Activité transactionnelle mensuelle")
        ax.set_ylabel("Nombre d'opérations")
        ax.set_xlabel("Mois")
        ax.grid(True, axis="y", color=C["border"], linestyle="--", linewidth=0.5)
        plt.xticks(rotation=30, ha="right")  # Rotation des labels de l'axe X

        self._mpl_style(fig, ax)
        fig.tight_layout()
        self._embed_figure(fig)

    def on_show(self):
        """Pré-remplit le champ avec le compte actif si disponible."""
        if self.app.compte_actif_num and MATPLOTLIB_OK:
            self.v_num.set(self.app.compte_actif_num)


# ─────────────────────────────────────────────────────────────
#  PAGE LISTE – TOUS LES COMPTES
# ─────────────────────────────────────────────────────────────

class ListePage(BasePage):
    """
    Page affichant tous les comptes de la banque dans un tableau.
    Fonctionnalités :
    - Recherche en temps réel par numéro ou nom
    - Sélection d'un compte pour l'activer dans la sidebar
    - Suppression d'un compte avec confirmation
    """
    PAGE_ID = "liste"

    def _build(self):
        self._page_title("Tous les comptes", "Vue globale de la banque")

        # Barre d'outils : recherche + bouton supprimer
        top = tk.Frame(self, bg=C["navy_mid"])
        top.pack(fill="x", padx=30, pady=(0, 8))

        tk.Label(top, text="🔍 Rechercher :", bg=C["navy_mid"],
                 fg=C["muted"], font=FONT_SMALL).pack(side="left")

        # Variable de recherche : quand elle change, _filtrer() est appelée automatiquement
        self.v_recherche = tk.StringVar()
        # trace_add("write", ...) = exécute la fonction à chaque modification du texte
        self.v_recherche.trace_add("write", lambda *_: self._filtrer())
        make_entry(top, textvariable=self.v_recherche, width=22).pack(side="left", padx=8)

        # Bouton de suppression (à droite)
        btn_del = make_btn(top, "Supprimer le compte sélectionné",
                           self._supprimer, primary=False, width=26, icon="✘")
        btn_del.pack(side="right")
        tooltip(btn_del, "Supprime définitivement le compte et tout son historique")

        # Tableau de tous les comptes
        card = self._card(self)
        card.pack(fill="both", expand=True, padx=30, pady=(0, 8))

        cols   = ("Numéro", "Titulaire", "Solde (DA)", "Nb opérations", "Ouvert le")
        widths = [90, 200, 140, 110, 140]
        self.tree = self._make_tree(card, cols, widths, "List", 12)

        # Quand on clique sur une ligne → active le compte
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Pied de tableau : résumé
        self.footer = tk.Label(card, text="", bg=C["navy_light"],
                               fg=C["muted"], font=FONT_SMALL, anchor="e")
        self.footer.pack(fill="x", padx=16, pady=8)

    def on_show(self):
        """Rafraîchit la liste à chaque fois que la page est affichée."""
        self._filtrer()

    def _filtrer(self):
        """
        Remplit le tableau avec les comptes filtrés selon le terme de recherche.
        Si le champ de recherche est vide, affiche tous les comptes.
        """
        # Vide le tableau
        for row in self.tree.get_children():
            self.tree.delete(row)

        terme = self.v_recherche.get().strip()

        # Choisit la bonne méthode selon s'il y a un terme ou non
        if terme:
            comptes = self.db.rechercher_comptes(terme)
        else:
            comptes = self.db.tous_les_comptes()

        total = 0.0
        for c in comptes:
            # Compte le nombre de transactions pour ce compte
            nb = len(self.db.historique(c["numero"]))

            self.tree.insert("", "end", values=(
                c["numero"],
                c["titulaire"],
                f"{c['solde']:,.2f}",   # Formatage avec séparateur de milliers
                nb,
                c["cree_le"]
            ))
            total += c["solde"]

        # Met à jour le résumé en bas
        n = len(comptes)
        self.footer.config(text=f"{n} compte(s)  ·  Total : {total:,.2f} DA")

    def _on_select(self, _):
        """
        Appelée quand l'utilisateur clique sur une ligne du tableau.
        Active le compte sélectionné dans la sidebar.
        """
        sel = self.tree.selection()  # Récupère la (ou les) ligne(s) sélectionnée(s)
        if not sel:
            return

        # Récupère les valeurs de la première ligne sélectionnée
        vals = self.tree.item(sel[0])["values"]

        # La première colonne contient le numéro de compte
        self.app.set_compte_actif(str(vals[0]))

    def _supprimer(self):
        """
        Supprime le compte sélectionné après confirmation de l'utilisateur.
        Demande une confirmation car l'action est irréversible.
        """
        sel = self.tree.selection()
        if not sel:
            self.app.notify("Sélectionnez un compte dans la liste.", "warn")
            return

        vals = self.tree.item(sel[0])["values"]
        num  = str(vals[0])  # Numéro du compte à supprimer
        tit  = str(vals[1])  # Nom du titulaire (pour le message)

        # Boîte de confirmation : l'utilisateur doit dire "Oui" pour continuer
        ok = messagebox.askyesno(
            "Confirmation",
            f"Supprimer définitivement le compte {num} ({tit}) ?\n"
            f"Tout l'historique sera perdu.",
            icon="warning")

        if ok:  # L'utilisateur a confirmé
            self.db.supprimer_compte(num)

            # Si le compte supprimé était le compte actif, on le désélectionne
            if self.app.compte_actif_num == num:
                self.app.compte_actif_num = None
                self.app.refresh_compte_card()

            self._filtrer()  # Rafraîchit le tableau
            self.app.set_status(f"Compte {num} supprimé.")


# ─────────────────────────────────────────────────────────────
#  POINT D'ENTRÉE DU PROGRAMME
# ─────────────────────────────────────────────────────────────

# Ce bloc s'exécute uniquement quand on lance ce fichier directement
# (pas quand il est importé dans un autre script)
if __name__ == "__main__":
    # Crée et démarre l'application
    app = BanquePyApp()

    # mainloop() = boucle principale de Tkinter.
    # Elle attend les événements (clics, frappes...) et y répond.
    # Le programme reste ouvert jusqu'à la fermeture de la fenêtre.
    app.mainloop()