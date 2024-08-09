import sys
import sqlite3, ctypes, os
from datetime import datetime, date, timedelta
from PyQt6.QtWidgets import (QApplication, QMenuBar, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QListWidget, QPushButton, QLabel, QLineEdit, QFormLayout, 
                             QDateEdit, QFileDialog, QMessageBox, QTabWidget, QComboBox, QDialog, QDialogButtonBox, QSpinBox, QSystemTrayIcon, QMenu, QTableWidget, QTableWidgetItem, QListWidgetItem)
from PyQt6.QtCore import Qt, QDate, QBuffer, QByteArray, QIODevice, QTimer
from PyQt6.QtGui import QPixmap, QIcon, QImage, QColor, QAction
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import csv
from openpyxl import Workbook

# Nasconde il prompt dei comandi
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# Funzione che converte una data in stringa formato ISO per memorizzazione nel database
def adapt_date(val):
    return val.isoformat()

# Funzione che converte una stringa ISO in un oggetto data
def convert_date(val):
    if isinstance(val, str):
        return date.fromisoformat(val)
    elif isinstance(val, date):
        return val
    elif isinstance(val, bytes):
        return date.fromisoformat(val.decode())
    else:
        raise ValueError("Valore non valido per la data")

# Registra le funzioni di conversione per la data
sqlite3.register_adapter(date, adapt_date)
sqlite3.register_converter("DATE", convert_date)

class Database:
    """
    Classe per gestire il database delle valvole.

    Attributes:
        conn (sqlite3.Connection): Connessione al database.
        cursor (sqlite3.Cursor): Cursor per eseguire le query.
        alerts_paused (bool): Indica se le notifiche sono state messe in pausa.
        pause_end_date (date): Data di fine della pausa delle notifiche.
    """

    def __init__(self):
        """
        Inizializza il database e crea le tabelle se non esistono.
        """
        try:
            self.conn = sqlite3.connect('valves.db', detect_types=sqlite3.PARSE_DECLTYPES)
            self.cursor = self.conn.cursor()
            self.cursor.execute('''CREATE TABLE IF NOT EXISTS valves
                                (id TEXT PRIMARY KEY,
                                 costruttore TEXT,
                                 tag TEXT,
                                 posizione TEXT,
                                 nominal_pressure TEXT,
                                 inlet_diameter TEXT,
                                 outlet_diameter TEXT,
                                 last_collaud_date DATE,
                                 years_until_collaud INTEGER,
                                 avviso_anticipo INTEGER)''')
            self.cursor.execute('''CREATE TABLE IF NOT EXISTS valve_images
                                (id INTEGER PRIMARY KEY,
                                 valve_id TEXT,
                                 image BLOB)''')
            self.conn.commit()
            self.alerts_paused = False
            self.pause_end_date = None
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")

    def close(self):
        """
        Chiude la connessione al database.
        """
        try:
            self.conn.close()
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")

    def get_valves(self):
        """
        Restituisce la lista delle valvole.

        Returns:
            list: Lista delle valvole.
        """
        try:
            self.cursor.execute("SELECT * FROM valves")
            rows = self.cursor.fetchall()
            valves = []
            for row in rows:
                valve = list(row)
                valve[7] = convert_date(valve[7])  # Converti solo la colonna della data
                valves.append(tuple(valve))
            return valves
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")
            return []

    def get_valve(self, id):
        """
        Restituisce la valvola con l'ID specificato.

        Args:
            id (str): ID della valvola.

        Returns:
            tuple: Valvola con l'ID specificato.
        """
        try:
            self.cursor.execute("SELECT * FROM valves WHERE id=?", (id,))
            valve = self.cursor.fetchone()
            if valve:
                self.cursor.execute("SELECT image FROM valve_images WHERE valve_id=?", (id,))
                images = [row[0] for row in self.cursor.fetchall()]
                return valve + (images,)
            else:
                return None
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")
            return None

    def insert_valve(self, valve):
        """
        Inserisce una nuova valvola nel database.

        Args:
            valve (tuple): Valvola da inserire.

        Returns:
            bool: True se l'inserimento è stato eseguito con successo, False altrimenti.
        """
        try:
            self.cursor.execute("SELECT * FROM valves WHERE id=?", (valve[0],))
            if self.cursor.fetchone():
                return False
            self.cursor.execute("""INSERT INTO valves (id, costruttore, tag, posizione, nominal_pressure, inlet_diameter, outlet_diameter, last_collaud_date, years_until_collaud, avviso_anticipo)
                VALUES (?,?,?,?,?,?,?,?,?,?)""", valve)
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")
            return False

    def update_valve(self, id, valve):
        """
        Aggiorna la valvola con l'ID specificato.

        Args:
            id (str): ID della valvola.
            valve (tuple): Valvola da aggiornare.
        """
        try:
            images = valve[-1]
            if images:
                images = [sqlite3.Binary(image) for image in images]
            else:
                images = []
            self.cursor.execute("""UPDATE valves SET costruttore=?, tag=?, posizione=?, nominal_pressure=?, inlet_diameter=?, outlet_diameter=?, last_collaud_date=?, years_until_collaud=?, avviso_anticipo=?
                WHERE id=?""", valve[:-1] + (id,))
            self.cursor.execute("DELETE FROM valve_images WHERE valve_id=?", (id,))
            for image in images:
                self.cursor.execute("INSERT INTO valve_images (valve_id, image) VALUES (?,?)", (id, image))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")

    def delete_valve(self, id):
        """
        Cancella la valvola con l'ID specificato.

        Args:
            id (str): ID della valvola.
        """
        try:
            self.cursor.execute("DELETE FROM valves WHERE id=?", (id,))
            self.cursor.execute("DELETE FROM valve_images WHERE valve_id=?", (id,))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")

    def update_valve_image(self, id, image):
        """
        Aggiorna l'immagine della valvola con l'ID specificato.

        Args:
            id (str): ID della valvola.
            image (bytes): Immagine da aggiornare.
        """
        try:
            self.cursor.execute("INSERT INTO valve_images (valve_id, image) VALUES (?,?)", (id, image))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")

class ExportFormatDialog(QDialog):
    """
    Classe per la finestra di dialogo per la scelta del formato di esportazione.

    Attributes:
        format_combo (QComboBox): Combo box per la scelta del formato di esportazione.
    """

    def __init__(self, parent=None):
        """
        Inizializza la finestra di dialogo.
        """
        super().__init__(parent)
        self.setWindowTitle("Seleziona il formato di esportazione")
        self.layout = QVBoxLayout()

        self.format_combo = QComboBox()
        self.format_combo.addItems(["PDF", "CSV", "Excel"])
        self.layout.addWidget(self.format_combo)

        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addWidget(self.buttons)

        self.setLayout(self.layout)

    def get_selected_format(self):
        """
        Restituisce il formato di esportazione selezionato.

        Returns:
            str: Formato di esportazione selezionato.
        """
        return self.format_combo.currentText()

class ValveManager(QMainWindow):

    def closeEvent(self, event):
        dialog = QDialog(self)
        dialog.setWindowTitle("Chiudi programma")
        layout = QVBoxLayout(dialog)

        # Messaggio di conferma
        label = QLabel("Vuoi chiudere il programma?")
        layout.addWidget(label)

        # Pulsanti per le scelte
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        minimize_button = QPushButton("Minimizza nel tray")

        # Connetti i segnali
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        minimize_button.clicked.connect(lambda: dialog.done(2))  # Usa 2 per minimizzare

        # Aggiungi pulsanti al layout
        button_layout = QHBoxLayout()
        button_layout.addWidget(button_box)
        button_layout.addWidget(minimize_button)
        layout.addLayout(button_layout)

        # Mostra il dialogo
        result = dialog.exec()

        if result == QDialog.DialogCode.Accepted:
            # L'utente ha scelto di chiudere
            self.db.close()
            self.destroy()
            event.accept()
            QApplication.quit()
        elif result == QDialog.DialogCode.Rejected:
            # L'utente ha scelto di annullare
            event.ignore()
        elif result == 2:
            # L'utente ha scelto di minimizzare
            self.db.close()  # Chiudi la connessione al database anche se si minimizza
            self.hide()
            event.ignore()

    def __init__(self):
        """
        Inizializza la finestra principale.
        """
        super().__init__()
        # Crea un menu "File"
        menu = QMenu("File", self)

        # Crea un'azione "Exit"
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(QApplication.quit)

        # Aggiunge l'azione "Exit" al menu "File"
        menu.addAction(exit_action)

        # Crea una barra dei menu
        barra_menu = QMenuBar(self)

        # Aggiunge il menu "File" alla barra dei menu
        barra_menu.addMenu(menu)

        # Imposta la barra dei menu come barra dei menu principale
        self.setMenuBar(barra_menu)
        
        # Crea un menu "Opzioni"
        opzioni_menu = QMenu("Opzioni", self)

        # Crea un'azione "Percorso database"
        percorso_database_action = QAction("Percorso database", self)
        percorso_database_action.triggered.connect(self.modifica_percorso_database)

        # Aggiunge l'azione "Percorso database" al menu "Opzioni"
        opzioni_menu.addAction(percorso_database_action)

        # Aggiunge il menu "Opzioni" alla barra dei menu
        barra_menu.addMenu(opzioni_menu)

        self.setWindowTitle("Gestione Collaudi Valvole di Sicurezza")
        self.setGeometry(100, 100, 1000, 600)
        self.setWindowIcon(QIcon('icona.ico'))

        self.db = Database()
        self.alerts_paused = False
        self.pause_end_date = None
        self.init_ui()
        self.init_tray()
        self.setup_collaud_check()
        self.image_list.resizeEvent = self.resize_image_list

    def modifica_percorso_database(self):
        # Crea una finestra di dialogo per selezionare il percorso del database
        dialog = QFileDialog(self)
        dialog.setWindowTitle("Seleziona il percorso del database")
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        dialog.setOption(QFileDialog.Option.ShowDirsOnly, True)

        # Mostra la finestra di dialogo
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Ottieni il percorso selezionato
            percorso = dialog.selectedFiles()[0]

            # Verifica se il database esiste già nel percorso selezionato
            db_path = os.path.join(percorso, 'valves.db')
            self.db.conn.close()
            self.db.conn = sqlite3.connect(db_path, detect_types=sqlite3.PARSE_DECLTYPES)
            self.db.cursor = self.db.conn.cursor()

            # Crea le tabelle del database
            self.db.cursor.execute('''CREATE TABLE IF NOT EXISTS valves
                                (id TEXT PRIMARY KEY,
                                costruttore TEXT,
                                tag TEXT,
                                posizione TEXT,
                                nominal_pressure TEXT,
                                inlet_diameter TEXT,
                                outlet_diameter TEXT,
                                last_collaud_date DATE,
                                years_until_collaud INTEGER,
                                avviso_anticipo INTEGER)''')
            self.db.cursor.execute('''CREATE TABLE IF NOT EXISTS valve_images
                                (id INTEGER PRIMARY KEY,
                                valve_id TEXT,
                                image BLOB)''')
            self.db.conn.commit()

            # Aggiorna la lista delle valvole
            self.load_valves()

    def resize_image_list(self, event):
        """
        Ridimensiona la lista delle immagini.
        """
        for i in range(self.image_list.count()):
            item = self.image_list.item(i)
            image_label = self.image_list.itemWidget(item)
            if image_label:
                image_label.setFixedSize(self.image_list.width(), self.image_list.height())

    def init_ui(self):
        """
        Inizializza l'interfaccia utente.
        """

        main_layout = QHBoxLayout()

        list_layout = QVBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Cerca valvole...")
        self.search_input.textChanged.connect(self.search_valves)
        list_layout.addWidget(self.search_input)

        ricerca_avanzata_button = QPushButton("Ricerca Avanzata")
        ricerca_avanzata_button.clicked.connect(self.ricerca_avanzata)
        list_layout.addWidget(ricerca_avanzata_button)

        self.valve_list = QListWidget()
        self.valve_list.itemClicked.connect(self.show_valve_details)
        list_layout.addWidget(self.valve_list)

        add_button = QPushButton("Inserisci Valvola")
        add_button.clicked.connect(self.insert_valve)
        list_layout.addWidget(add_button)

        delete_button = QPushButton("Elimina Valvola")
        delete_button.clicked.connect(self.delete_valve)
        list_layout.addWidget(delete_button)

        main_layout.addLayout(list_layout, 1)

        tab_widget = QTabWidget()

        details_widget = QWidget()
        details_layout = QVBoxLayout(details_widget)

        form_layout = QFormLayout()
        self.id_input = QLineEdit()
        self.costruttore_input = QLineEdit()
        self.tag_input = QLineEdit()
        self.posizione_input = QLineEdit()
        self.nominal_pressure_input = QLineEdit()
        self.inlet_diameter_input = QLineEdit()
        self.outlet_diameter_input = QLineEdit()
        self.last_collaud_date_input = QDateEdit()
        self.last_collaud_date_input.setCalendarPopup(True)
        self.years_until_collaud_input = QSpinBox()
        self.years_until_collaud_input.setRange(1, 10)
        self.avviso_anticipo_input = QSpinBox()
        self.avviso_anticipo_input.setRange(1, 365)
        self.avviso_anticipo_input.setValue(90)
        self.image_list = QListWidget()
        self.image_list.itemClicked.connect(self.show_selected_image)
        self.image_list.itemDoubleClicked.connect(self.remove_selected_image)

        form_layout.addRow("Numero Seriale:", self.id_input)
        form_layout.addRow("Costruttore:", self.costruttore_input)
        form_layout.addRow("Tag:", self.tag_input)
        form_layout.addRow("Posizione:", self.posizione_input)
        form_layout.addRow("Pressione di taratura:", self.nominal_pressure_input)
        form_layout.addRow("Diametro ingresso:", self.inlet_diameter_input)
        form_layout.addRow("Diametro uscita:", self.outlet_diameter_input)
        form_layout.addRow("Ultimo collaudo:", self.last_collaud_date_input)
        form_layout.addRow("Anni fino al prossimo collaudo:", self.years_until_collaud_input)
        form_layout.addRow("Avviso scadenza anticipo (giorni):", self.avviso_anticipo_input)
        form_layout.addRow("Immagini:", self.image_list)

        details_layout.addLayout(form_layout)

        button_layout = QHBoxLayout()
        save_button = QPushButton("Salva modifiche")
        save_button.clicked.connect(self.save_valve)
        delete_card_button = QPushButton("Cancella Scheda")
        delete_card_button.clicked.connect(self.prepare_new_valve)
        add_image_button = QPushButton("Aggiungi Immagine")
        add_image_button.clicked.connect(self.add_image)
        remove_image_button = QPushButton("Elimina Immagine")
        remove_image_button.clicked.connect(self.remove_image)
        export_image_button = QPushButton("Esporta Immagine")
        export_image_button.clicked.connect(self.export_image)
        button_layout.addWidget(save_button)
        button_layout.addWidget(delete_card_button)
        button_layout.addWidget(add_image_button)
        button_layout.addWidget(remove_image_button)
        button_layout.addWidget(export_image_button)

        details_layout.addLayout(button_layout)

        tab_widget.addTab(details_widget, "Dettagli Valvola")

        report_widget = QWidget()
        report_layout = QVBoxLayout(report_widget)
        self.report_table = QTableWidget()
        report_layout.addWidget(self.report_table)

        generate_report_button = QPushButton("Genera Report")
        generate_report_button.clicked.connect(self.generate_report)
        report_layout.addWidget(generate_report_button)

        export_report_button = QPushButton("Esporta Report")
        export_report_button.clicked.connect(self.export_report)
        report_layout.addWidget(export_report_button)

        tab_widget.addTab(report_widget, "Report")

        main_layout.addWidget(tab_widget, 2)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        self.load_valves()

    def init_tray(self):
        """
        Inizializza la tray icon.
        """
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon('icona.ico'))
        self.tray_icon.setToolTip("Gestione Collaudi Valvole di Sicurezza")  # Imposta il titolo dell'alert
        self.tray_icon.setVisible(True)

        tray_menu = QMenu()
        show_action = tray_menu.addAction("Mostra")
        show_action.triggered.connect(self.show)
        quit_action = tray_menu.addAction("Esci")
        quit_action.triggered.connect(QApplication.quit)

        pause_menu = tray_menu.addMenu("Pausa Alert")
        pause_day_action = pause_menu.addAction("1 giorno")
        pause_day_action.triggered.connect(lambda: self.pause_alerts(1))
        pause_month_action = pause_menu.addAction("1 mese")
        pause_month_action.triggered.connect(lambda: self.pause_alerts(30))
        pause_year_action = pause_menu.addAction("1 anno")
        pause_year_action.triggered.connect(lambda: self.pause_alerts(365))
        resume_action = pause_menu.addAction("Annulla pausa")
        resume_action.triggered.connect(self.resume_alerts)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def pause_alerts(self, days):
        """
        Mette in pausa le notifiche per il numero di giorni specificato.

        Args:
            days (int): Numero di giorni per cui mettere in pausa le notifiche.
        """
        self.alerts_paused = True
        self.pause_end_date = date.today() + timedelta(days=days)
        self.tray_icon.showMessage("Pausa Alert", f"Gli alert sono stati messi in pausa per {days} giorni.")

    def resume_alerts(self):
        """
        Annulla la pausa delle notifiche.
        """
        self.alerts_paused = False
        self.tray_icon.showMessage("Pausa Alert", "La pausa degli alert è stata annullata.")

    def update_valve_colors(self):
        for i in range(self.valve_list.count()):
            item = self.valve_list.item(i)
            valve_id = item.text().split(":")[0]
            valve = self.db.get_valve(valve_id)
            next_collaud_date = valve[7] + timedelta(days=valve[8]*365)
            today = date.today()
            if next_collaud_date <= today:
                item.setBackground(QColor("red"))  # Rosso se scaduta
            elif (next_collaud_date - today).days <= valve[9]:
                item.setBackground(QColor(204, 153, 0))  # Giallo più scuro se in preavviso
            else:
                item.setBackground(QColor(0, 0, 0, 0))  # Nessun colore di sfondo

    def load_valves(self):
        self.valve_list.clear()
        valves = self.db.get_valves()
        for valve in valves:
            item = QListWidgetItem()
            item.setText(f"{valve[0]}: {valve[2]}")
            self.valve_list.addItem(item)
        self.update_valve_colors()

    def search_valves(self):
        """
        Cerca le valvole in base al testo inserito nella barra di ricerca.
        """
        search_text = self.search_input.text().lower()
        for i in range(self.valve_list.count()):
            item = self.valve_list.item(i)
            if search_text in item.text().lower():
                item.setHidden(False)
            else:
                item.setHidden(True)

    def show_valve_details(self, item):
        """
        Mostra i dettagli della valvola selezionata.

        Args:
            item (QListWidgetItem): Item selezionato.
        """
        valve_id = item.text().split(":")[0]
        valve = self.db.get_valve(valve_id)
        if valve:
            self.id_input.setText(valve[0])
            self.costruttore_input.setText(valve[1])
            self.tag_input.setText(valve[2])
            self.posizione_input.setText(valve[3])
            self.nominal_pressure_input.setText(valve[4])
            self.inlet_diameter_input.setText(valve[5])
            self.outlet_diameter_input.setText(valve[6])
            self.last_collaud_date_input.setDate(QDate(valve[7]))
            self.years_until_collaud_input.setValue(valve[8])
            self.avviso_anticipo_input.setValue(valve[9])
            next_collaud_date = valve[7] + timedelta(days=valve[8]*365)
            for image in valve[10]:
                image_label = QLabel()
                pixmap = QPixmap()
                pixmap.loadFromData(image)
                image_label.setPixmap(pixmap.scaled(100, 100))
                item = QListWidgetItem()
                item.setSizeHint(image_label.size())
                self.image_list.addItem(item)
                self.image_list.setItemWidget(item, image_label)
        self.id_input.setEnabled(False)  # Disabilita la modifica del codice seriale

    def save_valve(self):
        """
        Salva le modifiche alla valvola.
        """
        try:
            valve_id = self.id_input.text()
            costruttore = self.costruttore_input.text()
            tag = self.tag_input.text()
            posizione = self.posizione_input.text()
            nominal_pressure = self.nominal_pressure_input.text()
            inlet_diameter = self.inlet_diameter_input.text()
            outlet_diameter = self.outlet_diameter_input.text()
            last_collaud_date = self.last_collaud_date_input.date().toPyDate()
            years_until_collaud = self.years_until_collaud_input.value()
            avviso_anticipo = self.avviso_anticipo_input.value()

            # Ottieni il codice seriale originale
            original_valve = self.db.get_valve(self.valve_list.currentItem().text().split(':')[0])
            if original_valve:
                original_id = original_valve[0]
            else:
                original_id = None

            # Validazione degli input
            if not valve_id:
                QMessageBox.warning(self, "Errore", "Il codice seriale è obbligatorio.")
                return
            if not costruttore:
                QMessageBox.warning(self, "Errore", "Il costruttore è obbligatorio.")
                return
            if not tag:
                QMessageBox.warning(self, "Errore", "Il tag è obbligatorio.")
                return
            if not posizione:
                QMessageBox.warning(self, "Errore", "La posizione è obbligatoria.")
                return
            if not nominal_pressure:
                QMessageBox.warning(self, "Errore", "La Pressione di taratura è obbligatoria.")
                return
            if not inlet_diameter:
                QMessageBox.warning(self, "Errore", "Il diametro di ingresso è obbligatorio.")
                return
            if not outlet_diameter:
                QMessageBox.warning(self, "Errore", "Il diametro di uscita è obbligatorio.")
                return
            if not last_collaud_date:
                QMessageBox.warning(self, "Errore", "La data dell'ultimo collaudo è obbligatoria.")
                return
            if not years_until_collaud:
                QMessageBox.warning(self, "Errore", "Gli anni fino al prossimo collaudo sono obbligatori.")
                return

            # Controllo se il codice seriale è stato modificato
            if original_id and valve_id!= original_id:
                QMessageBox.warning(self, "Errore", "Il codice seriale non può essere modificato.")
                return

            # Salva le modifiche
            reply = QMessageBox.question(self, 'Conferma salvataggio', 'Sei sicuro di voler salvare le modifiche?',
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)

            if reply == QMessageBox.StandardButton.Yes:
                # Ottieni le immagini
                images = []
                for i in range(self.image_list.count()):
                    item = self.image_list.item(i)
                    image_label = self.image_list.itemWidget(item)
                    if image_label:
                        image = QByteArray()
                        buffer = QBuffer(image)
                        buffer.open(QIODevice.OpenModeFlag.WriteOnly)
                        image_label.pixmap().save(buffer, "PNG")
                        buffer.close()
                        images.append(image)

                self.db.update_valve(original_id, (costruttore, tag, posizione, nominal_pressure, inlet_diameter, outlet_diameter, last_collaud_date, years_until_collaud, avviso_anticipo, images))
                QMessageBox.information(self, 'Modifiche salvate', 'Le modifiche sono state salvate correttamente.')
                # Aggiorna i colori
                self.update_valve_colors()
        except Exception as e:
            print(f"Errore: {e}")

    def prepare_new_valve(self):
        """
        Prepara una nuova valvola.
        """
        self.id_input.clear()
        self.costruttore_input.clear()
        self.tag_input.clear()
        self.posizione_input.clear()
        self.nominal_pressure_input.clear()
        self.inlet_diameter_input.clear()
        self.outlet_diameter_input.clear()
        self.last_collaud_date_input.setDate(QDate.currentDate())
        self.years_until_collaud_input.setValue(1)
        self.avviso_anticipo_input.setValue(90)
        self.image_list.clear()
        self.id_input.setEnabled(True)  # Abilita la modifica del codice seriale

    def insert_valve(self):
        """
        Inserisce una nuova valvola.
        """
        try:
            # Legge i dati dalla scheda
            valve_id = self.id_input.text()
            costruttore = self.costruttore_input.text()
            tag = self.tag_input.text()
            posizione = self.posizione_input.text()
            nominal_pressure = self.nominal_pressure_input.text()
            inlet_diameter = self.inlet_diameter_input.text()
            outlet_diameter = self.outlet_diameter_input.text()
            last_collaud_date = self.last_collaud_date_input.date().toPyDate()
            years_until_collaud = self.years_until_collaud_input.value()
            avviso_anticipo = self.avviso_anticipo_input.value()

            # Validazione degli input
            if not valve_id:
                QMessageBox.warning(self, "Errore", "Il codice seriale è obbligatorio.")
                return
            if not costruttore:
                QMessageBox.warning(self, "Errore", "Il costruttore è obbligatorio.")
                return
            if not tag:
                QMessageBox.warning(self, "Errore", "Il tag è obbligatorio.")
                return
            if not posizione:
                QMessageBox.warning(self, "Errore", "La posizione è obbligatoria.")
                return
            if not nominal_pressure:
                QMessageBox.warning(self, "Errore", "La Pressione di taratura è obbligatoria.")
                return
            if not inlet_diameter:
                QMessageBox.warning(self, "Errore", "Il diametro di ingresso è obbligatorio.")
                return
            if not outlet_diameter:
                QMessageBox.warning(self, "Errore", "Il diametro di uscita è obbligatorio.")
                return
            if not last_collaud_date:
                QMessageBox.warning(self, "Errore", "La data dell'ultimo collaudo è obbligatoria.")
                return
            if not years_until_collaud:
                QMessageBox.warning(self, "Errore", "Gli anni fino al prossimo collaudo sono obbligatori.")
                return

            # Inserisce la valvola nel database
            if self.db.insert_valve((valve_id, costruttore, tag, posizione, nominal_pressure, inlet_diameter, outlet_diameter, last_collaud_date, years_until_collaud, avviso_anticipo)):
                self.load_valves()
            else:
                QMessageBox.warning(self, "Errore", "La valvola con questo ID già esiste.")
        except Exception as e:
            print(f"Errore: {e}")

    def delete_valve(self):
        """
        Cancella la valvola selezionata.
        """
        try:
            valve_id = self.valve_list.currentItem().text().split(":")[0]
            reply = QMessageBox.question(self, 'Conferma eliminazione', f'Sei sicuro di voler eliminare la valvola {valve_id}?',
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)

            if reply == QMessageBox.StandardButton.Yes:
                self.db.delete_valve(valve_id)
                self.load_valves()
        except Exception as e:
            print(f"Errore: {e}")

    def add_image(self):
        """
        Aggiunge un'immagine alla valvola.
        """
        try:
            file_name, _ = QFileDialog.getOpenFileName(self, "Seleziona immagine", "", "Immagini (*.png *.xpm *.jpg)")
            if file_name:
                image = QImage(file_name)
                image_bytes = QByteArray()
                buffer = QBuffer(image_bytes)
                buffer.open(QIODevice.OpenModeFlag.WriteOnly)
                image.save(buffer, "PNG")
                buffer.close()
                self.db.update_valve_image(self.id_input.text(), image_bytes)
                image_label = QLabel()
                pixmap = QPixmap()
                pixmap.loadFromData(image_bytes)
                image_label.setPixmap(pixmap.scaled(100, 100))
                item = QListWidgetItem()
                item.setSizeHint(image_label.size())
                self.image_list.addItem(item)
                self.image_list.setItemWidget(item, image_label)
        except Exception as e:
            print(f"Errore: {e}")

    def show_selected_image(self, item):
        """
        Mostra l'immagine selezionata.

        Args:
            item (QListWidgetItem): Item selezionato.
        """
        try:
            image_label = self.image_list.itemWidget(item)
            if image_label:
                image_label.show()
        except Exception as e:
            print(f"Errore: {e}")

    def remove_selected_image(self, item):
        """
        Rimuove l'immagine selezionata.

        Args:
            item (QListWidgetItem): Item selezionato.
        """
        try:
            reply = QMessageBox.question(self, 'Conferma rimozione', 'Sei sicuro di voler rimuovere l\'immagine?',
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.image_list.takeItem(self.image_list.row(item))
        except Exception as e:
            print(f"Errore: {e}")

    def remove_image(self):
        """
        Rimuove l'immagine selezionata.
        """
        try:
            selected_item = self.image_list.currentItem()
            if selected_item:
                reply = QMessageBox.question(self, 'Conferma rimozione', 'Sei sicuro di voler rimuovere l\'immagine?',
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.Yes:
                    self.image_list.takeItem(self.image_list.row(selected_item))
            else:
                QMessageBox.warning(self, "Errore", "Seleziona un'immagine da rimuovere.")
        except Exception as e:
            print(f"Errore: {e}")

    def export_image(self):
        """
        Esporta l'immagine selezionata.
        """
        try:
            selected_item = self.image_list.currentItem()
            if selected_item:
                image_label = self.image_list.itemWidget(selected_item)
                if image_label:
                    image = image_label.pixmap().toImage()
                    file_name, _ = QFileDialog.getSaveFileName(self, "Salva immagine", "", "Immagini (*.png *.xpm *.jpg)")
                    if file_name:
                        image.save(file_name, "PNG")
            else:
                QMessageBox.warning(self, "Errore", "Seleziona un'immagine da esportare.")
        except Exception as e:
            print(f"Errore: {e}")

    def generate_report(self):
        """
        Genera il report delle valvole.
        """
        try:
            valves = self.db.get_valves()
            self.report_table.setRowCount(len(valves))
            self.report_table.setColumnCount(10)
            self.report_table.setHorizontalHeaderLabels(["ID", "Costruttore", "Tag", "Posizione", "Pressione di taratura", "Diametro ingresso", "Diametro uscita", "Ultimo collaudo", "Prossimo collaudo", "Avviso anticipo"])
            for i, valve in enumerate(valves):
                self.report_table.setItem(i, 0, QTableWidgetItem(str(valve[0])))
                self.report_table.setItem(i, 1, QTableWidgetItem(str(valve[1])))
                self.report_table.setItem(i, 2, QTableWidgetItem(str(valve[2])))
                self.report_table.setItem(i, 3, QTableWidgetItem(str(valve[3])))
                self.report_table.setItem(i, 4, QTableWidgetItem(str(valve[4])))
                self.report_table.setItem(i, 5, QTableWidgetItem(str(valve[5])))
                self.report_table.setItem(i, 6, QTableWidgetItem(str(valve[6])))
                self.report_table.setItem(i, 7, QTableWidgetItem(str(valve[7])))
                next_collaud_date = valve[7] + timedelta(days=valve[8]*365)
                self.report_table.setItem(i, 8, QTableWidgetItem(str(next_collaud_date)))
                self.report_table.setItem(i, 9, QTableWidgetItem(str(valve[9])))

                # Aggiunta delle immagini
                image_label = QLabel()
                images = self.db.get_valve(valve[0])[10]
                if images:
                    pixmap = QPixmap()
                    pixmap.loadFromData(images[0])
                    image_label.setPixmap(pixmap.scaled(100, 100))
                else:
                    image_label.setText("Nessuna immagine")
                self.report_table.setCellWidget(i, 10, image_label)

            self.report_table.resizeColumnsToContents()
        except Exception as e:
            print(f"Errore: {e}")

    def export_report(self):
        """
        Esporta il report delle valvole.
        """
        try:
            dialog = ExportFormatDialog(self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                export_format = dialog.get_selected_format()
                valves = self.db.get_valves()
                if export_format:
                    if export_format == "PDF":
                        self.export_to_pdf(valves)
                    elif export_format == "CSV":
                        self.export_to_csv(valves)
                    elif export_format == "Excel":
                        self.export_to_excel(valves)
            else:
                print("Esportazione annullata")
        except Exception as e:
            print(f"Errore: {e}")

    def export_to_pdf(self, valves):
        """
        Esporta il report in formato PDF.

        Args:
            valves (list): Lista delle valvole.
        """
        try:
            file_name, _ = QFileDialog.getSaveFileName(self, "Salva PDF", "", "PDF Files (*.pdf)")
            if file_name:
                c = canvas.Canvas(file_name, pagesize=letter)
                width, height = letter
                c.drawString(100, height - 100, "Report Valvole di Sicurezza")
                y = height - 150
                for valve in valves:
                    next_collaud_date = valve[7] + timedelta(days=valve[8]*365)
                    valve_details = f"ID: {valve[0]}, Costruttore: {valve[1]}, Tag: {valve[2]}, Posizione: {valve[3]}, Pressione di taratura: {valve[4]}, Diametro ingresso: {valve[5]}, Diametro uscita: {valve[6]}, Ultimo collaudo: {valve[7]}, Prossimo collaudo: {next_collaud_date}, Avviso anticipo: {valve[9]}"
                    c.drawString(100, y, valve_details)
                    y -= 30
                    if y < 100:
                        c.showPage()
                        y = height - 100
                c.save()
        except Exception as e:
            print(f"Errore: {e}")

    def export_to_csv(self, valves):
        """
        Esporta il report in formato CSV.

        Args:
            valves (list): Lista delle valvole.
        """
        try:
            file_name, _ = QFileDialog.getSaveFileName(self, "Salva CSV", "", "CSV Files (*.csv)")
            if file_name:
                with open(file_name, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(["ID", "Costruttore", "Tag", "Posizione", "Pressione di taratura", "Diametro ingresso", "Diametro uscita", "Ultimo collaudo", "Prossimo collaudo", "Avviso anticipo"])
                    for valve in valves:
                        next_collaud_date = valve[7] + timedelta(days=valve[8]*365)
                        writer.writerow([valve[0], valve[1], valve[2], valve[3], valve[4], valve[5], valve[6], valve[7], next_collaud_date, valve[9]])
        except Exception as e:
            print(f"Errore: {e}")

    def export_to_excel(self, valves):
        """
        Esporta il report in formato Excel.

        Args:
            valves (list): Lista delle valvole.
        """
        try:
            file_name, _ = QFileDialog.getSaveFileName(self, "Salva Excel", "", "Excel Files (*.xlsx)")
            if file_name:
                wb = Workbook()
                ws = wb.active
                ws.append(["ID", "Costruttore", "Tag", "Posizione", "Pressione di taratura", "Diametro ingresso", "Diametro uscita", "Ultimo collaudo", "Prossimo collaudo", "Avviso anticipo"])
                for valve in valves:
                    next_collaud_date = valve[7] + timedelta(days=valve[8]*365)
                    ws.append([valve[0], valve[1], valve[2], valve[3], valve[4], valve[5], valve[6], valve[7], next_collaud_date, valve[9]])
                wb.save(file_name)
        except Exception as e:
            print(f"Errore: {e}")

    def giorni_rimanenti(last_collaud_date, years_until_collaud, avviso_anticipo):
        today = date.today()
        next_collaud_date = last_collaud_date + timedelta(days=years_until_collaud*365)
        giorni_rimanenti = (next_collaud_date - today).days
        if giorni_rimanenti < 0:
            return 0
        elif giorni_rimanenti <= avviso_anticipo:
            return giorni_rimanenti
        else:
            return avviso_anticipo

    def setup_collaud_check(self):
        """
        Imposta il controllo della scadenza dei collaudi.
        """
        try:
            timer = QTimer(self)
            timer.timeout.connect(self.check_collauds)
            timer.start(600000)  # Controlla ogni 10 minuti (in millisecondi)
        except Exception as e:
            print(f"Errore: {e}")

    def check_collauds(self):
        if self.alerts_paused and self.pause_end_date is not None and date.today() < self.pause_end_date:
            return
        try:
            self.db.cursor.execute("SELECT id, costruttore, tag, posizione, nominal_pressure, inlet_diameter, outlet_diameter, last_collaud_date, years_until_collaud, avviso_anticipo FROM valves")
            valves = self.db.cursor.fetchall()
            today = date.today()
            for valve in valves:
                next_collaud_date = valve[7] + timedelta(days=valve[8]*365)
                avviso_anticipo = valve[9]
                if next_collaud_date <= today:
                    self.tray_icon.showMessage(
                        "Promemoria Collaudo",
                        f"La valvola {valve[1]} (ID: {valve[0]}) è scaduta.",
                        QSystemTrayIcon.MessageIcon.Critical
                    )
                elif (next_collaud_date - today).days <= avviso_anticipo:
                    giorni_rimanenti = (next_collaud_date - today).days
                    self.tray_icon.showMessage(
                        "Promemoria Collaudo",
                        f"La valvola {valve[1]} (ID: {valve[0]}) deve essere collaudata entro {giorni_rimanenti} giorni.",
                        QSystemTrayIcon.MessageIcon.Warning
                    )
            # Aggiorna i colori
            self.update_valve_colors()
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")

    def ricerca_avanzata(self):
        """
        Crea una finestra di dialogo per la ricerca avanzata.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Ricerca Avanzata")
        dialog.setLayout(QVBoxLayout())

        # Aggiungi i campi di ricerca
        numero_seriale_label = QLabel("Numero Seriale:")
        numero_seriale_input = QLineEdit()
        costruttore_label = QLabel("Costruttore:")
        costruttore_input = QLineEdit()
        tag_label = QLabel("Tag:")
        tag_input = QLineEdit()
        posizione_label = QLabel("Posizione:")
        posizione_input = QLineEdit()
        pressione_nominale_label = QLabel("Pressione di taratura:")
        pressione_nominale_input = QLineEdit()
        diametro_ingresso_label = QLabel("Diametro ingresso:")
        diametro_ingresso_input = QLineEdit()
        diametro_uscita_label = QLabel("Diametro uscita:")
        diametro_uscita_input = QLineEdit()

        # Aggiungi i pulsanti di ricerca e annulla
        ricerca_button = QPushButton("Ricerca")
        annulla_button = QPushButton("Annulla")

        # Connetti i pulsanti alle funzioni di ricerca e annulla
        ricerca_button.clicked.connect(lambda: self.esegui_ricerca_avanzata(numero_seriale_input.text(), costruttore_input.text(), tag_input.text(), posizione_input.text(), pressione_nominale_input.text(), diametro_ingresso_input.text(), diametro_uscita_input.text()))
        annulla_button.clicked.connect(dialog.reject)

        # Aggiungi i campi di ricerca e i pulsanti alla finestra di dialogo
        dialog.layout().addWidget(numero_seriale_label)
        dialog.layout().addWidget(numero_seriale_input)
        dialog.layout().addWidget(costruttore_label)
        dialog.layout().addWidget(costruttore_input)
        dialog.layout().addWidget(tag_label)
        dialog.layout().addWidget(tag_input)
        dialog.layout().addWidget(posizione_label)
        dialog.layout().addWidget(posizione_input)
        dialog.layout().addWidget(pressione_nominale_label)
        dialog.layout().addWidget(pressione_nominale_input)
        dialog.layout().addWidget(diametro_ingresso_label)
        dialog.layout().addWidget(diametro_ingresso_input)
        dialog.layout().addWidget(diametro_uscita_label)
        dialog.layout().addWidget(diametro_uscita_input)
        dialog.layout().addWidget(ricerca_button)
        dialog.layout().addWidget(annulla_button)

        # Mostra la finestra di dialogo
        dialog.exec()

    def esegui_ricerca_avanzata(self, numero_seriale, costruttore, tag, posizione, pressione_nominale, diametro_ingresso, diametro_uscita):
        """
        Esegue la ricerca avanzata.
        """
        try:
            # Ottieni le valvole dal database
            valves = self.db.get_valves()

            # Filtra le valvole in base ai criteri di ricerca
            filtered_valves = []
            for valve in valves:
                if (numero_seriale and numero_seriale not in valve[0]) or \
                (costruttore and costruttore not in valve[1]) or \
                (tag and tag not in valve[2]) or \
                (posizione and posizione not in valve[3]) or \
                (pressione_nominale and pressione_nominale not in valve[4]) or \
                (diametro_ingresso and diametro_ingresso not in valve[5]) or \
                (diametro_uscita and diametro_uscita not in valve[6]):
                    continue
                filtered_valves.append(valve)

            # Aggiorna la lista delle valvole
            self.valve_list.clear()
            for valve in filtered_valves:
                self.valve_list.addItem(f"{valve[0]}: {valve[2]}")
        except Exception as e:
            print(f"Errore: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    manager = ValveManager()
    manager.show()
    sys.exit(app.exec())