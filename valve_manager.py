"""
Gestione Collaudi Valvole di Sicurezza

Questo programma gestisce la manutenzione e il controllo delle valvole di sicurezza.
Consente di inserire, modificare e cancellare le valvole, nonché di esportare i dati in formato PDF, CSV o Excel.
Inoltre, il programma controlla automaticamente la scadenza dei collaudi e invia notifiche tramite il sistema di notifiche del sistema operativo.
"""

import sys
import sqlite3
from datetime import datetime, date, timedelta
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QListWidget, QPushButton, QLabel, QLineEdit, QFormLayout, 
                             QDateEdit, QFileDialog, QMessageBox, QTabWidget, QComboBox, QDialog, QDialogButtonBox, QSpinBox, QSystemTrayIcon, QTextEdit, QMenu, QTableWidget, QTableWidgetItem, QListWidgetItem)
from PyQt6.QtCore import Qt, QDate, QBuffer, QByteArray, QIODevice, QTimer
from PyQt6.QtGui import QPixmap, QIcon, QImage
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import csv
from openpyxl import Workbook

# Funzione per convertire la data in formato ISO
def adapt_date(val):
    """
    Converte la data in formato ISO.

    Args:
        val (date): Data da convertire.

    Returns:
        str: Data in formato ISO.
    """
    return val.isoformat()

# Funzione per convertire la data da formato ISO
def convert_date(val):
    """
    Converte la data da formato ISO.

    Args:
        val (str): Data in formato ISO.

    Returns:
        date: Data convertita.
    """
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
                                 name TEXT,
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
                valve[5] = convert_date(valve[5])  # Converti solo la colonna della data
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
            self.cursor.execute("""INSERT INTO valves (id, name, nominal_pressure, inlet_diameter, outlet_diameter, last_collaud_date, years_until_collaud, avviso_anticipo)
                VALUES (?,?,?,?,?,?,?,?)""", valve)
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
            self.cursor.execute("""UPDATE valves SET name=?, nominal_pressure=?, inlet_diameter=?, outlet_diameter=?, last_collaud_date=?, years_until_collaud=?, avviso_anticipo=?
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
    """
    Classe per la gestione delle valvole.

    Attributes:
        db (Database): Oggetto per la gestione del database.
        alerts_paused (bool): Indica se le notifiche sono state messe in pausa.
        pause_end_date (date): Data di fine della pausa delle notifiche.
    """

    def __init__(self):
        """
        Inizializza la finestra principale.
        """
        super().__init__()
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
        self.name_input = QLineEdit()
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

        form_layout.addRow("Codice seriale:", self.id_input)
        form_layout.addRow("Nome:", self.name_input)
        form_layout.addRow("Pressione nominale:", self.nominal_pressure_input)
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

    def load_valves(self):
        """
        Carica la lista delle valvole.
        """
        self.valve_list.clear()
        valves = self.db.get_valves()
        for valve in valves:
            self.valve_list.addItem(f"{valve[0]}: {valve[1]}")
        self.search_valves()

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
        valve_id = item.text().split(':')[0]
        valve = self.db.get_valve(valve_id)
        if valve:
            self.id_input.setText(valve[0])
            self.name_input.setText(valve[1])
            self.nominal_pressure_input.setText(valve[2])
            self.inlet_diameter_input.setText(valve[3])
            self.outlet_diameter_input.setText(valve[4])
            self.last_collaud_date_input.setDate(QDate(valve[5]))
            self.years_until_collaud_input.setValue(valve[6])
            self.avviso_anticipo_input.setValue(valve[7])
            next_collaud_date = valve[5] + timedelta(days=valve[6]*365)
            for image in valve[8]:
                image_label = QLabel()
                pixmap = QPixmap()
                pixmap.loadFromData(image)
                image_label.setPixmap(pixmap)
                image_label.setScaledContents(True)
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
            name = self.name_input.text()
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
            if not name:
                QMessageBox.warning(self, "Errore", "Il nome è obbligatorio.")
                return
            if not nominal_pressure:
                QMessageBox.warning(self, "Errore", "La pressione nominale è obbligatoria.")
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

                self.db.update_valve(original_id, (name, nominal_pressure, inlet_diameter, outlet_diameter, last_collaud_date, years_until_collaud, avviso_anticipo, images))
                QMessageBox.information(self, 'Modifiche salvate', 'Le modifiche sono state salvate correttamente.')
        except Exception as e:
            print(f"Errore: {e}")

    def prepare_new_valve(self):
        """
        Prepara una nuova valvola.
        """
        self.id_input.clear()
        self.name_input.clear()
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
            name = self.name_input.text()
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
            if not name:
                QMessageBox.warning(self, "Errore", "Il nome è obbligatorio.")
                return
            if not nominal_pressure:
                QMessageBox.warning(self, "Errore", "La pressione nominale è obbligatoria.")
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
            if self.db.insert_valve((valve_id, name, nominal_pressure, inlet_diameter, outlet_diameter, last_collaud_date, years_until_collaud, avviso_anticipo)):
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
            valve_id = self.valve_list.currentItem().text().split(':')[0]
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
            self.report_table.setColumnCount(9)
            self.report_table.setHorizontalHeaderLabels(["ID", "Nome", "Pressione Nominale", "Diametro Ingresso", "Diametro Uscita", "Ultimo Collaudo", "Prossimo Collaudo", "Avviso Anticipo", "Immagini"])
            for i, valve in enumerate(valves):
                self.report_table.setItem(i, 0, QTableWidgetItem(str(valve[0])))
                self.report_table.setItem(i, 1, QTableWidgetItem(str(valve[1])))
                self.report_table.setItem(i, 2, QTableWidgetItem(str(valve[2])))
                self.report_table.setItem(i, 3, QTableWidgetItem(str(valve[3])))
                self.report_table.setItem(i, 4, QTableWidgetItem(str(valve[4])))
                self.report_table.setItem(i, 5, QTableWidgetItem(str(valve[5])))
                next_collaud_date = valve[5] + timedelta(days=valve[6]*365)
                self.report_table.setItem(i, 6, QTableWidgetItem(str(next_collaud_date)))
                self.report_table.setItem(i, 7, QTableWidgetItem(str(valve[7])))

                # Aggiunta delle immagini
                image_label = QLabel()
                images = self.db.get_valve(valve[0])[8]
                if images:
                    pixmap = QPixmap()
                    pixmap.loadFromData(images[0])
                    image_label.setPixmap(pixmap.scaled(100, 100))
                else:
                    image_label.setText("Nessuna immagine")
                self.report_table.setCellWidget(i, 8, image_label)

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
                    next_collaud_date = valve[5] + timedelta(days=valve[6]*365)
                    valve_details = f"ID: {valve[0]}, Nome: {valve[1]}, Pressione Nominale: {valve[2]}, Diametro Ingresso: {valve[3]}, Diametro Uscita: {valve[4]}, Ultimo Collaudo: {valve[5]}, Prossimo Collaudo: {next_collaud_date}, Avviso Anticipo: {valve[7]}"
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
                    writer.writerow(["ID", "Nome", "Pressione Nominale", "Diametro Ingresso", "Diametro Uscita", "Ultimo Collaudo", "Prossimo Collaudo", "Avviso Anticipo"])
                    for valve in valves:
                        next_collaud_date = valve[5] + timedelta(days=valve[6]*365)
                        writer.writerow([valve[0], valve[1], valve[2], valve[3], valve[4], valve[5], next_collaud_date, valve[7]])
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
                ws.append(["ID", "Nome", "Pressione Nominale", "Diametro Ingresso", "Diametro Uscita", "Ultimo Collaudo", "Prossimo Collaudo", "Avviso Anticipo"])
                for valve in valves:
                    next_collaud_date = valve[5] + timedelta(days=valve[6]*365)
                    ws.append([valve[0], valve[1], valve[2], valve[3], valve[4], valve[5], next_collaud_date, valve[7]])
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
        
    def check_collauds(self):
        if self.alerts_paused and self.pause_end_date is not None and date.today() < self.pause_end_date:
            return
        try:
            self.db.cursor.execute("SELECT id, name, last_collaud_date, years_until_collaud, avviso_anticipo FROM valves")
            valves = self.db.cursor.fetchall()
            today = date.today()
            for valve in valves:
                next_collaud_date = valve[2] + timedelta(days=valve[3]*365)
                avviso_anticipo = valve[4]
                if next_collaud_date <= today:
                    self.tray_icon.showMessage(
                        "Promemoria Collaudo",
                        f"La valvola {valve[1]} (ID: {valve[0]}) è scaduta.",
                        QSystemTrayIcon.MessageIcon.Critical
                    )
                    continue
                elif (next_collaud_date - today).days <= avviso_anticipo:
                    self.tray_icon.showMessage(
                        "Promemoria Collaudo",
                        f"La valvola {valve[1]} (ID: {valve[0]}) deve essere collaudata entro {avviso_anticipo} giorni.",
                        QSystemTrayIcon.MessageIcon.Warning
                    )
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")


    def setup_collaud_check(self):
        """
        Imposta il controllo della scadenza dei collaudi.
        """
        try:
            timer = QTimer(self)
            timer.timeout.connect(self.check_collauds)
            timer.start(600)  # Controlla ogni 24 ore (in millisecondi)
        except Exception as e:
            print(f"Errore: {e}")

    def ricerca_avanzata(self):
        """
        Crea una finestra di dialogo per la ricerca avanzata.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Ricerca Avanzata")
        dialog.setLayout(QVBoxLayout())

        # Aggiungi i campi di ricerca
        nome_label = QLabel("Nome:")
        nome_input = QLineEdit()
        pressione_nominale_label = QLabel("Pressione nominale:")
        pressione_nominale_input = QLineEdit()
        diametro_ingresso_label = QLabel("Diametro ingresso:")
        diametro_ingresso_input = QLineEdit()
        diametro_uscita_label = QLabel("Diametro uscita:")
        diametro_uscita_input = QLineEdit()

        # Aggiungi i pulsanti di ricerca e annulla
        ricerca_button = QPushButton("Ricerca")
        annulla_button = QPushButton("Annulla")

        # Connetti i pulsanti alle funzioni di ricerca e annulla
        ricerca_button.clicked.connect(lambda: self.esegui_ricerca_avanzata(nome_input.text(), pressione_nominale_input.text(), diametro_ingresso_input.text(), diametro_uscita_input.text()))
        annulla_button.clicked.connect(dialog.reject)

        # Aggiungi i campi di ricerca e i pulsanti alla finestra di dialogo
        dialog.layout().addWidget(nome_label)
        dialog.layout().addWidget(nome_input)
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

    def esegui_ricerca_avanzata(self, nome, pressione_nominale, diametro_ingresso, diametro_uscita):
        """
        Esegue la ricerca avanzata.
        """
        try:
            # Ottieni le valvole dal database
            valves = self.db.get_valves()

            # Filtra le valvole in base ai criteri di ricerca
            filtered_valves = []
            for valve in valves:
                if (nome and nome not in valve[1]) or \
                (pressione_nominale and pressione_nominale not in valve[2]) or \
                (diametro_ingresso and diametro_ingresso not in valve[3]) or \
                (diametro_uscita and diametro_uscita not in valve[4]):
                    continue
                filtered_valves.append(valve)

            # Aggiorna la lista delle valvole
            self.valve_list.clear()
            for valve in filtered_valves:
                self.valve_list.addItem(f"{valve[0]}: {valve[1]}")
        except Exception as e:
            print(f"Errore: {e}")

    def check_collauds(self):
        if self.alerts_paused and self.pause_end_date is not None and date.today() < self.pause_end_date:
            return
        try:
            self.db.cursor.execute("SELECT id, name, last_collaud_date, years_until_collaud, avviso_anticipo FROM valves")
            valves = self.db.cursor.fetchall()
            today = date.today()
            for valve in valves:
                next_collaud_date = valve[2] + timedelta(days=valve[3]*365)
                avviso_anticipo = valve[4]
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
        except sqlite3.Error as e:
            print(f"Errore di database: {e}")

    def setup_collaud_check(self):
        """
        Imposta il controllo della scadenza dei collaudi.
        """
        try:
            timer = QTimer(self)
            timer.timeout.connect(self.check_collauds)
            timer.start(6000)  # Controlla ogni 24 ore (in millisecondi)
        except Exception as e:
            print(f"Errore: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    manager = ValveManager()
    manager.show()
    sys.exit(app.exec())
