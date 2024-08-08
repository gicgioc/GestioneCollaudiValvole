Il programma "Gestore Collaudi Valvole di Sicurezza" è un'applicazione desktop sviluppata in Python utilizzando le librerie PyQt6, sqlite3 e reportlab. Il programma consente di gestire le valvole di sicurezza, inserendo e modificando le loro caratteristiche, eseguendo ricerche avanzate e generando report in diversi formati.

Funzionalità principali:

Inserimento e modifica valvole: l'utente può inserire nuove valvole e modificare quelle esistenti, specificando il codice seriale, il nome, la pressione nominale, il diametro di ingresso e di uscita, la data dell'ultimo collaudo e gli anni fino al prossimo collaudo.
Ricerca avanzata: l'utente può eseguire ricerche avanzate sulle valvole in base a diversi criteri, come il nome, la pressione nominale e i diametri di ingresso e di uscita.
Generazione report: il programma consente di generare report in diversi formati (PDF, CSV e Excel) contenenti le informazioni sulle valvole.
Controllo scadenza collaudi: il programma verifica automaticamente la scadenza dei collaudi e mostra una notifica se una valvola è prossima alla scadenza o è già scaduta.
Pausa alert: l'utente può mettere in pausa le notifiche per un periodo di tempo specificato.
Guida per la prima esecuzione e compilazione del codice sorgente:

Installazione delle librerie: prima di eseguire il programma, è necessario installare le librerie richieste (PyQt6, sqlite3 e reportlab). Ciò può essere fatto utilizzando pip, il gestore dei pacchetti Python:

pip install pyqt6 sqlite3 reportlab
Esecuzione del programma: una volta installate le librerie, è possibile eseguire il programma utilizzando Python:

python gestore_collaudo_valvole.py
Compilazione del codice sorgente: se si desidera creare un file eseguibile del programma, è possibile utilizzare strumenti come PyInstaller o cx_Freeze.

Nota: il codice sorgente fornito è già pronto per l'uso e non richiede modifiche per la prima esecuzione. Tuttavia, potrebbe essere necessario adattarlo alle specifiche esigenze dell'utente.
