# Outlook/Office 16.x Credential Cleaner

Εργαλείο για τη διαγραφή αποθηκευμένων Windows Credentials που σχετίζονται με το Outlook και Office 16.x.

## Το Πρόβλημα

Το Outlook έχει συχνά πρόβλημα με λογαριασμούς που χρησιμοποιούν OAuth token για authentication (Gmail, Microsoft 365, κ.ά.). Όταν το token αλλάζει ή λήγει, το Outlook δεν μπορεί να συνδεθεί και εμφανίζει σφάλματα authentication. Η λύση είναι να διαγραφούν τα παλιά credentials ώστε το Outlook να ζητήσει νέο login και να πάρει φρέσκο token.

## Χαρακτηριστικά

- GUI με λίστα όλων των Windows Credentials
- Checkboxes για επιλογή ποια θα διαγραφούν
- Αυτόματη προ-επιλογή Outlook/Office credentials
- Preview mode για να δείτε τι θα διαγραφεί πριν το κάνετε
- Progress bar και status

## Εγκατάσταση

```bash
pip install PyQt6
```

## Χρήση

```bash
python outlook_credential_cleaner.py
```

## Δημιουργία EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icons/icon.ico --name "OutlookCredentialCleaner" --add-data "icons;icons" outlook_credential_cleaner.py
```

Το exe θα βρεθεί στο `dist/OutlookCredentialCleaner.exe`

## Developed by

**Διονύσης Πρόκος**

Download: https://github.com/DProkos/OutlookCredentialCleaner/releases/tag/v1.0
