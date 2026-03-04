"""
Outlook/Office 16.x Credential Cleaner - PyQt6 GUI
Deletes ONLY Outlook/Office 16.x related Windows Credentials for CURRENT user.
"""
import subprocess
import sys
import re
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QProgressBar, QTextEdit, QCheckBox, QMessageBox,
    QTreeWidget, QTreeWidgetItem, QHeaderView, QSplitter, QMenuBar, QMenu
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
import os

PATTERNS = [
    r'MicrosoftOffice16',
    r'Office16',
    r'16\.0',
    r'Outlook',
    r'MSOID',
    r'ADAL',
    r'WAM',
    r'OneAuth',
    r'Identity',
    r'microsoftidentity',
    r'AzureAD'
]

def get_cmdkey_targets():
    """Get all credential targets from cmdkey."""
    try:
        result = subprocess.run(['cmdkey', '/list'], capture_output=True, text=True, shell=True)
        raw = result.stdout
    except Exception:
        return []
    
    targets = []
    for line in raw.splitlines():
        match = re.match(r'^\s*(Target|Στόχος)\s*:\s*(.+)\s*$', line, re.IGNORECASE)
        if match:
            targets.append(match.group(2).strip())
    return sorted(set(targets))

def matches_patterns(text, patterns):
    """Check if text matches any pattern."""
    for p in patterns:
        if re.search(p, text, re.IGNORECASE):
            return True
    return False

def delete_credential(target):
    """Delete a single credential target."""
    try:
        result = subprocess.run(['cmdkey', f'/delete:{target}'], capture_output=True, text=True, shell=True)
        return result.returncode == 0
    except Exception:
        return False


class CleanerThread(QThread):
    """Worker thread for credential cleaning."""
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    result = pyqtSignal(str, bool)
    finished_all = pyqtSignal(int, int)

    def __init__(self, targets, whatif=False):
        super().__init__()
        self.targets = targets
        self.whatif = whatif

    def run(self):
        success = 0
        fail = 0
        total = len(self.targets)
        
        for i, target in enumerate(self.targets):
            self.status.emit(f"Processing: {target}")
            self.progress.emit(i + 1, total)
            
            if self.whatif:
                self.result.emit(target, True)
                success += 1
            else:
                ok = delete_credential(target)
                self.result.emit(target, ok)
                if ok:
                    success += 1
                else:
                    fail += 1
        
        self.finished_all.emit(success, fail)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Outlook/Office 16.x Credential Cleaner")
        self.setMinimumSize(700, 600)
        self.all_targets = []
        self.init_ui()
        self.scan_credentials()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        # Title
        title = QLabel("Outlook/Office 16.x Credential Cleaner")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Status label
        self.status_label = QLabel("Scanning credentials...")
        self.status_label.setFont(QFont("Segoe UI", 10))
        layout.addWidget(self.status_label)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Splitter for tree and log
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Tree widget for credentials
        tree_container = QWidget()
        tree_layout = QVBoxLayout(tree_container)
        tree_layout.setContentsMargins(0, 0, 0, 0)
        
        tree_label = QLabel("Select credentials to delete:")
        tree_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        tree_layout.addWidget(tree_label)
        
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Credential Target", "Type"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tree.setFont(QFont("Consolas", 9))
        tree_layout.addWidget(self.tree)
        
        # Select buttons
        select_layout = QHBoxLayout()
        self.select_outlook_btn = QPushButton("Select Outlook/Office")
        self.select_outlook_btn.clicked.connect(self.select_outlook_only)
        select_layout.addWidget(self.select_outlook_btn)
        
        self.select_all_btn = QPushButton("Select All")
        self.select_all_btn.clicked.connect(self.select_all)
        select_layout.addWidget(self.select_all_btn)
        
        self.select_none_btn = QPushButton("Select None")
        self.select_none_btn.clicked.connect(self.select_none)
        select_layout.addWidget(self.select_none_btn)
        
        tree_layout.addLayout(select_layout)
        splitter.addWidget(tree_container)

        # Log area
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setFont(QFont("Consolas", 9))
        self.log_area.setMaximumHeight(150)
        splitter.addWidget(self.log_area)
        
        layout.addWidget(splitter)

        # WhatIf checkbox
        self.whatif_check = QCheckBox("Preview only (WhatIf) - Don't actually delete")
        self.whatif_check.setChecked(True)
        layout.addWidget(self.whatif_check)

        # Buttons
        btn_layout = QHBoxLayout()
        self.scan_btn = QPushButton("🔍 Scan")
        self.scan_btn.clicked.connect(self.scan_credentials)
        btn_layout.addWidget(self.scan_btn)

        self.clean_btn = QPushButton("🧹 Delete Selected")
        self.clean_btn.clicked.connect(self.clean_credentials)
        self.clean_btn.setEnabled(False)
        btn_layout.addWidget(self.clean_btn)

        layout.addLayout(btn_layout)
        
        # Menu bar
        menubar = self.menuBar()
        help_menu = menubar.addMenu("Help")
        about_action = help_menu.addAction("About")
        about_action.triggered.connect(self.show_about)

    def show_about(self):
        QMessageBox.about(
            self,
            "About",
            "<h3>Outlook/Office 16.x Credential Cleaner</h3>"
            "<p>Version 1.0</p>"
            "<p>Developed by <b>Διονύσης Πρόκος</b></p>"
            "<hr>"
            "<p><b>Τι κάνει:</b><br>"
            "Διαγράφει τα αποθηκευμένα Windows Credentials που σχετίζονται "
            "με το Outlook και Office 16.x.</p>"
            "<p><b>Γιατί το χρειάζεστε:</b><br>"
            "Δημιουργήθηκε για να λύσει το πρόβλημα του Outlook με "
            "λογαριασμούς που χρησιμοποιούν OAuth token (Gmail, Microsoft 365, κ.ά.). "
            "Όταν το token λήξει ή αλλάξει, το Outlook δεν μπορεί να συνδεθεί. "
            "Διαγράφοντας τα παλιά credentials, το Outlook θα ζητήσει νέο login "
            "και θα πάρει φρέσκο token.</p>"
            "<p><b>Preview Mode:</b><br>"
            "Όταν είναι ενεργοποιημένο, δείχνει τι θα διαγραφεί χωρίς να "
            "κάνει πραγματική διαγραφή. Απενεργοποιήστε το για να διαγράψετε.</p>"
        )

    def log(self, msg, color=None):
        if color:
            self.log_area.append(f'<span style="color:{color}">{msg}</span>')
        else:
            self.log_area.append(msg)

    def scan_credentials(self):
        self.log_area.clear()
        self.tree.clear()
        self.status_label.setText("Scanning credentials...")
        self.progress_bar.setValue(0)
        
        self.all_targets = get_cmdkey_targets()
        
        for target in self.all_targets:
            item = QTreeWidgetItem(self.tree)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setText(0, target)
            
            is_outlook = matches_patterns(target, PATTERNS)
            item.setText(1, "Outlook/Office" if is_outlook else "Other")
            item.setCheckState(0, Qt.CheckState.Checked if is_outlook else Qt.CheckState.Unchecked)
        
        outlook_count = sum(1 for t in self.all_targets if matches_patterns(t, PATTERNS))
        other_count = len(self.all_targets) - outlook_count
        
        self.log(f"Total credentials: {len(self.all_targets)}", "cyan")
        self.log(f"Outlook/Office: {outlook_count} (pre-selected)", "cyan")
        self.log(f"Other: {other_count}", "cyan")
        
        self.status_label.setText(f"Found {len(self.all_targets)} credentials ({outlook_count} Outlook/Office)")
        self.clean_btn.setEnabled(len(self.all_targets) > 0)
        self.progress_bar.setValue(100)

    def get_selected_targets(self):
        selected = []
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if item.checkState(0) == Qt.CheckState.Checked:
                selected.append(item.text(0))
        return selected

    def select_outlook_only(self):
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            is_outlook = matches_patterns(item.text(0), PATTERNS)
            item.setCheckState(0, Qt.CheckState.Checked if is_outlook else Qt.CheckState.Unchecked)

    def select_all(self):
        for i in range(self.tree.topLevelItemCount()):
            self.tree.topLevelItem(i).setCheckState(0, Qt.CheckState.Checked)

    def select_none(self):
        for i in range(self.tree.topLevelItemCount()):
            self.tree.topLevelItem(i).setCheckState(0, Qt.CheckState.Unchecked)

    def clean_credentials(self):
        targets = self.get_selected_targets()
        if not targets:
            QMessageBox.warning(self, "No Selection", "Please select at least one credential to delete.")
            return
        
        whatif = self.whatif_check.isChecked()
        
        if not whatif:
            reply = QMessageBox.question(
                self, "Confirm Delete",
                f"Are you sure you want to delete {len(targets)} credentials?\n\n"
                "You may need to re-login to affected applications.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        
        self.log_area.clear()
        self.scan_btn.setEnabled(False)
        self.clean_btn.setEnabled(False)
        self.whatif_check.setEnabled(False)
        self.tree.setEnabled(False)
        
        mode = "PREVIEW" if whatif else "DELETING"
        self.log(f"=== {mode} MODE ===", "cyan")
        self.log("")
        
        self.thread = CleanerThread(targets, whatif)
        self.thread.progress.connect(self.on_progress)
        self.thread.status.connect(self.on_status)
        self.thread.result.connect(self.on_result)
        self.thread.finished_all.connect(self.on_finished)
        self.thread.start()

    def on_progress(self, current, total):
        self.progress_bar.setValue(int(current / total * 100))

    def on_status(self, msg):
        self.status_label.setText(msg)

    def on_result(self, target, success):
        if self.whatif_check.isChecked():
            self.log(f"[PREVIEW] Would delete: {target}", "yellow")
        elif success:
            self.log(f"✅ Deleted: {target}", "green")
        else:
            self.log(f"❌ FAILED: {target}", "red")

    def on_finished(self, success, fail):
        self.log("")
        if self.whatif_check.isChecked():
            self.log("=== PREVIEW COMPLETE ===", "yellow")
            self.log("Uncheck 'Preview only' and click Delete to actually delete.", "yellow")
            self.status_label.setText("Preview complete - no changes made")
        elif fail == 0:
            self.log("✅ Done! Re-login to affected apps if prompted.", "green")
            self.status_label.setText(f"Successfully deleted {success} credentials")
        else:
            self.log(f"⚠ Completed with {fail} failures.", "orange")
            self.status_label.setText(f"Deleted {success}, Failed {fail}")
        
        self.scan_btn.setEnabled(True)
        self.clean_btn.setEnabled(True)
        self.whatif_check.setEnabled(True)
        self.tree.setEnabled(True)
        
        # Refresh list after actual deletion
        if not self.whatif_check.isChecked():
            self.scan_credentials()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # Set app icon
    base_path = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_path, "icons", "icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
