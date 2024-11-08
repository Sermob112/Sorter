import sys
from PySide6.QtWidgets import QApplication
from ui import DuplicateChecker


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DuplicateChecker()
    window.show()
    sys.exit(app.exec()) 
