import sys
from PySide6.QtWidgets import QApplication
from ui import DuplicateChecker

#pyinstaller --onefile --windowed  --name=Sorter_v8  main.py 

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DuplicateChecker()
    window.show()
    sys.exit(app.exec()) 
