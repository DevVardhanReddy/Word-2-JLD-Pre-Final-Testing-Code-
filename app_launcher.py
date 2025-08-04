import sys
import os
import threading
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, Qt
from PyQt5.QtGui import QIcon

# Import your Flask application
from app import app as flask_app

class FlaskThread(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.daemon = True

    def run(self):
        # Run Flask app on a local port that won't interfere with other services
        flask_app.run(host='127.0.0.1', port=5001, debug=False)

class WordJLDDesktopApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Set window properties
        self.setWindowTitle('WORDJLD')
        self.setGeometry(100, 100, 1200, 800)

        # Create WebEngineView to display Flask application
        self.web_view = QWebEngineView()
        self.setCentralWidget(self.web_view)

        # Load the local Flask application
        self.web_view.load(QUrl('http://127.0.0.1:5001/'))

        # Optional: Set window icon if you have one
        # self.setWindowIcon(QIcon('path/to/your/icon.png'))

def main():
    # Start Flask in a separate thread
    flask_thread = FlaskThread()
    flask_thread.start()

    # Create Qt Application
    app = QApplication(sys.argv)
    
    # Create and show the desktop application
    desktop_app = WordJLDDesktopApp()
    desktop_app.show()

    # Exit the app when the window is closed
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()