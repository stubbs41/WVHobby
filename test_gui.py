#!/usr/bin/env python3
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget

class TestWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PySide6 Test")
        self.setGeometry(100, 100, 400, 200)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Add a label
        label = QLabel("PySide6 is working correctly!")
        layout.addWidget(label)
        
        # Add a button
        button = QPushButton("Click Me")
        button.clicked.connect(self.on_button_click)
        layout.addWidget(button)
    
    def on_button_click(self):
        print("Button clicked!")

def main():
    app = QApplication(sys.argv)
    window = TestWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 