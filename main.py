# -*- coding: utf-8 -*-
import sys
import os

# Добавляем путь к проекту
sys.path.insert(0, os.path.dirname(__file__))

from config.config_manager import ConfigManager
from gui.main_window import MainWindow

def main():
    config_manager = ConfigManager()
    app = MainWindow(config_manager)
    app.mainloop()

if __name__ == "__main__":
    main()
