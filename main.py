import sys
import os

# Zorg dat tools/ gevonden wordt, zowel in development als in de .exe
if getattr(sys, 'frozen', False):
    base = sys._MEIPASS
else:
    base = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, os.path.join(base, "tools"))

from weekrapport_gui import main

if __name__ == "__main__":
    main()
