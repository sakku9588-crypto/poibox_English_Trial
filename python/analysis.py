# start.py
import tkinter as tk
from ananeko_gemini_en_t import ArchiveImporterGUI

# 起動するにゃ！
root = tk.Tk()
app = ArchiveImporterGUI(root)
root.mainloop()