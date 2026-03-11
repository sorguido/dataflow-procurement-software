"""
Simple test to verify main window persistence and taskbar icon visibility
"""

import sys
import os

# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui_dialogs import get_main_window, select_config_file
import tkinter as tk

def test_window_persistence():
    """Test that main window persists and taskbar icon remains visible"""
    
    print("Creating main window...")
    main_window = get_main_window()
    root = main_window.get_root()
    
    print("Main window created and hidden")
    print("Taskbar icon should now be visible")
    print()
    
    # Create a test label to show status
    status_label = tk.Label(root, text="Main window is active (hidden)")
    status_label.pack()
    
    print("Opening file dialog...")
    print("Note: If you click outside the dialog, you can click the taskbar icon to return")
    
    # Show file selection dialog
    selected_file = select_config_file()
    
    if selected_file:
        print(f"Selected: {selected_file}")
    else:
        print("No file selected or dialog cancelled")
    
    print()
    print("Test complete. The main window will close in 3 seconds...")
    root.after(3000, lambda: main_window.destroy())
    root.mainloop()

if __name__ == "__main__":
    test_window_persistence()
