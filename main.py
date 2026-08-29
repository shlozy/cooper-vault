"""
Cooper Vault — application entry point.

Run this file to start the app:

    python main.py

This just imports the Tkinter application class (app.py, which itself
pulls in theme.py, database.py, and authentication.py) and starts the
Tk event loop. It intentionally contains no page/UI/database logic of
its own — that all lives in the modules it imports.
"""
from app import CooperAppMaster

if __name__ == "__main__":
    app = CooperAppMaster()
    app.mainloop()
