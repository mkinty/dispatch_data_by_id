"""Point d'entrée du Dispatcher SNA : python main.py (ou uv run main.py)."""
import warnings

from dispatch_sna.gui.app import App


def main():
    warnings.filterwarnings("ignore")
    App().mainloop()


if __name__ == "__main__":
    main()
