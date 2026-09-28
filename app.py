"""Standalone local BRD and solution slide inventory prompt builder."""

import tkinter as tk

from inventory_ui import InventoryTab


def main():
    root = tk.Tk()
    root.title("BRD and Solution Inventory")
    root.geometry("1050x780")
    root.minsize(850, 620)
    InventoryTab(root)
    root.mainloop()


if __name__ == "__main__":
    main()
