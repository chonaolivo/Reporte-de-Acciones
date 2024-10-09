from datetime import datetime
import time
import pygetwindow as gw
from openpyxl import Workbook

def get_window_titles():
    titles = []
    for window in gw.getAllTitles():
        if isinstance(window, str):  # Comprobamos si el objeto es una cadena
            titles.append(window)
    return titles

def log_window_titles():
    wb = Workbook()
    ws = wb.active
    ws.append(["Fecha y Hora", "Ventana", "Estado"])
    previous_titles = set()
    while True:
        current_titles = set(get_window_titles())
        opened_titles = current_titles - previous_titles
        closed_titles = previous_titles - current_titles
        if opened_titles or closed_titles:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for title in opened_titles:
                ws.append([current_time, title, "Abierta"])
            for title in closed_titles:
                ws.append([current_time, title, "Cerrada"])
            wb.save('window_log.xlsx')
            previous_titles = current_titles
        time.sleep(5)  # Check every 5 seconds for changes

if __name__ == "__main__":
    log_window_titles()
