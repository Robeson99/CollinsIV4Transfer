import socket
import pandas as pd
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog
import re
import xlwings as xw

# Global variables
running = False
file_path = "camera_data.xlsx"  # Default save location

def create_socket_connection(ip, port):
    try:
        print(f"[*] create_socket_connection → {ip}:{port}")
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((ip, port))
        print("[*] Socket connected")
        return sock
    except socket.error as e:
        print(f"[!] Socket connection error: {e}")
        return None

def send_command(sock, command):
    try:
        print(f"[*] send_command → {repr(command)}")
        sock.sendall(command.encode('ascii'))
        response = sock.recv(1024)
        resp_text = response.decode('ascii', errors='ignore').strip()
        print(f"    ← Response: {repr(resp_text)}")
        return resp_text
    except socket.error as e:
        print(f"[!] Socket error during send: {e}")
        return None

def listen_for_data(sock):
    global running
    print("[*] listen_for_data entered")
    buffer = ""
    sock.settimeout(1.0)

    try:
        while running:
            try:
                chunk = sock.recv(1024)
            except socket.timeout:
                print("[DEBUG] recv timed out, still waiting…")
                continue

            if not chunk:
                print("[!] Connection closed by remote.")
                break

            print(f"[RAW CHUNK] {repr(chunk)}")
            text = chunk.decode('ascii', errors='ignore')
            buffer += text

            # split on any newline
            parts = re.split(r'[\r\n]+', buffer)
            lines = parts[:-1]
            buffer = parts[-1]  # leftover

            # fallback if no newline but looks like full record
            if not lines and buffer.count(',') >= 9:
                print(f"[FALLBACK] treating buffer as one record: {repr(buffer)}")
                lines = [buffer]
                buffer = ""

            for line in lines:
                line = line.strip()
                if not line:
                    continue
                print(f"[LINE ▶] {repr(line)}")
                parsed = parse_response(line)
                print(f"[PARSED ▶] {parsed}")
                save_to_excel(parsed, file_path)

                # update UI
                ts = parsed['Timestamp'][0]
                t1 = parsed['Tool1'][0]
                window.after(0, update_status, f"{ts} | Tool1: {t1 or '<empty>'}")

    except socket.error as e:
        print(f"[!] Socket error while receiving data: {e}")

    finally:
        print("[*] listen_for_data exiting")
        sock.close()
        running = False
        window.after(0, lambda: button.config(text="Run"))
        window.after(0, running_label.pack_forget)

def parse_response(response):
    line = response.strip()
    print(f"[parse_response] raw: {repr(line)}")
    parts = [p.strip() for p in line.split(',')]
    print(f"    parts[{len(parts)}]: {parts}")

    data = {'Timestamp': [pd.Timestamp.now()]}
    # extract up to 11 read-text fields at idx 9,13,17,...
    for i in range(11):
        idx = 9 + i*4
        key = f"Tool{i+1}"
        data[key] = [parts[idx] if idx < len(parts) else '']
    return data

def ensure_workbook(path):
    if not os.path.exists(path):
        print(f"[*] ensure_workbook creating: {path}")
        wb = xw.Book()
        sht = wb.sheets[0]
        sht.name = 'Sheet1'
        sht.range('A1').value = ['Timestamp'] + [f'Tool{i+1}' for i in range(11)]
        wb.save(path)
        wb.close()
        print(f"[+] Created new workbook: {path}")

def save_to_excel(data_dict, file_name):
    print(f"[*] save_to_excel → {file_name}")
    os.makedirs(os.path.dirname(file_name) or '.', exist_ok=True)

    # attach to COM
    app = xw.apps.active if xw.apps else xw.App(visible=False)
    # open or create
    if file_name in [b.name for b in app.books]:
        wb = app.books[file_name]
    elif os.path.exists(file_name):
        wb = app.books.open(file_name)
    else:
        wb = app.books.add()
        sht = wb.sheets[0]
        sht.name = 'Sheet1'
        sht.range('A1').value = list(data_dict.keys())
        wb.save(file_name)

    sht = wb.sheets['Sheet1']
    used = sht.api.UsedRange.Rows.Count
    first = sht.range('A1').value
    if used == 1 and not first:
        next_row = 1
        sht.range('A1').value = list(data_dict.keys())
    else:
        next_row = used + 1

    headers = list(data_dict.keys())
    row = [data_dict[h][0] for h in headers]
    print(f"[DEBUG] Writing row {next_row}: {row}")

    sht.range(f'A{next_row}').value = row
    wb.save()
    print(f"[+] Appended to {file_name}")

def select_file():
    global file_path
    chosen = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                          filetypes=[("Excel Files", "*.xlsx")])
    if chosen:
        file_path = chosen
        file_label.config(text=f"Save Path: {file_path}")

def run():
    global running
    if not running:
        ensure_workbook(file_path)
        running = True
        button.config(text="Stop")
        running_label.pack()
        threading.Thread(target=start_socket_communication, daemon=True).start()
    else:
        running = False
        button.config(text="Run")
        running_label.pack_forget()

def start_socket_communication():
    ip = ip_entry.get().strip()
    port = int(port_entry.get().strip())
    print(f"[*] start_socket_communication → {ip}:{port}")
    sock = create_socket_connection(ip, port)
    if sock:
        send_command(sock, 'OF,01\r')
        send_command(sock, 'OE,1\r')
        listen_for_data(sock)
    else:
        print("[!] Failed to connect")
        window.after(0, lambda: button.config(text="Run"))
        window.after(0, running_label.pack_forget)

def update_status(text):
    status_label.config(text=text)

# --- GUI Setup (unchanged) ---
window = tk.Tk()
window.title('IV4 Text to Excel')
window.geometry('500x350')

title_label = ttk.Label(window, text='IV4 Excel Transfer', font='calibri 18 bold')
title_label.pack(pady=10)

ip_label = ttk.Label(window, text='IP Address:')
ip_label.pack()
ip_entry = ttk.Entry(window)
ip_entry.pack(pady=5)

port_label = ttk.Label(window, text='Port:')
port_label.pack()
port_entry = ttk.Entry(window)
port_entry.pack(pady=5)

browse_button = ttk.Button(window, text="Browse Save Location", command=select_file)
browse_button.pack(pady=5)

file_label = ttk.Label(window, text=f"Save Path: {file_path}")
file_label.pack()

button = ttk.Button(window, text='Run', command=run)
button.pack(pady=10)

running_label = ttk.Label(window, text='Running', foreground='green')

status_label = ttk.Label(window, text="Status: Waiting for data...", foreground="blue")
status_label.pack(pady=5)

window.mainloop()
