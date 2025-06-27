import socket
import pandas as pd
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog
import re

# Global variables
running = False
file_path = "camera_data.xlsx"  # Default save location

def create_socket_connection(ip, port):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((ip, port))
        return sock
    except socket.error as e:
        print(f"Socket connection error: {e}")
        return None

def send_command(sock, command):
    try:
        sock.sendall(command.encode('ascii'))
        response = sock.recv(1024)
        print(f"Sent: {command.strip()} | Received: {response.decode('ascii').strip()}")
        return response.decode('ascii')
    except socket.error as e:
        print(f"Socket error: {e}")
        return None

def listen_for_data(sock):
    global running
    buffer = ""
    sock.settimeout(1.0)  # so we can print “waiting” periodically
    try:
        while running:
            try:
                chunk = sock.recv(1024)
            except socket.timeout:
                # no data for 1s
                continue

            if not chunk:
                print("Connection closed by remote.")
                break

            # accumulate and split on any newline
            buffer += chunk.decode('ascii', errors='ignore')
            parts = re.split(r'[\r\n]+', buffer)
            # all but last are complete
            for line in parts[:-1]:
                if not line:
                    continue
                print(f"→ RAW LINE: {repr(line)}")
                parsed = parse_response(line)
                save_to_excel(parsed, file_path)
                # update UI with Tool1
                ts = parsed['Timestamp'][0]
                t1 = parsed['Tool1'][0]
                window.after(0, update_status, f"{ts} | Tool1: {t1 or '<empty>'}")
            buffer = parts[-1]

    except socket.error as e:
        print(f"Socket error while receiving data: {e}")
    finally:
        sock.close()
        running = False
        window.after(0, lambda: button.config(text="Run"))
        window.after(0, running_label.pack_forget)

def parse_response(response):
    # strip any stray whitespace or control chars
    line = response.strip()
    print(f"Parsing: {repr(line)}")
    parts = [p.strip() for p in line.split(',')]
    print(f"  => parts[{len(parts)}]: {parts}")

    data = {'Timestamp': [pd.Timestamp.now()]}
    # grab up to 11 read-text fields at idx 9,13,17,...
    for i in range(11):
        idx = 9 + i*4
        key = f"Tool{i+1}"
        data[key] = [parts[idx] if idx < len(parts) else '']
    return data

def save_to_excel(data_dict, file_name):
    df = pd.DataFrame(data_dict)
    if os.path.exists(file_name):
        with pd.ExcelWriter(file_name, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            start_row = writer.sheets['Sheet1'].max_row
            df.to_excel(writer, index=False, header=False,
                        startrow=start_row, sheet_name='Sheet1')
    else:
        df.to_excel(file_name, index=False, sheet_name='Sheet1')
    print(f"Saved row to {file_name}")

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
        running = True
        button.config(text="Stop")
        running_label.pack()
        threading.Thread(target=start_socket_communication, daemon=True).start()
    else:
        running = False
        button.config(text="Run")
        running_label.pack_forget()

def start_socket_communication():
    ip = ip_entry.get()
    port = int(port_entry.get())
    sock = create_socket_connection(ip, port)
    if sock:
        send_command(sock, 'OF,01\r')
        send_command(sock, 'OE,1\r')
        listen_for_data(sock)
    else:
        print("Failed to connect")
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
