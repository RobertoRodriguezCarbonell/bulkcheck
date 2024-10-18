import csv
import pandas as pd
import requests
import json
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import time
import math
import re
import io

def is_valid_ip(ip):
    """Verifica si la dirección IP es válida."""
    ip_regex = re.compile(
        r'^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$'  # IPv4
        r'|^(?:[0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}$'  # IPv6
    )
    return re.match(ip_regex, ip) is not None

def bulk_check(excel_path, api_key, export_path, progress, output_box):
    start_time = time.time()
    json_temp_path = os.path.join(os.path.dirname(export_path), 'aipdbulkchecktempfile.json')

    try:
        # Leer el archivo Excel
        df = pd.read_excel(excel_path, engine='openpyxl')
        
        print("Columnas en el Excel:", df.columns.tolist())

        # Verifica que existan las columnas 'email' y 'ip'
        if 'email' not in df.columns or 'ip' not in df.columns:
            raise ValueError("El archivo Excel no tiene las columnas esperadas: 'email' o 'ip'.")

        results = []
        total_rows = len(df)

        for i, row in df.iterrows():
            email = str(row['email']).strip()  # Extrae y limpia el email
            ips = str(row['ip']).strip().split()  # Divide las IPs si están separadas por espacios

            for ip in ips:  # Procesa cada IP por separado
                ip = ip.strip()  # Limpia la IP

                if is_valid_ip(ip):  # Verifica si la IP es válida
                    response = requests.get(
                        f"https://api.abuseipdb.com/api/v2/check?ipAddress={ip}",
                        headers={
                            'Accept': 'application/json',
                            'Key': api_key
                        }
                    )

                    print(f"Response for {ip}: {response.status_code}, {response.text}")

                    if response.status_code == 200:
                        data = response.json()
                        results.append({"email": email, "ip": ip, "data": data})  # Almacena los datos obtenidos
                    else:
                        output_box.insert(tk.END, f"{ip} is not a valid IP or failed to check!\n")
                else:
                    output_box.insert(tk.END, f"{ip} is not a valid IP format.\n")

            progress['value'] = (i + 1) / total_rows * 100
            output_box.delete('1.0', tk.END)
            output_box.insert(tk.END, f"Processing {i + 1} of {total_rows}\n")
            root.update_idletasks()

        if results:
            with open(json_temp_path, 'w') as json_file:
                for result in results:
                    json_file.write(json.dumps(result) + "\n")

            # Guardar los resultados en un CSV
            with open(export_path, 'w', newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                # Añadimos la columna 'email' al archivo de salida
                csv_writer.writerow(["email", "ipAddress", "abuseConfidenceScore", "isp", "domain", "countryCode", "totalReports", "lastReportedAt"])

                for result in results:
                    email = result["email"]
                    ip = result["ip"]
                    data = result["data"]
                    csv_writer.writerow([
                        email, 
                        data["data"]["ipAddress"], 
                        data["data"]["abuseConfidenceScore"], 
                        data["data"]["isp"], 
                        data["data"]["domain"], 
                        data["data"]["countryCode"], 
                        data["data"]["totalReports"], 
                        data["data"]["lastReportedAt"]
                    ])
        else:
            output_box.insert(tk.END, "No valid results to export.\n")

        end_time = time.time()
        elapsed_time = end_time - start_time
        elapsed_minutes, elapsed_seconds = divmod(elapsed_time, 60)
        elapsed_minutes = math.floor(elapsed_minutes)
        elapsed_seconds = round(elapsed_seconds, 1)

        avg_time_per_ip = elapsed_time / total_rows if total_rows > 0 else 0
        avg_time_per_ip = round(avg_time_per_ip, 1)

        output_box.insert(tk.END, f"Started check of {total_rows} IPs at {time.strftime('%b %d %H:%M:%S', time.localtime(start_time))}\n")
        output_box.insert(tk.END, f"Completed check at {time.strftime('%b %d %H:%M:%S', time.localtime(end_time))}\n")
        output_box.insert(tk.END, f"Time elapsed was {elapsed_minutes} minutes and {elapsed_seconds} seconds\n")
        output_box.insert(tk.END, f"Average time per IP checked was {avg_time_per_ip} seconds\n\n")
        output_box.insert(tk.END, "The Admiralty commends you for your efforts!\n")

    except Exception as e:
        output_box.delete('1.0', tk.END)
        output_box.insert(tk.END, f"An error occurred: {str(e)}\n")

    finally:
        if os.path.exists(json_temp_path):
            os.remove(json_temp_path)

def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def browse_save_file(entry):
    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filename:
        if os.path.exists(filename):
            if messagebox.askokcancel("Warning", "The file already exists. Do you want to overwrite it?"):
                entry.delete(0, tk.END)
                entry.insert(0, filename)
        else:
            entry.delete(0, tk.END)
            entry.insert(0, filename)

def main():
    global root
    root = tk.Tk()
    root.title("Admiral SYN-ACKbar's AbuseIPDB Bulk Checker")
    root.geometry("500x600")

    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    title_label1 = ttk.Label(frame, text="Admiral SYN-ACKbar's", font=("Sylfaen", 14, "italic"))
    title_label1.grid(row=0, column=0, columnspan=3)
    title_label2 = ttk.Label(frame, text="AbuseIPDB Bulk Checker", font=("Sylfaen", 18, "bold"))
    title_label2.grid(row=1, column=0, columnspan=3)

    api_label = ttk.Label(frame, text="API Key:")
    api_label.grid(row=2, column=0, sticky=tk.W)
    api_entry = ttk.Entry(frame, width=30, show="*")
    api_entry.grid(row=2, column=1, sticky=(tk.W, tk.E))

    excel_label = ttk.Label(frame, text="Excel Input File Path / Name:")
    excel_label.grid(row=3, column=0, sticky=tk.W)
    excel_entry = ttk.Entry(frame, width=30)
    excel_entry.grid(row=3, column=1, sticky=(tk.W, tk.E))
    excel_button = ttk.Button(frame, text="Browse", command=lambda: browse_file(excel_entry))
    excel_button.grid(row=3, column=2)

    export_label = ttk.Label(frame, text="CSV Output File Path / Name:")
    export_label.grid(row=4, column=0, sticky=tk.W)
    export_entry = ttk.Entry(frame, width=30)
    export_entry.grid(row=4, column=1, sticky=(tk.W, tk.E))
    export_button = ttk.Button(frame, text="Save As", command=lambda: browse_save_file(export_entry))
    export_button.grid(row=4, column=2)

    submit_button = ttk.Button(frame, text="ENGAGE", command=lambda: bulk_check(excel_entry.get(), api_entry.get(), export_entry.get(), progress, output_box), style='Engage.TButton')
    submit_button.grid(row=5, column=0, columnspan=3)

    output_label = ttk.Label(frame, text="Output:")
    output_label.grid(row=6, column=0, sticky=tk.W)

    output_box = tk.Text(frame, height=15, width=50)
    output_box.grid(row=7, column=0, columnspan=3)

    progress = ttk.Progressbar(frame, orient="horizontal", length=400, mode="determinate")
    progress.grid(row=8, column=0, columnspan=3, pady=10)

    for i in range(3):
        frame.columnconfigure(i, weight=1)

    root.mainloop()

if __name__ == "__main__":
    main()
