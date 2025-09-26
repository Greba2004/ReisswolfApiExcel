import requests
import openpyxl
from pathlib import Path
from datetime import datetime, timedelta
import os
import json
from tkinter import Tk, Label, Button, messagebox, filedialog
from tkcalendar import DateEntry
from requests.auth import HTTPBasicAuth

# ---- API pristup ----
subdomain = "rwrs"
username = "ALEKSA"
password = "AleksaRW1234"

# ---- Podrazumevani folder ----
base_folder = Path(os.path.expanduser("~")) / "Desktop" / "RADNI_NALOZI"
base_folder.mkdir(exist_ok=True)

# ---- Konfiguracioni fajl ----
config_file = Path(os.path.expanduser("~")) / ".rwam_config.json"

# učitavanje poslednjeg foldera
if config_file.exists():
    try:
        with open(config_file, "r") as f:
            cfg = json.load(f)
            last_folder = Path(cfg.get("last_folder", base_folder))
            if last_folder.exists():
                selected_folder = last_folder
            else:
                selected_folder = base_folder
    except Exception:
        selected_folder = base_folder
else:
    selected_folder = base_folder

def choose_folder():
    global selected_folder
    folder = filedialog.askdirectory(initialdir=selected_folder)
    if folder:
        selected_folder = Path(folder)
        folder_label.config(text=str(selected_folder))
        # sačuvaj u konfiguracioni fajl
        try:
            with open(config_file, "w") as f:
                json.dump({"last_folder": str(selected_folder)}, f)
        except Exception as e:
            messagebox.showwarning("Upozorenje", f"Ne mogu da sačuvam folder:\n{e}")

def format_dt(dt):
    if isinstance(dt, str):
        return dt[:16].replace("T", " ")
    return dt

def fetch_orders_for_date(date_obj, ws, existing_ids):
    """Preuzima naloge za jedan dan i dodaje u dati worksheet"""
    api_url = f"https://{subdomain}.rwam.reisswolf.com/api/order"
    params = {
        "max": 500,
        "dateOrderedFrom": date_obj.isoformat(),
        "dateOrderedTo": date_obj.isoformat()
    }

    try:
        response = requests.get(api_url, auth=HTTPBasicAuth(username, password), params=params)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        messagebox.showerror("Greška", f"Problem pri preuzimanju naloga za {date_obj}:\n{e}")
        return 0

    new_orders_count = 0

    for order in data.get("orders", []):
        order_id = order.get("id")
        if order_id in existing_ids:
            continue

        account = order.get("account", {})
        account_str = f"{account.get('name','')} ({account.get('acronym','')})"

        barcode = order.get("barcode", {}).get("value", "")
        date_created = format_dt(order.get("dateCreated"))
        orderer = order.get("orderer", {}).get("fullName", "")
        status = order.get("status", "")
        service_level_name = order.get("serviceLevelName", "")
        additional_info = order.get("additionalInfo") if status == "CANCELLED" else ""
        pickup = order.get("pickupAddress", {})
        pickup_info = f"{pickup.get('extension','')}, {pickup.get('number','')}, {pickup.get('city','')}"

        row = [
            account_str,
            barcode,
            date_created,
            orderer,
            status,
            service_level_name,
            additional_info,
            pickup_info
        ]

        ws.append(row)
        existing_ids.add(order_id)
        new_orders_count += 1

    return new_orders_count

def get_unique_excel_path(base_path):
    """Ako postoji fajl sa istim imenom, dodaje _1, _2, ..."""
    counter = 1
    path = base_path
    while path.exists():
        path = base_path.stem + f"_{counter}" + base_path.suffix
        path = base_path.parent / path
        counter += 1
    return path

def run_export():
    d_from = start_cal.get_date()
    d_to = end_cal.get_date()

    if d_from > d_to:
        messagebox.showwarning("Greška", "Početni datum ne može biti posle krajnjeg")
        return

    base_excel_path = selected_folder / f"Orders_{d_from}_{d_to}.xlsx"
    excel_path = get_unique_excel_path(base_excel_path)

    headers = ["account", "barcode", "dateCreated", "orderer", 
               "status", "serviceLevelName", "additionalInfo", "pickupAddress"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(headers)

    existing_ids = set()
    total_orders = 0
    current = d_from
    while current <= d_to:
        if current.weekday() < 5:  # preskoči subotu(5) i nedelju(6)
            new_count = fetch_orders_for_date(current, ws, existing_ids)
            total_orders += new_count
        current += timedelta(days=1)

    wb.save(excel_path)
    messagebox.showinfo("Gotovo", f"Preuzeto {total_orders} naloga\nFajl: {excel_path}")

# ---- Tkinter GUI ----
root = Tk()
root.title("Preuzimanje radnih naloga RWAM")

Label(root, text="Početni datum:").grid(row=0, column=0, padx=10, pady=5)
start_cal = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
start_cal.grid(row=0, column=1, padx=10, pady=5)

Label(root, text="Krajnji datum:").grid(row=1, column=0, padx=10, pady=5)
end_cal = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
end_cal.grid(row=1, column=1, padx=10, pady=5)

Label(root, text="Folder:").grid(row=2, column=0, padx=10, pady=5)
folder_label = Label(root, text=str(selected_folder))
folder_label.grid(row=2, column=1, padx=10, pady=5)
Button(root, text="Browse", command=choose_folder).grid(row=2, column=2, padx=5, pady=5)

Button(root, text="Preuzmi naloge", command=run_export).grid(row=3, column=0, columnspan=3, pady=15)

root.mainloop()
