import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pdfplumber
from datetime import datetime, timedelta
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import re
from tkcalendar import Calendar


def initialize_root(title, width, height):
    root = tk.Tk()
    root.title(title)
    root.iconbitmap('iconopythonprueba.ico')
    center_window(root, width, height)
    return root


def center_window(root, width, height):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int((screen_width - width) / 2)
    center_y = int((screen_height - height) / 2)
    root.geometry(f'{width}x{height}+{center_x}+{center_y}')


def get_pdf_path():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    messagebox.showinfo("Selecciona tu horario", "Por favor selecciona el PDF con tu horario.")
    file_path = filedialog.askopenfilename(title="Selecciona tu horario", filetypes=[("PDF files", "*.pdf")])
    if file_path == "":  # Si el usuario no selecciona un archivo y cierra la ventana
        root.deiconify()  # Muestra la ventana oculta para poder mostrar otra messagebox
        user_choice = messagebox.askquestion("No has escogido tu horario", "¿Te gustaría volver a la selección?")
        if user_choice == 'yes':
            file_path = filedialog.askopenfilename(title="Selecciona tu horario", filetypes=[("PDF files", "*.pdf")])
        root.destroy()
        return file_path
    else:
        root.destroy()
        return file_path


def extract_all_ids(pdf_path):
    ids = set()
    with pdfplumber.open(pdf_path) as pdf:
        headers = None
        for page_number, page in enumerate(pdf.pages):
            print(f"Processing page {page_number + 1}")
            table = page.extract_table()
            if table:
                if headers is None:
                    headers = table[0]
                    if 'ID' not in headers:
                        headers = None
                        continue
                id_index = headers.index('ID')
                for row in table[1:]:
                    if row[id_index]:
                        ids.add(row[id_index].strip())
            else:
                print(f"No table found on page {page_number + 1}")
    return sorted(list(ids))


def select_id_from_list(ids):
    root = initialize_root("¿Quién eres?", 250, 140)
    root.focus_force()  # Force the window to take focus
    selected_id = tk.StringVar(root)
    tk.Label(root, text="Por favor escoge tu nombre en la lista:").pack(pady=10)
    combo_box = ttk.Combobox(root, values=ids, textvariable=selected_id, state="readonly")
    combo_box.pack(pady=5)
    combo_box.set("Elige tu nombre")
    button = ttk.Button(root, text="Confirmar", command=root.destroy)
    button.pack(pady=20)
    root.mainloop()
    return selected_id.get()


def extract_row_by_id(pdf_path, target_id):
    headers = ["ID", "LUNES IN", "LUNES OUT", "MARTES IN", "MARTES OUT", "MIÉRCOLES IN", "MIÉRCOLES OUT",
               "JUEVES IN", "JUEVES OUT", "VIERNES IN", "VIERNES OUT", "SÁBADO IN", "SÁBADO OUT", "DOMINGO IN",
               "DOMINGO OUT"]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table[1:]:
                    if row[0] == target_id:
                        return dict(zip(headers, row))
    return None


def normalize_schedule(data):
    days = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]
    schedule = {}
    time_pattern = re.compile(r'^\d{2}:\d{2}$')

    def format_time(time_str):
        if time_str and ':' in time_str:
            parts = time_str.split(':')
            if len(parts[0]) == 1:
                parts[0] = '0' + parts[0]
            return ':'.join(parts)
        return time_str

    for day in days:
        in_key = f'{day} IN'
        out_key = f'{day} OUT'
        in_time = format_time(data.get(in_key, ''))
        out_time = format_time(data.get(out_key, ''))
        in_time = in_time if in_time and time_pattern.match(in_time) else None
        out_time = out_time if out_time and time_pattern.match(out_time) else None
        schedule[day] = {'IN': in_time, 'OUT': out_time}
    return schedule


def get_week_start_from_calendar():
    root = initialize_root("Selecciona la semana", 400, 350)
    today = datetime.now()  # Get today's date
    cal = Calendar(root, selectmode='day', year=today.year, month=today.month, day=today.day)
    cal.pack(pady=20, fill="both", expand=True)

    def grab_date():
        selected_date = cal.selection_get()
        start_of_week = selected_date - timedelta(days=selected_date.weekday())
        user_date.set(start_of_week.strftime('%Y-%m-%d'))
        root.quit()

    user_date = tk.StringVar(root)
    select_button = ttk.Button(root, text="Selecciona la semana", command=grab_date)
    select_button.pack(pady=20)

    # Make sure the window pops up and stays on top until a date is selected
    root.attributes('-topmost', True)
    root.mainloop()
    root.attributes('-topmost', False)
    root.destroy()
    return user_date.get()


def format_time(time_str):
    if time_str and ':' in time_str:
        parts = time_str.split(':')
        if len(parts[0]) == 1:
            parts[0] = '0' + parts[0]
        return ':'.join(parts)
    return time_str


def format_schedule_for_calendar(normalized_schedule, start_date):
    formatted_schedule = {}
    for day, times in normalized_schedule.items():
        if times['IN'] and times['OUT']:
            day_index = list(normalized_schedule.keys()).index(day)
            in_datetime = datetime.strptime(f'{start_date} {times["IN"]}', '%Y-%m-%d %H:%M')
            out_datetime = datetime.strptime(f'{start_date} {times["OUT"]}', '%Y-%m-%d %H:%M')
            in_datetime += timedelta(days=day_index)
            out_datetime += timedelta(days=day_index)
            formatted_schedule[day] = {'IN': in_datetime, 'OUT': out_datetime}
    return formatted_schedule


def get_resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def get_calendar_service():
    scopes = ['https://www.googleapis.com/auth/calendar.events']
    credentials_path = get_resource_path('credentials.json')
    flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes=scopes)
    credentials = flow.run_local_server(port=0)
    service = build('calendar', 'v3', credentials=credentials)
    return service


def create_calendar_event(service, summary, start_datetime, end_datetime):
    event = {
        'summary': summary,
        'start': {'dateTime': start_datetime.isoformat(), 'timeZone': 'Europe/Madrid'},
        'end': {'dateTime': end_datetime.isoformat(), 'timeZone': 'Europe/Madrid'},
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10}
            ],
        },
    }
    event = service.events().insert(calendarId='primary', body=event).execute()
    print(f"Event created: {event.get('htmlLink')}")
    return summary, event.get('htmlLink')


def display_success_message():
    root = initialize_root("Confirmación", 300, 100)

    label = tk.Label(root, text="Su horario laboral se ha añadido a Google Calendar con éxito", wraplength=250)
    label.pack(pady=10, padx=10)

    ok_button = ttk.Button(root, text="Ok", command=root.destroy)
    ok_button.pack(pady=10)

    root.mainloop()


# def display_events(events):
#     root = initialize_root("Scheduled Events", 900, 300)
#     frame = tk.Frame(root)
#     frame.pack(fill=tk.BOTH, expand=True)
#     scrollbar = tk.Scrollbar(frame)
#     scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
#     listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, width=50, height=15)
#     for event in events:
#         listbox.insert(tk.END, f"{event[0]}: {event[1]}")
#     listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
#     scrollbar.config(command=listbox.yview)
#     exit_button = ttk.Button(root, text="Close", command=root.destroy)
#     exit_button.pack(side=tk.BOTTOM, pady=10)
#     root.mainloop()


def main():
    pdf_path = get_pdf_path()  # Solicita al usuario seleccionar el archivo PDF
    if not pdf_path:
        print("No file selected.")
        return

    # Extrae todos los IDs del archivo PDF
    ids = extract_all_ids(pdf_path)
    if not ids:
        print("No IDs found in the document.")
        return

    # Permite al usuario seleccionar un ID de la lista
    target_id = select_id_from_list(ids)
    if not target_id:
        print("No ID selected.")
        return

    # Extrae la fila correspondiente al ID seleccionado
    matched_row = extract_row_by_id(pdf_path, target_id)
    if matched_row:
        normalized_schedule = normalize_schedule(matched_row)
        start_date = get_week_start_from_calendar()
        if not start_date:
            print("No date selected.")
            return
        formatted_schedule = format_schedule_for_calendar(normalized_schedule, start_date)
        service = get_calendar_service()
        event_links = []
        for day, times in formatted_schedule.items():
            if times['IN'] and times['OUT']:
                summary = f"Work Shift on {day}"
                event_details = create_calendar_event(service, summary, times['IN'], times['OUT'])
                event_links.append(event_details)
        if event_links:
            display_success_message()
        else:
            print("No events were created.")
    else:
        print(f"No data found for ID: {target_id}")


if __name__ == "__main__":
    main()
