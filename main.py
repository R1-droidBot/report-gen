import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm  # For setting image size
import os

# Database Setup
def setup_database():
    with sqlite3.connect("events.db") as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_number TEXT,
                event_name TEXT,
                event_ic TEXT,
                date TEXT,
                event_type TEXT,
                report_doc TEXT,
                geo_photo TEXT,
                attendees INTEGER,
                resource_person TEXT,
                designation TEXT,
                address TEXT,
                funding TEXT,
                days INTEGER,
                audience TEXT,
                mission_mapping TEXT,
                po_pso_mapping TEXT,
                attendance_check TEXT,
                permission_docs TEXT,
                co_po_link TEXT,
                remarks TEXT,
                image_path TEXT
            )
        ''')
        conn.commit()

# Function to upload an image
def upload_image():
    file_path = filedialog.askopenfilename(
        title="Select an Image",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.gif")]
    )
    if file_path:
        image_var.set(file_path)

# Function to add an event to the database
def add_event():
    try:
        with sqlite3.connect("events.db") as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO events (event_number, event_name, event_ic, date, event_type, 
                report_doc, geo_photo, attendees, resource_person, designation, address, 
                funding, days, audience, mission_mapping, po_pso_mapping, attendance_check, 
                permission_docs, co_po_link, remarks, image_path) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', tuple(var.get() for var in entry_vars) + (image_var.get(),))
            conn.commit()
        messagebox.showinfo("Success", "Event added successfully!")
        clear_fields()
    except Exception as e:
        messagebox.showerror("Database Error", f"An error occurred: {e}")

# Function to clear input fields
def clear_fields():
    for var in entry_vars:
        var.set("")
    image_var.set("")

# Function to get the template path
def get_template_path():
    template_path = "event_template.docx"
    if not os.path.exists(template_path):
        messagebox.showerror("Error", "Template file not found! Make sure event_template.docx exists.")
        return None
    return template_path

# Function to generate event report
def generate_report():
    try:
        with sqlite3.connect("events.db") as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM events")
            events = cursor.fetchall()

        if not events:
            messagebox.showerror("Error", "No events found in the database.")
            return

        template_path = get_template_path()
        if not template_path:
            return

        doc = DocxTemplate(template_path)
        event_list = []

        for e in events:
            image_path = e[21]  # Assuming image_path is the last column
            event_image = None
            if image_path and os.path.exists(image_path):
                try:
                    event_image = InlineImage(doc, image_path, width=Cm(6))
                except Exception as img_err:
                    print(f"Error loading image: {img_err}")
                    event_image = "Error loading image"
            
            event_data = {
                "event_number": e[1], "event_name": e[2], "event_ic": e[3], "date": e[4],
                "event_type": e[5], "report_doc": e[6], "geo_photo": e[7], "attendees": e[8],
                "resource_person": e[9], "designation": e[10], "address": e[11], "funding": e[12],
                "days": e[13], "audience": e[14], "mission_mapping": e[15], "po_pso_mapping": e[16],
                "attendance_check": e[17], "permission_docs": e[18], "co_po_link": e[19],
                "remarks": e[20], "event_image": event_image
            }
            event_list.append(event_data)

        context = {"events": event_list}
        print (context)
        doc.render(context)

        report_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")]
        )
        if report_path:
            doc.save(report_path)
            messagebox.showinfo("Success", "Report generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI Setup
root = tk.Tk()
root.title("Event Report Generator")

entry_vars = [tk.StringVar() for _ in range(20)]
image_var = tk.StringVar()

labels = ["Event Number", "Event Name", "Event IC", "Date", "Event Type",
          "Report Doc", "Geo Photo", "Attendees", "Resource Person", "Designation",
          "Address", "Funding", "Days", "Audience", "Mission Mapping", "PO-PSO Mapping",
          "Attendance Check", "Permission Docs", "CO-PO Link", "Remarks"]

for i, label in enumerate(labels):
    ttk.Label(root, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="w")
    ttk.Entry(root, textvariable=entry_vars[i]).grid(row=i, column=1, padx=5, pady=5)

ttk.Label(root, text="Upload Image:").grid(row=len(labels), column=0, padx=5, pady=5, sticky="w")
ttk.Entry(root, textvariable=image_var).grid(row=len(labels), column=1, padx=5, pady=5)
ttk.Button(root, text="Browse", command=upload_image).grid(row=len(labels), column=2, padx=5, pady=5)

ttk.Button(root, text="Add Event", command=add_event).grid(row=len(labels) + 1, column=0, padx=5, pady=10)
ttk.Button(root, text="Generate Report", command=generate_report).grid(row=len(labels) + 1, column=1, padx=5, pady=10)
ttk.Button(root, text="Clear", command=clear_fields).grid(row=len(labels) + 2, column=0, columnspan=2, pady=10)

setup_database()
root.mainloop()
