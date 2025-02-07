import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
from docxtpl import DocxTemplate
import os

# Database Setup
def setup_database():
    conn = sqlite3.connect("events.db")
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
            remarks TEXT
        )
    ''')
    conn.commit()
    conn.close()

# Function to add event to database
def add_event():
    conn = sqlite3.connect("events.db")
    cursor = conn.cursor()
    cursor.execute('''INSERT INTO events (event_number, event_name, event_ic, date, event_type, 
                   report_doc, geo_photo, attendees, resource_person, designation, address, 
                   funding, days, audience, mission_mapping, po_pso_mapping, attendance_check, 
                   permission_docs, co_po_link, remarks) 
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                   (event_number_var.get(), event_name_var.get(), event_ic_var.get(), date_var.get(),
                    event_type_var.get(), report_doc_var.get(), geo_photo_var.get(), attendees_var.get(),
                    resource_person_var.get(), designation_var.get(), address_var.get(),
                    funding_var.get(), days_var.get(), audience_var.get(), mission_mapping_var.get(),
                    po_pso_mapping_var.get(), attendance_check_var.get(), permission_docs_var.get(),
                    co_po_link_var.get(), remarks_var.get()))
    conn.commit()
    conn.close()
    messagebox.showinfo("Success", "Event added successfully!")
    clear_fields()

# Function to clear input fields
def clear_fields():
    for var in [event_number_var, event_name_var, event_ic_var, date_var, event_type_var, 
                report_doc_var, geo_photo_var, attendees_var, resource_person_var, 
                designation_var, address_var, funding_var, days_var, audience_var, 
                mission_mapping_var, po_pso_mapping_var, attendance_check_var, 
                permission_docs_var, co_po_link_var, remarks_var]:
        var.set("")

# Function to generate event report
def generate_report():
    conn = sqlite3.connect("events.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM events")
    events = cursor.fetchall()
    conn.close()
    print(list(events[0]))
    doc = DocxTemplate("event_template.docx")
    context = {"events": list(events[0])}
    doc.render(context)
    report_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if report_path:
        doc.save(report_path)
        messagebox.showinfo("Success", "Report generated successfully!")

# UI Setup
root = tk.Tk()
root.title("Event Management System")
root.geometry("900x600")

notebook = ttk.Notebook(root)
frame_add = ttk.Frame(notebook)
frame_view = ttk.Frame(notebook)
frame_export = ttk.Frame(notebook)
notebook.add(frame_add, text="Add Event")
notebook.add(frame_view, text="View Events")
notebook.add(frame_export, text="Export Report")
notebook.pack(expand=True, fill="both")

# Form Fields
fields = ["Event Number", "Name of Event", "Event I/C", "Date (DD/MM/YYYY)", "Type of Event",
          "Report Doc Link", "GeoTag Photo Link", "No. of Attendees", "Resource Person",
          "Designation", "Address", "Funding Received", "No. of Days", "Organised For",
          "Mission Mapping", "PO/PSO Mapping", "Attendance Check", "Permission Docs",
          "CO-PO Link", "Remarks"]

entry_vars = []
for i, field in enumerate(fields):
    ttk.Label(frame_add, text=field).grid(row=i, column=0, padx=5, pady=5, sticky="w")
    var = tk.StringVar()
    entry_vars.append(var)
    ttk.Entry(frame_add, textvariable=var, width=40).grid(row=i, column=1, padx=5, pady=5)

event_number_var, event_name_var, event_ic_var, date_var, event_type_var, report_doc_var, geo_photo_var, \
attendees_var, resource_person_var, designation_var, address_var, funding_var, days_var, audience_var, \
mission_mapping_var, po_pso_mapping_var, attendance_check_var, permission_docs_var, co_po_link_var, \
remarks_var = entry_vars

# Submit Button
submit_btn = ttk.Button(frame_add, text="Add Event", command=add_event)
submit_btn.grid(row=len(fields), columnspan=2, pady=10)

# Report Export Button
export_btn = ttk.Button(frame_export, text="Generate Report", command=generate_report)
export_btn.pack(pady=20)

setup_database()
root.mainloop()
