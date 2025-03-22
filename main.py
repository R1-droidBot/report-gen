import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docxtpl import DocxTemplate
import json
import os

EVENTS_FILE = "events.json"

# Function to load events from JSON file
def load_events():
    if os.path.exists(EVENTS_FILE):
        with open(EVENTS_FILE, "r") as file:
            return json.load(file)
    return []

# Function to save events to JSON file
def save_events(events):
    with open(EVENTS_FILE, "w") as file:
        json.dump(events, file, indent=4)

# Function to add an event
def add_event():
    event_data = {label: var.get() for label, var in zip(labels, entry_vars)}

    events = load_events()
    events.append(event_data)
    save_events(events)

    messagebox.showinfo("Success", "Event added successfully!")
    clear_fields()
    update_event_list()

# Function to clear input fields
def clear_fields():
    for var in entry_vars:
        var.set("")

# Function to update event list in the UI
def update_event_list():
    for row in tree.get_children():
        tree.delete(row)

    events = load_events()
    for event in events:
        tree.insert("", "end", values=(event["Event Number"], event["Event Name"], event["Date"]))

# Function to generate report for the selected event
def generate_report():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Error", "Please select an event to generate a report.")
        return

    selected_event_number = tree.item(selected_item, "values")[0]

    events = load_events()
    selected_event = next((e for e in events if e["Event Number"] == selected_event_number), None)
    
    if not selected_event:
        messagebox.showerror("Error", "Event not found.")
        return

    template_path = "event_template.docx"
    if not os.path.exists(template_path):
        messagebox.showerror("Error", "Template file not found!")
        return

    doc = DocxTemplate(template_path)
    doc.render({"event": selected_event})

    report_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if report_path:
        doc.save(report_path)
        messagebox.showinfo("Success", "Report generated successfully!")

# GUI Setup
root = tk.Tk()
root.title("Event Report Generator")

labels = ["Event Number", "Event Name", "Event IC", "Date", "Event Type",
          "Report Doc", "Geo Photo", "Attendees", "Resource Person", "Designation",
          "Address", "Funding", "Days", "Audience", "Mission Mapping", "PO-PSO Mapping",
          "Attendance Check", "Permission Docs", "CO-PO Link", "Remarks"]

entry_vars = [tk.StringVar() for _ in labels]

for i, label in enumerate(labels):
    ttk.Label(root, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="w")
    ttk.Entry(root, textvariable=entry_vars[i]).grid(row=i, column=1, padx=5, pady=5)

# Buttons
ttk.Button(root, text="Add Event", command=add_event).grid(row=len(labels), column=0, padx=5, pady=10)
ttk.Button(root, text="Generate Report", command=generate_report).grid(row=len(labels), column=1, padx=5, pady=10)

# Event List
tree = ttk.Treeview(root, columns=("Event Number", "Event Name", "Date"), show="headings")
tree.heading("Event Number", text="Event Number")
tree.heading("Event Name", text="Event Name")
tree.heading("Date", text="Date")
tree.grid(row=len(labels) + 1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

update_event_list()
root.mainloop()
