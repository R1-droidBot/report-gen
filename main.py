import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from PIL import Image, ImageTk
import json
import os
import shutil
import io

EVENTS_FILE = "events.json"
IMAGES_FOLDER = "event_images"

# Ensure images folder exists
os.makedirs(IMAGES_FOLDER, exist_ok=True)

def load_events():
    if os.path.exists(EVENTS_FILE):
        with open(EVENTS_FILE, "r") as file:
            return json.load(file)
    return []

def save_events(events):
    with open(EVENTS_FILE, "w") as file:
        json.dump(events, file, indent=4)

def add_event():
    event_data = {label: var.get() for label, var in zip(labels, entry_vars)}
    
    # Handle images
    event_data["images"] = []
    for img_path in image_paths:
        if img_path:
            dest_path = os.path.join(IMAGES_FOLDER, os.path.basename(img_path))
            shutil.copy(img_path, dest_path)
            event_data["images"].append(dest_path)
    
    if not event_data["Event Number"] or not event_data["Event Name"]:
        messagebox.showerror("Error", "Event Number and Event Name are required!")
        return

    events = load_events()
    if any(e["Event Number"] == event_data["Event Number"] for e in events):
        messagebox.showerror("Error", "Event with this number already exists!")
        return
        
    events.append(event_data)
    save_events(events)
    messagebox.showinfo("Success", "Event added successfully!")
    clear_fields()
    update_event_list()

def clear_fields():
    for var in entry_vars:
        var.set("")
    image_paths.clear()
    update_image_preview()

def update_event_list(search_term=None):
    for row in tree.get_children():
        tree.delete(row)

    events = load_events()
    for event in events:
        if search_term:
            search_term = search_term.lower()
            if (search_term in event["Event Number"].lower() or 
                search_term in event["Event Name"].lower() or 
                search_term in event["Date"].lower() or
                search_term in event["Resource Person"].lower() or
                search_term in event["Event Type"].lower()):
                tree.insert("", "end", values=(
                    event["Event Number"],
                    event["Event Name"],
                    event["Date"],
                    event["Resource Person"],
                    event["Event Type"],
                    f"{len(event.get('images', []))} images"
                ))
        else:
            tree.insert("", "end", values=(
                event["Event Number"],
                event["Event Name"],
                event["Date"],
                event["Resource Person"],
                event["Event Type"],
                f"{len(event.get('images', []))} images"
            ))

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

    try:
        doc = DocxTemplate(template_path)
        
        # Create context with all fields
        context = {
            'event': {
                'number': selected_event.get("Event Number", ""),
                'name': selected_event.get("Event Name", ""),
                'ic': selected_event.get("Event IC", ""),
                'date': selected_event.get("Date", ""),
                'type': selected_event.get("Event Type", ""),
                'report_doc': selected_event.get("Report Doc", ""),
                'geo_photo': selected_event.get("Geo Photo", ""),
                'attendees': selected_event.get("Attendees", ""),
                'resource_person': selected_event.get("Resource Person", ""),
                'designation': selected_event.get("Designation", ""),
                'address': selected_event.get("Address", ""),
                'funding': selected_event.get("Funding", ""),
                'days': selected_event.get("Days", ""),
                'audience': selected_event.get("Audience", ""),
                'mission_mapping': selected_event.get("Mission Mapping", ""),
                'po_pso_mapping': selected_event.get("PO-PSO Mapping", ""),
                'attendance_check': selected_event.get("Attendance Check", ""),
                'permission_docs': selected_event.get("Permission Docs", ""),
                'co_po_link': selected_event.get("CO-PO Link", ""),
                'remarks': selected_event.get("Remarks", "")
            }
        }
        
        # Handle images
        if "images" in selected_event and selected_event["images"]:
            context['images'] = []
            for idx, img_path in enumerate(selected_event["images"]):
                if os.path.exists(img_path):
                    try:
                        context['images'].append(
                            InlineImage(doc, img_path, width=Mm(50))
                        )
                        context[f'image_{idx}'] = context['images'][-1]
                    except Exception as img_error:
                        print(f"Skipping image {img_path}: {img_error}")
                        continue
        
        doc.render(context)
        
        report_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")]
        )
        if report_path:
            doc.save(report_path)
            messagebox.showinfo("Success", "Report generated successfully!")
    
    except Exception as e:
        messagebox.showerror("Error", f"Report generation failed: {str(e)}")

def delete_event():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Error", "Please select an event to delete.")
        return

    selected_event_number = tree.item(selected_item, "values")[0]
    
    if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete event {selected_event_number}?"):
        events = load_events()
        events = [e for e in events if e["Event Number"] != selected_event_number]
        save_events(events)
        update_event_list()
        messagebox.showinfo("Success", "Event deleted successfully!")

def search_events():
    search_term = search_var.get()
    update_event_list(search_term)

def add_images():
    files = filedialog.askopenfilenames(
        title="Select Images",
        filetypes=[("Image Files", "*.jpg *.jpeg *.png *.gif")]
    )
    if files:
        image_paths.extend(files)
        update_image_preview()

def update_image_preview():
    # Clear previous previews
    for widget in image_preview_frame.winfo_children():
        widget.destroy()
    
    # Display new previews in a grid (4 columns)
    for i, img_path in enumerate(image_paths):
        try:
            img = Image.open(img_path)
            img.thumbnail((100, 100))
            photo = ImageTk.PhotoImage(img)
            label = tk.Label(image_preview_frame, image=photo)
            label.image = photo  # Keep a reference
            label.grid(row=i//4, column=i%4, padx=5, pady=5)
        except Exception as e:
            print(f"Error loading image preview: {e}")

# GUI Setup
root = tk.Tk()
root.title("Event Report Generator")
root.geometry("1200x800")

# Image handling variables
image_paths = []

# Create Notebook (tabbed interface)
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# Create tabs
input_tab = ttk.Frame(notebook)
list_tab = ttk.Frame(notebook)
upload_tab = ttk.Frame(notebook)  # New tab for image uploads

notebook.add(input_tab, text="Add Event")
notebook.add(list_tab, text="Event List")
notebook.add(upload_tab, text="Upload Images")

# Input Tab Content - Organized in columns
labels = ["Event Number", "Event Name", "Event IC", "Date", "Event Type",
          "Report Doc", "Geo Photo", "Attendees", "Resource Person", "Designation",
          "Address", "Funding", "Days", "Audience", "Mission Mapping", "PO-PSO Mapping",
          "Attendance Check", "Permission Docs", "CO-PO Link", "Remarks"]

entry_vars = [tk.StringVar() for _ in labels]

# Create 3 columns in the input tab
col1_frame = ttk.Frame(input_tab)
col2_frame = ttk.Frame(input_tab)
col3_frame = ttk.Frame(input_tab)

col1_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
col2_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
col3_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)

# Distribute fields across 3 columns
for i, (label, var) in enumerate(zip(labels, entry_vars)):
    if i < 7:  # First 7 fields in column 1
        frame = col1_frame
    elif i < 14:  # Next 7 fields in column 2
        frame = col2_frame
    else:  # Remaining fields in column 3
        frame = col3_frame
    
    ttk.Label(frame, text=label).pack(anchor='w', padx=5, pady=2)
    ttk.Entry(frame, textvariable=var).pack(fill='x', padx=5, pady=2)

# Buttons frame at bottom of input tab
button_frame = ttk.Frame(input_tab)
button_frame.pack(fill='x', pady=10)

ttk.Button(button_frame, text="Add Event", command=add_event).pack(side='left', padx=20)
ttk.Button(button_frame, text="Clear Fields", command=clear_fields).pack(side='left', padx=20)
ttk.Button(button_frame, text="Generate Report", command=generate_report).pack(side='left', padx=20)

# Image Upload Tab Content
upload_frame = ttk.Frame(upload_tab)
upload_frame.pack(fill='both', expand=True, padx=10, pady=10)

# Add Images Button
ttk.Button(upload_frame, text="Select Images", command=add_images).pack(pady=10)

# Image Preview Frame
image_preview_frame = ttk.Frame(upload_frame)
image_preview_frame.pack(fill='both', expand=True)

# Event List Tab Content
# Search Frame at top
search_frame = ttk.Frame(list_tab)
search_frame.pack(fill='x', padx=5, pady=5)

search_var = tk.StringVar()
ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
search_entry = ttk.Entry(search_frame, textvariable=search_var)
search_entry.pack(side='left', fill='x', expand=True, padx=5)
ttk.Button(search_frame, text="Search", command=search_events).pack(side='left', padx=5)
ttk.Button(search_frame, text="Clear", command=lambda: [search_var.set(""), update_event_list()]).pack(side='left', padx=5)

# Treeview Frame
tree_frame = ttk.Frame(list_tab)
tree_frame.pack(fill='both', expand=True, padx=5, pady=5)

# Define columns with widths
columns = ("Event Number", "Event Name", "Date", "Resource Person", "Event Type", "Images")
tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
tree.heading("Event Number", text="Event Number")
tree.heading("Event Name", text="Event Name")
tree.heading("Date", text="Date")
tree.heading("Resource Person", text="Resource Person")
tree.heading("Event Type", text="Event Type")
tree.heading("Images", text="Images")

# Set column widths
tree.column("Event Number", width=120, anchor='center')
tree.column("Event Name", width=200, anchor='w')
tree.column("Date", width=100, anchor='center')
tree.column("Resource Person", width=200, anchor='w')
tree.column("Event Type", width=150, anchor='w')
tree.column("Images", width=100, anchor='center')

# Add scrollbars
v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

tree.pack(side='left', fill='both', expand=True)
v_scroll.pack(side='right', fill='y')
h_scroll.pack(side='bottom', fill='x')

# Button Frame at bottom
list_button_frame = ttk.Frame(list_tab)
list_button_frame.pack(fill='x', pady=5)

ttk.Button(list_button_frame, text="Generate Report", command=generate_report).pack(side='left', padx=20)
ttk.Button(list_button_frame, text="Delete Event", command=delete_event).pack(side='left', padx=20)

# Initialize the event list
update_event_list()

# Bind Enter key to search
search_entry.bind('<Return>', lambda event: search_events())

root.mainloop()
