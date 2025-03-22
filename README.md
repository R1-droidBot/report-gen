📌 Event Management System
    A Tkinter-based desktop application that allows users to manage events, store event details in an SQLite database, and generate event reports in Word format using docxtpl.

🛠 Features
  ✅ Add New Events – Store event details in a local SQLite database.
  ✅ View Saved Events – Display stored events in a structured format.
  ✅ Upload Images – Attach event-related images.
  ✅ Generate Event Report – Export event details into a formatted Word document (.docx).
  ✅ User-Friendly UI – Built with Tkinter and ttk Notebook.

🏗️ Tech Stack
  Python – Core logic
  Tkinter – GUI framework
  SQLite – Database for storing events
  docxtpl – Generate Word reports
  ttk Notebook – Tabbed UI

📜 Usage
  1️⃣ Open the application.
  2️⃣ Fill in event details and click "Add Event".
  3️⃣ Upload an image if needed.
  4️⃣ View saved events under the "View Events" tab.
  5️⃣ Generate a report using the "Export Report" tab.

📄 Word Report Template (event_template.docx)
  Your report will be generated using this format:
         Event Report

Event Number: {{ event["Event Number"] }}
Event Name: {{ event["Event Name"] }}
Event In-Charge: {{ event["Event IC"] }}
Date of Conduction: {{ event["Date"] }}
Event Type: {{ event["Event Type"] }}

Report File: {{ event["Report Doc"] }}
GeoTag Photo: {{ event["Geo Photo"] }}

Number of Attendees: {{ event["Attendees"] }}
Resource Person: {{ event["Resource Person"] }}
Designation: {{ event["Designation"] }}
Address: {{ event["Address"] }}

Funding Received: {{ event["Funding"] }}
Number of Days: {{ event["Days"] }}
Organized For: {{ event["Audience"] }}

Mapping with Institute Mission: {{ event["Mission Mapping"] }}
PO-PSO Mapping: {{ event["PO-PSO Mapping"] }}

Remarks: {{ event["Remarks"] }}
📌 Ensure event_template.docx is present in the same folder!



🏆 Credits
Ved Patil
Pushpak Aher
Tejas Chandankar
Suryaprakash Yadav
Harsh Pardeshi
Mentor- Prof. Nilam Khairnar
