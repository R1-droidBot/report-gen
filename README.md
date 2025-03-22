<br>📌 Event Management System<br>
    <br>A Tkinter-based desktop application that allows users to manage events, store event details in an SQLite database, and generate event reports in Word format using docxtpl.<br>

<br>🛠 Features<br>
  <br>✅ Add New Events – Store event details in a local SQLite database.<br>
  <br>✅ View Saved Events – Display stored events in a structured format.<br>
  <br>✅ Upload Images – Attach event-related images.<br>
  <br>✅ Generate Event Report – Export event details into a formatted Word document (.docx).<br>
  <br>✅ User-Friendly UI – Built with Tkinter and ttk Notebook.<br>

<br>🏗️ Tech Stack<br>
  <br>Python – Core logic<br>
  Tkinter – GUI framework<br>
  SQLite – Database for storing events<br>
  docxtpl – Generate Word reports<br>
  ttk Notebook – Tabbed UI<br>

📜 Usage<br>
  1️⃣ Open the application.<br>
  2️⃣ Fill in event details and click "Add Event".<br>
  3️⃣ Upload an image if needed.<br>
  4️⃣ View saved events under the "View Events" tab.<br>
  5️⃣ Generate a report using the "Export Report" tab.<br>

📄 Word Report Template (event_template.docx)<br>
  Your report will be generated using this format:<br>
         Event Report<br>

Event Number: {{ event["Event Number"] }}<br>
Event Name: {{ event["Event Name"] }}<br>
Event In-Charge: {{ event["Event IC"] }}<br>
Date of Conduction: {{ event["Date"] }}<br>
Event Type: {{ event["Event Type"] }}<br>

Report File: {{ event["Report Doc"] }}<br>
GeoTag Photo: {{ event["Geo Photo"] }}<br>

Number of Attendees: {{ event["Attendees"] }}<br>
Resource Person: {{ event["Resource Person"] }}<br>
Designation: {{ event["Designation"] }}<br>
Address: {{ event["Address"] }}<br>

Funding Received: {{ event["Funding"] }}<br>
Number of Days: {{ event["Days"] }}<br>
Organized For: {{ event["Audience"] }}<br>

Mapping with Institute Mission: {{ event["Mission Mapping"] }}<br>
PO-PSO Mapping: {{ event["PO-PSO Mapping"] }}<br>

Remarks: {{ event["Remarks"] }}<br>
📌 Ensure event_template.docx is present in the same folder!<br>



🏆 Credits<br>
Ved Patil<br>
Pushpak Aher<br>
Tejas Chandankar<br>
Suryaprakash Yadav<br>
Harsh Pardeshi<br>
Mentor- Prof. Nilam Khairnar<br>
