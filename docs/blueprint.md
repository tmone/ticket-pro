# **App Name**: TicketCheck Pro

## Core Features:

- Admin Login: Secure the webapp behind password protection, so that only authorized staff can use the webapp
- Excel Data Upload: Allow the admin user to upload an Excel (.xlsx) file containing ticket data: name, phone, email, seat (Row, Number), and unique code.
- QR Code Scanner: Use the device's camera to scan QR codes printed on tickets.
- Ticket Lookup: Search the imported Excel data for the scanned QR code. Display the customer's information (name, seat) if found, or display 'Ticket Not Found'.
- Check-in Recording: Record the check-in time when a valid ticket is scanned. Check all tickets for a "Checked In" state; prevent duplicated check-ins. If the ticket was already used display an alert stating the time of first usage.
- Data Export: Generate and export an Excel (.xlsx) report containing the list of attendees and their check-in times.

## Style Guidelines:

- Primary color: Pink (#d63384) to convey trust and efficiency.
- Background color: Very light gray (#F5F5F5), almost white, to keep the design clean and unobtrusive.
- Accent color: Green (#198754) to highlight important actions like scanning or exporting.
- Body and headline font: 'Inter' sans-serif font for clear and modern readability. Its neutral appearance will assure good legibility in longer passages of text.
- Use clear and recognizable icons for actions like scanning, exporting, and displaying errors.
- Design a responsive layout to ensure the application works well on various devices (desktops, tablets, and mobile phones).
- Use subtle animations to provide feedback during the scanning and data processing stages. An example of such a feedback is to flash the border of the code while scanning it.