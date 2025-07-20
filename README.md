# Invoice-Parser-Access-Database-Automation


This repository showcases an end-to-end invoice processing automation system using **VBA (Visual Basic for Applications)** and **Microsoft Access**. The solution extracts, parses, and stores invoice data received via email into a relational Access database.

---

## 📧 Project Overview

Invoices were regularly received as **PDF attachments** via email. The goal was to automate the entire process from extraction to structured storage.

### Key Steps:

1. **Email Integration:**

   - Used VBA to scan an Outlook inbox.
   - Filtered relevant emails with PDF invoice attachments.
   - Saved attachments to a local folder.

2. **Text Extraction:**

   - Converted PDF invoices to **.txt files** for easier parsing.
   - Stored the original PDFs and their text versions side-by-side for reference.

3. **Data Parsing & Storage:**

   - Parsed key fields from invoice text (e.g., Invoice Number, Product Details, Client Info).
   - Inserted parsed data into a structured **MS Access database** using VBA macros.

4. **Relational Schema:**

   - Built a normalized relational schema in Access with the following tables:
     - **Client Details**: Stores client company info.
     - **Invoices**: Invoice header data.
     - **Invoice Details**: Line items for each invoice.
     - **Payments**: Associated payment records.

---

## 📁 Repository Structure

```
|-- /PDFs           # Original PDF invoices
|-- /TextFiles      # Converted .txt versions of invoices
|-- /AccessDB       # MS Access database file (InvoiceDB.accdb)
|-- /Screenshots    # Images of table relationships, macros, and file structure
|-- README.md       # Project documentation (this file)
```

---

## 📈 Macros Implemented

Here are some of the macros built within Access:

- `ExtractTextFromPDFs`
- `ParseAllInvoices`
- `ParseAndInsertInvoiceDetails_UsingTwoVars`
- `PopulateClientDetails`
- `PopulateInvoices`
- `SaveAttachmentsFromEmails`
- `SaveFilteredAttachments`

These automate the end-to-end data flow from email to Access.

---

## ⚙️ Technologies Used

- Microsoft Access
- VBA (Visual Basic for Applications)
- Outlook (for email automation)
- PDF to Text conversion tools (external or script-based)

---

## 🚀 Future Enhancements

- Integrate error logging and validation checks
- Add UI form for manual review or corrections
- Export invoice data to Excel or Power BI for visualization
