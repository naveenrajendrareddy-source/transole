# Transol VMS - Vendor Management System

A comprehensive Django-based Vendor Management System for invoice generation, document management, and GST compliance.

## 🚀 Features

### Core Functionality
- **Bulk Invoice Creation** - Upload Excel files to create multiple invoices at once
- **Automatic PDF Generation** - Generate professional tax invoices, delivery challans, and transport bills
- **GST Compliance** - Automatic CGST/SGST/IGST calculation based on Place of Supply
- **Document Bundling** - Combine invoices, delivery notes, transport charges, email approvals, and images into single PDF
- **Soft Delete** - Safe deletion with restore capability
- **Sequential Invoice Numbering** - Auto-generated invoice numbers (Tsol-00001, Tsol-00002, etc.)

### Document Types
1. **Tax Invoice** - Complete GST-compliant invoice with tax matrix
2. **Delivery Challan** - Delivery note with item details
3. **Transport Charges Bill** - Separate billing for transport with GST
4. **Email Approval** - Attach approval emails
5. **Packed Images** - Product/delivery images

### Bulk Upload Features
- Excel template with dropdowns and auto-fill
- Support for multiple items per invoice
- File attachment support (PO, Email, Images)
- Automatic DC and Transport document generation
- Comprehensive error logging
- Duplicate detection and prevention

## 📋 Requirements

- Python 3.9+
- Django 4.2+
- SQLite (development) / PostgreSQL (production recommended)
- See `requirements.txt` for full dependencies

## 🛠️ Installation

### Quick Start (Windows)

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd transole-main
   ```

2. **Run the portable launcher**
   ```bash
   run_transole_portable.bat
   ```
   This will automatically:
   - Create virtual environment
   - Install dependencies
   - Run migrations
   - Start the server
   - Open browser to http://127.0.0.1:8000/

### Manual Setup

1. **Create virtual environment**
   ```bash
   python -m venv venv
   ```

2. **Activate virtual environment**
   - Windows: `venv\Scripts\activate`
   - Mac/Linux: `source venv/bin/activate`

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Create .env file**
   ```bash
   copy .env.example .env
   ```
   Edit `.env` and add your configuration:
   ```
   SECRET_KEY=your-secret-key-here
   DEBUG=True
   ALLOWED_HOSTS=localhost,127.0.0.1
   ```

5. **Run migrations**
   ```bash
   python manage.py migrate
   ```

6. **Create superuser**
   ```bash
   python manage.py createsuperuser
   ```

7. **Run server**
   ```bash
   python manage.py runserver
   ```

8. **Access the application**
   - Main app: http://127.0.0.1:8000/
   - Admin panel: http://127.0.0.1:8000/admin/

## 📊 Usage

### Setting Up Master Data

1. **Company Profile** - Add your company details in Admin panel
2. **Buyers** - Add buyer/customer information
3. **Store Locations** - Add delivery locations (Ship To addresses)
4. **Items** - Add products/services with HSN codes and GST rates

### Bulk Upload Process

1. **Download Template**
   - Go to Bulk Upload page
   - Click "Download Template"
   - Excel file with pre-configured columns and dropdowns

2. **Fill Excel File**
   - Column A: Buyer Name (dropdown)
   - Column B: Location Name (dropdown)
   - Column C: Item Name (dropdown)
   - Column D: Item Description (optional)
   - Column E: Quantity
   - Column J: Transport Charges (if applicable)
   - Column L: Generate Invoice (Yes/No)
   - Column M: Generate PDF (Yes/No)
   - Column N: Tally Invoice Number (groups items)
   - Column U: Delivery Note Number
   - Column V: Delivery Note Date
   - Columns AD-AK: File paths for attachments

3. **Upload and Process**
   - Upload filled Excel file
   - System processes and generates PDFs
   - Download bundled PDF from confirmation list

### Excel Column Guide

Hover over column headers in the template to see tooltips with usage instructions.

**Key Columns:**
- **Column N (Tally Invoice No.)**: Use same number for multiple rows to group items into one invoice
- **Column U (Delivery Note)**: Auto-generates DC if filled
- **Column J (Transport Charges)**: Auto-generates transport bill if filled
- **Column AG (Email Approval)**: Path to approval email PDF
- **Columns AH-AL (Images)**: Paths to product/delivery images

## 🔧 Configuration

### Database (Production)

For production, switch to PostgreSQL:

```python
# settings.py
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'transol_db',
        'USER': 'your_db_user',
        'PASSWORD': 'your_db_password',
        'HOST': 'localhost',
        'PORT': '5432',
    }
}
```

### Email Configuration

Update `.env` file:
```
EMAIL_HOST_USER=your_email@gmail.com
EMAIL_HOST_PASSWORD=your_app_password
```

## 📁 Project Structure

```
transole-main/
├── clientdoc/              # Main application
│   ├── models.py          # Database models
│   ├── views.py           # View functions
│   ├── pdf_generator.py  # PDF generation logic
│   ├── forms.py           # Django forms
│   └── templates/         # HTML templates
├── transol/               # Project settings
│   ├── settings.py        # Configuration
│   └── urls.py            # URL routing
├── media/                 # Uploaded files
├── static/                # CSS, JS, images
├── requirements.txt       # Python dependencies
├── manage.py              # Django management
└── run_transole_portable.bat  # Quick launcher
```

## 🐛 Troubleshooting

### Common Issues

1. **UNIQUE constraint failed**
   - Fixed in latest version
   - System now checks all invoices including soft-deleted ones

2. **Transport charges not showing**
   - Ensure Column J has numeric value
   - Check that "Generate PDF" is set to "Yes"

3. **Images appearing twice**
   - Fixed in latest version with deduplication logic

4. **File not found errors**
   - Check file paths in Excel are absolute paths
   - Ensure files exist at specified locations
   - Check bulk upload log for specific errors

## 📝 Recent Updates (Latest Commit)

- ✅ Fixed UNIQUE constraint error in bulk uploads
- ✅ DC Number now uses Excel Column U value
- ✅ Transport Bill number format: TC-{TallyInvoiceNumber}
- ✅ Removed duplicate images in PDF
- ✅ Updated Transport HSN code to 997619
- ✅ Improved PDF bundling order
- ✅ Enhanced error logging
- ✅ Added Excel column guidance

## 🤝 Contributing

This is a private project. For any issues or suggestions, please contact the development team.

## 📄 License

Proprietary - All rights reserved

## 👥 Credits

Developed for Transol VMS
Last Updated: January 2026

## 📞 Support

For support and queries, please contact the project administrator.

---

**Note**: This system is designed for Indian GST compliance. Ensure all HSN/SAC codes and tax rates are updated according to current regulations.
