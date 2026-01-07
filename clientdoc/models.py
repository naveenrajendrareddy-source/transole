from django.db import models
from django.utils import timezone
from decimal import Decimal
from django.db.models import Max 
from django.db.models import Sum 
from django.db import transaction 
from num2words import num2words # New library
from django.conf import settings
from .constants import INDIAN_STATE_CODES


class SoftDeleteManager(models.Manager):
    def get_queryset(self):
        return super().get_queryset().filter(is_deleted=False)

    def trash(self):
        return super().get_queryset().filter(is_deleted=True)

class SoftDeleteModel(models.Model):
    is_deleted = models.BooleanField(default=False)
    
    objects = SoftDeleteManager()
    all_objects = models.Manager()

    class Meta:
        abstract = True

    def delete(self, using=None, keep_parents=False):
        self.is_deleted = True
        self.save()

    def restore(self):
        self.is_deleted = False
        self.save()
    
    def hard_delete(self):
        super().delete()

class ActivityLog(models.Model):
    action = models.CharField(max_length=255)
    timestamp = models.DateTimeField(auto_now_add=True)
    details = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"{self.timestamp} - {self.action}"




# --- COMPANY AND LOCATION MODELS ---

class OurCompanyProfile(models.Model):
    """Stores the company details (used for PDF headers)."""
    name = models.CharField(max_length=200)
    address = models.TextField()
    contact_number = models.CharField(max_length=20, default='', blank=True)
    email = models.EmailField(default='', blank=True)
    gstin = models.CharField(max_length=15, default='', blank=True)
    state = models.CharField(max_length=50, default="Karnataka")
    state_code = models.CharField(max_length=2, default="29")
    
    # Bank Details
    bank_name = models.CharField(max_length=100, default="BANK OF BARODA")
    account_holder_name = models.CharField(max_length=100, default="Transcend Digital Solutions")
    account_number = models.CharField(max_length=50, default="89440200002020")
    ifsc_code = models.CharField(max_length=20, default="BARB0VJULSO")
    branch_name = models.CharField(max_length=100, default="ULSOOR")
    
    # Signature
    signature = models.ImageField(upload_to='company_signatures/', blank=True, null=True, help_text="Upload signature image for PDFs")

    class Meta:
        verbose_name = "Our Company Profile (Only One Entry)"
        verbose_name_plural = "Our Company Profile (Only One Entry)"

    def __str__(self):
        return self.name

# --- SHARED CHOICES ---
STATE_CHOICES = [
    ('Karnataka', 'Karnataka (29)'),
    ('Maharashtra', 'Maharashtra (27)'),
    ('Tamil Nadu', 'Tamil Nadu (33)'),
    ('Kerala', 'Kerala (32)'),
    ('Andhra Pradesh', 'Andhra Pradesh (37)'),
    ('Telangana', 'Telangana (36)'),
    ('Delhi', 'Delhi (07)'),
    ('Uttar Pradesh', 'Uttar Pradesh (09)'),
    ('Gujarat', 'Gujarat (24)'),
    ('Rajasthan', 'Rajasthan (08)'),
    ('West Bengal', 'West Bengal (19)'),
    ('Punjab', 'Punjab (03)'),
    ('Haryana', 'Haryana (06)'),
    ('Madhya Pradesh', 'Madhya Pradesh (23)'),
    ('Bihar', 'Bihar (10)'),
    ('Odisha', 'Odisha (21)'),
    ('Assam', 'Assam (18)'),
    # Add other states as needed
    ('Other', 'Other')
]

STATE_CODE_MAP = {
    'Karnataka': '29', 'Maharashtra': '27', 'Tamil Nadu': '33', 'Kerala': '32',
    'Andhra Pradesh': '37', 'Telangana': '36', 'Delhi': '07', 'Uttar Pradesh': '09',
    'Gujarat': '24', 'Rajasthan': '08', 'West Bengal': '19', 'Punjab': '03',
    'Haryana': '06', 'Madhya Pradesh': '23', 'Bihar': '10', 'Odisha': '21', 'Assam': '18',
    'Other': ''
}

class Buyer(SoftDeleteModel):
    """Represents a Buyer (Bill To)."""
    name = models.CharField(max_length=255, unique=True)
    address = models.TextField()
    gstin = models.CharField(max_length=15, blank=True, null=True)
    state = models.CharField(max_length=50, choices=STATE_CHOICES, default="Karnataka")
    state_code = models.CharField(max_length=2, default="29")
    pincode = models.CharField(max_length=10, blank=True, null=True)
    
    def save(self, *args, **kwargs):
        if self.state in STATE_CODE_MAP:
            self.state_code = STATE_CODE_MAP[self.state]
        super().save(*args, **kwargs)

    def __str__(self):
        return self.name

class StoreLocation(SoftDeleteModel):
    """Represents a client store location (Ship To/Consignee)."""
    name = models.CharField(max_length=255, unique=True, verbose_name="Site Name")
    site_code = models.CharField(max_length=50, blank=True, null=True)
    address = models.TextField()
    city = models.CharField(max_length=100, blank=True, null=True)
    state = models.CharField(max_length=50, choices=STATE_CHOICES, default="Karnataka")
    state_code = models.CharField(max_length=2, default="29")
    pincode = models.CharField(max_length=10, blank=True, null=True)
    gstin = models.CharField(max_length=15, blank=True, null=True)
    priority = models.CharField(max_length=10, blank=True, null=True, verbose_name="Priority (P1-P4)")
    
    def save(self, *args, **kwargs):
        if self.state in STATE_CODE_MAP:
            self.state_code = STATE_CODE_MAP[self.state]
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.name} ({self.city or 'No City'})"

class ItemCategory(models.Model):
    """Category for items (e.g., Acrylic, Vinyl)."""
    name = models.CharField(max_length=255, unique=True)
    
    def __str__(self):
        return self.name
    
    class Meta:
        verbose_name_plural = "Item Categories"

class Item(SoftDeleteModel):
    """Represents a product or service item for invoicing."""
    category = models.ForeignKey(ItemCategory, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Category/Details")
    name = models.CharField(max_length=255, unique=True, verbose_name="Item Name/Particular")
    description = models.TextField(blank=True, null=True)
    article_code = models.CharField(max_length=100, blank=True, null=True, verbose_name="Article/SKU")
    hsn_sac = models.CharField(max_length=20, default="844311", verbose_name="HSN/SAC")
    price = models.DecimalField(max_digits=10, decimal_places=2, default=0.00)
    unit = models.CharField(max_length=50, default="Nos", verbose_name="Unit (e.g. Nos, Kg)")
    
    GST_CHOICES = [
        (Decimal('0.00'), '0%'),
        (Decimal('0.05'), '5%'),
        (Decimal('0.12'), '12%'),
        (Decimal('0.18'), '18%'),
        (Decimal('0.28'), '28%'),
        (Decimal('0.40'), '40%'),
    ]
    gst_rate = models.DecimalField(max_digits=5, decimal_places=2, default=0.18, choices=GST_CHOICES, verbose_name="GST Rate") 
    # GST Enhancements
    hsn_code = models.CharField(max_length=20, blank=True, null=True, verbose_name="HSN Code")
    
    def save(self, *args, **kwargs):
        # Sync older hsn_sac to new hsn_code if needed, or vice-versa
        if not self.hsn_code and self.hsn_sac:
             self.hsn_code = self.hsn_sac
        super().save(*args, **kwargs)

    def __str__(self):
        return self.name

# --- INVOICE AND RELATED MODELS ---

class SalesInvoice(SoftDeleteModel):
    STATUS_CHOICES = [
        ('DRF', 'Draft (Invoice Created)'),
        ('DC', 'Delivery Challan Logged'),
        ('TRP', 'Transport Charges Logged'),
        ('FIN', 'Finalized (PDF Generated)')
    ]

    buyer = models.ForeignKey(Buyer, on_delete=models.PROTECT, null=True, blank=True, verbose_name="Buyer (Bill To)")
    location = models.ForeignKey(StoreLocation, on_delete=models.PROTECT, verbose_name="Location (Ship To/Consignee)")
    date = models.DateTimeField(default=timezone.now)
    created_at = models.DateTimeField(default=timezone.now)
    total = models.DecimalField(max_digits=10, decimal_places=2, default=0.00)
    status = models.CharField(max_length=3, choices=STATUS_CHOICES, default='DRF')
    
    tally_invoice_number = models.CharField(max_length=50, blank=True, null=True, verbose_name="Tally Invoice No.")
    app_invoice_number = models.CharField(max_length=50, unique=True, blank=True, null=True, verbose_name="App Invoice No.") # Tsol-XXXXX
    
    # New Fields matching the Invoice Image
    delivery_note = models.CharField(max_length=100, blank=True, null=True)
    mode_terms_payment = models.CharField(max_length=100, default="30 Days", verbose_name="Mode/Terms of Payment")
    reference_no_date = models.CharField(max_length=100, blank=True, null=True, verbose_name="Reference No. & Date")
    other_references = models.CharField(max_length=100, default="EMAIL Approval")
    buyers_order_no = models.CharField(max_length=100, blank=True, null=True, verbose_name="Buyer's Order No.")
    buyers_order_date = models.DateTimeField(default=timezone.now, verbose_name="Buyer's Order Date")
    dispatch_doc_no = models.CharField(max_length=100, blank=True, null=True, verbose_name="Dispatch Doc No.")
    delivery_note_date = models.DateTimeField(default=timezone.now, verbose_name="Delivery Note Date")
    dispatched_through = models.CharField(max_length=100, blank=True, null=True, verbose_name="Dispatched through")
    destination = models.CharField(max_length=100, blank=True, null=True)
    terms_of_delivery = models.TextField(blank=True, null=True)
    remark = models.TextField(blank=True, null=True, verbose_name="Remarks")
    
    # GST Calculated Fields (Decimal Integrity)
    customer_gstin = models.CharField(max_length=15, blank=True, null=True)
    place_of_supply = models.CharField(max_length=2, choices=INDIAN_STATE_CODES, blank=True, null=True)
    
    cgst_total = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0.00'))
    sgst_total = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0.00'))
    igst_total = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0.00'))
    
    # Store calculated Words
    amount_in_words = models.CharField(max_length=255, blank=True, null=True)
    tax_amount_in_words = models.CharField(max_length=255, blank=True, null=True)
        
    def calculate_gst_totals(self):
        """Calculates Taxes based on Place of Supply vs Company State."""
        # 1. Fetch Company State
        company_state_code = getattr(settings, 'COMPANY_STATE_CODE', '29')
        # Check specific model if available override
        try:
            profile = OurCompanyProfile.objects.first()
            if profile and profile.state_code:
                company_state_code = profile.state_code
        except Exception:
            pass
            
        # 2. Determine POS
        if not self.place_of_supply:
            # Fallback to location state code
             self.place_of_supply = self.location.state_code if self.location else '29'
             
        # Auto-populate GSTIN if missing
        if not self.customer_gstin:
            if self.buyer and self.buyer.gstin:
                self.customer_gstin = self.buyer.gstin
            elif self.location and self.location.gstin:
                self.customer_gstin = self.location.gstin
        
        # 3. Determine Tax Type
        is_inter_state = (self.place_of_supply != company_state_code)
        
        total_cgst = Decimal('0.00')
        total_sgst = Decimal('0.00')
        total_igst = Decimal('0.00')
        grand_total = Decimal('0.00')
        total_tax = Decimal('0.00')

        for item in self.invoiceitem_set.all():
            taxable = item.taxable_value
            gst_rate = item.item.gst_rate # Use item rate or snapshot?? Models say calculate_total used snapshot.
            # Use snapshot if available
            if item.gst_rate is not None:
                gst_rate = item.gst_rate

            tax_amount = (taxable * gst_rate).quantize(Decimal('0.01'))
            
            if is_inter_state:
                total_igst += tax_amount
            else:
                half_tax = (tax_amount / Decimal('2.00')).quantize(Decimal('0.01'))
                total_cgst += half_tax
                total_sgst += half_tax
            
            grand_total += taxable + tax_amount
            total_tax += tax_amount

        # --- Add Transport Charges if Any ---
        if hasattr(self, 'transportcharges'):
            trp = self.transportcharges
            if trp and trp.charges > 0:
                # Assuming 18% GST on Transport for now (Standard Service Rate)
                # Or check if specific config exists. User said "along with gst".
                trp_taxable = trp.charges
                trp_gst_rate = Decimal('0.18') 
                
                trp_tax_amt = (trp_taxable * trp_gst_rate).quantize(Decimal('0.01'))
                
                if is_inter_state:
                    total_igst += trp_tax_amt
                else:
                    half_tax = (trp_tax_amt / Decimal('2.00')).quantize(Decimal('0.01'))
                    total_cgst += half_tax
                    total_sgst += half_tax
                
                grand_total += trp_taxable + trp_tax_amt
                total_tax += trp_tax_amt

        self.cgst_total = total_cgst
        self.sgst_total = total_sgst
        self.igst_total = total_igst
        self.total = grand_total
        
        # Word Conversion
        try:
             self.amount_in_words = "INR " + num2words(self.total, lang='en_IN').title() + " Only"
             self.tax_amount_in_words = "INR " + num2words(total_tax, lang='en_IN').title() + " Only"
        except Exception:
             self.amount_in_words = "Error generating words"

        self.save()

    def save(self, *args, **kwargs):
        if not self.app_invoice_number:
            # FIX: Robust sequential number generation with retry logic
            for attempt in range(5): # Retry up to 5 times
                try:
                    with transaction.atomic():
                        # Find the highest existing app_invoice_number sequence
                        last_invoice = SalesInvoice.objects.all().order_by('-app_invoice_number').only('app_invoice_number').first()
                        
                        new_seq = 1
                        if last_invoice and last_invoice.app_invoice_number:
                            try:
                                last_seq = int(last_invoice.app_invoice_number.split('-')[1])
                                new_seq = last_seq + 1
                            except (ValueError, IndexError):
                                new_seq = 1
        
                        self.app_invoice_number = f"Tsol-{new_seq:05d}"
                        super().save(*args, **kwargs)
                        break # Success!
                except Exception: # Catch IntegrityError or other save issues
                    if attempt == 4: raise # Re-raise if last attempt
                    continue # Try again
        else:
            super().save(*args, **kwargs)

    def calculate_total(self):
        """Wrapper for new calculate_gst_totals to maintain compatibility."""
        self.calculate_gst_totals()
        
    def get_status_color(self):
        if self.status == 'FIN':
            return 'success'
        if self.status == 'TRP':
            return 'info'
        if self.status == 'DC':
            return 'warning'
        return 'secondary' 

    def __str__(self):
        return f"Invoice {self.tally_invoice_number or self.app_invoice_number} - {self.location.name}"

class InvoiceItem(models.Model):
    """Junction model for SalesInvoice and Item (the line items)."""
    invoice = models.ForeignKey(SalesInvoice, on_delete=models.CASCADE)
    item = models.ForeignKey(Item, on_delete=models.PROTECT)
    quantity = models.IntegerField(default=1)
    
    # Snapshot fields for historical accuracy
    price = models.DecimalField(max_digits=10, decimal_places=2, default=0.00)
    description = models.TextField(blank=True, null=True, verbose_name="Item Description (Snapshot)")
    quantity_shipped = models.IntegerField(default=1)
    quantity_billed = models.IntegerField(default=1)
    
    DISCOUNT_CHOICES = [
        ('Percentage', 'Percentage (%)'),
        ('Amount', 'Amount (Fixed)'),
    ]
    discount_type = models.CharField(max_length=20, choices=DISCOUNT_CHOICES, default='Percentage')
    discount_value = models.DecimalField(max_digits=10, decimal_places=2, default=0.00)

    GST_CHOICES = [
        (Decimal('0.00'), '0%'),
        (Decimal('0.05'), '5%'),
        (Decimal('0.12'), '12%'),
        (Decimal('0.18'), '18%'),
        (Decimal('0.28'), '28%'),
        (Decimal('0.40'), '40%'),
    ]
    gst_rate = models.DecimalField(max_digits=5, decimal_places=2, default=0.18, choices=GST_CHOICES)

    def save(self, *args, **kwargs):
        # Auto-populate from item if not set
        if self._state.adding:
            if self.price == 0:
                 self.price = self.item.price
            # If gst_rate is not explicitly set (or default), try to set from item
            # We check if it is the default 0.18, but maybe user wanted 18%. 
            # Better: if not set in kwargs (hard to check here), but checking vs item.
            pass # We rely on view or manual input.
        super().save(*args, **kwargs)

    @property
    def gross_amount(self):
        """Returns Quantity Billed * Price"""
        return (Decimal(self.quantity_billed) * self.price).quantize(Decimal('0.01'))

    @property
    def discount_amount(self):
        """Returns calculated discount amount"""
        gross = self.gross_amount
        if self.discount_type == 'Percentage':
            return (gross * (self.discount_value / Decimal('100.00'))).quantize(Decimal('0.01'))
        return self.discount_value.quantize(Decimal('0.01'))

    @property
    def taxable_value(self):
        """Returns Gross - Discount"""
        val = self.gross_amount - self.discount_amount
        return val if val > 0 else Decimal('0.00')

    def __str__(self):
        return f"{self.item.name} for Invoice {self.invoice.id}"

class DeliveryChallan(SoftDeleteModel):
    """Stores data for the Delivery Challan document."""
    invoice = models.OneToOneField(SalesInvoice, on_delete=models.CASCADE)
    date = models.DateTimeField(default=timezone.now)
    created_at = models.DateTimeField(default=timezone.now)
    notes = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"DC for Invoice {self.invoice.id}"

class TransportCharges(SoftDeleteModel):
    """Stores data for the Transport Charges document."""
    invoice = models.OneToOneField(SalesInvoice, on_delete=models.CASCADE)
    date = models.DateTimeField(default=timezone.now) 
    created_at = models.DateTimeField(default=timezone.now)
    charges = models.DecimalField(max_digits=10, decimal_places=2, default=0.00) 
    description = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"Transport for Invoice {self.invoice.id}"

class ConfirmationDocument(SoftDeleteModel):
    """Stores uploaded files and final combined PDF."""
    invoice = models.OneToOneField(SalesInvoice, on_delete=models.CASCADE)
    
    # This field resolves the date formatting issue in the confirmation_list template
    date = models.DateTimeField(default=timezone.now) 
    created_at = models.DateTimeField(default=timezone.now)
    
    # Uploaded Documents
    po_file = models.FileField(upload_to='confirmation_docs/po/', blank=True, null=True, verbose_name="PO Copy")
    approval_email_file = models.FileField(upload_to='confirmation_docs/email/', blank=True, null=True, verbose_name="Approval Email PDF")
    
    # Custom Uploads (Overrides generated ones if present)
    uploaded_invoice = models.FileField(upload_to='confirmation_docs/inv/', blank=True, null=True, verbose_name="Custom Invoice PDF")
    uploaded_dc = models.FileField(upload_to='confirmation_docs/dc/', blank=True, null=True, verbose_name="Custom DC PDF")

    # Final Output
    combined_pdf = models.FileField(upload_to='confirmations/', blank=True, null=True)
    
    def __str__(self):
        return f"Confirmation for Invoice {self.invoice.id}"

class PackedImage(models.Model):
    """Stores multiple images of packed goods linked to a ConfirmationDocument."""
    confirmation = models.ForeignKey(ConfirmationDocument, on_delete=models.CASCADE, null=True) 
    image = models.ImageField(upload_to='packed_images/')
    notes = models.TextField(blank=True, null=True, verbose_name="Image Notes") 

    def __str__(self):
        invoice_id = self.confirmation.invoice.id if self.confirmation and self.confirmation.invoice else "N/A"
        return f"Image for Confirmation {invoice_id}"

class BulkInvoiceUpload(models.Model):
    """Tracks bulk excel uploads for invoice generation."""
    file = models.FileField(upload_to='bulk_uploads/')
    uploaded_at = models.DateTimeField(auto_now_add=True)
    status = models.CharField(max_length=20, default='Pending', choices=[('Pending', 'Pending'), ('Processed', 'Processed'), ('Failed', 'Failed')])
    log = models.TextField(blank=True, null=True, help_text="Log of success/errors during processing")

    def __str__(self):
        return f"Upload {self.id} at {self.uploaded_at}"