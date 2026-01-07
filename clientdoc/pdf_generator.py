from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
from io import BytesIO
from decimal import Decimal

# Register Font for INR Symbol if available
# User requested fallback to Rs. if issues persist.
INR_SYMBOL = 'Rs.'

def clean(val): return str(val) if val else "-"
def clean_date(d): return d.strftime('%d-%b-%y') if d else ""

def create_header_table(title, company):
    styles = getSampleStyleSheet()
    
    # Handle missing company profile - Try to fetch if not passed
    if not company:
        from .models import OurCompanyProfile
        company = OurCompanyProfile.objects.first()

    # If still no company, use blank placeholders to avoid "Dummy Value" confusion
    if not company:
        class DefaultCompany:
            name = ""
            address = ""
            gstin = ""
            email = ""
            contact_number = ""
            state = ""
            state_code = ""
            bank_name = ""
            account_holder_name = ""
            account_number = ""
            ifsc_code = ""
            branch_name = ""
            signature = None # Fix missing attribute
        company = DefaultCompany()
    
    header_data = [
        [Paragraph(title, styles['Title'])],
        [Paragraph(company.name, styles['Heading3'])],
        [Paragraph(company.address.replace('\n', '<br/>'), styles['Normal'])],
    ]
    t = Table(header_data, colWidths=[180*mm])
    t.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    return t, company # Return company object in case it was the default one



def create_footer_with_signature(company, notes=""):
    styles = getSampleStyleSheet()
    style_small = ParagraphStyle('Small', parent=styles['Normal'], fontSize=8)
    
    footer_left = f"""
    <br/><u>Remarks/Notes:</u><br/>
    {notes or '-'}<br/><br/>
    <u>Declaration</u><br/>
    We declare that this document shows the actual details described and that all particulars are true and correct.
    """
    
    # Right side text
    footer_right_text = f"""
    Company's Bank Details<br/>
    A/c Holder's Name: <b>{company.account_holder_name}</b><br/>
    Bank Name: <b>{company.bank_name}</b><br/>
    A/c No.: <b>{company.account_number}</b><br/>
    Branch & IFS Code: <b>{company.branch_name} & {company.ifsc_code}</b><br/><br/>
    <b>for {company.name}</b><br/>
    """
    
    # Signature Image logic
    # Signature Image logic
    signature_img = None
    img_path = None
    
    # 1. Try Company Signature (Database)
    if hasattr(company, 'signature') and company.signature:
        try:
            img_path = company.signature.path
        except Exception:
            img_path = None
            
    # 2. Universal Fallback (Local File)
    if not img_path or not os.path.exists(img_path):
        fallback_path = os.path.join(os.path.dirname(__file__), 'signature.png')
        if os.path.exists(fallback_path):
            img_path = fallback_path

    # 3. Create Image Object if path found
    if img_path:
        try:
            # Check if file effectively exists before trying to open to avoid errors
            if os.path.exists(img_path):
                signature_img = Image(img_path, width=40*mm, height=15*mm)
                signature_img.hAlign = 'RIGHT'
        except Exception as e:
            print(f"Error loading signature: {e}")
            pass

    auth_sig_text = "<br/>Authorised Signatory"

    # Assemble Right Cell Content
    right_elements = [Paragraph(footer_right_text, style_small)]
    if signature_img:
        right_elements.append(signature_img)
    else:
        right_elements.append(Spacer(1, 15*mm)) # Space for manual sig
        
    right_elements.append(Paragraph(auth_sig_text, style_small))
    
    footer_data = [[Paragraph(footer_left, style_small), right_elements]]
    t_foot = Table(footer_data, colWidths=[95*mm, 95*mm])
    t_foot.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    return t_foot

def generate_invoice_pdf(invoice, company_input):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    elements = []
    styles = getSampleStyleSheet()
    
    t_header, company = create_header_table("TAX INVOICE", company_input)
    # elements.append(t_header) # Using a custom header structure for Invoice as per original
    
    # Custom styles
    style_normal = styles['Normal']
    style_normal.fontSize = 9
    style_bold = ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9)
    style_small = ParagraphStyle('Small', parent=styles['Normal'], fontSize=8)
    style_header = ParagraphStyle('Header', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, alignment=1) # Center

    # --- Title ---
    elements.append(Paragraph("TAX INVOICE", style_header))
    elements.append(Spacer(1, 5*mm))

    # --- Header Section (Seller, Consignee, Buyer vs Invoice Details) ---
    
    # LEFT COLUMN CONTENT
    seller_text = f"""
    <b>{company.name}</b><br/>
    {company.address}<br/>
    GSTIN/UIN: {company.gstin}<br/>
    State Name: {company.state}, Code: {company.state_code}<br/>
    Contact: {company.contact_number}<br/>
    E-Mail: {company.email}
    """
    
    consignee_text = f"""
    <br/><br/><b>Consignee (Ship to)</b><br/>
    <b>{invoice.location.name}</b><br/>
    {invoice.location.address}<br/>
    GSTIN/UIN: {clean(invoice.location.gstin)}<br/>
    State Name: {invoice.location.state}, Code: {invoice.location.state_code}
    """
    
    buyer_obj = invoice.buyer if invoice.buyer else invoice.location
    buyer_text = f"""
    <br/><br/><b>Buyer (Bill to)</b><br/>
    <b>{buyer_obj.name}</b><br/>
    {buyer_obj.address}<br/>
    GSTIN/UIN: {clean(buyer_obj.gstin)}<br/>
    State Name: {buyer_obj.state}, Code: {buyer_obj.state_code}<br/>
    """
    
    left_content = [Paragraph(seller_text + consignee_text + buyer_text, style_normal)]

    # RIGHT COLUMN CONTENT
    # Show both Tally and App Invoice No
    inv_no_str = ""
    if invoice.tally_invoice_number:
        inv_no_str += f"Tally Inv: {invoice.tally_invoice_number}<br/>"
    
    right_data = [
        [Paragraph("<b>Invoice No.</b>", style_small), Paragraph(f"<b>{inv_no_str}</b>", style_bold), Paragraph("<b>Dated</b>", style_small), Paragraph(f"<b>{clean_date(invoice.date)}</b>", style_bold)],
        [Paragraph("<b>Delivery Note</b>", style_small), Paragraph(clean(invoice.delivery_note), style_normal), Paragraph("<b>Mode/Terms of Payment</b>", style_small), Paragraph(clean(invoice.mode_terms_payment), style_normal)],
        [Paragraph("<b>Reference No. & Date.</b>", style_small), Paragraph(clean(invoice.reference_no_date), style_normal), Paragraph("<b>Other References</b>", style_small), Paragraph(clean(invoice.other_references), style_normal)],
        [Paragraph("<b>Buyer's Order No.</b>", style_small), Paragraph(clean(invoice.buyers_order_no), style_normal), Paragraph("<b>Dated</b>", style_small), Paragraph(clean_date(invoice.buyers_order_date), style_normal)],
        [Paragraph("<b>Dispatch Doc No.</b>", style_small), Paragraph(clean(invoice.dispatch_doc_no), style_normal), Paragraph("<b>Delivery Note Date</b>", style_small), Paragraph(clean_date(invoice.delivery_note_date), style_normal)],
        [Paragraph("<b>Dispatched through</b>", style_small), Paragraph(clean(invoice.dispatched_through), style_normal), Paragraph("<b>Destination</b>", style_small), Paragraph(clean(invoice.destination), style_normal)],
        [Paragraph("<b>Terms of Delivery</b>", style_small), Paragraph(clean(invoice.terms_of_delivery), style_normal), '', ''] 
    ]
    
    right_table = Table(right_data, colWidths=[25*mm, 25*mm, 25*mm, 25*mm])
    right_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('SPAN', (1,6), (3,6)), 
    ]))

    main_header_data = [[left_content, right_table]]
    main_table = Table(main_header_data, colWidths=[95*mm, 100*mm]) 
    main_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(main_table)
    
    # --- Items Table ---
    item_header = ['Sl No.', 'Description of Goods', 'HSN/SAC', 'Quantity', 'Rate', 'per', 'Amount']
    item_data = [item_header]
    
    total_qty = 0
    invoice.calculate_total()
    
    for idx, item in enumerate(invoice.invoiceitem_set.all(), 1):
        total_qty += item.quantity
        item_data.append([
            str(idx),
            Paragraph(f"<b>{item.item.name}</b><br/>{item.description or item.item.description or ''}", style_normal),
            item.item.hsn_sac,
            f"{item.quantity} Nos",
            f"Rs. {item.price}", # Snapshot price
            item.item.unit or "Nos", # Use actual unit from Item master
            f"Rs. {item.quantity * item.price}" # Snapshot price
        ])
    
    taxable_value = sum(i.quantity * i.price for i in invoice.invoiceitem_set.all())
    
    # We can't display a single rate if items have different rates. 
    # For the item table, we can show "See Tax Tbl" or the specific rate if passed.
    # The current loop logic (not shown above, but in original code) iterates items.
    
    bill_details = f"Bill Details: New Ref {clean(invoice.tally_invoice_number or invoice.app_invoice_number)} 30 Days {invoice.total} Dr"
    
    # Calculate totals for summary (simplified for view)
    # Calculate totals for summary
    total_cgst = sum(i.quantity * i.price * (i.gst_rate / 2) for i in invoice.invoiceitem_set.all())
    total_sgst = sum(i.quantity * i.price * (i.gst_rate / 2) for i in invoice.invoiceitem_set.all())
    
    # Transport Charges Injection
    if hasattr(invoice, 'transportcharges') and invoice.transportcharges and invoice.transportcharges.charges > 0:
        trp = invoice.transportcharges
        trp_val = trp.charges
        item_data.append([
            str(len(invoice.invoiceitem_set.all()) + 1),
            Paragraph(f"<b>Transport Charges</b><br/>{trp.description or ''}", style_normal),
            '9967',
            "1",
            f"Rs. {trp_val}",
            "",
            f"Rs. {trp_val}"
        ])
        
        # Add to Tax Calculation
        total_cgst += trp_val * Decimal('0.09') # 9%
        total_sgst += trp_val * Decimal('0.09') # 9%
        taxable_value += trp_val
        
    # We display the TOTAL CGST/SGST in the item table now, instead of per rate, 
    # because the detailed breakdown is in the Tax Analysis Matrix below.
    item_data.append(['', Paragraph(f"<b>Output CGST (Total)</b>", style_normal), '', '', '', '', f"Rs. {total_cgst:.2f}"])
    item_data.append(['', Paragraph(f"<b>Output SGST (Total)</b>", style_normal), '', '', '', '', f"Rs. {total_sgst:.2f}"])
    item_data.append(['', Paragraph(f"<br/><b>Bill Details:</b><br/>{bill_details}", style_small), '', '', '', '', ''])

    item_data.append(['', 'Total', '', f"{total_qty} Nos", '', '', f"{INR_SYMBOL} {invoice.total}"])
    
    col_widths = [10*mm, 78*mm, 20*mm, 25*mm, 20*mm, 10*mm, 25*mm]
    
    t_items = Table(item_data, colWidths=col_widths)
    t_items.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (1,0), (1,-1), 'LEFT'), 
        ('ALIGN', (-1,0), (-1,-1), 'RIGHT'), 
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), 
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTNAME', (-1,-1), (-1,-1), 'Helvetica-Bold'), 
        ('SPAN', (1,-2), (6,-2)), 
        ('SPAN', (1,-1), (2,-1)), 
    ]))
    elements.append(t_items)
    
    elements.append(Paragraph(f"Amount Chargeable (in words)<br/><b>{invoice.amount_in_words or ''}</b>", style_normal))
    elements.append(Spacer(1, 2*mm))

    # --- Tax Analysis Matrix ---
    tax_data = [
        ['HSN/SAC', 'Taxable Value', 'CGST', '', 'SGST', '', 'Total Tax Amount'],
        ['', '', 'Rate', 'Amount', 'Rate', 'Amount', '']
    ]
    
    hsn_map = {} 
    for item in invoice.invoiceitem_set.all():
        h = item.item.hsn_sac
        val = item.quantity * item.price
        rate = item.gst_rate
        
        # Key by HSN AND Rate to separate different tax rates for same HSN (rare but possible)
        key = (h, rate)
        hsn_map[key] = hsn_map.get(key, 0) + val
        
    # Add Transport to HSN Map
    if hasattr(invoice, 'transportcharges') and invoice.transportcharges and invoice.transportcharges.charges > 0:
        k = ('9967', Decimal('0.18'))
        hsn_map[k] = hsn_map.get(k, 0) + invoice.transportcharges.charges
        
    total_tax_amt = 0
    for (hsn, rate), val in hsn_map.items():
        c_rate = rate / 2
        s_rate = rate / 2
        c_amt = val * c_rate
        s_amt = val * s_rate
        tot_t = c_amt + s_amt
        total_tax_amt += tot_t
        
        tax_data.append([
            hsn, f"Rs. {val:.2f}", f"{c_rate*100:.1f}%", f"Rs. {c_amt:.2f}", f"{s_rate*100:.1f}%", f"Rs. {s_amt:.2f}", f"Rs. {tot_t:.2f}"
        ])
    
    tax_data.append([
        'Total', f"Rs. {taxable_value:.2f}", '', f"Rs. {total_cgst:.2f}", '', f"Rs. {total_sgst:.2f}", f"Rs. {total_tax_amt:.2f}"
    ])
    
    t_tax = Table(tax_data, colWidths=[25*mm, 35*mm, 15*mm, 30*mm, 15*mm, 30*mm, 40*mm])
    t_tax.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
        ('ALIGN', (0,0), (0,-1), 'LEFT'), 
        ('SPAN', (2,0), (3,0)), 
        ('SPAN', (4,0), (5,0)), 
        ('SPAN', (0,0), (0,1)), 
        ('SPAN', (1,0), (1,1)), 
        ('SPAN', (6,0), (6,1)), 
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
    ]))
    elements.append(t_tax)
    
    elements.append(Paragraph(f"Tax Amount (in words) : <b>{invoice.tax_amount_in_words or ''}</b>", style_normal))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(create_footer_with_signature(company, invoice.delivery_note))
    
    elements.append(Paragraph("This is a Computer Generated Invoice", ParagraphStyle('Center', parent=styles['Normal'], alignment=1, fontSize=8)))

    doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_dc_pdf(invoice, dc, company_input):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    elements = []
    styles = getSampleStyleSheet()
    
    t_header, company = create_header_table("DELIVERY CHALLAN", company_input)
    # elements.append(t_header) # Using a custom header structure for consistent look
    
    style_normal = styles['Normal']
    style_normal.fontSize = 9
    style_bold = ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9)
    style_header = ParagraphStyle('Header', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, alignment=1)

    elements.append(Paragraph("DELIVERY CHALLAN", style_header))
    elements.append(Spacer(1, 5*mm))
    
    # Left: Seller details + Consignee
    seller_text = f"""
    <b>{company.name}</b><br/>
    {company.address}<br/>
    GSTIN/UIN: {company.gstin}<br/>
    State Name: {company.state}, Code: {company.state_code}<br/>
    Contact: {company.contact_number}
    """
    
    consignee_text = f"""
    <br/><br/><b>Consignee (Ship to)</b><br/>
    <b>{invoice.location.name}</b><br/>
    {invoice.location.address}<br/>
    GSTIN/UIN: {clean(invoice.location.gstin)}<br/>
    """
    
    left_content = [Paragraph(seller_text + consignee_text, style_normal)]
    
    # Right: DC Details
    dc_data = [
        [Paragraph("<b>DC No.</b>", style_normal), Paragraph(f"<b>DC-{invoice.tally_invoice_number or invoice.app_invoice_number}</b>", style_bold)],
        [Paragraph("<b>Date</b>", style_normal), Paragraph(f"<b>{clean_date(dc.date)}</b>", style_bold)],
        [Paragraph("<b>Ref Invoice No.</b>", style_normal), Paragraph(clean(invoice.tally_invoice_number or invoice.app_invoice_number), style_normal)],
        [Paragraph("<b>Vehicle No.</b>", style_normal), Paragraph(clean(dc.notes), style_normal)],
        [Paragraph("<b>Dispatched through</b>", style_normal), Paragraph(clean(invoice.dispatched_through), style_normal)],
        [Paragraph("<b>Destination</b>", style_normal), Paragraph(clean(invoice.destination), style_normal)],
    ]
    
    right_table = Table(dc_data, colWidths=[35*mm, 55*mm])
    right_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    
    main_header_data = [[left_content, right_table]]
    main_table = Table(main_header_data, colWidths=[95*mm, 95*mm])
    main_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    elements.append(main_table)
    
    # Items
    item_header = ['Sl No', 'Description of Goods', 'HSN/SAC', 'Quantity', 'Remarks']
    item_data = [item_header]
    total_qty = 0
    for idx, item in enumerate(invoice.invoiceitem_set.all(), 1):
        total_qty += item.quantity
        item_data.append([
            str(idx),
            Paragraph(f"<b>{item.item.name}</b><br/>{item.item.description or ''}", style_normal),
            item.item.hsn_sac,
            f"{item.quantity} Nos",
            ''
        ])
    
    item_data.append(['', 'Total', '', f"{total_qty} Nos", ''])

    t_items = Table(item_data, colWidths=[15*mm, 85*mm, 30*mm, 30*mm, 30*mm])
    t_items.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (1,0), (1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (-1,-1), (-1,-1), 'Helvetica-Bold'),
        ('SPAN', (1,-1), (2,-1)), 
    ]))
    elements.append(t_items)
    elements.append(Spacer(1, 15*mm))
    
    # Footer
    footer_text = f"""
    <br/><br/>
    Received the above goods in good condition.<br/><br/><br/>
    <b>Receiver's Signature</b>
    """
    
    # Signature Image logic for DC
    signature_img = None
    img_path = None
    
    # 1. Try Company Signature (Database)
    if hasattr(company, 'signature') and company.signature:
        try:
            img_path = company.signature.path
        except Exception:
            img_path = None
            
    # 2. Universal Fallback (Local File)
    if not img_path or not os.path.exists(img_path):
        fallback_path = os.path.join(os.path.dirname(__file__), 'signature.png')
        if os.path.exists(fallback_path):
            img_path = fallback_path

    if img_path:
        try:
            if os.path.exists(img_path):
                # Center the image in the signature block
                signature_img = Image(img_path, width=40*mm, height=15*mm) 
                signature_img.hAlign = 'CENTER'
        except Exception:
            pass

    auth_sign_header_text = f"<b>for {company.name}</b>"
    auth_sign_footer_text = "Authorised Signatory" # Removed <br/> to control spacing via Table

    # Create a nested table for the signature block to ensure centering
    sign_data = []
    sign_data.append([Paragraph(auth_sign_header_text, ParagraphStyle('SignHead', parent=styles['Normal'], alignment=1))]) # Center
    
    if signature_img:
        sign_data.append([signature_img])
    else:
        sign_data.append([Spacer(1, 15*mm)])
        
    sign_data.append([Paragraph(auth_sign_footer_text, ParagraphStyle('SignFoot', parent=styles['Normal'], alignment=1))]) # Center
    
    t_sign = Table(sign_data, colWidths=[90*mm])
    t_sign.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))

    left_cell = [Paragraph(footer_text, style_normal)]
    
    # Main Footer Table
    foot_data = [[left_cell, t_sign]]
    
    t_foot = Table(foot_data, colWidths=[95*mm, 95*mm])
    t_foot.setStyle(TableStyle([
        ('ALIGN', (0,0), (0,0), 'LEFT'),  # Left cell left aligned
        ('ALIGN', (1,0), (1,0), 'RIGHT'), # Right cell content right aligned (the table itself)
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    elements.append(t_foot)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_transport_pdf(invoice, transport, company_input):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    elements = []
    styles = getSampleStyleSheet()
    
    t_header, company = create_header_table("TRANSPORT CHARGES", company_input)
    # elements.append(t_header)
    
    style_normal = styles['Normal']
    style_normal.fontSize = 9
    style_bold = ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9)
    style_header = ParagraphStyle('Header', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, alignment=1)

    elements.append(Paragraph("TRANSPORT BILL", style_header))
    elements.append(Spacer(1, 5*mm))
    
    # Left: Seller details + Bill To
    seller_text = f"""
    <b>{company.name}</b><br/>
    {company.address}<br/>
    GSTIN/UIN: {company.gstin}<br/>
    Contact: {company.contact_number}
    """
    
    bill_to_text = f"""
    <br/><br/><b>Billed To</b><br/>
    <b>{invoice.location.name}</b><br/>
    {invoice.location.address}<br/>
    GSTIN/UIN: {clean(invoice.location.gstin)}<br/>
    """
    
    left_content = [Paragraph(seller_text + bill_to_text, style_normal)]
    
    # Right: Bill Details
    trp_data = [
        [Paragraph("<b>Bill No.</b>", style_normal), Paragraph(f"<b>TRP-{transport.id}</b>", style_bold)],
        [Paragraph("<b>Date</b>", style_normal), Paragraph(f"<b>{clean_date(transport.date)}</b>", style_bold)],
        [Paragraph("<b>Ref Invoice No.</b>", style_normal), Paragraph(clean(invoice.tally_invoice_number or invoice.app_invoice_number), style_normal)],
        [Paragraph("<b>Vehicle/Ref</b>", style_normal), Paragraph(clean(transport.description), style_normal)],
    ]
    
    right_table = Table(trp_data, colWidths=[35*mm, 55*mm])
    right_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    
    main_header_data = [[left_content, right_table]]
    main_table = Table(main_header_data, colWidths=[95*mm, 95*mm])
    main_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    elements.append(main_table)
    
    # Charges Table
    t_data = [['Description', 'HSN', 'Amount']]
    t_data.append([
        Paragraph(f"Transport Charges<br/>{transport.description or ''}", style_normal), 
        '996719', 
        f"{transport.charges}"
    ])
    
    t_data.append(['Total', '', f"{INR_SYMBOL} {transport.charges}"])
    
    t = Table(t_data, colWidths=[120*mm, 30*mm, 30*mm])
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('ALIGN', (-1,1), (-1,-1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (-1,-1), (-1,-1), 'Helvetica-Bold'),
        ('SPAN', (0,-1), (1,-1)),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 8),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 10*mm))
    
    elements.append(create_footer_with_signature(company, "Transport Charges"))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer
