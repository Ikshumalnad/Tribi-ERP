import io
import os
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

def number_to_words_indian(num):
    units = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN",
             "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
    tens = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]

    def conv_less_1000(n):
        res = ""
        if n >= 100:
            res += units[n // 100] + " HUNDRED "
            n %= 100
        if n >= 20:
            res += tens[n // 10] + ("-" + units[n % 10] if n % 10 != 0 else "")
        elif n > 0:
            res += units[n]
        return res.strip()

    if num == 0:
        return "ZERO"

    num_int = int(num)
    parts = []
    
    crore = num_int // 10000000
    num_int %= 10000000
    if crore > 0:
        parts.append(conv_less_1000(crore) + " CRORE")
        
    lakh = num_int // 100000
    num_int %= 100000
    if lakh > 0:
        parts.append(conv_less_1000(lakh) + " LAKH")
        
    thousand = num_int // 1000
    num_int %= 1000
    if thousand > 0:
        parts.append(conv_less_1000(thousand) + " THOUSAND")
        
    if num_int > 0:
        parts.append(conv_less_1000(num_int))
        
    return " ".join(parts).strip()


def amount_in_words(amount, currency_code='INR'):
    try:
        amt = float(amount)
    except (ValueError, TypeError):
        return "ZERO"
    
    rupees = int(amt)
    paise = int(round((amt - rupees) * 100))
    
    prefix = "Rs. "
    sub_unit = "PAISE"
    if currency_code == 'USD':
        prefix = "USD "
        sub_unit = "CENTS"
    elif currency_code == 'EUR':
        prefix = "EUR "
        sub_unit = "CENTS"
    elif currency_code == 'GBP':
        prefix = "GBP "
        sub_unit = "PENCE"
    elif currency_code == 'INR':
        prefix = "Rs. "
        sub_unit = "PAISE"
    
    words = prefix + number_to_words_indian(rupees)
    if paise > 0:
        words += f" AND {sub_unit} " + number_to_words_indian(paise)
    words += " ONLY."
    return words


class NumberedCanvas(canvas.Canvas):
    po_formatted_number = ""

    def __init__(self, *args, **kwargs):
        super(NumberedCanvas, self).__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setStrokeColor(colors.black)
        self.setLineWidth(0.5)
        self.line(30, 30, 565, 30)
        po_str = getattr(NumberedCanvas, 'po_formatted_number', '')
        self.drawString(30, 18, f"PO No. {po_str}")
        page_text = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(565, 18, page_text)
        self.restoreState()


def build_po_pdf(po, po_items, vendor, custom_terms=None, payment_terms=None, hide_signature=False, policy_line=None):
    if policy_line is None:
        try:
            from app import get_setting
            policy_line = get_setting('policy_line_text', 'COVERED UNDER NEW INDIA ASSURANCE POLICY NO. 67020021200200000030')
        except Exception:
            policy_line = 'COVERED UNDER NEW INDIA ASSURANCE POLICY NO. 67020021200200000030'

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=30,
        rightMargin=30,
        topMargin=25,
        bottomMargin=45
    )

    NumberedCanvas.po_formatted_number = po.get('formatted_po_number', '')

    styles = getSampleStyleSheet()

    # Custom styles matching image
    title_style = ParagraphStyle(
        'POTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=14,
        alignment=1, # Center
        spaceAfter=6
    )

    company_title = ParagraphStyle(
        'CompanyTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor('#1a2e4a')
    )

    company_address = ParagraphStyle(
        'CompanyAddress',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#333333')
    )

    logo_style = ParagraphStyle(
        'LogoText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=20,
        alignment=1,
        textColor=colors.HexColor('#000000')
    )

    logo_sub = ParagraphStyle(
        'LogoSub',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6.5,
        alignment=1,
        textColor=colors.HexColor('#888888')
    )

    right_header_style = ParagraphStyle(
        'RightHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=11,
        alignment=2 # Right
    )

    po_banner_style = ParagraphStyle(
        'POBanner',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        alignment=1,
        textColor=colors.HexColor('#000000')
    )

    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9,
        alignment=1,
        textColor=colors.HexColor('#000000')
    )

    cell_style = ParagraphStyle(
        'CellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#000000')
    )

    cell_bold = ParagraphStyle(
        'CellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#000000')
    )

    cell_center = ParagraphStyle(
        'CellCenter',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        alignment=1,
        textColor=colors.HexColor('#000000')
    )

    cell_right = ParagraphStyle(
        'CellRight',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        alignment=2,
        textColor=colors.HexColor('#000000')
    )

    section_header = ParagraphStyle(
        'SectionHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        alignment=1,
        textColor=colors.HexColor('#000000')
    )

    elements = []

    # 1. Main Title
    elements.append(Paragraph("PURCHASE ORDER", title_style))

    # 2. Top Header Box (Company Info, Logo, GSTIN)
    col1_content = [
        Paragraph("TRIBI Systems Pvt. Ltd.", company_title),
        Paragraph("Regd. Off : No. 55/B", company_address),
        Paragraph("1st Main Road", company_address),
        Paragraph("Electronics City - Phase 1", company_address),
        Paragraph("Bangalore - 560100", company_address),
        Paragraph("KARNATAKA", company_address),
    ]

    # Load the logo image dynamically and securely using absolute path
    logo_path = os.path.join(os.path.dirname(__file__), 'static', 'logo.png')
    if os.path.exists(logo_path):
        try:
            logo_img = Image(logo_path, width=110, height=55)
            logo_content = [
                Spacer(1, 2),
                logo_img,
                Spacer(1, 2)
            ]
        except Exception as e:
            logo_content = [
                Spacer(1, 4),
                Paragraph("<b>T R I B I</b>", logo_style),
                Paragraph('<font color="#84cc16"><b>━━━</b></font>', logo_sub),
            ]
    else:
        logo_content = [
            Spacer(1, 4),
            Paragraph("<b>T R I B I</b>", logo_style),
            Paragraph('<font color="#84cc16"><b>━━━</b></font>', logo_sub),
        ]

    col3_content = [
        Paragraph("GSTIN : 29AAAC04139M1Z3", right_header_style),
        Spacer(1, 6),
        Paragraph("STATE : Karnataka", right_header_style),
        Spacer(1, 4),
        Paragraph("STATE CODE : 29", right_header_style),
    ]

    top_table = Table(
        [[col1_content, logo_content, col3_content]],
        colWidths=[200, 135, 200]
    )
    top_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(top_table)

    # 3. PO Number and Date Banner
    date_str = ""
    if po.get('date_raised'):
        date_str = po['date_raised'].strftime('%d-%b-%Y') if hasattr(po['date_raised'], 'strftime') else str(po['date_raised'])
    else:
        date_str = date.today().strftime('%d-%b-%Y')

    formatted_po = po.get('formatted_po_number', '')
    banner_text = f"Purchase Order No. : {formatted_po}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Date : {date_str}"
    banner_table = Table(
        [[Paragraph(banner_text, po_banner_style)]],
        colWidths=[535]
    )
    banner_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('PADDING', (0,0), (-1,-1), 4),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#fcfcfc')),
    ]))
    elements.append(banner_table)

    # 4. Vendor & Prepared By Section
    vendor_lines = [
        Paragraph("<b>To</b>", cell_bold),
        Paragraph(f"<b>{(vendor.get('full_name') or vendor.get('short_name') or '').upper()}</b>", cell_bold),
    ]
    if vendor.get('address_line_1'):
        vendor_lines.append(Paragraph(vendor['address_line_1'], cell_style))
    city_pin = f"{vendor.get('city') or ''} {vendor.get('pincode') or ''}".strip()
    if city_pin:
        vendor_lines.append(Paragraph(city_pin, cell_style))
    if vendor.get('country'):
        vendor_lines.append(Paragraph(vendor['country'], cell_style))
    if vendor.get('gst_no'):
        vendor_lines.append(Spacer(1, 4))
        vendor_lines.append(Paragraph(f"<b>GSTIN : {vendor['gst_no']}</b>", cell_style))
    vendor_lines.append(Paragraph("Attn of : ", cell_style))

    meta_lines = [
        Paragraph(f"Prepared by : {po.get('raised_by') or 'Yathisha M N'}", cell_style),
        Spacer(1, 4),
        Paragraph(f"Reviewed by : {po.get('reviewed_by') or 'Yathisha M N'}", cell_style),
        Spacer(1, 10),
        Paragraph(f"Payment Terms : {payment_terms or '0 days'}", cell_style),
    ]

    party_table = Table(
        [[vendor_lines, meta_lines]],
        colWidths=[267.5, 267.5]
    )
    party_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.75, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(party_table)

    # 5. Line Items Table Header
    currency_code = po.get('currency', 'INR') or 'INR'
    currency_symbols = {
        'INR': 'Rs.',
        'USD': '$',
        'EUR': '€',
        'GBP': '£'
    }
    currency_symbol = currency_symbols.get(currency_code, 'Rs.')

    headers = [
        Paragraph("No.", table_header_style),
        Paragraph("Item Code", table_header_style),
        Paragraph("Description", table_header_style),
        Paragraph("Make", table_header_style),
        Paragraph("MPN", table_header_style),
        Paragraph("Qty", table_header_style),
        Paragraph("Units", table_header_style),
        Paragraph(f"Unit Price<br/>({currency_symbol})", table_header_style),
        Paragraph(f"Total Price<br/>({currency_symbol})", table_header_style),
        Paragraph(f"Tax<br/>({currency_symbol})", table_header_style),
        Paragraph(f"Total<br/>({currency_symbol})", table_header_style),
    ]

    col_widths = [20, 75, 95, 40, 55, 25, 27, 48, 50, 50, 50] # Sum = 535

    table_data = [headers]

    # How this vendor's tax is charged. It is a property of the VENDOR, not of
    # the goods: the HSN code says what the rate is, this says how it is split.
    #
    #   1  registered in Karnataka      -> CGST + SGST, half the rate each
    #   2  registered elsewhere in India-> IGST, the whole rate
    #   0  outside India                -> no Indian GST on the invoice at all
    #
    # Default 1 when the column is absent, which is what the form assumed
    # before the column existed.
    try:
        tax_mode = int(vendor.get('tax_mode'))
    except (TypeError, ValueError):
        tax_mode = 1
    if tax_mode not in (0, 1, 2):
        tax_mode = 1

    subtotal = 0.0
    total_cgst = 0.0
    total_sgst = 0.0
    total_igst = 0.0
    total_tax_amount = 0.0

    for idx, item in enumerate(po_items, 1):
        qty = float(item.get('qty_ordered', 0))
        unit_price = float(item.get('unit_price', 0))
        line_price = round(qty * unit_price, 4)
        subtotal += line_price

        # The HSN row carries cgst, sgst AND igst all filled in, because the
        # rate is the same fact expressed two ways. Only one way applies to
        # any given order, and the vendor decides which.
        #
        # `or 18.0` was wrong here: a genuine 0% HSN is falsy, so ten of the
        # codes in the table were printing at 18%. None means "no rate known",
        # zero means zero.
        raw_rate = item.get('tax_rate')
        tax_rate = 18.0 if raw_rate is None else float(raw_rate)

        if tax_mode == 0:
            cgst_rate = sgst_rate = igst_rate = 0.0
            tax_rate = 0.0
        elif tax_mode == 2:
            cgst_rate = sgst_rate = 0.0
            igst_rate = tax_rate
        else:
            cgst_rate = sgst_rate = round(tax_rate / 2.0, 2)
            igst_rate = 0.0

        line_cgst = round(line_price * (cgst_rate / 100.0), 4)
        line_sgst = round(line_price * (sgst_rate / 100.0), 4)
        line_igst = round(line_price * (igst_rate / 100.0), 4)
        # Charged = what is actually added, which is the sum of the lines
        # printed. Before, this was the full rate regardless, so the totals
        # box never added up to the Total row.
        line_tax = round(line_cgst + line_sgst + line_igst, 4)

        total_cgst += line_cgst
        total_sgst += line_sgst
        total_igst += line_igst
        total_tax_amount += line_tax

        line_total = line_price + line_tax

        # An unclassified item used to print as HSN 853400 — printed circuit
        # boards — whatever it actually was. Saying nothing is better than
        # saying something untrue on a tax document.
        hsn_code = item.get('hsn_code')
        item_code_text = (f"{item.get('item_code', '')}<br/>[HSN {hsn_code}]"
                          if hsn_code else f"{item.get('item_code', '')}<br/>[HSN not set]")

        if tax_mode == 0:
            tax_text = "0.0000<br/>[not taxed]"
        else:
            tax_text = f"{line_tax:.4f}<br/>[{tax_rate:.2f} %]"

        row = [
            Paragraph(str(idx), cell_center),
            Paragraph(item_code_text, cell_style),
            Paragraph(item.get('item_desc') or '', cell_style),
            Paragraph(item.get('manufacturer_name') or '', cell_style),
            Paragraph(item.get('mpn') or f"Tribi - {item.get('item_code', '')}", cell_style),
            Paragraph(f"{qty:g}", cell_center),
            Paragraph(item.get('unit') or 'nos.', cell_center),
            Paragraph(f"{unit_price:.4f}", cell_right),
            Paragraph(f"{line_price:.4f}", cell_right),
            Paragraph(tax_text, cell_center),
            Paragraph(f"{line_total:.4f}", cell_right),
        ]
        table_data.append(row)

    grand_total = subtotal + total_tax_amount

    items_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    items_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#999999')),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e5e7eb')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(items_table)

    # 6. Instructions & Summary Totals Box
    instructions_text = po.get('remarks') or "PO for supplied items as per requirement."
    left_instr_content = [
        Paragraph("<u>Instructions</u>", section_header),
        Spacer(1, 6),
        Paragraph(instructions_text, cell_style)
    ]

    # Only the lines that apply. All three used to print at once — an in-state
    # split and an inter-state charge side by side — and they summed to twice
    # the tax actually added, so the box never agreed with the Total beneath it.
    totals_table_data = [
        [Paragraph("Sub Total", cell_right), Paragraph(f"{subtotal:.4f}", cell_right)],
    ]
    if tax_mode == 1:
        totals_table_data.append([Paragraph("CGST", cell_right),
                                  Paragraph(f"{total_cgst:.4f}", cell_right)])
        totals_table_data.append([Paragraph("SGST", cell_right),
                                  Paragraph(f"{total_sgst:.4f}", cell_right)])
    elif tax_mode == 2:
        totals_table_data.append([Paragraph("IGST", cell_right),
                                  Paragraph(f"{total_igst:.4f}", cell_right)])
    else:
        totals_table_data.append([Paragraph("GST", cell_right),
                                  Paragraph("Not applicable", cell_right)])

    total_row = len(totals_table_data)
    totals_table_data.append(
        [Paragraph(f"<b>Total ({currency_symbol})</b>", cell_right),
         Paragraph(f"<b>{grand_total:.2f}</b>",
                   ParagraphStyle('GrandTot', parent=cell_right,
                                  fontName='Helvetica-Bold', fontSize=9))])

    totals_sub_table = Table(totals_table_data, colWidths=[100, 100])
    totals_sub_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        # The rule sat on row 4 whatever was above it. The box is a different
        # height per mode now, so it follows the Total row.
        ('LINEABOVE', (0,total_row), (1,total_row), 0.75, colors.black),
    ]))

    summary_block = Table(
        [[left_instr_content, totals_sub_table]],
        colWidths=[330, 205]
    )
    summary_block.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.75, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))

    elements.append(summary_block)
    elements.append(Spacer(1, 10))

    # 7. Value in Words Banner & Terms & Conditions
    words_banner = Table(
        [[Paragraph(f"<b>Value in words : {amount_in_words(grand_total, currency_code)}</b>", cell_style)]],
        colWidths=[535]
    )
    words_banner.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('PADDING', (0,0), (-1,-1), 6),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#ffffff')),
    ]))

    terms_lines = [
        Paragraph("<u>Terms & Conditions</u>", section_header),
        Spacer(1, 6),
    ]
    if tax_mode == 0:
        terms_lines.append(Paragraph(
            "<b>No Indian GST is charged on this order — the supplier is "
            "outside India.</b>", cell_style))
        terms_lines.append(Spacer(1, 3))
    if custom_terms:
        # Split terms by newline and filter out empty lines
        terms_list = [line.strip() for line in custom_terms.split('\n') if line.strip()]
    else:
        terms_list = [
            "Subject to Bangalore Jurisdiction",
            "Purchase Order No. should be mentioned in all the supporting documents along with the supply of material",
            "Tribi Item Code should be mentioned in all the supporting documents along with the supply of material"
        ]
    for term in terms_list:
        terms_lines.append(Paragraph(term, cell_style))
        terms_lines.append(Spacer(1, 3))

    signatory_name = "" if hide_signature else (po.get('approved_by') or 'Prashanth Alva')
    signatory_date = "" if hide_signature else date_str

    signatory_lines = [
        Paragraph("<b>For TRIBI Systems Pvt. Ltd.</b>", ParagraphStyle('SigHead', parent=cell_center, fontName='Helvetica-Bold')),
        Spacer(1, 18),
    ]
    if signatory_name:
        signatory_lines.extend([
            Paragraph(f"<b>{signatory_name}</b>", cell_center),
            Spacer(1, 2),
            Paragraph("<b>Authorised Signatory</b>", cell_center),
            Spacer(1, 4),
            Paragraph(signatory_date, cell_center),
        ])
    else:
        signatory_lines.extend([
            Spacer(1, 15),
            Paragraph("<b>Authorised Signatory</b>", cell_center),
        ])

    terms_block = Table(
        [[terms_lines, signatory_lines]],
        colWidths=[330, 205]
    )
    terms_block.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.75, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))

    footer_items = [
        words_banner,
        Spacer(1, 6),
        terms_block
    ]

    if policy_line and str(policy_line).strip():
        policy_style = ParagraphStyle(
            'PolicyLineStyle',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=7.5,
            leading=9.5,
            alignment=1, # Center
            textColor=colors.HexColor('#000000')
        )
        policy_banner = Table(
            [[Paragraph(str(policy_line).strip(), policy_style)]],
            colWidths=[535]
        )
        policy_banner.setStyle(TableStyle([
            #('BOX', (0,0), (-1,-1), 0.75, colors.black),
            ('PADDING', (0,0), (-1,-1), 4),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#ffffff')),
        ]))
        footer_items.extend([
            Spacer(1, 6),
            policy_banner
        ])

    footer_section = KeepTogether(footer_items)
    elements.append(footer_section)

    doc.build(elements, canvasmaker=NumberedCanvas)

    buffer.seek(0)
    return buffer
