"""
Requirements:
    pip install python-docx reportlab pywin32

Run:
    python student_invoice_generator.py
"""
import os
from decimal import Decimal
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph, SimpleDocTemplate, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Staff Member (CHANGE TO YOUR NAME)
STAFF_MEMBER = "Benjamin Li"

# Outlook automation
OUTLOOK_AVAILABLE = True
try:
    import win32com.client
except Exception:
    OUTLOOK_AVAILABLE = False

# File paths
LETTERHEAD_LOGO = "Logo.png"
FOOTER_IMAGE = "Footer.png"
SIGNATURE_IMAGE = "Signature.JPG"
CONTACT_INFO = (
    "23 Amsterdam Crescent,<br/>",
    "Salisbury Downs, SA<br/>",
    "PO Box 535, Salisbury, SA 5108<br/>",
    "E tmc@tmc.catholic.edu.au<br/>",
    "T (08) 8182 2600<br/>",
    "www.tmc.catholic.edu.au"
)

class InvoiceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Student Invoice Generator")
        self.geometry("700x560")
        self.resizable(False, False)
        self.items = []
        self.create_widgets()

    def create_widgets(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        topfrm = ttk.Frame(frm)
        topfrm.pack(fill=tk.X, pady=(0,8))

        ttk.Label(topfrm, text="Student Name:").grid(row=0, column=0, sticky=tk.W)
        self.student_entry = ttk.Entry(topfrm, width=30)
        self.student_entry.grid(row=0, column=1, padx=(6,20))

        ttk.Label(topfrm, text="Parent Name:").grid(row=0, column=2, sticky=tk.W)
        self.parent_entry = ttk.Entry(topfrm, width=30)
        self.parent_entry.grid(row=0, column=3, padx=(6,0))

        status_frame = ttk.Frame(frm)
        status_frame.pack(fill=tk.X, pady=(8,8))

        ttk.Label(status_frame, text="Device Status:").grid(row=0, column=0, sticky=tk.W)
        self.device_status = tk.StringVar(value="Missing")
        ttk.Radiobutton(status_frame, text="Missing", variable=self.device_status, value="Missing").grid(row=0, column=1, padx=6)
        ttk.Radiobutton(status_frame, text="Damaged", variable=self.device_status, value="Damaged").grid(row=0, column=2, padx=6)

        itemsfrm = ttk.LabelFrame(frm, text="Items")
        itemsfrm.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(itemsfrm)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=6)

        ttk.Label(left, text="Item name:").pack(anchor=tk.W)
        self.item_entry = ttk.Entry(left, width=30)
        self.item_entry.pack(anchor=tk.W)

        ttk.Label(left, text="Cost (e.g. 25.50):").pack(anchor=tk.W, pady=(8,0))
        self.cost_entry = ttk.Entry(left, width=15)
        self.cost_entry.pack(anchor=tk.W)

        ttk.Button(left, text="Add item", command=self.add_item).pack(anchor=tk.W, pady=(8,0))
        ttk.Button(left, text="Remove selected", command=self.remove_selected).pack(anchor=tk.W, pady=(6,0))

        right = ttk.Frame(itemsfrm)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=6, pady=6)

        self.tree = ttk.Treeview(right, columns=("item", "cost"), show="headings", selectmode='browse')
        self.tree.heading('item', text='Item')
        self.tree.heading('cost', text='Cost')
        self.tree.column('item', width=340)
        self.tree.column('cost', width=160, anchor=tk.E)
        self.tree.pack(fill=tk.BOTH, expand=True)

        totalfrm = ttk.Frame(frm)
        totalfrm.pack(fill=tk.X, pady=(8,0))
        ttk.Label(totalfrm, text="Total:", font=(None, 10, 'bold')).pack(side=tk.LEFT)
        self.total_var = tk.StringVar(value="0.00")
        ttk.Label(totalfrm, textvariable=self.total_var, font=(None, 10, 'bold')).pack(side=tk.LEFT, padx=6)

        btnfrm = ttk.Frame(frm)
        btnfrm.pack(fill=tk.X, pady=(12,0))
        ttk.Button(btnfrm, text="Generate PDF & Open Outlook Draft", command=self.generate_and_email).pack(side=tk.LEFT)
        ttk.Button(btnfrm, text="Save PDF to...", command=self.save_pdf_to).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(btnfrm, text="Quit", command=self.quit).pack(side=tk.RIGHT)

    # ------------------------ Item management ------------------------
    def add_item(self):
        name = self.item_entry.get().strip()
        cost_text = self.cost_entry.get().strip()
        if not name:
            messagebox.showwarning("Missing item", "Please enter an item name.")
            return
        try:
            cost = Decimal(cost_text)
        except Exception:
            messagebox.showwarning("Invalid cost", "Please enter a valid numeric cost (e.g. 25.50).")
            return
        cost = cost.quantize(Decimal('0.01'))
        self.items.append((name, cost))
        self.tree.insert('', tk.END, values=(name, f"{cost:.2f}"))
        self.item_entry.delete(0, tk.END)
        self.cost_entry.delete(0, tk.END)
        self.update_total()

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        self.tree.delete(sel[0])
        del self.items[idx]
        self.update_total()

    def update_total(self):
        total = sum(cost for (_n, cost) in self.items)
        self.total_var.set(f"{total:.2f}")

    # ------------------------ PDF Header ------------------------
    def draw_header(self, canvas, doc):
        page_width, page_height = A4
        max_logo_width_mm = 60
        max_logo_height_mm = 30
        left_padding_mm = 10
        top_padding_mm = 10

        if os.path.exists(LETTERHEAD_LOGO):
            canvas.drawImage(
                LETTERHEAD_LOGO,
                left_padding_mm*mm,
                page_height - max_logo_height_mm*mm - top_padding_mm*mm,
                width=max_logo_width_mm*mm,
                height=max_logo_height_mm*mm,
                preserveAspectRatio=True,
                mask='auto'
            )

        from reportlab.lib.enums import TA_RIGHT
        styleR = ParagraphStyle('Right', parent=getSampleStyleSheet()['Normal'], alignment=TA_RIGHT, fontSize=9)
        contact_text = ''.join(CONTACT_INFO)
        p = Paragraph(contact_text, styleR)
        w, h = p.wrap(page_width - 100*mm, 100)
        p.drawOn(canvas, page_width - left_padding_mm*mm - w, page_height - h - top_padding_mm*mm)

    # ------------------------ PDF Footer ------------------------
    def draw_footer(self, canvas, doc):
        page_width, page_height = A4
        if os.path.exists(FOOTER_IMAGE):
            max_footer_width_mm = 160
            max_footer_height_mm = 30
            x = (page_width - max_footer_width_mm*mm) / 2
            y = -5*mm
            canvas.drawImage(
                FOOTER_IMAGE,
                x,
                y,
                width=max_footer_width_mm*mm,
                height=max_footer_height_mm*mm,
                preserveAspectRatio=True,
                mask='auto'
            )

    # ------------------------ Build PDF ------------------------
    def build_pdf(self, pdf_path):
        normal = ParagraphStyle('Normal', parent=getSampleStyleSheet()['Normal'], fontName='Helvetica', fontSize=11, leading=16, spaceAfter=10)
        bold = ParagraphStyle('Bold', parent=getSampleStyleSheet()['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=16, spaceAfter=10)
        date_style = ParagraphStyle('Date', parent=getSampleStyleSheet()['Normal'], fontName='Helvetica', fontSize=11, leading=16, spaceAfter=10)

        from datetime import datetime
        today = datetime.today().strftime("%d %B %Y").lstrip("0")

        doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                                leftMargin=20*mm, rightMargin=20*mm,
                                topMargin=60*mm, bottomMargin=30*mm)

        elements = []

        parent = self.parent_entry.get().strip()
        student = self.student_entry.get().strip()
        status = self.device_status.get().lower()

        # Date & Subject
        elements.append(Paragraph(today, date_style))
        elements.append(Paragraph(f"RE: Thomas More College {status.capitalize()} Device", bold))
        elements.append(Spacer(1, 12))

        # Greeting
        elements.append(Paragraph(f"<br/><br/>Hi {parent},", normal))

        # Body text
        text1 = (f"I am writing to inform you that {student} has recently visited the ICT office "
                 f"with a {status} device. As per the TMC User Device Charter, any cost "
                 "relating to the repair or replacement of devices or accessories is passed on to the family.")
        elements.append(Paragraph(text1, normal))

        text2 = ("The cost for the repair and replacement is as below, and you will receive an invoice "
                 "from our Finance department for this amount.")
        elements.append(Paragraph(text2, normal))
        elements.append(Spacer(1, 12))

        # Items table
        if self.items:
            data = [["Item", "Cost ($)"]] + [[n, f"{c:.2f}"] for n, c in self.items]
            total = sum(c for (_n, c) in self.items)
            data.append(["Total", f"{total:.2f}"])
            tbl = Table(data, colWidths=[120*mm, 30*mm])
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('ALIGN', (-1,0), (-1,-1), 'RIGHT'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTNAME', (-2,-1), (-1,-1), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 11),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            elements.append(tbl)
            elements.append(Spacer(1, 12))

        # Closing paragraph
        elements.append(Paragraph("<br/>Feel free to contact us by replying to this email if you have any questions "
                                  "or if you would like to discuss this matter further.", normal))
        elements.append(Spacer(1, 12))

        # Signature
        elements.append(Paragraph("Yours Sincerely,", normal))
        if os.path.exists(SIGNATURE_IMAGE):
            sig_img = Image(SIGNATURE_IMAGE)
            sig_img.drawHeight = 25*mm
            sig_img.drawWidth = sig_img.drawHeight * sig_img.imageWidth / sig_img.imageHeight
            sig_img.hAlign = 'LEFT'  # <--- left align the image
            elements.append(sig_img)
            elements.append(Spacer(1, 6))
        elements.append(Paragraph("Angelo Anastasiadis", normal))
        elements.append(Paragraph("ICT Manager", bold))

        doc.build(elements,
                  onFirstPage=lambda c, d: (self.draw_header(c,d), self.draw_footer(c,d)),
                  onLaterPages=lambda c, d: self.draw_footer(c,d))

    # ------------------------ Generate / Email ------------------------
    def generate_and_email(self):
        if not self.items:
            messagebox.showwarning("No items", "Please add at least one item before generating.")
            return
        student = self.student_entry.get().strip()
        parent = self.parent_entry.get().strip()
        status = self.device_status.get()
        if not student or not parent:
            messagebox.showwarning("Missing details", "Please enter both Student Name and Parent Name.")
            return
        safe_student = student.replace(" ", "_")
        default_name = f"TMC_{status}_Device_{safe_student}.pdf"
        save_path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF files', '*.pdf')],
                                                 initialfile=default_name, title="Save PDF as...")
        if not save_path:
            messagebox.showinfo("Cancelled", "PDF generation cancelled.")
            return
        try:
            self.build_pdf(save_path)
        except Exception as e:
            messagebox.showerror("PDF error", f"Failed to build PDF: {e}")
            return

        if not OUTLOOK_AVAILABLE:
            messagebox.showerror("Outlook not available", "pywin32/win32com is not installed or Outlook automation is unavailable.")
            return
        try:
            app = win32com.client.Dispatch('Outlook.Application')
            mail = app.CreateItem(0)
            mail.Subject = f"Thomas More College {status} Device"
            mail.Body = (f"Hi {parent},\n\n"
                         f"I am writing in regard to a {status.lower()} device that {student} has brought into the office. "
                         "As per the TMC Device Charter Policy, any costs are passed onto the supporting family.\n\n"
                         "Please refer to the attached letter for full details.\n\n"
                         "Feel free to contact us by replying to this email if you have any questions or if you would like to discuss this matter further.\n\n"
                         f"Kind regards,\n{STAFF_MEMBER}")
            mail.Attachments.Add(save_path)
            mail.Display(True)
            messagebox.showinfo("Done", "PDF saved and Outlook draft opened.")
        except Exception as e:
            messagebox.showerror("Outlook error", f"Failed to create Outlook draft: {e}")

    # ------------------------ Save PDF ------------------------
    def save_pdf_to(self):
        if not self.items:
            messagebox.showwarning("No items", "Please add at least one item before saving.")
            return
        student = self.student_entry.get().strip()
        status = self.device_status.get()
        if not student:
            messagebox.showwarning("Missing details", "Please enter a Student Name before saving.")
            return
        safe_student = student.replace(" ", "_")
        default_name = f"TMC_{status}_Device_{safe_student}.pdf"
        fn = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF files', '*.pdf')], initialfile=default_name)
        if not fn:
            return
        try:
            self.build_pdf(fn)
            messagebox.showinfo("Saved", f"PDF saved to: {fn}")
        except Exception as e:
            messagebox.showerror("Save error", f"Failed to save PDF: {e}")

if __name__ == '__main__':
    app = InvoiceApp()
    app.mainloop()
