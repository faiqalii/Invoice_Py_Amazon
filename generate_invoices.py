import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
from datetime import datetime, timedelta
from arabic_reshaper import reshape
from bidi.algorithm import get_display

# Function to wrap text for table cells or addresses
def wrap_text(text, width, font_name, font_size, c):
    lines = []
    words = str(text).split()
    current_line = []
    current_width = 0

    for word in words:
        word_width = c.stringWidth(word + " ", font_name, font_size)
        if current_width + word_width <= width:
            current_line.append(word)
            current_width += word_width
        else:
            lines.append(" ".join(current_line))
            current_line = [word]
            current_width = word_width
    if current_line:
        lines.append(" ".join(current_line))
    return lines

# Function to prepare Arabic text for display
def prepare_arabic_text(text):
    if isinstance(text, str) and any(char >= '\u0600' and char <= '\u06FF' for char in text):
        reshaped_text = reshape(text)  # Reshape Arabic for proper rendering
        return get_display(reshaped_text)  # Convert to RTL display order
    return text

# Register an Arabic-supporting font with the correct path
pdfmetrics.registerFont(TTFont('ArabicFont', r"C:\Users\Faiqa\OneDrive\Desktop\Amazon\Invoice GEN bot\Invoice GEN bot\DejaVuSans.ttf"))  # Update with your font path

# Step 1: Load the CSV file with the full path
csv_file = r"C:\Users\Faiqa\OneDrive\Desktop\Amazon\Invoice GEN bot\Invoice GEN bot\orders.csv"  # Full path to the CSV file
output_folder = 'invoices'  # Folder to save generated invoices

# Check if the CSV file exists
if not os.path.exists(csv_file):
    print(f"Error: The file '{csv_file}' was not found. Please ensure the file exists and the path is correct.")
    exit(1)

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Read the CSV file
try:
    orders_df = pd.read_csv(csv_file)
except Exception as e:
    print(f"Error reading the CSV file: {e}")
    exit(1)

# Step 2: Generate invoices for each order
for index, row in orders_df.iterrows():
    # Extract data from the row with default values for missing columns
    order_id = row.get('Amazon Order Id', '')
    purchase_date = row.get('Purchase Date', '')
    buyer_name = prepare_arabic_text(row.get('Buyer Name', ''))
    product_title = prepare_arabic_text(row.get('Title', ''))
    quantity = row.get('Shipped Quantity', 0)
    item_price = float(row.get('Item Price', 0.0))  # Ensure numeric
    shipping_price = float(row.get('Shipping Price', 0.0))  # Ensure numeric
    shipping_address1 = prepare_arabic_text(row.get('Shipping Address 1', ''))
    shipping_address2 = prepare_arabic_text(row.get('Shipping Address 2', ''))
    shipping_city = prepare_arabic_text(row.get('Shipping City', ''))
    shipping_state = prepare_arabic_text(row.get('Shipping State', ''))
    shipping_country = prepare_arabic_text(row.get('Shipping Country Code', ''))
    billing_address1 = prepare_arabic_text(row.get('Billing Address 1', ''))
    billing_address2 = prepare_arabic_text(row.get('Billing Address 2', ''))
    billing_city = prepare_arabic_text(row.get('Billing City', ''))
    billing_state = prepare_arabic_text(row.get('Billing State', ''))
    billing_postal_code = prepare_arabic_text(row.get('Bill Postal Code', ''))  # Default to empty string
    billing_country = prepare_arabic_text(row.get('Bill Country', ''))
    item_promo_discount = float(row.get('Item Promo Discount', 0.0))  # Ensure numeric
    shipment_promo_discount = float(row.get('Shipment Promo Discount', 0.0))  # Ensure numeric

    # Debug: Print raw values to verify
    print(f"Order {order_id}: Item Price={item_price}, Shipping Price={shipping_price}, Item Promo Discount={item_promo_discount}, Shipment Promo Discount={shipment_promo_discount}")

    # Calculate total (Item Price + Shipping Price - absolute Discounts)
    total = item_price + shipping_price - abs(item_promo_discount) - abs(shipment_promo_discount)

    # Generate additional fields
    current_date = datetime.now().strftime("%b %d, %Y")  # e.g., "Mar 10, 2025"
    due_date = (datetime.now() + timedelta(days=30)).strftime("%b %d, %Y")  # Due date 30 days from now
    order_number = str(index + 1000)  # Simple order number (e.g., 1000, 1001, ...)

    # Step 3: Create PDF invoice
    pdf_file = os.path.join(output_folder, f'{order_id}.pdf')
    c = canvas.Canvas(pdf_file, pagesize=letter)

    # Header Section
    c.setFillColor(colors.lightgrey)
    c.rect(0, 720, 612, 80, fill=1)  # Background for header
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 750, "INVOICE")
    c.setFont("Helvetica", 12)
    c.drawString(400, 750, f"Invoice #: {order_id}")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 730, "Al Barkah LLC")
    c.setFont("Helvetica", 10)
    c.drawString(50, 710, "Al Khobar, Eastern Region, KSA")
    c.line(50, 690, 550, 690)  # Horizontal line below header



    # Bill To
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 660, "Bill To:")
    c.setFont("ArabicFont", 10)
    y_position = 640
    c.drawString(50, y_position, buyer_name)
    y_position -= 15

    # Wrap billing address
    billing_address = f"{billing_address1} {billing_address2}"
    wrapped_billing_address = wrap_text(billing_address, 150, "ArabicFont", 10, c)
    for line in wrapped_billing_address:
        c.drawString(50, y_position, prepare_arabic_text(line))
        y_position -= 15

    billing_location = f"{billing_city}, {billing_state} {billing_postal_code}"
    wrapped_billing_location = wrap_text(billing_location, 150, "ArabicFont", 10, c)
    for line in wrapped_billing_location:
        c.drawString(50, y_position, prepare_arabic_text(line))
        y_position -= 15

    c.drawString(50, y_position, billing_country)
    billing_end_y = y_position - 15

    # Ship To
    c.setFont("Helvetica-Bold", 12)
    c.drawString(200, 660, "Ship To:")
    c.setFont("ArabicFont", 10)
    y_position = 640
    c.drawString(200, y_position, buyer_name)
    y_position -= 15

    # Wrap shipping address
    shipping_address = f"{shipping_address1} {shipping_address2}"
    wrapped_shipping_address = wrap_text(shipping_address, 150, "ArabicFont", 10, c)
    for line in wrapped_shipping_address:
        c.drawString(200, y_position, prepare_arabic_text(line))
        y_position -= 15

    shipping_location = f"{shipping_city}, {shipping_state}"
    wrapped_shipping_location = wrap_text(shipping_location, 150, "ArabicFont", 10, c)
    for line in wrapped_shipping_location:
        c.drawString(200, y_position, prepare_arabic_text(line))
        y_position -= 15

    c.drawString(200, y_position, shipping_country)
    shipping_end_y = y_position - 15

    # Additional Info
    c.setFont("Helvetica-Bold", 12)
    c.drawString(350, 660, "Date:")
    c.setFont("Helvetica", 10)
    c.drawString(350, 640, current_date)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(350, 620, "Due Date:")
    c.setFont("Helvetica", 10)
    c.drawString(350, 600, due_date)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(350, 580, "Order Number:")
    c.setFont("Helvetica", 10)
    c.drawString(350, 560, order_number)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(350, 540, "Balance Due:")
    c.setFont("Helvetica", 10)
    c.drawString(350, 520, f"SAR {total:.2f}")

    # Determine the lowest y-position to draw the table
    table_start_y = min(billing_end_y, shipping_end_y, 500) - 30  # Adjusted spacing
    c.line(50, table_start_y, 550, table_start_y)  # Horizontal line above table

    # Order Details Table
    # Wrap product title to fit in the table
    wrapped_product = wrap_text(product_title, 200, "ArabicFont", 10, c)
    if len(wrapped_product) > 1:
        row_height = len(wrapped_product) * 15  # Adjust row height for wrapped text
    else:
        row_height = 20  # Default row height

    data = [
        ["Item", "Quantity", "Rate", "Amount"],
        ["\n".join(wrapped_product), str(quantity), f"SAR {item_price:.2f}", f"SAR {total:.2f}"]
    ]
    # Specify row heights: header row (25), data row (dynamic)
    table = Table(data, colWidths=[200, 70, 70, 70], rowHeights=[25, row_height])  # Adjusted colWidths and header height
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, 1), 'ArabicFont'),  # Use Arabic font for data row
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('FONTSIZE', (0, 1), (-1, 1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.black),  # Added inner grid for better separation
    ]))
    table_y_position = table_start_y - 25 - row_height  # Adjusted to avoid overlap
    table.wrapOn(c, 50, table_y_position)
    table.drawOn(c, 50, table_y_position)

    # Calculate the total table height (header + data row)
    table_height = 25 + row_height  # Header row (25) + data row (dynamic)
    table_bottom_y = table_y_position  # Bottom of the table
    print(f"Order {order_id}: Table height={table_height}, Table bottom y={table_bottom_y}")

    # Totals (dynamically position based on table height)
    c.setFont("Helvetica-Bold", 12)
    y_position = table_bottom_y - 70  # Increased spacing below table
    print(f"Order {order_id}: Total y_position={y_position}")
    c.drawString(350, y_position, "Total:")
    c.setFont("Helvetica", 10)
    c.drawString(400, y_position, f"SAR {total:.2f}")
    y_position -= 60  # Increased spacing between Total and Paid to 60
    print(f"Order {order_id}: Paid y_position={y_position}")
    # Use a simpler font and encode the string to avoid rendering issues
    c.setFont("Helvetica", 12)  # Switch to Helvetica instead of Helvetica-Bold
    c.drawString(350, y_position, "Paid:".encode('utf-8').decode('utf-8'))
    c.setFont("Helvetica", 10)
    c.drawString(400, y_position, f"SAR {total:.2f}")  # Assuming paid in full for FBA orders

    # Footer Section (ensure it doesn't overlap with content above)
    footer_y = 40  # Footer height
    if y_position - 20 < footer_y:
        print(f"Warning: Order {order_id} content is too close to footer. y_position={y_position}, footer_y={footer_y}")
    c.setFillColor(colors.lightgrey)
    c.rect(0, 0, 612, 40, fill=1)  # Background for footer
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8)
    c.drawCentredString(306, 20, "Thank you for your business!")

    # Save the PDF
    c.save()

    print(f"Generated invoice for Order #{order_id}")

print("Invoice generation complete! BY AARQ")