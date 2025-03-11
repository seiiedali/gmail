import sqlite3
import pandas as pd
from bs4 import BeautifulSoup
import re

# Database setup
def setup_database():
    conn = sqlite3.connect("orders.db")
    cursor = conn.cursor()
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS customers (
        name TEXT PRIMARY KEY,
        address TEXT,
        phone_number TEXT,
        email_address TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS products (
        item_code TEXT PRIMARY KEY,
        description TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders (
        po_number TEXT PRIMARY KEY,
        customer_name TEXT,
        sold_on TEXT,
        must_ship_by TEXT,
        ship_method TEXT,
        delivery_type TEXT,
        payment_method TEXT,
        FOREIGN KEY (customer_name) REFERENCES customers(name)
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS order_items (
        order_po_number TEXT,
        product_item_code TEXT,
        quantity INTEGER,
        price REAL,
        PRIMARY KEY (order_po_number, product_item_code),
        FOREIGN KEY (order_po_number) REFERENCES orders(po_number),
        FOREIGN KEY (product_item_code) REFERENCES products(item_code)
    )''')
    
    conn.commit()
    conn.close()


def find_target_table(soup, target_text="PO Number"):
    
    # Find all h5 elements containing the target text
    h5_elements = soup.find_all('h5', string=lambda text: target_text in text if text else False)
    
    for h5 in h5_elements:
        # Navigate up the DOM: h5 -> td -> tr -> tbody -> table
        td_parent = h5.find_parent('td')
        if not td_parent:
            continue
            
        tr_parent = td_parent.find_parent('tr')
        if not tr_parent:
            continue
            
        tbody_parent = tr_parent.find_parent('tbody')
        if not tbody_parent:
            continue
            
        table_parent = tbody_parent.find_parent('table')
        if not table_parent:
            continue
            
        # Check if this table doesn't contain other tables
        if not table_parent.find('table'):
            return table_parent
    
    return None


def extract_order_data(soup):
    
    # Find the h5 element with "PO Number" and navigate to the inner table
    inner_table = find_target_table(soup, target_text="PO Number")
    # Get all header elements (first row)
    header_row = inner_table.find('tr')
    headers = [h5.get_text().strip() for h5 in header_row.find_all('h5')]
    
    # Get the data row (second row)
    data_row = header_row.find_next_sibling('tr')
    data_cells = data_row.find_all('td')
    
    # Extract values for each field
    order_data = {
        "po_number": "",
        "customer_name": "",
        "sold_on": "",
        "must_ship_by": "",
        "ship_method": "",
        "delivery_type": "",
        "payment_method": ""
    }
    
    # Map headers to values
    header_to_value_map = {}
    for i, header in enumerate(headers):
        # Get all h5 texts in the corresponding data cell
        cell_texts = [h.get_text().strip() for h in data_cells[i].find_all('h5') if h.get_text().strip()]
        # Join non-empty texts with space
        value = " ".join(cell_texts)
        header_to_value_map[header] = value
    
    # Populate the order_data dictionary
    order_data["po_number"] = header_to_value_map.get("PO Number", "")
    order_data["sold_on"] = header_to_value_map.get("Sold On", "")
    order_data["must_ship_by"] = header_to_value_map.get("Must Ship By", "")
    order_data["ship_method"] = header_to_value_map.get("Ship Method", "")
    order_data["delivery_type"] = header_to_value_map.get("Delivery Type", "")
    order_data["payment_method"] = header_to_value_map.get("Payment Method", "")
    
    # Special case for customer_name (assuming "Sold On" value is the customer name)
    order_data["customer_name"] = header_to_value_map.get("Sold On", "")
    
    return order_data



def extract_customer_data(soup):
    # Find the h5 element with "PO Number" and navigate to the inner table
    inner_table = find_target_table(soup, target_text="Account # / Customer #")
    # Find the "Customer" column index first
    headers_row = inner_table.find('tr')
    headers = headers_row.find_all('td')
    
    customer_col_idx = None
    ship_to_col_idx = None
    
    for idx, header in enumerate(headers):
        header_text = header.find('h5').get_text().strip()
        if header_text == "Customer":
            customer_col_idx = idx
        elif header_text == "Ship To":
            ship_to_col_idx = idx
    
    if customer_col_idx is None and ship_to_col_idx is None:
        return {"name": "", "address": "", "phone_number": "", "email_address": ""}
    
    # Get the data row
    data_row = headers_row.find_next_sibling('tr')
    data_cells = data_row.find_all('td')
    
    # Initialize customer data
    customer_data = {
        "name": "",
        "address": "",
        "phone_number": "",
        "email_address": ""
    }
    
    # Try to extract from Customer column first, then Ship To if needed
    primary_col_idx = customer_col_idx if customer_col_idx is not None else ship_to_col_idx
    cell = data_cells[primary_col_idx]
    
    # Get all h5 elements in the cell
    h5_elements = cell.find_all('h5')
    h5_texts = [h.get_text().strip() for h in h5_elements]
    non_empty_h5s = [text for text in h5_texts if text]
    
    # Extract name - usually the first non-empty h5
    if non_empty_h5s:
        customer_data["name"] = non_empty_h5s[0]
    # Extract address - usually the next 2-3 non-empty h5s joined together

    stripped_address_components = [re.sub(r'\s{2,}', ' ', s).strip() for s in non_empty_h5s[1:]]
    customer_data["address"] = ", ".join(stripped_address_components)
    
    # If we didn't get complete information, check the Ship To column
    if customer_col_idx is not None and ship_to_col_idx is not None:
        ship_to_cell = data_cells[ship_to_col_idx]
        ship_to_h5s = [h.get_text().strip() for h in ship_to_cell.find_all('h5')]
        non_empty_ship_to_h5s = [text for text in ship_to_h5s if text]
        
        # Look for phone and email in the Ship To column
        for text in non_empty_ship_to_h5s:
            # Simple checks for phone number (all digits or with some formatting)
            if text.replace("-", "").replace("(", "").replace(")", "").replace(" ", "").replace("+", "").isdigit():
                customer_data["phone_number"] = text
            # Simple check for email
            elif "@" in text and "." in text:
                customer_data["email_address"] = text
    
    return customer_data

# Function to extract data from HTML
def extract_data_from_html(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    order_table  = extract_order_data(soup=soup)
    customer_data = extract_customer_data(soup=soup)
    # # Extract products and order items
    # products = []
    # order_items = []
    # product_table = soup.find("table", text=lambda x: x and "Item Code" in x)
    # if product_table:
    #     product_rows = product_table.find_all("tr")[1:]
    #     for row in product_rows:
    #         cols = row.find_all("td")
    #         item_code = cols[1].text.strip()
    #         description = cols[2].text.strip()
    #         quantity = int(cols[0].text.strip())
    #         price = float(cols[5].text.strip().replace("$", ""))
            
    #         products.append({"item_code": item_code, "description": description})
    #         order_items.append({"order_po_number": order_data["po_number"], "product_item_code": item_code, "quantity": quantity, "price": price})
    
    # return customer_data, order_data, products, order_items

# Function to save data to SQLite
def save_to_database(customer, order, products, order_items):
    conn = sqlite3.connect("orders.db")
    cursor = conn.cursor()
    
    cursor.execute('''INSERT INTO customers (name, address, phone_number, email_address)
                      VALUES (?, ?, ?, ?) ON CONFLICT(name) DO UPDATE SET 
                      address=excluded.address, phone_number=excluded.phone_number, email_address=excluded.email_address''',
                   (customer["name"], customer["address"], customer["phone_number"], customer["email_address"]))
    
    cursor.execute('''INSERT INTO orders (po_number, customer_name, sold_on, must_ship_by, ship_method, delivery_type, payment_method)
                      VALUES (?, ?, ?, ?, ?, ?, ?) ON CONFLICT(po_number) DO UPDATE SET
                      customer_name=excluded.customer_name, sold_on=excluded.sold_on, must_ship_by=excluded.must_ship_by,
                      ship_method=excluded.ship_method, delivery_type=excluded.delivery_type, payment_method=excluded.payment_method''',
                   (order["po_number"], order["customer_name"], order["sold_on"], order["must_ship_by"],
                    order["ship_method"], order["delivery_type"], order["payment_method"]))
    
    for product in products:
        cursor.execute('''INSERT INTO products (item_code, description)
                          VALUES (?, ?) ON CONFLICT(item_code) DO UPDATE SET description=excluded.description''',
                       (product["item_code"], product["description"]))
    
    for item in order_items:
        cursor.execute('''INSERT INTO order_items (order_po_number, product_item_code, quantity, price)
                          VALUES (?, ?, ?, ?) ON CONFLICT(order_po_number, product_item_code) DO UPDATE SET
                          quantity=excluded.quantity, price=excluded.price''',
                       (item["order_po_number"], item["product_item_code"], item["quantity"], item["price"]))
    
    conn.commit()
    conn.close()

# Function to export data to Excel
def export_to_excel():
    conn = sqlite3.connect("orders.db")
    
    customers_df = pd.read_sql_query("SELECT * FROM customers", conn)
    products_df = pd.read_sql_query("SELECT * FROM products", conn)
    orders_df = pd.read_sql_query("SELECT * FROM orders", conn)
    
    with pd.ExcelWriter("exported_data.xlsx") as writer:
        customers_df.to_excel(writer, sheet_name="Customers", index=False)
        products_df.to_excel(writer, sheet_name="Products", index=False)
        orders_df.to_excel(writer, sheet_name="Orders", index=False)
    
    conn.close()
