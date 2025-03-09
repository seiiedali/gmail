import sqlite3
import pandas as pd
from bs4 import BeautifulSoup

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

# Function to extract data from HTML
def extract_data_from_html(html_content):
    soup = BeautifulSoup(html_content, "html.parser")

    # Extract order details
    tables = soup.find_all("table")
    order_table = None
    for table in tables:
        if "PO Number" in table.get_text():
            order_table = table
            break
    # order_table = soup.find("table", text=lambda x: x and "PO Number" in x)
    order_rows = order_table.find_all("tr")[1] if order_table else []
    order_data = {
        "po_number": order_rows.find_all("td")[0].text.strip(),
        "customer_name": order_rows.find_all("td")[1].text.strip(),
        "sold_on": order_rows.find_all("td")[2].text.strip(),
        "must_ship_by": order_rows.find_all("td")[3].text.strip(),
        "ship_method": order_rows.find_all("td")[4].text.strip(),
        "delivery_type": order_rows.find_all("td")[5].text.strip(),
        "payment_method": order_rows.find_all("td")[6].text.strip()
    }
    
    # Extract customer details
    ship_to_table = soup.find("table", text=lambda x: x and "Ship To" in x)
    ship_to_rows = ship_to_table.find_all("tr")[1] if ship_to_table else []
    customer_data = {
        "name": ship_to_rows.find_all("td")[0].text.strip(),
        "address": ship_to_rows.find_all("td")[1].text.strip(),
        "phone_number": ship_to_rows.find_all("td")[2].text.strip(),
        "email_address": ship_to_rows.find_all("td")[3].text.strip()
    }
    
    # Extract products and order items
    products = []
    order_items = []
    product_table = soup.find("table", text=lambda x: x and "Item Code" in x)
    if product_table:
        product_rows = product_table.find_all("tr")[1:]
        for row in product_rows:
            cols = row.find_all("td")
            item_code = cols[1].text.strip()
            description = cols[2].text.strip()
            quantity = int(cols[0].text.strip())
            price = float(cols[5].text.strip().replace("$", ""))
            
            products.append({"item_code": item_code, "description": description})
            order_items.append({"order_po_number": order_data["po_number"], "product_item_code": item_code, "quantity": quantity, "price": price})
    
    return customer_data, order_data, products, order_items

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
