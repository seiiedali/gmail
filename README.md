> extract the information (from multiple html content) and save them to a SQLite database,

the database should have customer, order, product:

## `CUSTOMERS TABLE`

```
name (from ship to) (unique PK)
address (from ship to)
phone_number (from ship to)
email_address (from ship to)
```

## `PRODUCTS TABLE`

```
item_code (unique PK)
description
```

## `ORDERS TABLE`

```
po_number (unique PK)
customer_name (unique FK to CUSTOMERS)
sold_on
must_ship_by
ship_method
delivery_type
payment_method
```

## `ORDER_ITEMS`

```
order_po_number (FK to ORDERS)
product_item_code (FK to PRODUCTS)
quantity
price
(order_po_number + product_item_code = PK)
```

the code should have the ability to give out the list of

- customers
- products
- orders

as a excel file.

obviously the considering the unique filed (PK) in each of the tables there shouldn't be multiple instances of save customers or products or the orders but rather update them on each run of the script

the code also receives a json of credentials (file path) to know where it should connect to read emails.

here is the sample of content of the target emails that we want to extract info from:
