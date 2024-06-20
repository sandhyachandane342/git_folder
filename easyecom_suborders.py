import psycopg2
import requests
from openpyxl import Workbook
import datetime
import json
import time
start_time = time.time()
conn = psycopg2.connect(
    host='borosilcp.postgres.database.azure.com',
    database='easyecom',
    user='db_user',
    password='Hello123!'
)
cursor = conn.cursor()
sql_query = '''SELECT * FROM easyecom_suborders LIMIT 0; '''
cursor.execute(sql_query)
db_data = cursor.fetchall()

db_headers = [desc[0] for desc in cursor.description]
print(db_headers)

cursor.close()
conn.close()

def get_all_orders(url, headers, payload, base_url):
  order_data = []
  next_url = url
  while next_url:
    try:
      response = requests.get(next_url, headers=headers, data=payload, timeout = 30)
      response.raise_for_status()
      data = response.json()
      
    except requests.exceptions.ConnectionError as e:
      print(f"Connection error: {e}")
    except requests.exceptions.ValueError as e:
      print(f"Caught value error: {e}")
    except requests.exceptions.HTTPError as e:
      print("HTTP error")
    except requests.exceptions.MissingSchema as e:
      print("Missing Schema: including http or https")
    except requests.exceptions.Timeout as e:
      print("The request timed out. Please try again later.")
    except requests.exceptions.RequestException as e:
      print(f"Request failed: {e}")
      break
      
    for item in data["data"]["orders"]:
        invoice_id = item.get('invoice_id', '')
        print(invoice_id)
        order_data.append([invoice_id])

    next_url = data["data"].get('nextUrl', '')
    if next_url:
       next_url = base_url + next_url
   
  return order_data

def get_order_details(url, invoice_id, header):
  order_data_new = []
  url = f"https://api.easyecom.io/orders/V2/getOrderDetails?invoice_id={invoice_id}"
  try:
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()
    
  except requests.exceptions.ConnectionError as e:
      print(f"Connection error: {e}")
  except requests.exceptions.ValueError as e:
    print(f"Caught value error: {e}")
  except requests.exceptions.HTTPError as e:
      print("HTTP error")
  except requests.exceptions.MissingSchema as e:
      print("Missing Schema: including http or https")
  except requests.exceptions.Timeout as e:
      print("The request timed out. Please try again later.")
  except requests.exceptions.RequestException as e:
      print(f"Request failed: {e}")

  for item in data["data"]:
      invoice_id = item.get('invoice_id', '')
      order_id = item.get('order_id', '')
      reference_code = item.get('reference_code', '')
      warehouse_name = item.get('company_name', '')
      warehouse_id = item.get('warehouse_id', '')
      pickup_state = item.get('pickup_state', '')
      pickup_state_code = item.get('pickup_state_code', '')
      pickup_pin_code = item.get('pickup_pin_code', '')
      order_type = item.get('order_type', '')
      replacement_order = item.get('replacement_order', '')
      marketplace = item.get('marketplace', '')
      marketplace_id = item.get('marketplace_id', '')
      qc_passed = item.get('qcPassed', '')
      order_date = item.get('order_date', '')
      last_update_timestamp = item.get('last_update_date', '')
      manifest_date = item.get('manifest_date', '')
      invoice_number = item.get('invoice_number', '')
      marketplace_invoice_num = item.get('marketplace_invoice_num', '')
      shipping_last_update_timestamp = item.get('shipping_last_update_date', '')
      courier_aggregator = item.get('courier', '')
      courier_name = item.get('courier_aggregator_name', '')
      carrier_id = item.get('carrier_id', '')
      awb_number = item.get('awb_number', '')
      order_status = item.get('order_status', '')
      order_status_id = item.get('order_status_id', '')
      easyecom_order_history = json.dumps(item.get('easyecom_order_history', ''))
      shipping_status = item.get('shipping_status', '')
      shipping_status_id = item.get('shipping_status_id', '')
      shipping_history = json.dumps(item.get('shipping_history', ''))
      payment_mode = item.get('payment_mode', '')
      payment_mode_id = item.get('payment_mode_id', '')
      payment_gateway_transaction_number = item.get('payment_gateway_transaction_number', '')
      buyer_gst = item.get('buyer_gst', '')
      customer_name = item.get('customer_name', '')
      shipping_name = item.get('shipping_name', '')
      shipping_mobile = item.get('contact_num', '')
      shipping_address_1 = item.get('address_line_1', '')
      shipping_address_2 = item.get('address_line_2', '')
      shipping_city = item.get('city', '')
      shipping_pin_code = item.get('pin_code', '')
      shipping_state = item.get('state', '')
      shipping_state_code = item.get('state_code', '')
      shipping_country = item.get('country', '')
      customer_email = item.get('email', '')
      billing_name = item.get('billing_name', '')
      billing_address_1 = item.get('billing_address_1', '')
      billing_address_2 = item.get('billing_address_2', '')
      billing_city = item.get('billing_city', '')
      billing_state = item.get('billing_state', '')
      billing_state_code = item.get('billing_state_code', '')
      billing_pin_code = item.get('billing_pin_code', '')
      billing_country = item.get('billing_country', '')
      billing_mobile = item.get('billing_mobile', '')
      invoice_documents = json.dumps(item.get('invoice_documents', ''))
      order_item = item.get('order_items', [])
      for order_Item in order_item: 
        suborder_id = order_Item.get('suborder_id', '')
        suborder_num = order_Item.get('suborder_num', '')
        item_quantity = order_Item.get('item_quantity', '')
        returned_quantity = order_Item.get('returned_quantity', '')
        cancelled_quantity = order_Item.get('cancelled_quantity', '')
        shipped_quantity = order_Item.get('shipped_quantity', '')
        sku = order_Item.get('sku', '')
        product_name = order_Item.get('productName', '')
        category = order_Item.get('category', '')
        brand = order_Item.get('brand', '')
        product_tax_code = order_Item.get('product_tax_code', '')
        ean = order_Item.get('ean', '')
        custom_fields = json.dumps(order_Item.get('custom_fields', ''))
        tax_rate = order_Item.get('tax_rate', '')
        selling_price = order_Item.get('selling_price', '')
        invoice_code = order_Item.get('invoicecode', '')
        sku_type = order_Item.get('sku_type', '')
        breakup_types = json.dumps(item.get('breakup_types', ''))    
   
        order_data_new.append([invoice_id, order_id, reference_code, warehouse_name, warehouse_id, pickup_state, pickup_state_code, pickup_pin_code, order_type, replacement_order, marketplace, marketplace_id, qc_passed, order_date, last_update_timestamp, manifest_date, invoice_number, marketplace_invoice_num, shipping_last_update_timestamp, courier_aggregator, courier_name, carrier_id, awb_number, order_status, order_status_id, easyecom_order_history, shipping_status, shipping_status_id, shipping_history, payment_mode, payment_mode_id, payment_gateway_transaction_number, buyer_gst, customer_name, shipping_name, shipping_mobile, shipping_address_1, shipping_address_2, shipping_city, shipping_pin_code, shipping_state, shipping_state_code, shipping_country, customer_email, billing_name, billing_address_1, billing_address_2, billing_city, billing_state, billing_state_code, billing_pin_code, billing_country, billing_mobile, invoice_documents, suborder_id, suborder_num,  item_quantity, returned_quantity, cancelled_quantity, shipped_quantity, sku, product_name, category, brand, product_tax_code, ean,  custom_fields, tax_rate, selling_price, breakup_types, invoice_code, sku_type])
 
  return order_data_new

if __name__ == "__main__":
  current_date = datetime.datetime.now() 
  start_date = (current_date - datetime.timedelta(days=1)).strftime('%Y-%m-%d 00:00:00') 
  end_date = (current_date - datetime.timedelta(days=1)).strftime('%Y-%m-%d 23:59:59')
  url = f"https://api.easyecom.io/orders/V2/getAllOrders?manifest_end_date={end_date}&manifest_start_date={start_date}"

  payload = {}
  headers = {
      'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczpcL1wvbG9hZGJhbGFuY2VyLW0uZWFzeWVjb20uaW9cL2FjY2Vzc1wvdG9rZW4iLCJpYXQiOjE3MTQ0NjI0NTUsImV4cCI6MTcyMjM0NjQ1NSwibmJmIjoxNzE0NDYyNDU1LCJqdGkiOiJtTzI3T0NvMHlLUVlTUWI1Iiwic3ViIjoxMDIwNTgsInBydiI6ImE4NGRlZjY0YWQwMTE1ZDVlY2NjMWY4ODQ1YmNkMGU3ZmU2YzRiNjAiLCJ1c2VyX2lkIjoxMDIwNTgsImNvbXBhbnlfaWQiOjE0NDM2LCJyb2xlX3R5cGVfaWQiOjExLCJwaWlfYWNjZXNzIjowLCJyb2xlcyI6bnVsbCwiY19pZCI6MTQ0MzYsInVfaWQiOjEwMjA1OCwibG9jYXRpb25fcmVxdWVzdGVkX2ZvciI6MTQ0MzZ9._mjo1xPJcCK1UuEfHsNdwU804ebMnaY7GO6ThSE4_os'
  }
  base_url = "https://api.easyecom.io"
  order_data = get_all_orders(url, headers, payload, base_url)

try:
  workbook = Workbook()
  worksheet = workbook.active
  worksheet.append(["invoice_id", "order_id", "reference_code", "warehouse_name", "warehouse_id", "pickup_state", "pickup_state_code", "pickup_pin_code", "order_type", "replacement_order", "marketplace", "marketplace_id", "qc_passed", "order_date", "last_update_timestamp", "manifest_date", "invoice_number", "marketplace_invoice_num", "shipping_last_update_timestamp", "courier_aggregator", "courier_name", "carrier_id", "awb_number", "order_status", "order_status_id", "easyecom_order_history", "shipping_status", "shipping_status_id", "shipping_history", "payment_mode", "payment_mode_id", "payment_gateway_transaction_number", "buyer_gst", "customer_name", "shipping_name", "shipping_mobile", "shipping_address_1", "shipping_address_2", "shipping_city", "shipping_pin_code", "shipping_state", "shipping_state_code", "shipping_country", "customer_email", "billing_name", "billing_address_1", "billing_address_2", "billing_city", "billing_state", "billing_state_code", "billing_pin_code", "billing_country", "billing_mobile", "invoice_documents", "suborder_id", "suborder_num", " item_quantity", "returned_quantity", "cancelled_quantity", "shipped_quantity", "sku", "product_name", "category", "brand" ,"product_tax_code", "ean", " custom_fields", "tax_rate", "selling_price", "breakup_types", "invoice_code", "sku_type"])
  
  for order in order_data:
    invoice_id = order[0]
    order_data_new = get_order_details(url, invoice_id, headers)
        
    for row in order_data_new:
      worksheet.append(row)
      
except requests.exceptions.AttributeError as e:
      print("Attribute error")
except requests.exceptions.TypeError as e:
      print("Type error")     
except requests.exceptions.ImportError as e:
    print("Import error")
    
try:
    workbook.save('easyecom_data.xlsx')
    
except requests.exceptions.PermissionError as e:
   print("Permission error")  
except requests.exceptions.FileNotFoundError as e:
   print("File not found error")
 
end_time = time.time()
execution_time = end_time - start_time
print('Total runtime of the program is:',execution_time, 'seconds')