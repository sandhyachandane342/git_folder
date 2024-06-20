import requests
import json
from openpyxl import Workbook
def count_total_pages(url_base, headers, limit = 100):
    total_pages = 0
    request_order_shipping_details = []
    page = 1
    while True:
        url = f"{url_base}?page={page}&limit={limit}"
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()
            total_pages +=1
            print("Page:", page)
            for item in data['data']['list']:
                request_numbers = item['request_number']
                print(request_numbers)
                order = item.get('order',{})
                order_id = order.get('id','')
                order_name = order.get('name','')
                line_Items = item.get('line_items',[])
                for line_item in line_Items:
                    print(line_item)
                    shipping = line_item.get('shipping',{})
                    shipping_companys = shipping[0].get('shipping_company','')
                    shipping_awb = shipping[0].get('awb','')
                    #shipping_quantity = shipping[0].get('quantity','')
                    #shipping_reason = shipping[0].get('reason','')
                    quantities = line_item.get('quantity')
                    reasons = line_item.get('reason')
                    
                    request_order_shipping_details.append([request_numbers, order_id, order_name, shipping_companys, shipping_awb, quantities, reasons])         
            if not data["data"]["hasNextPage"]:
                print(data)
                break
            page +=1
            if (page==5):
                break
        except requests.RequestException as e:
            print(f"Error occurred: {e}")
            break
    return total_pages, request_order_shipping_details
if __name__ == "__main__":
    url_base = "https://admin.returnprime.com/return-exchange/v2"
    headers = {
  'x-rp-token': '91171f32099b9cc98497a6f139d3f9582bc23ffe2c42c4b9dadce79adfc630ea'
}
    total_pages, request_order_shipping_details = count_total_pages(url_base,headers)
    wb = Workbook()
    ws = wb.active
    ws.append(["Request Number", "Order Id", "Order Name", "Shipping Company", "Shipping awb", "Shipping Quantity", "Shipping Reason"])
    for row in request_order_shipping_details:
        ws.append(row)
    wb.save("request_order_shipping_details.xlsx")
    print("Total Number of Pages:", total_pages)
    print("Request Number save request_order_shipping_details.xlsx")

