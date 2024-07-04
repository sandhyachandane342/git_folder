import requests
import json
from openpyxl import Workbook 
import datetime

query = '''
{
    products(first: 250, query: "tag:larah") {
      edges {
        node {
          id
          title
          variants(first: 1) {
            edges {
              node {
                sku
                createdAt
              }
            }
          }
        }
      }
    }
}
'''

def get_previous_timestamp(days_ago):
    previous_date = datetime.datetime.now() - datetime.timedelta(days = days_ago)
    return int(previous_date.timestamp())

def get_shopify_details(shopify_url, headers):
    product_details = []
    try:
        response = requests.post(shopify_url, headers=headers, json={'query': query})
        response.raise_for_status()
        data = response.json() 
        products = data['data']['products']['edges']
        
        for product in products:
            product_node = product['node']
            product_id = product_node['id']
            product_title = product_node['title']
            variants = product_node['variants']['edges']
            
            previous_timestamp = get_previous_timestamp(1)
            
            for variant in variants:
                variant_node = variant['node']
                variant_sku = variant_node['sku']
                
                variant_created_at = previous_timestamp
                
                product_details.append([product_id, product_title, variant_sku, variant_created_at])
    
    except requests.exceptions.RequestException as e:
        print("Request Exception Error")
    
    return product_details

if __name__ == "__main__":
    shopify_url = "https://my-borosil.myshopify.com/admin/api/2024-07/graphql.json"
    access_token = "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"
    headers = {
        'Content-Type': 'application/json',
        'X-Shopify-Access-Token': access_token
    }
    product_details = get_shopify_details(shopify_url, headers)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["Product ID", "Product Title", "Variant SKU", "Variant Created At"])

    for row in product_details:
        variant_created_at = datetime.datetime.utcfromtimestamp(row[3]).strftime('%Y-%m-%d %H:%M:%S')
        worksheet.append([row[0], row[1], row[2], variant_created_at])
    workbook.save('shopify_data.xlsx')