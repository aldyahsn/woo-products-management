import csv
import pandas as pd
from woocommerce import API
import json

def create_update_product(product_data):
    # Check if product exists by SKU
    response = wcapi.get("products", params={
        'meta_data': [
            {'key': 'id_product', 'value': product_data["ID PRODUCT"]}  # Storing custom ID in meta data
        ]
    })
    products = response.json()
    
    product_payload = {
        "name": product_data["NAME"],
        "type": product_data["TYPE"],
        "regular_price": str(product_data["PRICE"]),
        "description": product_data["DESCRIPTION"],
        "sku": product_data["SKU"],
        'meta_data': [
            {'key': 'id_product', 'value': product_data["ID PRODUCT"]}  # Storing custom ID in meta data
        ]
    }
    
    if products:  # If product exists, update it
        product_id = products[0]['id']
        update_response = wcapi.put(f"products/{product_id}", product_payload)
        print(f"Updated Product: {product_data['SKU']}, Response: {update_response.status_code}")
    else:  # If product does not exist, create it
        create_response = wcapi.post("products", product_payload)
        print(f"Created Product: {product_data['SKU']}, Response: {create_response.status_code}")
    print("---------------")


# WooCommerce REST API credentials
wcapi = API(
    url="https://obsoletecoat.s4-tastewp.com/",  # Your store URL
    consumer_key="ck_b8c753fb28e215b99343b877343b9f529eaef9e0",  # Your consumer key
    consumer_secret="cs_190dc3d5651ef98fc943555da0e52ec6cd95a9db",  # Your consumer secret
    version="wc/v3"  # API version
)

# Read Excel file
excel_file_path = 'Data.xlsx'
df = pd.read_excel(excel_file_path, engine='openpyxl')

# Process each row in the DataFrame
for index, row in df.iterrows():
    product_info = row.to_dict()
    # print(product_info)
    create_update_product(product_info)

