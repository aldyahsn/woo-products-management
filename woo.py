import csv
import pandas as pd
from woocommerce import API
import json
import sys
import requests
from datetime import datetime, timedelta

start_time = datetime.now()

# WooCommerce REST API credentials
site_url = "https://hamesha.rocketweb.id/"
media_endpoint = "/wp-json/wp/v2/media"
wcapi = API(
    url=site_url,  # Your store URL
    consumer_key="ck_e30a83a4853fe0c72425e794317fa8a47e79c8b5",  # Your consumer key
    consumer_secret="cs_79876babb4634488847c4bc70aabf57c98ac5540",  # Your consumer secret
    version="wc/v3",  # API version
    timeout=30,
    wp_api=True,
)
auth = requests.auth.HTTPBasicAuth('admin', 'admin')

# Make the GET request to fetch media items
media_items = []
response = requests.get(f"{site_url}{media_endpoint}", auth=auth)
if response.status_code == 200:
    media_items = response.json()
    # for item in media_items:
    #     # print(json.dumps(item, indent=4))
    #     # Print details about each media item
    #     print(f"ID: {item['id']}, URL: {item['source_url']}, filename: {item['title']['rendered']}")
else:
    print(f"Failed to retrieve media items. Status code: {response.status_code}")


categories = wcapi.get('products/categories').json()
# print(json.dumps(categories, indent=4))

# Read Excel file
excel_file_path = 'Data.xlsx'
scanned_products = []

df = pd.read_excel(excel_file_path, engine='openpyxl')

total = len(df)
scanned = 0


for index, row in df.iterrows():
    data_row = row.to_dict()

    # Looping data di file excel, jika sudah discanned, maka skip
    if data_row['NAME'] not in scanned_products:
        
        # lakukan semua pendataan di dataframe excel nama produk yang terkait
        datafound_df = df[df['NAME'].str.lower() == data_row['NAME'].lower()]
        scanned_products.append(data_row['NAME'])
        
        data_variations = []
        data_dimensions = []
        data_area = []
        data_subcategory = {}
        data_image = {}

        # cari data category
        for _category in categories:
            _name = _category.get("name")
            if _name == data_row['SUBCATEGORY']:
                data_subcategory = _category
                break
        # jika kategori tidak ketemu maka create subcategori baru
        if not data_subcategory:
            print('Ada data yang kategorinya masih kosong')
            sys.exit()  


        # cari data image
        for _image in media_items:
            _image_id = _image['id']
            _image_url = _image['source_url']
            _image_name = _image['title']['rendered']
            if _image_name == data_row['NAME']:
                data_image['id'] = _image_id



        # add variants 
        for _variant_index, _variant_row in datafound_df.iterrows():
            _data_variant = _variant_row.to_dict()

            if _data_variant['DIMENSION'] not in data_dimensions:
                data_dimensions.append(_data_variant['DIMENSION'])
            if _data_variant['AREA'] not in data_area:
                data_area.append(_data_variant['AREA'])
            
            _v = {
                "regular_price": str(_data_variant['PRICE']),
                "sku": _data_variant["SKU"],
                "attributes": [
                    {"name": "Dimension", "option": _data_variant['DIMENSION']},
                    {"name": "Area", "option": _data_variant['AREA']}
                ],
                # "weight": "",
                # "dimensions": {
                #     "length": "",
                #     "width": "",
                #     "height": ""
                # },
            }
            data_variations.append(_v)
        

        data_parent = {
            "name": data_row['NAME'],
            "categories": [
                {"id": data_subcategory['id']},
            ],
            "images": [data_image],
            "type": "variable",
            "tags": [],
            "attributes": [
                {
                    "name": "Dimension",
                    "visible": True,
                    "variation": True,
                    "options": data_dimensions,
                },
                {
                    "name": "Area",
                    "visible": True,
                    "variation": True,
                    "options": data_area
                }
            ]
        }

        # PARENT
        # Cek nama untuk parent jika exist
        # print(json.dumps(data_parent, indent=4))
        response_parent = wcapi.get("products", params={'search':data_parent['name']}).json()
        if response_parent:
            id_parent = response_parent[0]['id']
            response_parent = wcapi.put(f"products/{id_parent}", data_parent).json()
            print(f"Product {response_parent['name']} updated: {response_parent['id']}")
            # print(json.dumps(data_parent, indent=4))
        else:
            response_parent = wcapi.post("products", data_parent).json()
            print(f"Product {response_parent['name']} created: {response_parent['id']}")
            # print(json.dumps(data_parent, indent=4))

        # VARIATIONS
        for variation in data_variations:
            
            # CEK SKU pada variant jika exist
            response_variant = wcapi.get(f"products/{response_parent['id']}/variations", params={'sku':variation['sku']}).json()
            if response_variant:
                id_variant = response_variant[0]['id']
                wcapi.put(f"products/{response_parent['id']}/variations/{id_variant}", variation).json()
                print(f"- Variant {variation['sku']} updated: {id_variant}")        
            else:
                response_variant = wcapi.post(f"products/{response_parent['id']}/variations", variation).json()
                print(f"- Variant {response_variant['sku']} created: {response_variant['id']}")   

        scanned = scanned + len(data_variations)
        print(f"----> Scanned: {(scanned/total)*100}%")     

endtime = datetime.now()
time_difference = endtime - start_time
days = time_difference.days
seconds = time_difference.seconds
hours = seconds // 3600
minutes = (seconds % 3600) // 60
seconds = seconds % 60
print(f"To scan {total} data, it takes {days}:{hours}:{minutes}:{seconds}")