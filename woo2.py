import csv
import pandas as pd
from woocommerce import API
import json


# WooCommerce REST API credentials
wcapi = API(
    url="https://hamesha.rocketweb.id/",  # Your store URL
    consumer_key="ck_e30a83a4853fe0c72425e794317fa8a47e79c8b5",  # Your consumer key
    consumer_secret="cs_79876babb4634488847c4bc70aabf57c98ac5540",  # Your consumer secret
    version="wc/v3"  # API version
)

products = wcapi.get('products/19/variations', params={'search':'PRODUK 01', 'sku':'P01-D'}).json()
print(json.dumps(products, indent=4))

