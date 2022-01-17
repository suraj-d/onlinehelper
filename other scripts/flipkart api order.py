import requests
import base64
import json

client_id = '8191b171a937795714745572633601551757'
client_secret = '250c534e779105d610c3bfb42645478ee'

token_base_url = 'https://api.flipkart.net'
token_end_point = '/oauth-service/oauth/token'

token_data = {
    'grant_type': "client_credentials",
    "scope": "Seller_Api"
}

client_creds = f'{client_id}:{client_secret}'
client_creds_b64 = base64.b64encode(client_creds.encode())

token_header = {
    'Authorization': f"Basic {client_creds_b64.decode()}"
}

token_request_json = requests.get(
    token_base_url+token_end_point, headers=token_header, params=token_data).json()

access_token = token_request_json["access_token"]
print(f'access token {access_token}')

# get order from flipkart api
base_url = f'{token_base_url}/sellers'
shipment_end_point = '/v3/shipments/filter'
headers = {
    "Authorization": f"Bearer {access_token}",
    'accept': "application/json",
    'Content-Type': "application/json"
}
page_size = {
    "pageSize": 20
}
json_data = {
    "filter": {
        "type": "postDispatch",
        "states": [
            "SHIPPED"
        ],
        "serviceProfiles": [
            "FBF_LITE"
        ]
    }
}

# get data from flipkart
shipment_request = requests.post(
    base_url+shipment_end_point, headers=headers, json=json_data, params=page_size)
if shipment_request.status_code >= 400:
    print(shipment_request.json().get("message"))

# convert to json and get python dictonary
order_json = shipment_request.json()

print(order_json.get('nextPageUrl'))
# print in json format for better visual of data
print(json.dumps(order_json, indent=4))


i = 0
# while i < len(order_json):
#     shipment = order_json.get('shipments')[i].get('orderItems')[0]
#     print(i+1)
#     print(order_json.get('shipments')[i].get("locationId"))
#     print("Order item id: "+shipment.get('orderItemId'))
#     print("Sku: "+shipment.get('sku'))
#     print('quantity: '+str(shipment.get('quantity')))
#     print('serviceProfile: '+shipment.get('serviceProfile'))

#     i += 1

# order in handover count
hoc_endpoint = '/v3/shipments/handover/counts'
location_id = {
    'locationId': order_json.get('shipments')[0].get("locationId")}


hoc_data = requests.get(base_url+hoc_endpoint,
                        params=location_id, headers=headers)
# print(hoc_data.content)

return_endpoint='/v2/returns'
return_id = {
    ''
}
