import requests
import base64
import json


# # get order from flipkart api
# base_url = f'{token_base_url}/sellers'
# shipment_end_point = '/v3/shipments/filter'
# headers = {
#     "Authorization": f"Bearer {access_token}",
#     'accept': "application/json",
#     'Content-Type': "application/json"
# }
# page_size = {
#     "pageSize": 20
# }
# json_data = {
#     "filter": {
#         "type": "postDispatch",
#         "states": [
#             "SHIPPED"
#         ],
#         "serviceProfiles": [
#             "FBF_LITE"
#         ]
#     }
# }
#
# # get data from flipkart
# shipment_request = requests.post(
#     base_url+shipment_end_point, headers=headers, json=json_data, params=page_size)
# if shipment_request.status_code >= 400:
#     print(shipment_request.json().get("message"))
#
# # convert to json and get python dictonary
# order_json = shipment_request.json()
#
# print(order_json.get('nextPageUrl'))
# # print in json format for better visual of data
# print(json.dumps(order_json, indent=4))


# i = 0
# while i < len(order_json):
#     shipment = order_json.get('shipments')[i].get('orderItems')[0]
#     print(i+1)
#     print(order_json.get('shipments')[i].get("locationId"))
#     print("Order item id: "+shipment.get('orderItemId'))
#     print("Sku: "+shipment.get('sku'))
#     print('quantity: '+str(shipment.get('quantity')))
#     print('serviceProfile: '+shipment.get('serviceProfile'))

#     i += 1

# # order in handover count
# hoc_endpoint = '/v3/shipments/handover/counts'
# location_id = {
#     'locationId': order_json.get('shipments')[0].get("locationId")}
#
#
# hoc_data = requests.get(base_url+hoc_endpoint,
#                         params=location_id, headers=headers)
# # print(hoc_data.content)

# get order from flipkart api


def get_access_token():
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
        token_base_url + token_end_point, headers=token_header, params=token_data).json()

    access_token: str = token_request_json["access_token"]
    print(f'access token {access_token}')

    return {'access_token': access_token,
            'base_url': token_base_url
            }


def get_return_data(data_id: str):
    auth_data = get_access_token()

    base_url = auth_data.get('base_url')
    access_token = auth_data.get('access_token')

    base_url = f'{base_url}/sellers'
    end_point = '/v2/returns'
    headers = {
        "Authorization": f"Bearer {access_token}",
        'accept': "application/json",
        'Content-Type': "application/json"
    }
    params = {
        "returnIds": {data_id}
    }

    # get data from flipkart
    shipment_request = requests.get(
        base_url + end_point, params=params, headers=headers)

    if shipment_request.status_code != 200:
        return shipment_request.json()[0].get('message')
        # print(shipment_request.json().get("message"))
    else:
        # convert to json and get python dictonary
        return_order_json = shipment_request.json()

        # print in json format for better visual of data
        print(json.dumps(return_order_json, indent=4))
        order_item_id: str = return_order_json.get('returnItems')[0].get('orderItemId')
        return order_item_id


order_item_id_get = get_return_data("8982944177")
print(order_item_id_get)


def get_order_data(order_item_id):
    auth_data = get_access_token()

    base_url = auth_data.get('base_url')
    access_token = auth_data.get('access_token')

    base_url = f'{base_url}/sellers'
    end_point = '/v2/orders'
    headers = {
        "Authorization": f"Bearer {access_token}",
        'accept': "application/json",
        'Content-Type': "application/json"
    }
    params = {
        "order_item_id": order_item_id
    }
    url = f"{base_url}{end_point}/{params.get('order_item_id')}"
    print(url)
    # get data from flipkart
    shipment_request = requests.get(url=url, headers=headers)

    if shipment_request.status_code != 200:
        print(shipment_request.status_code)
        return shipment_request.json()
        # print(shipment_request.json().get("message"))
    else:
        # convert to json and get python dictionary
        order_json = shipment_request.json()

        # # print in json format for better visual of data
        # print(json.dumps(order_json, indent=4))
        return order_json


a = get_order_data("22333140838125600")
print(a)
