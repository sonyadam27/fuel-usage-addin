import json, urllib.request

# Load credentials
myg = json.load(open('C:/Users/sonyadam/myg-cli.json'))
profile = myg['profiles']['alamui']
creds = {
    'database': 'alamui_indonesia',
    'userName': 'sonyadam@geotab.com',
    'sessionId': profile['keycloakSession']
}

# Load the add-in config
addin_config = json.load(open('C:/Users/sonyadam/Desktop/fuel-usage-addin/config-inline.json'))

def api_call(method, params):
    params['credentials'] = creds
    payload = json.dumps({'method': method, 'params': params}).encode()
    req = urllib.request.Request('https://my.geotab.com/apiv1', data=payload, headers={'Content-Type': 'application/json'})
    resp = urllib.request.urlopen(req)
    result = json.loads(resp.read())
    if 'error' in result:
        raise Exception(result['error']['message'])
    return result.get('result')

# Step 1: Get current SystemSettings
print("Getting current SystemSettings...")
ss = api_call('Get', {'typeName': 'SystemSettings'})
system_settings = ss[0]

# Step 2: Add the new add-in to customerPages
customer_pages = system_settings.get('customerPages', [])

# Check if Fuel Usage Per Day already exists
existing = [a for a in customer_pages if a.get('name') == 'Fuel Usage Per Day']
if existing:
    print(f"Removing existing Fuel Usage Per Day add-in(s): {len(existing)}")
    customer_pages = [a for a in customer_pages if a.get('name') != 'Fuel Usage Per Day']

# Add the new add-in config
print(f"Adding Fuel Usage Per Day add-in with {len(addin_config.get('files', {}))} files...")
customer_pages.append(addin_config)

# Step 3: Update SystemSettings
system_settings['customerPages'] = customer_pages
print("Saving SystemSettings...")
result = api_call('Set', {'typeName': 'SystemSettings', 'entity': system_settings})
print(f"Saved! Result: {result}")

# Step 4: Verify
ss2 = api_call('Get', {'typeName': 'SystemSettings'})
pages = ss2[0].get('customerPages', [])
fuel = [a for a in pages if a.get('name') == 'Fuel Usage Per Day']
if fuel:
    f = fuel[0]
    print(f"\nVerified: Fuel Usage Per Day installed!")
    print(f"  Files: {list(f.get('files', {}).keys())}")
    print(f"  Items path: {f['items'][0]['path'] if f.get('items') else 'none'}")
    print(f"  Config length: {len(json.dumps(f))} chars")
else:
    print("\nWARNING: Fuel Usage Per Day NOT found after save!")
