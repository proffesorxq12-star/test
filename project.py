import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw()  # Hide the root window

print("Select Mapping File")
mapping_file = askopenfilename(
    title="Select Mapping File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

print("Select Inventory File")
inventory_file = askopenfilename(
    title="Select Inventory File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

df = pd.read_excel(mapping_file)
dK = pd.read_excel(inventory_file)


duplicate_mapping = {}

for _, row in df.iterrows():

    original = row.iloc[0]

    duplicates = []

    for sku in row.iloc[1:]:
        if pd.notna(sku):
            duplicates.append(str(sku))

    duplicate_mapping[original] = duplicates


inventory_update = dict(zip(dK["sku"], dK["quantity"]))

final_dict = {}

for inv_sku,inv_stock in inventory_update.items():
    final_dict[inv_sku] = inv_stock
    k = duplicate_mapping.get(inv_sku, [])
    for i in k:
        final_dict[i] = inv_stock

for key,values in final_dict.items():
    print(f"{key}:{values}")

rows = []

for sku, qty in final_dict.items():

    rows.append({
        "sku": sku,
        "price": "",
        "minimum-seller-allowed-price": "",
        "maximum-seller-allowed-price": "",
        "quantity": qty,
        "leadtime-to-ship": "",
        "fulfillment-channel": "",
        "merchant_shipping_group_name": ""
    })

df = pd.DataFrame(rows)
p
df.to_excel("amazon_inventory.xlsx", index=False)

print("Excel file created successfully.")

