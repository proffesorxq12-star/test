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

dK = pd.read_excel(r"C:\Users\Atal\Desktop\inventory.xlsx")

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







# inventory_update = inventory = {
#     "Toothbrush Holder-01": 0,
#     "Ear Wax Removal Tool 001": 2999,
#     "Bathmat-grey-60x40 001": 246,
#     "Screwdriver-01": 1108,
#     "Door Bottom Stripping Seal Strip 001": 0,
#     "VT-PXM1-HKUS": 0,
#     "3 In 1 Car Polish 001": 50,
#     "LED Motion Sensor Night Light 001": 0,
#     "Ice Roller Face Massager 001": 0,
#     "DO-O5AV-RSVC": 150,
#     "6mm Heat Insulated Cover 001": 100,
#     "Microfibre Cleaning Cloth 001": 9841,
#     "Polka Dot Fabric Swim Cap 001": 0,
#     "DE00L03532-DO-1": 0,
#     "1": 0,
#     "EM00KB1166-DO-1": 110,
#     "Window Squeegee Cleaner-01": 99,
#     "Fruit Peeler Set-01": 100,
#     "Kitchen Cleaner Spray-01": 101,
#     "Car Duster Mop-01": 2000,
#     "Manual Push Chopper-01": 98,
#     "Screwdriver-02": 3000,
#     "Non-Slip Bath Mat-01": 0,
#     "Ice Roller for Face Massager-01": 0,
#     "Stainless Steel Y-Peeler-01": 1500,
#     "EM00KB1198-DO-1": 111,
#     "USLA001-DF-1": 10,
#     "SDCA01-DF-1": 12,
#     "EM00KB1183-DO-1": 103,
#     "SP012-DF-1": 10,
#     "EM00KB1179-DO-1": 107,
#     "EM00KB1197-DO-1": 101,
#     "DE00L03564-DO-1": 0,
#     "EM00KB1167-DO-1": 105,
#     "EM00KB1180-DO-1": 101,
#     "EM00KB1187-DO-1": 103,
#     "EM00KB1189-DO-1": 107,
#     "EM00KB978-DO-1": 100,
#     "FK-X87T-GWQW": 150,
#     "F2-8ZV2-ZOVT": 100,
#     "7J-75D9-DIHY": 100,
#     "NB-PHF3-DMC5": 100,
#     "3B-ZTVL-KYXE": 70,
#     "LW-A2IW-HSC8": 84,
#     "RI-A7BZ-Y0IM": 99,
#     "R7-GEKP-T5DR": 100,
#     "54-5ZN0-KBHA": 111,
#     "EM00KB498-DO-1": 80,
#     "S4-O1CB-2P7N": 100,
#     "DU-YVKQ-9Z9S": 100,
#     "SV-V7KH-CJ29": 500,
#     "MR-DZYY-GRVD": 0,
#     "49-YCT5-4G6Z": 100,
#     "EM00KB417-DO-1": 30,
#     "6E-6HM2-0FEM": 40,
#     "9Q-SV3L-TNHN": 0,
#     "7H-DSFX-5OF9": 100,
#     "EM00KB1223-DO-1": 111,
#     "DE00L03652-DO-1": 110,
#     "EM00KB1234-DO-1": 102,
#     "DE00L03687-DO-1": 100,
#     "DE00L03692-DO-1": 104,
#     "EM00KB1252-PO-1": 103,
#     "EM00KB1255-DO-1": 115,
#     "DE00L03735-DO-1": 102,
#     "EM00KB1261-YO-1": 105,
#     "EM00KB1263-DO-1": 109,
#     "EM00KB1262-DO-1": 107,
#     "EM00KB1264-DO-1": 111,
#     "EM00KB1265-DO-1": 113,
#     "DE00L03607-RO-1": 119,
#     "EM00KB1208-DO-1": 103,
#     "EM00KB1214-DO-1": 100,
#     "DE00L03638-DO-1": 109
# }

# duplicate_mapping = {
#     "Toothbrush Holder-01": ["TBH-001-DUP", "TBH-002-DUP"],
#     "Ear Wax Removal Tool 001": ["EWR-001-DUP", "EWR-002-DUP", "EWR-003-DUP"],
#     "Bathmat-grey-60x40 001": ["BMG-001-DUP", "BMG-002-DUP"],
#     "Screwdriver-01": ["SD01-DUP-1", "SD01-DUP-2", "SD01-DUP-3"],
#     "Door Bottom Stripping Seal Strip 001": ["DBS-001-DUP", "DBS-002-DUP"],
#     "VT-PXM1-HKUS": ["VT-PXM1-HKUS-DUP1", "VT-PXM1-HKUS-DUP2"],
#     "3 In 1 Car Polish 001": ["CP-001-DUP1", "CP-001-DUP2"],
#     "LED Motion Sensor Night Light 001": ["LED-001-DUP1", "LED-001-DUP2"],
#     "Ice Roller Face Massager 001": ["IRM-001-DUP1", "IRM-001-DUP2"],
#     "DO-O5AV-RSVC": ["DO-O5AV-RSVC-DUP1", "DO-O5AV-RSVC-DUP2"],
#     "6mm Heat Insulated Cover 001": ["HIC-001-DUP1", "HIC-001-DUP2"],
#     "Microfibre Cleaning Cloth 001": ["MCC-001-DUP1", "MCC-001-DUP2", "MCC-001-DUP3"],
#     "Polka Dot Fabric Swim Cap 001": ["PSC-001-DUP1", "PSC-001-DUP2"],
#     "DE00L03532-DO-1": ["DE00L03532-DUP1", "DE00L03532-DUP2"],
#     "1": ["SKU1-DUP1", "SKU1-DUP2"],
#     "EM00KB1166-DO-1": ["EM1166-DUP1", "EM1166-DUP2"],
#     "Window Squeegee Cleaner-01": ["WSC-001-DUP1", "WSC-001-DUP2"],
#     "Fruit Peeler Set-01": ["FPS-001-DUP1", "FPS-001-DUP2"],
#     "Kitchen Cleaner Spray-01": ["KCS-001-DUP1", "KCS-001-DUP2", "KCS-001-DUP3"],
#     "Car Duster Mop-01": ["CDM-001-DUP1", "CDM-001-DUP2", "CDM-001-DUP3"],
#     "Manual Push Chopper-01": ["MPC-001-DUP1", "MPC-001-DUP2"],
#     "Screwdriver-02": ["SD02-DUP1", "SD02-DUP2"],
#     "Non-Slip Bath Mat-01": ["NSBM-001-DUP1", "NSBM-001-DUP2"],
#     "Ice Roller for Face Massager-01": ["IRFM-001-DUP1", "IRFM-001-DUP2"],
#     "Stainless Steel Y-Peeler-01": ["SSYP-001-DUP1", "SSYP-001-DUP2", "SSYP-001-DUP3"],
#     "EM00KB1198-DO-1": ["EM1198-DUP1", "EM1198-DUP2"],
#     "USLA001-DF-1": ["USLA001-DUP1", "USLA001-DUP2"],
#     "SDCA01-DF-1": ["SDCA01-DUP1", "SDCA01-DUP2"],
#     "EM00KB1183-DO-1": ["EM1183-DUP1", "EM1183-DUP2"],
#     "SP012-DF-1": ["SP012-DUP1", "SP012-DUP2"],
#     "EM00KB1179-DO-1": ["EM1179-DUP1", "EM1179-DUP2"],
#     "EM00KB1197-DO-1": ["EM1197-DUP1", "EM1197-DUP2"],
#     "DE00L03564-DO-1": ["DE00L03564-DUP1", "DE00L03564-DUP2"],
#     "EM00KB1167-DO-1": ["EM1167-DUP1", "EM1167-DUP2"],
#     "EM00KB1180-DO-1": ["EM1180-DUP1", "EM1180-DUP2"],
#     "EM00KB1187-DO-1": ["EM1187-DUP1", "EM1187-DUP2"],
#     "EM00KB1189-DO-1": ["EM1189-DUP1", "EM1189-DUP2"],
#     "EM00KB978-DO-1": ["EM978-DUP1", "EM978-DUP2"],
#     "FK-X87T-GWQW": ["FKX87T-DUP1", "FKX87T-DUP2"],
#     "F2-8ZV2-ZOVT": ["F28ZV2-DUP1", "F28ZV2-DUP2"],
#     "7J-75D9-DIHY": ["7J75D9-DUP1", "7J75D9-DUP2"],
#     "NB-PHF3-DMC5": ["NBPHF3-DUP1", "NBPHF3-DUP2"],
#     "3B-ZTVL-KYXE": ["3BZTVL-DUP1", "3BZTVL-DUP2"],
#     "LW-A2IW-HSC8": ["LWA2IW-DUP1", "LWA2IW-DUP2"],
#     "RI-A7BZ-Y0IM": ["RIA7BZ-DUP1", "RIA7BZ-DUP2"],
#     "R7-GEKP-T5DR": ["R7GEKP-DUP1", "R7GEKP-DUP2"],
#     "54-5ZN0-KBHA": ["545ZN0-DUP1", "545ZN0-DUP2"],
#     "EM00KB498-DO-1": ["EM498-DUP1", "EM498-DUP2"],
#     "S4-O1CB-2P7N": ["S4O1CB-DUP1", "S4O1CB-DUP2"],
#     "DU-YVKQ-9Z9S": ["DUYVKQ-DUP1", "DUYVKQ-DUP2"],
#     "SV-V7KH-CJ29": ["SVV7KH-DUP1", "SVV7KH-DUP2"],
#     "MR-DZYY-GRVD": ["MRDZYY-DUP1", "MRDZYY-DUP2"],
#     "49-YCT5-4G6Z": ["49YCT5-DUP1", "49YCT5-DUP2"],
#     "EM00KB417-DO-1": ["EM417-DUP1", "EM417-DUP2"],
#     "6E-6HM2-0FEM": ["6E6HM2-DUP1", "6E6HM2-DUP2"],
#     "9Q-SV3L-TNHN": ["9QSV3L-DUP1", "9QSV3L-DUP2"],
#     "7H-DSFX-5OF9": ["7HDSFX-DUP1", "7HDSFX-DUP2"],
#     "EM00KB1223-DO-1": ["EM1223-DUP1", "EM1223-DUP2"],
#     "DE00L03652-DO-1": ["DE03652-DUP1", "DE03652-DUP2"],
#     "EM00KB1234-DO-1": ["EM1234-DUP1", "EM1234-DUP2"],
#     "DE00L03687-DO-1": ["DE03687-DUP1", "DE03687-DUP2"],
#     "DE00L03692-DO-1": ["DE03692-DUP1", "DE03692-DUP2"],
#     "EM00KB1252-PO-1": ["EM1252-DUP1", "EM1252-DUP2"],
#     "EM00KB1255-DO-1": ["EM1255-DUP1", "EM1255-DUP2"],
#     "DE00L03735-DO-1": ["DE03735-DUP1", "DE03735-DUP2"],
#     "EM00KB1261-YO-1": ["EM1261-DUP1", "EM1261-DUP2"],
#     "EM00KB1263-DO-1": ["EM1263-DUP1", "EM1263-DUP2"],
#     "EM00KB1262-DO-1": ["EM1262-DUP1", "EM1262-DUP2"],
#     "EM00KB1264-DO-1": ["EM1264-DUP1", "EM1264-DUP2"],
#     "EM00KB1265-DO-1": ["EM1265-DUP1", "EM1265-DUP2"],
#     "DE00L03607-RO-1": ["DE03607-DUP1", "DE03607-DUP2"],
#     "EM00KB1208-DO-1": ["EM1208-DUP1", "EM1208-DUP2"],
#     "EM00KB1214-DO-1": ["EM1214-DUP1", "EM1214-DUP2"],
#     "DE00L03638-DO-1": ["DE03638-DUP1", "DE03638-DUP2"]
# }



    




