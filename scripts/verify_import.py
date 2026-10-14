import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, '..'))
if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

import dataverse as dv

client = dv.DataverseClient()

print("Checking Dataverse LOA Monitoring records...")

provinces = [
    'ZAMBALES', 'PANGASINAN', 'CAMARINES SUR', 'CAVITE',
    'BATANGAS', 'NUEVA ECIJA', 'TARLAC', 'QUEZON', 'OCCIDENTAL MINDORO',
    'BULACAN', 'LEYTE', 'MISAMIS ORIENTAL', 'PALAWAN', 'SOUTH COTABATO',
    'NUEVA ECIJA - NE2 NAGSAAG', 'NUEVA ECIJA - CABANATUAN ROW', 'TARLAC ROW', 'SARANGGANI'
]

total_counted = 0
print(f"{'Province':<22} | {'Count':<8} | Sample Record")
print("-" * 75)

for prov in provinces:
    # FetchXML count per province
    fetch_xml = f"<fetch aggregate='true'><entity name='cr63f_loamonitoring'><attribute name='cr63f_loamonitoringid' aggregate='count' alias='record_count'/><filter><condition attribute='cr63f_province' operator='eq' value='{prov}'/></filter></entity></fetch>"
    res = client.get('cr63f_loamonitorings', params={'fetchXml': fetch_xml})
    items = res.get('value', [])
    count = items[0].get('record_count', 0) if items else 0
    total_counted += count

    # Sample record
    s_res = client.get(
        'cr63f_loamonitorings',
        params={
            '$filter': f"cr63f_province eq '{prov}'",
            '$select': "cr63f_control1,cr63f_municipality,cr63f_registeredowner,cr63f_totalcontractprice,cr63f_paid,cr63f_balance",
            '$top': 1
        }
    )
    s_items = s_res.get('value', [])
    sample_str = ""
    if s_items:
        rec = s_items[0]
        sample_str = f"{rec.get('cr63f_control1')} ({rec.get('cr63f_municipality')}) - Owner: {rec.get('cr63f_registeredowner')}"

    print(f"{prov:<22} | {count:<8,} | {sample_str}")

print("-" * 75)
print(f"{'Total Counted':<22} | {total_counted:<8,}")
