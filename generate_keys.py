import csv
import uuid
import firebase_admin
from firebase_admin import credentials, firestore

# 1. Connect to Firebase using your service account key
cred = credentials.Certificate("serviceAccountKey.json")
firebase_admin.initialize_app(cred)
db = firestore.client()

keys_list = []
batch = db.batch()
keys_ref = db.collection("license_keys")

print("Generating 2,000 unique license keys...")

# 2. Generate 2,000 keys and format them
for i in range(1, 2001):
    raw = str(uuid.uuid4()).upper().split("-")
    key_code = f"ILAW-{raw[0]}-{raw[1]}-{raw[2]}"
    
    keys_list.append([i, key_code, "Unused", ""])

    doc_ref = keys_ref.document(key_code)
    batch.set(doc_ref, {
        "is_used": False,
        "used_by": None,
        "used_at": None
    })

    # Firestore batch limit is 500 documents
    if i % 500 == 0:
        batch.commit()
        batch = db.batch()
        print(f"Uploaded {i}/2000 keys to Firestore...")

batch.commit()

# 3. Export to a CSV file so you can send keys to your clients
with open("license_keys.csv", mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(["ID", "License Key", "Status", "Assigned To Email"])
    writer.writerows(keys_list)

print("Done! All 2,000 keys are saved in 'license_keys.csv' and uploaded to Firebase.")