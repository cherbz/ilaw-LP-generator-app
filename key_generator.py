import uuid
import json
import firebase_admin
from firebase_admin import credentials, firestore

# Initialize Firebase with your local credentials file
cred = credentials.Certificate("serviceAccountKey.json")
firebase_admin.initialize_app(cred)
db = firestore.client()

def generate_key():
    # Format: ILAW-XXXX-XXXX-XXXX
    raw_id = uuid.uuid4().hex.upper()
    return f"ILAW-{raw_id[:4]}-{raw_id[4:8]}-{raw_id[8:12]}"

def generate_and_upload_keys(count=1000):
    keys_list = []
    batch = db.batch()
    batch_counter = 0

    print(f"Generating and uploading {count} unique license keys to Firebase...")

    for i in range(count):
        key = generate_key()
        keys_list.append(key)
        
        # Firestore document reference
        key_ref = db.collection("license_keys").document(key)
        
        # Document data structure
        key_data = {
            "key": key,
            "is_used": False,
            "used_by": None,
            "duration_days": 365
        }
        
        batch.set(key_ref, key_data)
        batch_counter += 1

        # Firestore allows up to 500 operations per batch commit
        if batch_counter == 500:
            batch.commit()
            print("Successfully uploaded batch 1 (500 keys)...")
            batch = db.batch()
            batch_counter = 0

    if batch_counter > 0:
        batch.commit()
        print(f"Successfully uploaded batch 2 ({batch_counter} keys)...")

    # Save a local backup file so you can sell/distribute these keys
    with open("license_keys_backup.json", "w") as f:
        json.dump(keys_list, f, indent=4)

    print(f"\n✅ All {count} license keys uploaded to Firebase Firestore!")
    print("📁 Saved local backup list to 'license_keys_backup.json'.")

if __name__ == "__main__":
    generate_and_upload_keys(1000)