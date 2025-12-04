import json
import re

# -----------------------------------------------------
# INPUT JSON FILES
# -----------------------------------------------------
BILLS_FILE = "bills.json"
PARTY_FILE = "party_data.json"
PRODUCT_FILE = "product_data.json"

# -----------------------------------------------------
# OUTPUT JSON (FINAL)
# -----------------------------------------------------
OUTPUT = "final_firebase_ready.json"


# =====================================================
# Helper: Load JSON
# =====================================================
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# =====================================================
# Fix-1: Convert arrays → dict and remove null indexes
# =====================================================
def array_to_object(arr):
    result = {}
    for idx, item in enumerate(arr):
        if item is None:
            continue
        result[str(idx)] = fix_structure(item)
    return result


# =====================================================
# Fix-2: Sanitize KEYS (Firebase safe)
# =====================================================
def sanitize_key(key):
    key = str(key).strip()

    key = key.replace('"', '')      # remove double quotes
    key = key.replace("/", "_")     # replace slash
    key = key.replace(" ", "_")     # replace spaces
    key = key.replace("\\", "_")    # replace backslash

    key = re.sub(r'[^A-Za-z0-9_]', '', key)  # remove all special chars

    if key == "":
        key = "UNKNOWN_KEY"

    return key


# =====================================================
# Fix-3: Recursively fix values and keys
# =====================================================
def fix_structure(data):
    if isinstance(data, list):
        return array_to_object(data)

    if isinstance(data, dict):
        new_dict = {}
        for k, v in data.items():
            safe_key = sanitize_key(k)
            new_dict[safe_key] = fix_structure(v)
        return new_dict

    if isinstance(data, str):
        return data.replace("\\", "/")  # Fix file paths for Firebase

    return data


# =====================================================
# LOAD all 3 JSON files
# =====================================================
bills_data = load_json(BILLS_FILE)
party_data = load_json(PARTY_FILE)
product_data = load_json(PRODUCT_FILE)


# =====================================================
# FIX structures and sanitize everything
# =====================================================
final_json = {
    "bills": fix_structure(bills_data),
    "party_data": fix_structure(party_data),
    "product_data": fix_structure(product_data)
}


# =====================================================
# SAVE final Firebase-compatible JSON
# =====================================================
with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump(final_json, f, indent=2, ensure_ascii=False)

print("✔ FINAL Firebase-ready JSON created successfully!")
print("➡ OUTPUT FILE:", OUTPUT)
