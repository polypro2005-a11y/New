import json
import sys

with open('models.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

for m in data['data']:
    name = m.get('id', '').lower()
    if 'nemotron' in name or ('deepseek' in name and ('v4' in name or 'flash' in name or 'free' in name)):
        print(f"ID: {m.get('id')}")
        print(f"  Name: {m.get('name', 'N/A')}")
        print(f"  Context: {m.get('context_length', 'N/A')}")
        print(f"  Pricing: {m.get('pricing', {})}")
        print(f"  Description: {m.get('description', 'N/A')[:200]}...")
        print()