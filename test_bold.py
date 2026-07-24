"""Test <b> inside <pre> vs <code> in Telegram."""
import json, requests

with open(r'C:\Users\Admin\workspace\email_report_config.json', 'r') as f:
    config = json.load(f)

bt = config['telegram']['bot_token']
cid = config['telegram']['chat_id']

tests = [
    ("<b> in <pre>", '<pre><b>жирный</b> обычный</pre>'),
    ("<b> in <code>", '<code><b>жирный</b> обычный</code>'),
    ("<b> alone", '<b>жирный</b> обычный'),
]

for label, text in tests:
    r = requests.post(f'https://api.telegram.org/bot{bt}/sendMessage', json={
        'chat_id': cid,
        'text': text,
        'parse_mode': 'HTML'
    }, timeout=15)
    print(f'{label}: {r.status_code} {"OK" if r.status_code==200 else r.text[:60]}')
