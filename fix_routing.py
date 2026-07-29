#!/usr/bin/env python3
import yaml, os, shutil

config_path = os.path.expanduser('~/AppData/Local/hermes/config.yaml')
backup_path = config_path + '.bak'

# Backup
shutil.copy2(config_path, backup_path)

with open(config_path, 'r', encoding='utf-8') as f:
    data = yaml.safe_load(f)

# 1. Disable mixture (already done via hermes config set, but ensure)
data.setdefault('model', {})['mixture'] = False

# 2. Fix fallback_providers: remove ollama, fix model name, deduplicate
data['fallback_providers'] = [
    {'provider': 'gemini', 'model': 'gemini-2.5-flash'},
    {'provider': 'openai', 'model': 'gpt-4o-mini'},
    {'provider': 'openrouter', 'model': 'openrouter/free'},
]

with open(config_path, 'w', encoding='utf-8') as f:
    yaml.dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False, indent=2)

print('✓ fallback_providers fixed')
print(f'  Backup: {backup_path}')
