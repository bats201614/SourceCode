"""
Configuration Validator for fill-template skill
Validates all YAML config files and checks required sections
"""

import yaml
from pathlib import Path
from typing import List, Tuple

def get_skill_dir() -> Path:
    """Get the skills/fill-template directory"""
    return Path(__file__).parent.parent

def validate_yaml_file(file_path: Path) -> Tuple[bool, str]:
    """Validate a single YAML file"""
    if not file_path.exists():
        return False, f"File not found: {file_path}"

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            yaml.safe_load(f)
        return True, "OK"
    except yaml.YAMLError as e:
        return False, f"YAML syntax error: {e}"
    except Exception as e:
        return False, f"Error reading file: {e}"

def check_base_yaml(config: dict) -> List[str]:
    """Check base.yaml has required sections"""
    errors = []

    required_sku_cols = ['亚马逊SKU', '型号', '规格长', '规格宽']
    for col in required_sku_cols:
        if col not in config.get('sku_columns', {}):
            errors.append(f"Missing sku_columns: {col}")

    required_listing = ['title_search', 'content_offsets']
    for key in required_listing:
        if key not in config.get('listing_search', {}):
            errors.append(f"Missing listing_search: {key}")

    required_template = ['sheet', 'header_row', 'data_start_row']
    for key in required_template:
        if key not in config.get('template', {}):
            errors.append(f"Missing template: {key}")

    return errors

def check_category_yaml(config: dict) -> List[str]:
    """Check category YAML has required sections"""
    errors = []

    required_category = ['name', 'product_type', 'item_type_keyword']
    for key in required_category:
        if key not in config.get('category', {}):
            errors.append(f"Missing category: {key}")

    required_fixed = ['core', 'product_details', 'multi_value', 'frame', 'units', 'fulfillment', 'compliance']
    for key in required_fixed:
        if key not in config.get('fixed_values', {}):
            errors.append(f"Missing fixed_values: {key}")

    # Check core fields
    core_fields = ['brand_name', 'product_id_type', 'listing_action', 'manufacturer']
    for field in core_fields:
        if field not in config.get('fixed_values', {}).get('core', {}):
            errors.append(f"Missing fixed_values.core: {field}")

    return errors

def main():
    skill_dir = get_skill_dir()
    config_dir = skill_dir / 'config'

    print("=" * 50)
    print("Fill Template - Configuration Validator")
    print("=" * 50)
    print()

    all_valid = True

    # Check base.yaml
    print("Checking base.yaml...")
    base_path = config_dir / 'base.yaml'
    valid, msg = validate_yaml_file(base_path)
    if valid:
        with open(base_path, 'r', encoding='utf-8') as f:
            base_config = yaml.safe_load(f)
        errors = check_base_yaml(base_config)
        if errors:
            for err in errors:
                print(f"  [ERROR] {err}")
                all_valid = False
        else:
            print(f"  [OK] base.yaml is valid")
    else:
        print(f"  [ERROR] {msg}")
        all_valid = False
    print()

    # Check wall_art_paintings.yaml
    print("Checking wall_art_paintings.yaml...")
    category_path = config_dir / 'wall_art_paintings.yaml'
    valid, msg = validate_yaml_file(category_path)
    if valid:
        with open(category_path, 'r', encoding='utf-8') as f:
            category_config = yaml.safe_load(f)
        errors = check_category_yaml(category_config)
        if errors:
            for err in errors:
                print(f"  [ERROR] {err}")
                all_valid = False
        else:
            print(f"  [OK] wall_art_paintings.yaml is valid")
    else:
        print(f"  [ERROR] {msg}")
        all_valid = False
    print()

    # Check main script exists
    print("Checking fill_template-General.py...")
    script_path = skill_dir / 'scripts' / 'fill_template-General.py'
    if script_path.exists():
        print("  [OK] Script found")
    else:
        print("  [ERROR] Script not found")
        all_valid = False
    print()

    # Summary
    print("=" * 50)
    if all_valid:
        print("All validations passed!")
        return 0
    else:
        print("Validation failed. Please fix the errors above.")
        return 1

if __name__ == '__main__':
    exit(main())
