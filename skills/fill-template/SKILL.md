---
name: fill-template
description: Amazon product listing template filler for batch uploading products. Activates when user mentions filling Amazon template, uploading products, linking uploads, product publishing, or similar tasks involving Amazon seller template operations.
---

# Fill Template Skill

Fills Amazon batch upload templates with product data.

## Instructions

1. First, ask user to select a product category (e.g., Wall Art - Paintings)
2. Load the corresponding category config file
3. Ask user for file paths: SKU application form, listing content table, target template
4. Ask user for the listing title to search in the listing content table
5. Read product data from SKU application form using configured column indices
6. Read listing content from listing table by finding the user-specified title
7. Fill the template with combined data
8. Ask user to confirm copying values to Template sheet

## Usage

Run `python skills/fill-template/fill_template-General.py`

## Configuration

- `config/base.yaml` - Base column mappings and search logic
- `config/wall_art_paintings.yaml` - Wall Art category fixed values
