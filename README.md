# LocalizeXpert

This tool offers an automated method to replace translations in .mo files according to the order established in an Excel file.

## Installation

1. Clone this repository to your local machine:

    ```
    git clone https://github.com/NN0414/LocalizeXpert.git
    ```

2. Install the required Python packages:

    ```
    pip install polib
    pip install openpyxl
    ```

## Usage

### 1. Prepare Excel File

Prepare an Excel file with the following structure:

| Original Text | Context | Translated Text |
|---------------|---------|-----------------|
| Hello         | Greeting| 你好             |
| ...

### 2. Convert .mo to .po

```python
mo_to_po(mo_file_path, po_file_path)
```

Converts a .mo file to a .po file.

### 3. Apply Translations from Excel to .po

```python
apply_translations(excel_file_path, po_file_path)
```

Reads translations from an Excel file and applies them to the corresponding entries in a .po file.

### 4. Convert .po to .mo

```python
po_to_mo(po_file_path, mo_file_path)
```

Compiles a .po file back into a .mo file.

## Example

```python
mo_file_path = 'global.mo'
excel_file_path = 'translations.xlsx'
output_po_file_path = 'global.po'
output_mo_file_path = 'global_new.mo'

# Before conversion, remove existing files
remove_existing_files(output_po_file_path, output_mo_file_path)

# 1. Convert .mo to .po
mo_to_po(mo_file_path, output_po_file_path)

# 2. Apply translations from Excel to .po
apply_translations(excel_file_path, output_po_file_path)

# 3. Convert .po back to .mo
po_to_mo(output_po_file_path, output_mo_file_path)
```
