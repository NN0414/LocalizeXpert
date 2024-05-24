import polib
import openpyxl
import os

def mo_to_po(mo_file_path, po_file_path):
    print(f"Reverse {mo_file_path} to {po_file_path}")
    mo = polib.mofile(mo_file_path)
    mo.save_as_pofile(po_file_path)
    print("Reverse Complete.")

def po_to_mo(po_file_path, mo_file_path):
    print(f"Compiling {mo_file_path} from {po_file_path}")
    po = polib.pofile(po_file_path)
    po.save_as_mofile(mo_file_path)
    print("Compiling Complete.")

def build_translation_dict(excel_file):
    # Read the Excel file
    workbook = openpyxl.load_workbook(excel_file)

    translation_dict = {}

    print("Building translation replacement associations...")

    # Iterate over all worksheets and collect original and translated text pairs
    all_texts = (
        (original_text_cell.value, translated_text_cell.value)
        for sheet_name in workbook.sheetnames
        for original_text_cell, translated_text_cell in zip(
            workbook[sheet_name]['A'], workbook[sheet_name]['C']
        )
        if original_text_cell.value is not None and translated_text_cell.value is not None
    )

    # Build the translation dictionary directly from the worksheet data
    translation_dict.update(all_texts)

    return translation_dict

def apply_translations(excel_file, po_file):
    # Build the dictionary of original text to translated text
    translation_dict = build_translation_dict(excel_file)

    # Convert translation_dict to a set for faster lookups
    translation_set = set(translation_dict.keys())

    # Read the .po file
    po = polib.pofile(po_file)

    print("The following IDs have corrected translations:")

    for entry in po:
        # If a corresponding translation is found in the dictionary, replace the translation
        if entry.msgid in translation_set:
            print(entry.msgid)

            if entry.msgid_plural:
                # For msgid_plural, replace all msgstrs
                for idx, msgstr in enumerate(entry.msgstr_plural.values()):
                    entry.msgstr_plural[idx] = translation_dict[entry.msgid]
            else:
                entry.msgstr = translation_dict[entry.msgid]

            # Remove the processed entry from the set
            translation_set.remove(entry.msgid)

        # Check if the translation set is empty
        if not translation_set:
            break

    # Save the modified .po file
    po.save(po_file)

    # Print completion message
    print("Translations replacement Complete.")

def remove_existing_files(po_file_path, mo_file_path):
    if os.path.exists(po_file_path):
        os.remove(po_file_path)
        print(f"Removed existing .po file: {po_file_path}")

    if os.path.exists(mo_file_path):
        os.remove(mo_file_path)
        print(f"Removed existing .mo file: {mo_file_path}")

# Specify file paths
mo_file_path = 'global.mo'
excel_file_path = 'example.xlsx'
output_po_file_path = 'global.po'
output_mo_file_path = 'global_new.mo'

# Before conversion, remove existing files
remove_existing_files(output_po_file_path, output_mo_file_path)

# 1. Convert .mo to .po
mo_to_po(mo_file_path, output_po_file_path)

# 2. Read data from Excel file and replace translations in .po
apply_translations(excel_file_path, output_po_file_path)

# 3. Convert .po back to .mo
po_to_mo(output_po_file_path, output_mo_file_path)