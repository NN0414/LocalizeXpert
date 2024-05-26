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

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Get the range of msgid, msgid_plural and msgstr columns
        msgid = sheet['A']
        msgid_plural = sheet['B']
        msgstr = sheet['D']

        for msgid_text_cell, msgid_plural_text_cell, msgstr_text_cell in zip(msgid, msgid_plural, msgstr):
            msgid_text = msgid_text_cell.value
            msgid_plural_text = msgid_plural_text_cell.value
            msgstr_text = msgstr_text_cell.value

            if msgid_plural_text is None:
                msgid_plural_text = ' '

            # Skip the line if either original or translated text is empty
            if msgid_text is None or msgstr_text is None:
                continue

            PoKey = (msgid_text, msgid_plural_text)
            translation_dict[PoKey] = msgstr_text

    return translation_dict

def apply_translations(excel_file, po_file):
    # Build dictionary of original text to translated text
    translation_dict = build_translation_dict(excel_file)

    # Read the .po file
    po = polib.pofile(po_file)

    #print("The following IDs have corrected translations:")

    for entry in po:
        PoKey = (entry.msgid, entry.msgid_plural if entry.msgid_plural else ' ')

        # If a corresponding translation is found in the dictionary, replace the translation
        if PoKey in translation_dict:
            #print(PoKey)

            if entry.msgid_plural:
                # For msgid_plural, replace all msgstrs
                for idx, msgstr in enumerate(entry.msgstr_plural.values()):
                    entry.msgstr_plural[idx] = translation_dict[PoKey]
            else:
                entry.msgstr = translation_dict[PoKey]

            # Remove the processed entry from the dictionary
            del translation_dict[PoKey]

        # Check if the translation dictionary is empty
        if not translation_dict:
            break

    # Save the modified .po file
    po.save(po_file)

    # Print completion message
    print("Apply translations Complete.")

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
