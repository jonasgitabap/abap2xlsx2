import os
import re


def process_files_in_directory(directory_path):
    for root, dirs, files in os.walk(directory_path):
        for filename in files:
            file_path = os.path.join(root, filename)

            if filename.endswith((".abap", ".xml")) and os.path.isfile(file_path):
                with open(file_path, 'r', encoding='utf-8') as file:
                    file_content = file.read()
                    modified_content = re.sub(r'zif', '/nag/if', file_content)
                    modified_content = re.sub(r'ZIF', '/NAG/IF', modified_content)

                    modified_content = re.sub(r'zcl', '/nag/cl', modified_content)
                    modified_content = re.sub(r'ZCL', '/NAG/CL', modified_content)

                    modified_content = re.sub(r'zcx', '/nag/cx', modified_content)
                    modified_content = re.sub(r'ZCX', '/NAG/CX', modified_content)

                    modified_content = re.sub(r'zexcel', '/nag/excel', modified_content)
                    modified_content = re.sub(r'ZEXCEL', '/NAG/EXCEL', modified_content)

                    modified_content = re.sub(r'zabap', '/nag/abap', modified_content)
                    modified_content = re.sub(r'ZABAP', '/NAG/ABAP', modified_content)

                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(modified_content)

                print(f"Datei {filename} wurde bearbeitet.")


if __name__ == "__main__":
    folder_path = "/Users/jonasdenisch/abap/abap2xlsx/src/"
    process_files_in_directory(folder_path)
