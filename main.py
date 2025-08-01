import os, csv
import openpyxl
from typing import List


def main():
    base_dir = os.getcwd()
    base_folder_name = os.path.basename(base_dir)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"

    csv_files = list()

    for item in os.listdir(base_dir):
        item_path = os.path.join(base_dir, item)
        if os.path.isfile(item_path) and item.lower().endswith(".csv"):
            csv_files.append((base_folder_name, item))

    for item in os.listdir(base_dir):
        item_path = os.path.join(base_dir, item)
        if os.path.isdir(item_path):
            try:
                for sub_item in os.listdir(item_path):
                    sub_item_path = os.path.join(item_path, sub_item)
                    if os.path.isfile(sub_item_path) and sub_item.lower().endswith(".csv"):
                        csv_files.append((item, sub_item))
            except PermissionError:
                continue

    for col_idx, (folder_name, filename) in enumerate(csv_files, 1):
        file_path = os.path.join(base_dir, folder_name, filename) if folder_name != base_folder_name else os.path.join(
            base_dir, filename)

        ws.cell(row=1, column=col_idx, value=folder_name)

        file_base = os.path.splitext(filename)[0]
        words = file_base.split()
        last_three = " ".join(words[-3:]) if len(words) >= 3 else file_base
        ws.cell(row=2, column=col_idx, value=last_three)

        institution_text = ""
        try:
            with open(file_path, "r", encoding="UTF-8") as f:
                content = f.read()
                start_marker = "; - "
                end_marker = "; Email - "
                start_idx = content.find(start_marker)
                if start_idx != -1:
                    start_idx += len(start_marker)
                    end_idx = content.find(end_marker, start_idx)
                    if end_idx != -1:
                        institution_text = content[start_idx:end_idx]
                    else:
                        institution_text = "MARKER_ERROR: end_marker not found"
                else:
                    institution_text = "MARKER_ERROR: start_marker not found"
        except Exception as e:
            institution_text = f"ERROR: {str(e)}"

        ws.cell(row=3, column=col_idx, value=institution_text)

        data_rows = []
        try:
            with open(file_path, "r", encoding="UTF-8") as f:
                reader = csv.reader(f)
                lines = [line for line in reader if any(field.strip() for field in line)][1:114]

            for line in lines:
                if len(line) > 0:
                    last_field = line[-1].strip()
                    if last_field:
                        data_rows.append(last_field[-2:])
                    elif len(line) > 1:
                        second_last = line[-2].strip()
                        if second_last:
                            data_rows.append(second_last[-2:])
                if len(data_rows) >= 101:
                    break
        except Exception:
            pass

        for row_idx, value in enumerate(data_rows[:101], 4):
            ws.cell(row=row_idx, column=col_idx, value=value)

    output_path = os.path.join(base_dir, "result.txt")
    wb.save(output_path)
    print(f"Файл успешно создан: {output_path}")


if __name__ == "__main__":
    main()
