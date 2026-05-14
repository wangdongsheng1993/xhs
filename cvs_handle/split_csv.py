import csv
import os

input_file = "xhs_export_excel20265_蒸烤一体机-4.04-4.16-1913_831751206_1778659581607.csv"
output_dir = "split_csv_output"
chunk_size = 500

os.makedirs(output_dir, exist_ok=True)

with open(input_file, 'r', encoding='utf-8-sig') as f:
    reader = csv.reader(f)
    header = next(reader)

    file_num = 1
    row_count = 0
    output_file = os.path.join(output_dir, f"part_{file_num}.csv")

    out_f = open(output_file, 'w', encoding='utf-8-sig', newline='')
    writer = csv.writer(out_f)
    writer.writerow(header)

    for row in reader:
        writer.writerow(row)
        row_count += 1

        if row_count >= chunk_size:
            out_f.close()
            print(f"Written {output_file}: {row_count} data rows")
            file_num += 1
            row_count = 0
            output_file = os.path.join(output_dir, f"part_{file_num}.csv")
            out_f = open(output_file, 'w', encoding='utf-8-sig', newline='')
            writer = csv.writer(out_f)
            writer.writerow(header)

    # Close last file
    if row_count > 0:
        out_f.close()
        print(f"Written {output_file}: {row_count} data rows")

print(f"\nTotal: {file_num} files created")