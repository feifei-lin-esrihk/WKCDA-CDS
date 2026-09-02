from pathlib import Path

root_folder = Path(r"D:\Project-WKCDA-DCS-082026\04_MAP\Data\25_WKCD_180124_SLPK_A3.eslpk\nodes")
output_file = "file_list.txt"

with open(output_file, "w", encoding="utf-8") as f:
    for file in root_folder.rglob("*"):
        if file.is_file():
            f.write(str(file.relative_to(root_folder)) + "\n")

print(f"File list written to {output_file}")
