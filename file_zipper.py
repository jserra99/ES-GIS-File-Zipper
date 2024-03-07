from zipfile import ZipFile
import os

def main() -> None:
    directory = str(input("Copy/Paste Directory Path and press Enter: "))
    os.makedirs(os.path.join(directory, "Zipped"), exist_ok=True)
    
    for root, directories, files, in os.walk(directory):
        for folder in directories:
            if folder != "Zipped":
                folder_path = os.path.join(root, folder)
                with ZipFile(os.path.join(root, f"Zipped/{folder}.zip"), 'w') as zip:
                    for kRoot, kDirectories, kFiles in os.walk(folder_path):
                        for filename in kFiles:
                            zip.write(os.path.join(kRoot, filename), arcname=filename)
                print(f"Folder {folder} Zipped")
    print("All directories zipped.")

if __name__ == "__main__":
    main()