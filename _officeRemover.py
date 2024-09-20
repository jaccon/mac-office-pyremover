import os
import shutil

paths_to_remove = [
    "/Applications/Microsoft Word.app",
    "/Applications/Microsoft Excel.app",
    "/Applications/Microsoft PowerPoint.app",
    "/Applications/Microsoft Outlook.app",
    "/Applications/Microsoft OneNote.app",
    os.path.expanduser("~/Library/Containers/com.microsoft.Word"),
    os.path.expanduser("~/Library/Containers/com.microsoft.Excel"),
    os.path.expanduser("~/Library/Containers/com.microsoft.Powerpoint"),
    os.path.expanduser("~/Library/Containers/com.microsoft.Outlook"),
    os.path.expanduser("~/Library/Containers/com.microsoft.onenote.mac"),
    os.path.expanduser("~/Library/Application Scripts/com.microsoft.Word"),
    os.path.expanduser("~/Library/Application Scripts/com.microsoft.Excel"),
    os.path.expanduser("~/Library/Application Scripts/com.microsoft.Powerpoint"),
    os.path.expanduser("~/Library/Application Scripts/com.microsoft.Outlook"),
    os.path.expanduser("~/Library/Application Scripts/com.microsoft.onenote.mac"),
    os.path.expanduser("~/Library/Group Containers/UBF8T346G9.Office"),
    os.path.expanduser("~/Library/Group Containers/UBF8T346G9.OneDriveSyncClientSuite"),
    os.path.expanduser("~/Library/Group Containers/UBF8T346G9.ms"),
    os.path.expanduser("~/Library/Preferences/com.microsoft.office.plist"),
    os.path.expanduser("~/Library/Caches/com.microsoft.Word"),
    os.path.expanduser("~/Library/Caches/com.microsoft.Excel"),
    os.path.expanduser("~/Library/Caches/com.microsoft.Powerpoint"),
    os.path.expanduser("~/Library/Caches/com.microsoft.Outlook"),
    os.path.expanduser("~/Library/Caches/com.microsoft.onenote.mac"),
]

def remove_office_files(paths):
    for path in paths:
        try:
            if os.path.exists(path):
                if os.path.isdir(path):
                    print(f"Removing directory: {path}")
                    shutil.rmtree(path)
                else:
                    print(f"Removing file: {path}")
                    os.remove(path)
            else:
                print(f"Path not found: {path}")
        except Exception as e:
            print(f"Error removing {path}: {e}")

if __name__ == "__main__":
    confirm = input("This will remove all Microsoft Office related files. Are you sure? (y/n): ")
    if confirm.lower() == 'y':
        remove_office_files(paths_to_remove)
        print("Microsoft Office and related files have been removed.")
    else:
        print("Operation canceled.")
