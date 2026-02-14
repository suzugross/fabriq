========================================
File Copy Config - Usage Guide
========================================

This module copies files and folders from the source/ directory
to specified destinations based on copy_list.csv definitions.

HOW IT WORKS
========================================

1. Place source files/folders in the "source/" directory
2. Define copy targets in "copy_list.csv"
3. Run the module from Fabriq menu
4. Review the copy list and confirm execution

CSV FORMAT (copy_list.csv)
========================================

Columns:
  Enabled   - 1=enabled, 0=disabled
  FileName  - Name of the file/folder in source/ directory
  DestPath  - Destination directory path
  Overwrite - 1=overwrite if exists, 0=skip if exists
  Description - Description (display only, not used in processing)

Example:
  1,config.ini,C:\Program Files\MyApp,1,App config file
  1,templates,C:\Users\Public\Documents,0,Template folder
  0,old_data.zip,D:\Backup,1,Disabled entry

OVERWRITE CONTROL
========================================

When Overwrite=1:
  - If the destination file/folder already exists, it will be
    overwritten with the source version.

When Overwrite=0:
  - If the destination file/folder already exists, the copy
    is skipped. This prevents accidental overwrites of
    user-modified files.

FOLDER COPY
========================================

You can copy entire folders by placing them in source/.
The module uses "Copy-Item -Recurse -Force" which copies
the folder and all its contents to the destination.

Example:
  source/
    my_folder/
      file1.txt
      file2.txt
      subfolder/
        file3.txt

  copy_list.csv:
    1,my_folder,C:\Users\Public,1,Copy entire folder

  Result: C:\Users\Public\my_folder\ (with all contents)

STATUS MARKERS
========================================

During the confirmation display:

  [Copy]      - New file (destination does not exist)
  [Overwrite] - File exists, will be overwritten (Overwrite=1)
  [Current]   - File exists, will be skipped (Overwrite=0)
  [Missing]   - Source file not found in source/ directory

DESTINATION DIRECTORY
========================================

If the destination directory does not exist, it will be
created automatically before copying.

TIPS
========================================

- Use absolute paths for DestPath (e.g., C:\Users\Public\Desktop)
- Environment variables in paths are NOT expanded automatically
- File names are case-insensitive on Windows
- Large files may take time to copy; be patient
