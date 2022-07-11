# xlsx-bulk-converter
This is a Python script that finds all .xlsx files in a specified directory and its subdirectories then converts them to CSV.

**This is very much a work in progress.** It does not yet account for pivot tables nor broken references.

## Usage
```python ./xlsx-convert.py [directory]```

Replace `[directory]` with the directory path of your choice; both absolute paths and paths relative to the script's location are supported.

It should be OS-agnostic, though Windows users should probably run this through PowerShell or Linux Subsystem instead of Command Prompt. On the off-chance this program generates filenames that run up against the Win32 character limit, [this article on the character limit](https://docs.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=cmd) has instructions for Windows 10 users who want to remove the limit.