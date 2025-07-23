# 🔢 Renumber Files Module – EslamHub

Welcome to the **Renumber Files** repository by the **EslamHub** channel! 🚀  
This repository includes an Excel file to easily and quickly rename folder file numbers using **VBA**,  
along with two text files containing the **VBA code** used in the workbook.

## Contents

- 📊 **RenumberFiles.xlsm**: A ready-to-use Excel file for renaming files.
- 🧾 **Workbook_Open-Event.txt**: The `Workbook_Open` event code (inside `ThisWorkbook`).
- 🧾 **RenumberFiles-Module.txt**: The main renumbering code (a standalone module inside Excel).

## Usage Instructions

### 🔁 Auto Run

To enable automatic execution when opening the file, open the `Workbook_Open` event inside the **ThisWorkbook** window,  
or open the `Workbook_Open-Event.txt` file, and look for the line:

```vba
#If False Then
```
Then change False to True, so the line becomes:
```vba
#If True Then
```
After execution is complete, a message will appear allowing you to choose whether to open the file or close it.

### ✂️ Customize the Separator

To change the separator (such as `_`) used before numbering, open the `RenumberFiles-Module.txt` file or the code module inside Excel,  
and look for the line:

```vba
Const del$ = "_"
```Then change `_` to any separator you want, such as `-` or `.`.

### 📁 Folder Selection

When you run the tool, it will prompt you to select a folder.  
The default path is the same as the Excel file location.

### 🧮 Data Input

You will be prompted to enter the following values:

1. **Start Number**: The beginning of the numbering.
2. **End Number**: The end of the numbering.
3. **Number Format**: Number of digits (e.g., 2 → `00`, 3 → `000`).
4. **Change Amount**: Such as `+1`, `-2`, or `0` if you just want to format numbers.


## 🌐 Connect with Me

📺 [YouTube](https://www.youtube.com/@eslamhub)
📱 [TikTok](https://www.tiktok.com/@eslamhub)
📢 [LinkedIn](https://www.linkedin.com/in/eslamhub)
🐦 [X](https://x.com/eslamhub)
📘 [Facebook](https://www.facebook.com/eslamhub1)
📸 [Instagram](https://www.instagram.com/eslam.hub)

#Excel #VBA #FileRenaming #EslamHub
