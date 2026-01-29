# Spreadsheet Manager

A professional console-based spreadsheet application written in C with Excel-compatible CSV support.

---

## About

This is a full-featured spreadsheet program that runs in your terminal. It can store data in a grid (like Excel), save it to files, and even export to CSV format that Excel can open.

**Perfect for:**
- Learning C programming
- Understanding data structures
- Practicing file operations
- Building console applications

---

## Features

**Core Functions:**
- 40 rows × 20 columns data grid
- Auto-save (your data is never lost)
- Export to CSV (open in Excel)
- Import from CSV files
- Search across all cells
- Add, edit, delete any cell
- Rename, add, or remove columns
- Quick batch data entry mode

**Technical Features:**
- Full input validation
- Binary file storage
- CSV comma handling (RFC 4180 compliant)
- Efficient memory usage
- Error handling throughout

---

## Quick Start

### Step 1: Download
Download the `spreadsheet.c` file from this repository.

### Step 2: Compile
Open your terminal and run:
```bash
gcc spreadsheet.c -o spreadsheet.exe
```

### Step 3: Run
```bash
./spreadsheet.exe
```

That's it! The program will start and show you a menu.

---

## How to Use

When you run the program, you'll see a menu with these options:

```
1. View Spreadsheet          - See your data in a table
2. Write to Cell            - Add data to a specific cell
3. Quick Data Entry         - Enter multiple rows quickly
4. Edit Cell                - Change existing data
5. Clear Cell               - Delete data from a cell
6. Search Data              - Find text in your spreadsheet
7. Customize Columns        - Rename/add/remove columns
8. Clear All Data           - Erase everything
9. Export to CSV            - Save as Excel file
10. Import from CSV         - Load Excel file
0. Exit                     - Close the program
```

### Example: Adding Data

1. Run the program
2. Choose option 2 (Write to Cell)
3. Enter row number: 1
4. Enter column number: 1
5. Type your data: "John Smith"
6. Data is saved automatically!

### Example: Exporting to Excel

1. Choose option 9 (Export to CSV)
2. Enter filename: "mydata"
3. File "mydata.csv" is created
4. Open it in Microsoft Excel!

---

## What's Inside (Technical Details)

### Data Structure
```c
typedef struct {
    char cells[40][20][50];      // Stores all your data
    char colHeaders[20][20];     // Column names
    int activeRows;              // Number of rows
    int activeCols;              // Number of columns
} Spreadsheet;
```

This structure holds:
- 40 rows × 20 columns = 800 cells
- Each cell can hold 50 characters
- Column names (like "Name", "Age", "City")
- Active row and column counts

### How It Works

**Saving Data:**
- Every change is saved to `spreadsheet_data.dat`
- Uses binary format for fast loading
- When you restart, your data is still there

**CSV Export:**
- Converts your data to CSV format
- Handles commas in your data properly
- Creates Excel-compatible files

**CSV Import:**
- Reads CSV files line by line
- Splits data by commas
- Loads into the spreadsheet

---

## Configuration

You can change these settings in the code:

```c
#define MAX_ROWS 40              // Change to 100 for more rows
#define MAX_COLS 20              // Change to 30 for more columns
#define MAX_CELL_LENGTH 50       // Change to 100 for longer text
#define DATA_FILE "spreadsheet_data.dat"  // Change filename
```

After changing, recompile:
```bash
gcc spreadsheet.c -o spreadsheet.exe
```

---

## System Requirements

**Minimum:**
- Windows 7 or later
- 1 MB free RAM
- Any C compiler (GCC recommended)

**Note:** Currently Windows-only because it uses:
- `conio.h` (for keyboard input)
- `windows.h` (for cursor positioning)

---

## What It Can't Do (Yet)

- Run on Linux or Mac (Windows only)
- Calculate formulas (like =SUM or =AVG)
- Format cells (no colors or bold text)
- Multiple sheets (only one sheet)
- Undo/Redo operations

These features might be added in future versions!

---

## Learning Resources

### What You'll Learn By Reading The Code

**Beginner Topics:**
- How to use 2D and 3D arrays
- Reading and writing files
- String manipulation
- User input handling

**Intermediate Topics:**
- Data structures (structs)
- Binary file operations
- CSV parsing
- Menu-driven programs

**Advanced Topics:**
- Memory management
- Error handling
- Input validation
- Data persistence

### Code Organization

```
Total Lines: ~1000
├─ Data structures: 50 lines
├─ File operations: 200 lines
├─ User interface: 300 lines
├─ Data manipulation: 300 lines
└─ Utility functions: 150 lines
```

---

## Common Issues & Solutions

### Problem: "Cannot compile"
**Solution:** Make sure you have GCC installed:
```bash
gcc --version
```
If not installed, download from: https://gcc.gnu.org/

### Problem: "File not found" when importing CSV
**Solution:** Put the CSV file in the same folder as spreadsheet.exe

### Problem: "Data not saving"
**Solution:** Make sure you have write permissions in the folder

### Problem: "Program closes immediately"
**Solution:** Run from terminal/command prompt, not by double-clicking

---

## How to Contribute

Want to improve this project? Here's how:

**Easy Contributions:**
- Fix spelling mistakes in comments
- Add more error messages
- Improve the user interface text

**Medium Contributions:**
- Add new features (like sorting)
- Optimize existing code
- Add more validation

**Hard Contributions:**
- Make it work on Linux/Mac
- Add formula support
- Implement undo/redo

**Steps:**
1. Fork this repository
2. Make your changes
3. Test thoroughly
4. Submit a pull request

---

## File Structure

After running the program, you'll have:

```
your-folder/
├── spreadsheet.c              (Your source code)
├── spreadsheet.exe            (Compiled program)
├── spreadsheet_data.dat       (Your saved data)
└── exported_file.csv          (Any CSV you export)
```

---

## Version History

**v1.0 (Current)**
- Initial release
- All basic features working
- CSV import/export
- Auto-save functionality

**Planned for v2.0:**
- Cross-platform support
- Basic formulas
- Improved error messages
- Better CSV handling

---

## License

MIT License - Free to use, modify, and share.

This means you can:
- Use it for personal projects
- Use it for commercial projects
- Modify the code however you want
- Share it with others

Just keep the copyright notice if you redistribute.

---

## Need Help?

**Found a bug?** 
Open an issue on GitHub with:
- What you were trying to do
- What happened instead
- Steps to reproduce

**Have a question?**
Check the code comments - they explain what each part does.

**Want to learn more?**
Read the source code! It has comments explaining each function.

---

## Credits

**Built with:**
- Pure C (no external libraries)
- Standard C Library functions
- Windows Console API

**Inspired by:**
- Microsoft Excel
- Google Sheets
- LibreOffice Calc

**Created for:**
- Learning C programming
- Understanding data structures
- Practicing file operations

---

## Keywords

c-programming, spreadsheet, csv-parser, console-app, data-structures, file-io, educational-project, beginner-friendly, excel-compatible, crud-operations

---

## Final Notes

This is a learning project that demonstrates:
- How spreadsheets work internally
- How to handle file operations in C
- How to build menu-driven applications
- How to parse CSV files
- How to validate user input

**Perfect for beginners** who want to understand:
- Arrays and pointers
- File I/O
- String manipulation
- Data structures

Feel free to use this code to learn, modify it, or build upon it!

---

## Contact & Connect

**Developer:** Abdullah Zahid

**Get in Touch:**
- Email: az756462@gmail.com
- GitHub: [@AbdullahZahid44](https://github.com/AbdullahZahid44)
- LinkedIn: Abdullah Zahid (www.linkedin.com/in/abdullah-zahid-35a796306)

**Response Time:**
- Issues: Usually within 24-48 hours
- Pull Requests: Within 3-5 days
- Questions: Feel free to open a discussion!

---

## Support This Project

If you find this project helpful:

**Star the repository** - It helps others discover it!
**Share it** - Tell your friends learning C programming
**Contribute** - Submit improvements or bug fixes
**Provide feedback** - Let me know what you think

---

**Star this repo if you found it helpful!**

Made with C programming | No external dependencies | Educational purpose

---
