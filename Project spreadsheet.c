#include<stdio.h>
#include<conio.h>
#include<string.h>
#include<stdlib.h>

#define MAX_ROWS 20
#define MAX_COLS 10
#define CELL_WIDTH 15
#define MAX_CELL_LENGTH 50
#define DATA_FILE "spreadsheet_data.dat" 

typedef struct {
    char cells[MAX_ROWS][MAX_COLS][MAX_CELL_LENGTH];
    char colHeaders[MAX_COLS][50];
    int activeRows;
    int activeCols;
} Spreadsheet;

	Spreadsheet sheet;
	
	
void printBorder(char ch, int length)
{
    for (int i = 0; i < length; i++) { 
        printf("%c", ch);
    }
    printf("\n");
}



void clearScreen(){
    system("cls");
}
	
	
	void showWelcomeScreen()
{
    clearScreen();

    printBorder('=', 70);
    printf("|                                                                    |\n");
    printf("|                EXCEL-STYLE SPREADSHEET EDITOR                      |\n");
    printf("|         A Console-Based Spreadsheet Management System              |\n");
    printf("|                                                                    |\n");
    printBorder('=', 70);
    printf("|                                                                    |\n");
    printf("|          FEATURES:                                                 |\n");
    printf("|          - Create and manage tabular data                          |\n");
    printf("|          - Enter, edit, search, and clear cell data                |\n");
    printf("|          - Customize columns with auto-save support                |\n");
    printf("|                                                                    |\n");
    printBorder('-', 70);
    printf("|                                                                    |\n");
    printf("|            Developed By : Amina Sawati & Abdullah Zahid            |\n");
    printf("|            Submitted to : Dr. Mazhar Ali                           |\n");
    printf("|                                                                    |\n");
    printBorder('=', 70);

    printf("\nPress any key to continue...");
    getch();
}



void displayMainMenu() 
{
    clearScreen();

    printBorder('=', 50);
    printf("                MAIN MENU\n");
    printBorder('=', 50);

    printf("\n");
    printf("  Sheet Info : %d Rows x %d Columns\n", sheet.activeRows, sheet.activeCols);
    printf("  Data File  : %s (Auto-Saved)\n", DATA_FILE);

    printBorder('-' , 50);

    printf("\n");
    printf("  1. View Spreadsheet\n");
    printf("  2. Write to Specific Cell\n");
    printf("  3. Quick Data Entry (Row-wise)\n");
    printf("  4. Edit Existing Cell\n");
    printf("  5. Clear Cell\n");
    printf("\n");
    printf("  6. Search Data\n");
    printf("  7. Customize Columns\n");
    printf("  8. Clear All Data\n");
    printf("\n");
    printf("  9. Export to CSV (Excel)\n");

    printf("  0. Exit Program\n");
    printBorder('=', 50);

    printf("\nEnter your choice: ");
}


	
	
//Initialize spreadsheet
void initializeSpreadsheet()
{
	for(int i = 0; i<MAX_ROWS; i++)
	{
		for(int j = 0; j<MAX_COLS; j++)
		{
			strcpy(sheet.cells[i][j], "");
		} 
	}
	
	
	//Custom Headers
	
	strcpy(sheet.colHeaders[0] , "Name");
	strcpy(sheet.colHeaders[1] , "Age");
	strcpy(sheet.colHeaders[2] , "Contact");
	strcpy(sheet.colHeaders[3] , "Address");
	strcpy(sheet.colHeaders[4] , "Email");
	sheet.activeRows = MAX_ROWS;
	sheet.activeCols = 5;
}


void displayMenu()
{
	printf("\n1. View Spreadsheet\n");
	printf("2. Write to Specific Cell (Select Row & Column)\n");
    printf("3. Quick Data Entry (Multiple Rows)\n");
    printf("4. Edit Cell\n");
    printf("5. Clear Cell\n");
    printf("6. Search Data\n");
    printf("7. Customize Columns\n");
    printf("8. Clear All Data\n");
    printf("9. Export to CSV (Excel) File\n");
    printf("10. Import from CSV (Excel) File\n");
    printf("0. Exit\n");
}

// Save to file function

void saveToFile()
 {
    FILE *file = fopen(DATA_FILE, "wb");
    if (file == NULL) {
        printf("\nError: Unable to save data to file!\n");
        return;
    }
    
    fwrite(&sheet, sizeof(Spreadsheet), 1, file);
    fclose(file);
}



// Load spreadsheet data from file
void loadFromFile() {
    FILE *file = fopen(DATA_FILE, "rb");
    if (file == NULL)
	{
        // File doesn't exist, use default initialization
        return;
    }
    
    fread(&sheet, sizeof(Spreadsheet), 1, file);
    fclose(file);
    
    printf("\n>>> Previous data loaded successfully! <<<\n");
}


void displaySpreadsheet(void)
{
    int totalWidth = 5 + sheet.activeCols * (CELL_WIDTH + 3);

    printBorder('=', totalWidth);

    printf("%-3s |", "Row");
    for (int j = 0; j < sheet.activeCols; j++) {
        printf(" %-*s |", CELL_WIDTH, sheet.colHeaders[j]);
    }
    printf("\n");

    printBorder('=', totalWidth);

    for (int i = 0; i < sheet.activeRows; i++) {
        printf("%-3d |", i + 1);
        for (int j = 0; j < sheet.activeCols; j++) {
            printf(" %-*s |", CELL_WIDTH, sheet.cells[i][j]);
        }
        printf("\n");
        printBorder('-', totalWidth);
    }
}



void clearInputBuffer(){
    int c;
    while ((c = getchar()) != '\n' && c != EOF); 
}

//Input data from scratch
void inputToCell()
{
    int row, col;
    char data[MAX_CELL_LENGTH];

    printf("--- Input data to any cell! ---\n");
    printf("Press any key to continue...\n");
    getch();

    printf("Enter number of Row (1-%d): ", sheet.activeRows);
    if (scanf("%d", &row) != 1)
    {
        clearInputBuffer();
        printf("Invalid input!\n");
        getch();
        return;
    }
    clearInputBuffer();

    if (row < 1 || row > sheet.activeRows)
    {
        printf("Invalid row number!\n");
        getch();
        return;
    }

    printf("Enter Column number (1-%d): ", sheet.activeCols);
    if (scanf("%d", &col) != 1)
    {
        clearInputBuffer();
        printf("Invalid input!\n");
        getch();
        return;
    }
    clearInputBuffer();

    if (col < 1 || col > sheet.activeCols)
    {
        printf("Invalid column number!\n");
        getch();
        return;
    }
    
    
    if (strlen(sheet.cells[row-1][col-1]) != 0) 
	{
        printf("Cell already contains data.\n");
        printf("Use EDIT option instead.\n");
        getch();
        return;
    }

    printf("Entering data to Cell [Row %d, Column %d] | %s\n",row,col,sheet.colHeaders[col-1]);
    printf("Enter new data: ");

    fgets(data, MAX_CELL_LENGTH, stdin);
    data[strcspn(data, "\n")] = '\0';

    strncpy(sheet.cells[row-1][col-1], data, MAX_CELL_LENGTH - 1);
    sheet.cells[row-1][col-1][MAX_CELL_LENGTH - 1] = '\0';
    saveToFile();
    
    printf("\nData updated successfully.\n");
    getch();
    clearScreen();
    displaySpreadsheet();
}



// edit cell function
void editCell()
{
    int row, col;
    char data[MAX_CELL_LENGTH];

    printf("--- Editiing data of any cell! ---\n");
    printf("Press any key to continue...\n");
    getch();

    printf("Enter number of Row (1-%d): ", sheet.activeRows);
    if (scanf("%d", &row) != 1)
    {
        clearInputBuffer();
        printf("Invalid input!\n");
        getch();
        return;
    }
    clearInputBuffer();

    if (row < 1 || row > sheet.activeRows)
    {
        printf("Invalid row number!\n");
        getch();
        return;
    }

    printf("Enter Column number (1-%d): ", sheet.activeCols);
    if (scanf("%d", &col) != 1)
    {
        clearInputBuffer();
        printf("Invalid input!\n");
        getch();
        return;
    }
    clearInputBuffer();

    if (col < 1 || col > sheet.activeCols)
    {
        printf("Invalid column number!\n");
        getch();
        return;
    }
    
    if (strlen(sheet.cells[row-1][col-1]) == 0) {
    printf("Cell is empty! Cannot edit an empty cell.\n");
    getch();
    return;
}


    printf("Editing data to Cell [Row %d, Column %d] | %s\n",row,col,sheet.colHeaders[col-1]);

    printf("Current value: %s\n", sheet.cells[row-1][col-1]);
    printf("Enter new data: ");

    fgets(data, MAX_CELL_LENGTH, stdin);
    data[strcspn(data, "\n")] = '\0';

    strncpy(sheet.cells[row-1][col-1], data, MAX_CELL_LENGTH - 1);
    sheet.cells[row-1][col-1][MAX_CELL_LENGTH - 1] = '\0';
    saveToFile();

    printf("\nData updated successfully.\n");
    getch();
    clearScreen();
    displaySpreadsheet();

    
  
    
}

//clear a cell
void clearCell()
{
    int row, col;
    char data[MAX_CELL_LENGTH];

    printf("--- Clear any cell! ---\n");
    printf("Press any key to continue...\n");
    getch();

    printf("Enter number of Row (1-%d): ", sheet.activeRows);
    if (scanf("%d", &row) != 1)
    {
        clearInputBuffer();
        printf("Invalid input!\n");
        getch();
        return;
    }
    clearInputBuffer();

    if (row < 1 || row > sheet.activeRows)
    {
        printf("Invalid row number!\n");
        getch();
        return;
    }

    printf("Enter Column number (1-%d): ", sheet.activeCols);
    if (scanf("%d", &col) != 1)
    {
        clearInputBuffer();
        printf("Invalid input!\n");
        getch();
        return;
    }
    clearInputBuffer();

    if (col < 1 || col > sheet.activeCols)
    {
        printf("Invalid column number!\n");
        getch();
        return;
    }
    
    
    if (strlen(sheet.cells[row-1][col-1]) == '\0') 
	{
        printf("Cell already Empty!\n");
        printf("----Use Enter to cell option instead---\n");
        getch();
        return;
    }
    else
    {
    	 strcpy(sheet.cells[row-1][col-1], "");
    	
	}
	printf("\nCell Cleared successfully.\n");
    getch();
    clearScreen();
    displaySpreadsheet();

    
    saveToFile(); //Enter thi when ready
    
}

// clear all data

void clearAlldata()
{
printBorder('=', 50);
printf("              CLEAR ALL DATA\n");
printBorder('=', 50);
printf(" \n!!! WARNING !!!\nThis action will permanently delete all data.\n");
printf("Are you sure you want to continue? (Y/N): ");

	  char confirm;
	  confirm = getche();
	  
	  if(confirm == 'y' || confirm ==  'Y')
	  {
	  	 
	  for(int i = 0; i<MAX_ROWS; i++)
	  {
	  	for(int j = 0; j<MAX_COLS; j++)
	  	{
	  		strcpy(sheet.cells[i][j], "");
		  }
	   }
	   
	   saveToFile();	   
	   printf("\nAll data Cleared successfully!\n");

	  }
	  else
	  {
	  	printf("\nOPERATION CANCELED!\n");
	  }
	  
	printf("Press any key to continue...");
	getch();
	  
}


//Quick Data Entry Function
void quickDataEntry()
{
	    // Professional header
    printBorder('=', 80);
    printf("                        QUICK DATA ENTRY MODE\n");
    printBorder('=', 80);
    printf("\n  Instructions:\n");
    printf("  1- Enter data for each field\n");
    printf("  2- Press Enter after each entry\n");
    printf("  3- Type 'DONE' to finish and return to menu\n");
    printBorder('=', 80);
    printf("\n");
	//Find which row is empty
	int currentrow = 0;
	for(int i = 0; i<sheet.activeRows; i++){
		int isEmpty = 1;
		for(int j = 0; j<sheet.activeCols; j++){
			if(strlen(sheet.cells[i][j]) > 0){
				isEmpty = 0;
				break;
			}
		}
		if(isEmpty)
		{
			currentrow = i; //This row is empty, data will added forward here
			break; 
		}
	}
	
	while(currentrow < sheet.activeRows){
		
		printf(" --- ROW %d ---\n\n", currentrow +1);
		for(int col = 0; col<sheet.activeCols; col++)
		{
			printf("%-10s-: ", sheet.colHeaders[col]);
			
			char data[MAX_CELL_LENGTH];
			fgets(data, MAX_CELL_LENGTH, stdin);
			data[strcspn(data, "\n")] = 0;
			
			if(strcmp(data, "DONE") == 0 || strcmp(data, "done") == 0)
			{
				printf("\n---Data Entry Complete---\n");
				printf("Press any key to continue---\n");
				getch();
				return;
			}
			
			strcpy(sheet.cells[currentrow][col], data);

		}
		
		saveToFile();
		currentrow++;
			printBorder('_', 50);
			printf("    Row %d saved. Continue to next row? (y/n):\n", currentrow);
			printBorder('_', 50);
	
	char choice;
	choice = getch();
	if(choice != 'y' && choice != 'Y')
	 {
		break;
     }
	}
	
			printf("\n---Data Entry Completed---\n");
		printf("Press any key to Continue---");
		getch();
		clearScreen();
		saveToFile();
		displaySpreadsheet();
}



void searchData()
{
	char search_term[MAX_CELL_LENGTH];
	printf("Search Data---\n");

	do
	{
		printf("\nEnter search term: ");
		fgets(search_term, MAX_CELL_LENGTH, stdin);
		search_term[strcspn(search_term, "\n")] = 0;

		if (strlen(search_term) == 0)
		{
			printf("\nError\n");
			printf("Press Enter to try again... ");
			getch();
		}

	} while (strlen(search_term) == 0);

	printf("Result for your search\n");

	int found = 0;
	int count = 0;

	printBorder('-', 40);

	for (int i = 0; i < sheet.activeRows; i++)
	{
		for (int j = 0; j < sheet.activeCols; j++)
		{
			if (strstr(sheet.cells[i][j], search_term) != NULL)
			{
				count++;
				printf("%-3d | %-7d | %-7d | %-10s | %-30s\n",
					   count, i + 1, j + 1, sheet.colHeaders[j], sheet.cells[i][j]);
				found = 1;
			}
		}
	}

	if (!found)
	{
		printf("Result couldn't be found!\n");
		printf("Press Enter to continue...");
		getch();
	}
	else
	{
		printf("\nTotal Matches found: %d\n", count);
	}
}




void manageColoumns()
{
	while(1){
	clearScreen();
	printBorder('=', 70);
	printf("               EXCEL-STYLE SPREADSHEET EDITOR\n");
	printf("                    COLUMN CUSTOMIZATION\n");
	printBorder('=', 70);

	
	printf("\n   Current Columns (%d Active):\n\n", sheet.activeCols);
	for (int i = 0; i < sheet.activeCols; i++)
	{
		printf("      [%d] %-20s\n", i + 1, sheet.colHeaders[i]);
	}

	
	printBorder('-', 70);
	printf("\n   MENU OPTIONS\n\n");
	printf("      1. Rename a column\n");
	printf("      2. Add new column   (Current: %d | Max: %d)\n", sheet.activeCols, MAX_COLS);
	printf("      3. Remove last column\n");
	printf("      0. Return to main menu\n");
	printBorder('-', 70);
	printf("\n   Enter your choice: ");

    
    int choice;
    if(scanf("%d", &choice) !=1)
    {
    	clearInputBuffer();
        printf("Invalid input!\n");
        printf("\n   Press any key to continue...");
        getch();
        continue;
	}
	clearInputBuffer();
	switch(choice)
	{
		case 1:
		{
		    int col_num;
		    char new_Name[20];
		
		    while (1)
		    {
		        printf("Enter column Number to rename (1 - %d): ", sheet.activeCols);
		
		        if (scanf("%d", &col_num) != 1)
		        {
		            clearInputBuffer();
		            printf("Invalid input! Please enter a number.\n");
		            continue;
		        }
		        clearInputBuffer();
		
		        if (col_num < 1 || col_num > sheet.activeCols)
		        {
		            printf("Invalid column number! Try again.\n");
		            continue;
		        }
		
		        printf("Current name: %s\n", sheet.colHeaders[col_num - 1]);
		        printf("Enter new name: ");
		        fgets(new_Name, sizeof(new_Name), stdin);
		        new_Name[strcspn(new_Name, "\n")] = 0;
		
		        if (strlen(new_Name) == 0)
		        {
		            printf("Column name cannot be empty! Try again.\n");
		            continue;
		        }
		
		        break; 
		    }

		    strcpy(sheet.colHeaders[col_num - 1], new_Name);
		    saveToFile();
		
		    printf("Column Renamed Successfully!\n");
		    printf("'%s' is now the new name for column %d.\n", new_Name, col_num);
		    printBorder('-', 50);

		    getch();
		    break;
}

			
        case 2: {
    if (sheet.activeCols >= MAX_COLS) {
        printf("\nMaximum columns reached!\n");
    } else {
        char name[50];

        sheet.activeCols++; 

        printf("Enter Column Name for column %d: ", sheet.activeCols);
        fgets(name, sizeof(name), stdin);
        name[strcspn(name, "\n")] = '\0';

        if (strlen(name) == 0) {
            sprintf(name, "Col%d", sheet.activeCols);
        }

        strcpy(sheet.colHeaders[sheet.activeCols - 1], name);

        saveToFile();
        printf("\nColumn added! Total columns: %d\n", sheet.activeCols);
    }

    printf("\nPress any key to continue...");
    getch();
    break;
}

		
		case 3:
			{
				if(sheet.activeCols <= 1)
				{
					printf("Cannot Remove: Atleast 1 column Rewquired!\n");
				}
				else
				{
				for(int i = 0; i < sheet.activeRows; i++)
				{
   					 strcpy(sheet.cells[i][sheet.activeCols - 1], "");
				}
				sheet.activeCols--;
				printf("Column Removed Successfully! Total Columns: %d\n", sheet.activeCols);
				printBorder('-', 50);

				saveToFile();
				}
				printf("\n   Press any key to continue...");
				getch();
				break;
			}
			
			case 0:
				{
					return;
				}
			
			default:
			printf("\nInvalid Choice!");
			printf("\n   Press any key to continue...");		
			getch();
       }
   }
}



void exportToCSV() {
    clearScreen();
    printBorder('=', 80);
    printf("               EXPORT DATA TO CSV (EXCEL) FILE\n");
    printBorder('=', 80);
    printf("\n");

    char filename[100];

    while (1) {
        printf("Enter filename (without extension): ");
        if (fgets(filename, sizeof(filename), stdin) == NULL) {
            printf("\nError reading filename!\n");
            continue;
        }

        filename[strcspn(filename, "\n")] = '\0';

        if (strlen(filename) == 0) {
            printf("\nError: Filename cannot be empty!\n\n");
            continue;
        }

        int onlySpaces = 1;
        for (int i = 0; filename[i] != '\0'; i++) {
            if (filename[i] != ' ') 
			{
                onlySpaces = 0;
                break;
            }
        }

        if (onlySpaces) {
            printf("\nError: Filename cannot contain only spaces!\n\n");
            continue;
        }

        break;
    }

    strcat(filename, ".csv");

    FILE* file = fopen(filename, "w");
    if (file == NULL) {
        printf("\nError: Unable to create file!\n");
        printf("\nPress any key to continue...");
        getch();
        return;
    }

    // Write headers
    for (int j = 0; j < sheet.activeCols; j++) {
        fprintf(file, "%s", sheet.colHeaders[j]);
        if (j < sheet.activeCols - 1) {
            fprintf(file, ",");
        }
    }
    fprintf(file, "\n");

    // Write data in rows
    for (int i = 0; i < sheet.activeRows; i++) {
        int hasData = 0;
        for (int j = 0; j < sheet.activeCols; j++) {
            if (strlen(sheet.cells[i][j]) > 0) {
                hasData = 1;
                break;
            }
        }
        if (!hasData)
		{
			continue;
		 }  

        for (int j = 0; j < sheet.activeCols; j++) {
            if (strchr(sheet.cells[i][j], ',') != NULL) {
                fprintf(file, "\"%s\"", sheet.cells[i][j]);
            } else {
                fprintf(file, "%s", sheet.cells[i][j]);
            }

            if (j < sheet.activeCols - 1) {
                fprintf(file, ",");
            }
        }
        fprintf(file, "\n");
    }

    fclose(file);

    printf("\nData exported successfully to '%s'!\n", filename);
    printf("You can open this file directly in Microsoft Excel.\n");
    printf("\nPress any key to continue...");
    getch();
}

int main()
{
	initializeSpreadsheet();
	loadFromFile();
	showWelcomeScreen();
	while(1){
	displayMainMenu();
	
	int choice;       
    if (scanf("%d", &choice) != 1) 
	{
        clearInputBuffer();
        printf("\nInvalid input!\n");
        printf("Press any key to continue...");
        getch();
        continue;
        }
        clearInputBuffer();
	
	switch(choice)
	{
		case 1:
			displaySpreadsheet();
			printf("\nPress any key to continue...");
            getch();
            break; 
        case 2:
        	inputToCell();
        	break;
        case 3:
        	quickDataEntry();
        	break;
        case 4:
        	editCell();
        	break;
        case 5:
			clearCell();
			break;
		case 6:
			searchData();
			break;
		case 7:
			manageColoumns();
			break;
		case 8:
			clearAlldata();
			break;
		case 9:
			exportToCSV();
			break;
		case 0:
			clearScreen();
            printBorder('=', 70);
            printf("\n             Thank you for using Spreadsheet Manager!\n");
            printf("       All your data has been saved to '%s'\n\n", DATA_FILE);
            printBorder('=', 70);
            printf("\nPress any key to exit...");
            getch();
            return 0;
        default:
			    printf("\nInvalid choice!\n");
                printf("Press any key to continue...");
                getch();    
				break;
	}
	
}

return 0;
}


