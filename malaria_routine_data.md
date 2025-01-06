### 1.0 Introduction:


### 1.1 Load and combine all input files locally

```python
import pandas as pd

def combine_excel_files_with_validation(input_files, output_file):
    try:
        # Read the first file to get column names and their count as a reference
        reference_df = pd.read_excel(input_files[0], sheet_name='Sheet1')
        reference_columns = list(reference_df.columns)
        reference_column_length = len(reference_columns)

        # Validate and combine files
        dfs = []
        for file in input_files:
            df = pd.read_excel(file, sheet_name='Sheet1')
            if list(df.columns) != reference_columns:
                raise ValueError(f"Column names do not match for file: {file}")
            if len(df.columns) != reference_column_length:
                raise ValueError(f"Column count does not match for file: {file}")
            dfs.append(df)

        # Concatenate dataframes
        combined_df = pd.concat(dfs, ignore_index=True)
        combined_df.to_excel(output_file, index=False)
        print("Files combined successfully!")
    except Exception as e:
        print(f"Error: {e}")

# List of input files
input_files = [
    "/content/Bo_District_2015-2023.xls",
    "/content/Bombali_District_2015-2023.xls",
    "/content/Bonthe_District_2015-2023.xls",
    "/content/Falaba_District_2015-2023.xls",
    "/content/Kailahun_District_2015-2023.xls",
    "/content/Kambia_District_2015-2023.xls",
    "/content/Karene_District_2015-2023.xls",
    "/content/Kenema_District_2015-2023.xls",
    "/content/Moyamba_District_2015-2023.xls",
    "/content/Kono_District_2015-2023.xls",
    "/content/Port Loko_District_2015-2023.xls",
    "/content/Pujehun_District_2015-2023.xls",
    "/content/Tonkolili_District_2015-2023.xls",
    "/content/Western Area Rural_District_2015-2023.xls",
    "/content/Western Area Urban_District_2015-2023.xls",
    "/content/Koinadugu_District_2015-2023.xls"
]

# Output file
output_file = "malaria_routine_data.xlsx"

# Combine files with validation
combine_excel_files_with_validation(input_files, output_file)
```


```python
import pandas as pd
```
- imports the pandas library, which is used for data manipulation and analysis.

```python
def combine_excel_files_with_validation(input_files, output_file):
```
- defines a function named `combine_excel_files_with_validation` that takes two arguments: `input_files` (a list of file paths for input Excel files) and `output_file` (the name of the output Excel file).

```python
    try:
```
- starts a `try` block to catch any exceptions that might occur during the function's execution.

```python
        reference_df = pd.read_excel(input_files[0], sheet_name='Sheet1')
```
- reads the first Excel file in the `input_files` list into a DataFrame called `reference_df`, specifically reading the sheet named 'Sheet1'.

```python
        reference_columns = list(reference_df.columns)
```
- extracts the column names of `reference_df` as a list and assigns them to `reference_columns`.

```python
        reference_column_length = len(reference_columns)
```
- calculates the number of columns in `reference_columns` and stores it in `reference_column_length`.

```python
        dfs = []
```
- initializes an empty list `dfs` to store validated DataFrames.

```python
        for file in input_files:
```
- starts a loop to iterate through each file in the `input_files` list.

```python
            df = pd.read_excel(file, sheet_name='Sheet1')
```
- reads the current Excel file in the loop into a DataFrame named `df`, specifically reading the sheet named 'Sheet1'.

```python
            if list(df.columns) != reference_columns:
```
- checks if the column names of the current DataFrame (`df`) are not the same as the `reference_columns`.

```python
                raise ValueError(f"Column names do not match for file: {file}")
```
- raises a `ValueError` with a descriptive message if the column names do not match the reference.

```python
            if len(df.columns) != reference_column_length:
```
- checks if the number of columns in the current DataFrame (`df`) is different from the `reference_column_length`.

```python
                raise ValueError(f"Column count does not match for file: {file}")
```
- raises a `ValueError` if the column count does not match the reference.

```python
            dfs.append(df)
```
- appends the validated DataFrame (`df`) to the list `dfs`.

```python
        combined_df = pd.concat(dfs, ignore_index=True)
```
- concatenates all DataFrames in the `dfs` list into a single DataFrame named `combined_df`. The `ignore_index=True` ensures the new DataFrame has a continuous index.

```python
        combined_df.to_excel(output_file, index=False)
```
- writes the combined DataFrame to an Excel file specified by `output_file` without saving the index as a column.

```python
        print("Files combined successfully!")
```
- prints a success message if the files are combined without any issues.

```python
    except Exception as e:
```
- starts an `except` block to catch any exceptions that might occur during the `try` block.

```python
        print(f"Error: {e}")
```
- prints an error message that includes the exception details.


```python
input_files = [
    "/content/Bo_District_2015-2023.xls",
    "/content/Bombali_District_2015-2023.xls",
    "/content/Bonthe_District_2015-2023.xls",
    "/content/Falaba_District_2015-2023.xls",
    "/content/Kailahun_District_2015-2023.xls",
    "/content/Kambia_District_2015-2023.xls",
    "/content/Karene_District_2015-2023.xls",
    "/content/Kenema_District_2015-2023.xls",
    "/content/Moyamba_District_2015-2023.xls",
    "/content/Kono_District_2015-2023.xls",
    "/content/Port Loko_District_2015-2023.xls",
    "/content/Pujehun_District_2015-2023.xls",
    "/content/Tonkolili_District_2015-2023.xls",
    "/content/Western Area Rural_District_2015-2023.xls",
    "/content/Western Area Urban_District_2015-2023.xls",
    "/content/Koinadugu_District_2015-2023.xls"
]
```
- defines a list of file paths for the input Excel files to be combined.

```python
output_file = "malaria_routine_data.xlsx"
```
- specifies the name of the output Excel file.

```python
combine_excel_files_with_validation(input_files, output_file)
```
- calls the function `combine_excel_files_with_validation`, passing `input_files` and `output_file` as arguments to validate and combine the Excel files.



#### 1.1.1 User guide: what to modify

1. **File Paths**
   - Replace the file paths in the `input_files` list with the paths to your Excel files. For example:
     ```python
     input_files = [
         "/path/to/your/file1.xlsx",
         "/path/to/your/file2.xlsx",
         "/path/to/your/file3.xlsx"
     ]
     ```

2. **Sheet Name**
   - If your Excel files have a different sheet name, update the `sheet_name` parameter in both `pd.read_excel` calls. For example:
     ```python
     pd.read_excel(file, sheet_name='YourSheetName')
     ```

3. **Output File**
   - Change the name and path of the output file by modifying the `output_file` variable. For example:
     ```python
     output_file = "/path/to/combined_file.xlsx"
     ```

4. **Column Validation**
   - If you do not need column validation, you can remove or comment out the validation checks:
     ```python
     if list(df.columns) != reference_columns:
         raise ValueError(f"Column names do not match for file: {file}")

     if len(df.columns) != reference_column_length:
         raise ValueError(f"Column count does not match for file: {file}")
     ```

5. **Index Inclusion**
   - If you want to include the index in the output file, change the `index=False` parameter in `to_excel` to `index=True`. For example:
     ```python
     combined_df.to_excel(output_file, index=True)
     ```

6. **File Format**
   - If your input files are `.csv` instead of `.xlsx`, update `pd.read_excel` to `pd.read_csv` for reading the files, and change `to_excel` to `to_csv` for writing the output.

---


### 1.2 Load and combine all input files from GitHub


```python
import pandas as pd
from urllib.parse import quote

# Base raw GitHub URL
base_url = "https://raw.githubusercontent.com/mohamedsillahkanu/SNT/f4408583752429926d6b7a41e8a0052aeee93d84/files/"

# List of file names
file_names = [
    "Bo_District_2015-2023.xls",
    "Bombali_District_2015-2023.xls",
    "Bonthe_District_2015-2023.xls",
    "Falaba_District_2015-2023.xls",
    "Kailahun_District_2015-2023.xls",
    "Kambia_District_2015-2023.xls",
    "Karene_District_2015-2023.xls",
    "Kenema_District_2015-2023.xls",
    "Moyamba_District_2015-2023.xls",
    "Kono_District_2015-2023.xls",
    "Port_Loko_District_2015-2023.xls",
    "Pujehun_District_2015-2023.xls",
    "Tonkolili_District_2015-2023.xls",
    "Western_Area_Rural_District_2015-2023.xls",
    "Western_Area_Urban_District_2015-2023.xls",
    "Koinadugu_District_2015-2023.xls"
]

# Encode file names to handle spaces and special characters
file_urls = [base_url + quote(file_name) for file_name in file_names]

# Function to read and validate Excel files
def read_excel_file(url):
    try:
        print(f"Reading file: {url}")
        df = pd.read_excel(url, engine='xlrd')  # Adjust engine if necessary
        print(f"Successfully read file: {url}")
        return df
    except Exception as e:
        print(f"Error reading file {url}: {e}")
        return None

# Function to combine files with validation
def combine_excel_files_with_validation(file_urls, output_file):
    try:
        dfs = []
        reference_df = read_excel_file(file_urls[0])
        if reference_df is None:
            raise ValueError("Reference file could not be read.")

        reference_columns = list(reference_df.columns)
        reference_column_length = len(reference_columns)

        # Validate and combine
        for url in file_urls:
            df = read_excel_file(url)
            if df is None:
                continue
            if list(df.columns) != reference_columns:
                raise ValueError(f"Column names do not match for file: {url}")
            if len(df.columns) != reference_column_length:
                raise ValueError(f"Column count does not match for file: {url}")
            dfs.append(df)

        # Combine all dataframes
        combined_df = pd.concat(dfs, ignore_index=True)
        combined_df.to_excel(output_file, index=False)
        print(f"Files combined successfully! Output saved to {output_file}")
    except Exception as e:
        print(f"Error: {e}")

# Combine and save the output
output_file = "malaria_routine_data.xlsx"
combine_excel_files_with_validation(file_urls, output_file)

```

**1. Import Libraries**
```python
import pandas as pd
from urllib.parse import quote
```
- **`pandas`** is imported to handle data manipulation and Excel file operations.
- **`quote`** from `urllib.parse` is imported to encode file names for safe use in URLs.


**2. Base GitHub URL**
```python
base_url = "https://raw.githubusercontent.com/mohamedsillahkanu/SNT/f4408583752429926d6b7a41e8a0052aeee93d84/files/"
```
- The base URL points to the directory containing the Excel files in the GitHub repository.



**3. List of File Names**
```python
file_names = [
    "Bo_District_2015-2023.xls",
    "Bombali_District_2015-2023.xls",
    ...
    "Koinadugu_District_2015-2023.xls"
]
```
- A list of file names that need to be fetched from the GitHub repository.


**4. Encode File Names**
```python
file_urls = [base_url + quote(file_name) for file_name in file_names]
```
- Encodes the file names to handle spaces and special characters.
- Combines each file name with the base URL to create a list of fully qualified file URLs.



**5. Define a Function to Read Excel Files**
```python
def read_excel_file(url):
    try:
        print(f"Reading file: {url}")
        df = pd.read_excel(url, engine='xlrd')  # Adjust engine if necessary
        print(f"Successfully read file: {url}")
        return df
    except Exception as e:
        print(f"Error reading file {url}: {e}")
        return None
```
- **Purpose**: Reads an Excel file from a URL.
- **How It Works**:
  1. Prints a message indicating the file being read.
  2. Reads the file into a DataFrame using `pd.read_excel`.
  3. Returns the DataFrame if successful, or `None` if an error occurs.



**6. Define a Function to Combine Files**
```python
def combine_excel_files_with_validation(file_urls, output_file):
    try:
        dfs = []
        reference_df = read_excel_file(file_urls[0])
        if reference_df is None:
            raise ValueError("Reference file could not be read.")
```
- **Initial Setup**:
  - Creates an empty list `dfs` to store the DataFrames.
  - Reads the first file as the **reference file** and validates it is readable.


**7. Validate and Extract Reference Columns**
```python
        reference_columns = list(reference_df.columns)
        reference_column_length = len(reference_columns)
```
- Extracts the column names and the count of columns from the reference file to ensure consistency across all files.


**8. Loop Through Files and Validate**
```python
        for url in file_urls:
            df = read_excel_file(url)
            if df is None:
                continue
            if list(df.columns) != reference_columns:
                raise ValueError(f"Column names do not match for file: {url}")
            if len(df.columns) != reference_column_length:
                raise ValueError(f"Column count does not match for file: {url}")
            dfs.append(df)
```
- Loops through each file URL:
  1. Reads the file into a DataFrame.
  2. Skips the file if it cannot be read.
  3. Validates that the column names and count match the reference file.
  4. Appends the validated DataFrame to `dfs`.


**9. Combine and Save**
```python
        combined_df = pd.concat(dfs, ignore_index=True)
        combined_df.to_excel(output_file, index=False)
        print(f"Files combined successfully! Output saved to {output_file}")
```
- Combines all validated DataFrames into a single DataFrame using `pd.concat`.
- Saves the combined DataFrame to the specified `output_file` in Excel format.



**10. Handle Errors**
```python
    except Exception as e:
        print(f"Error: {e}")
```
- Catches and prints any errors encountered during the file combination process.



**11. Specify Output File and Execute**
```python
output_file = "malaria_routine_data.xlsx"
combine_excel_files_with_validation(file_urls, output_file)
```
- Sets the name of the output file.
- Calls the `combine_excel_files_with_validation` function, passing the file URLs and the output file name.


#### 1.2.1 User guide: what to modify

To adapt this code to your own requirements, make the following changes:

1. **Base URL**  
   Replace the `base_url` with the URL of your GitHub repository or the location where your Excel files are stored.  
   ```python
   base_url = "https://raw.githubusercontent.com/your-username/your-repo/branch-name/your-folder/"
   ```

2. **File Names**  
   Update the `file_names` list with the names of the files you want to combine. Ensure these file names match exactly with the files in your repository or storage location.  
   ```python
   file_names = [
       "Your_File1.xlsx",
       "Your_File2.xlsx",
       ...
   ]
   ```

3. **Sheet Name**  
   If your Excel files have a sheet name other than `'Sheet1'`, update the `sheet_name` parameter in the `read_excel_file` function:  
   ```python
   df = pd.read_excel(url, sheet_name='YourSheetName', engine='xlrd')
   ```

4. **Output File Name**  
   Change the `output_file` variable to set the desired name and location for the combined output file:  
   ```python
   output_file = "your_combined_file.xlsx"
   ```

5. **Excel Engine**  
   If your Excel files use a format incompatible with `xlrd`, use the appropriate engine (`openpyxl` for `.xlsx`, `xlrd` for older `.xls`):  
   ```python
   df = pd.read_excel(url, engine='openpyxl')  # For .xlsx files
   ```

6. **Column Validation (Optional)**  
   If column validation is not needed, comment out or remove the validation checks in the loop:  
   ```python
   # if list(df.columns) != reference_columns:
   #     raise ValueError(f"Column names do not match for file: {url}")

   # if len(df.columns) != reference_column_length:
   #     raise ValueError(f"Column count does not match for file: {url}")
   ```

---


### 1.3 Load the combined excel file

```
import pandas as pd

def read_combined_excel(file_path):
    """
    Reads a combined Excel file and loads it into a pandas DataFrame.

    Parameters:
        file_path (str): Path to the Excel file.

    Returns:
        pd.DataFrame: DataFrame containing the data from the Excel file.
    """
    try:
        # Read the Excel file
        combined_df = pd.read_excel(file_path, sheet_name=0)  # Read the first sheet
        print("Combined DataFrame loaded successfully!")
        return combined_df
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
    except Exception as e:
        print(f"Error reading the Excel file: {e}")

# Call the function
output_file = "malaria_routine_data.xlsx"
combined_data = read_combined_excel(output_file)

if combined_data is not None:
    print(combined_data.head())  # Display the first few rows

df = combined_data.copy()

```

#### 1.3.1 User guide: what to change

**1. File Path**  
```python
output_file = "malaria_routine_data.xlsx"
```
- Replace `"malaria_routine_data.xlsx"` with the path to your Excel file that you want to load.
- Example:
  ```python
  output_file = "path/to/your_file.xlsx"
  ```

**2. Sheet Name**  
```python
combined_df = pd.read_excel(file_path, sheet_name=0)
```
- Replace `sheet_name=0` if your Excel file has a different sheet name or if you want to specify a sheet by name.
- Example:
  ```python
  combined_df = pd.read_excel(file_path, sheet_name="YourSheetName")
  ```

**3. Error Handling Messages**  
```python
print(f"Error: The file '{file_path}' was not found.")
```
- Update the error messages to be more descriptive or relevant to your use case.
- Example:
  ```python
  print(f"Error: Unable to find the specified Excel file at '{file_path}'. Please check the file path.")
  ```


**4. Data Validation (Optional)**  
- If you want to perform additional checks on the loaded DataFrame (e.g., ensuring specific columns exist), add validation logic after loading the data.
- Example:
  ```python
  if "RequiredColumn" not in combined_df.columns:
      print("Error: The required column 'RequiredColumn' is missing.")
      return None
  ```


**5. Output Display**  
```python
print(combined_data.head())
```
- Modify how much of the data you want to display or the specific rows/columns to preview.
- Example:
  ```python
  print(combined_data.head(10))  # Display the first 10 rows
  print(combined_data[['Column1', 'Column2']])  # Display specific columns
  ```

**6. Copying the DataFrame (Optional)**  
```python
df = combined_data.copy()
```
- If you don’t need to create a copy of the DataFrame, you can skip this step or directly work with `combined_data`.


### 1.4 Create the date column

```
def process_dataframe_with_integer_month(df, column_to_split, drop_column):
    """
    Processes the DataFrame by splitting a column into 'month' and 'year',
    converting 'month' to integer, creating a new 'Date' column in 'YYYY-MM' format,
    and dropping specified columns.

    Args:
        df (pd.DataFrame): The DataFrame to process.
        column_to_split (str): The name of the column to split into 'month' and 'year'.
        drop_column (str): The name of the column to drop.

    Returns:
        pd.DataFrame: The processed DataFrame.
    """
    try:
        # Dictionary to map month names to integers
        month_map = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04',
            'May': '05', 'June': '06', 'July': '07', 'August': '08',
            'September': '09', 'October': '10', 'November': '11', 'December': '12'
        }

        # Split the specified column into 'month' and 'year'
        df[['month', 'year']] = df[column_to_split].str.split(' ', expand=True)

        # Map month names to integers
        df['month'] = df['month'].map(month_map)

        # Convert 'year' column to numeric
        df['year'] = pd.to_numeric(df['year'], errors='raise')

        # Create a new 'Date' column in 'YYYY-MM' format as a string
        df['Date'] = df['year'].astype(str) + '-' + df['month']

        # Drop the specified columns (both the column to split and the additional column)
        df.drop(columns=[column_to_split, drop_column], inplace=True)

        return df
    except Exception as e:
        print(f"Error processing DataFrame: {e}")
        return None



# Call the function
df = process_dataframe_with_integer_month(df, 'periodname', 'orgunitlevel5')

```


**1. Define the function `process_dataframe_with_integer_month`**  
```python
def process_dataframe_with_integer_month(df, column_to_split, drop_column):
    """
    Processes the DataFrame by splitting a column into 'month' and 'year',
    converting 'month' to integer, creating a new 'Date' column in 'YYYY-MM' format,
    and dropping specified columns.

    Args:
        df (pd.DataFrame): The DataFrame to process.
        column_to_split (str): The name of the column to split into 'month' and 'year'.
        drop_column (str): The name of the column to drop.

    Returns:
        pd.DataFrame: The processed DataFrame.
    """
```
- Defines a function named `process_dataframe_with_integer_month`.
- Includes a docstring to describe the function's purpose, input parameters, and return value.


**2. Use a `try` block to handle errors**  
```python
    try:
```
- Starts a `try` block to process the DataFrame while catching any exceptions.



**3. Define a dictionary to map month names to integers**  
```python
        month_map = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04',
            'May': '05', 'June': '06', 'July': '07', 'August': '08',
            'September': '09', 'October': '10', 'November': '11', 'December': '12'
        }
```
- Creates a dictionary (`month_map`) that maps month names (e.g., `'January'`) to their corresponding numeric values as strings (e.g., `'01'`).



**4. Split the specified column into 'month' and 'year'**  
```python
        df[['month', 'year']] = df[column_to_split].str.split(' ', expand=True)
```
- Splits the values in the `column_to_split` column into two separate columns: `'month'` and `'year'`.
- The `str.split(' ', expand=True)` method splits the string at spaces and expands the result into multiple columns.



**5. Map month names to integers**  
```python
        df['month'] = df['month'].map(month_map)
```
- Maps the values in the `'month'` column to their corresponding numeric values using the `month_map` dictionary.



**6. Convert the 'year' column to numeric**  
```python
        df['year'] = pd.to_numeric(df['year'], errors='raise')
```
- Converts the `'year'` column to numeric values using `pd.to_numeric`. 
- If the conversion fails, an error is raised.

---

**7. Create a new 'Date' column in 'YYYY-MM' format**  
```python
        df['Date'] = df['year'].astype(str) + '-' + df['month']
```
- Combines the `'year'` and `'month'` columns to create a new `'Date'` column in the format `YYYY-MM`.



**8. Drop the specified columns**  
```python
        df.drop(columns=[column_to_split, drop_column], inplace=True)
```
- Drops the original column used for splitting (`column_to_split`) and the additional column (`drop_column`) from the DataFrame.



**9. Return the processed DataFrame**  
```python
        return df
```
- Returns the modified DataFrame.



**10. Handle errors during processing**  
```python
    except Exception as e:
        print(f"Error processing DataFrame: {e}")
        return None
```
- Catches any exceptions raised during processing.
- Prints the error message and returns `None`.



**11. Call the function**  
```python
df = process_dataframe_with_integer_month(df, 'periodname', 'orgunitlevel5')
```
- Calls the `process_dataframe_with_integer_month` function.
- Passes the DataFrame (`df`) along with the column to split (`'periodname'`) and the column to drop (`'orgunitlevel5'`).
- Stores the returned processed DataFrame back in the variable `df`.


#### 1.4.1 User guide: what to modify

**1. Column to Split**  
```python
df[['month', 'year']] = df[column_to_split].str.split(' ', expand=True)
```
- Replace `column_to_split` with the name of the column in your DataFrame that contains both the month and year (e.g., `'January 2023'`).
- Example:
  ```python
  column_to_split = 'YourColumnName'
  ```

**2. Column to Drop**  
```python
df.drop(columns=[column_to_split, drop_column], inplace=True)
```
- Replace `drop_column` with the name of the column you want to remove after processing.
- Example:
  ```python
  drop_column = 'YourColumnToRemove'
  ```


**3. Month and Year Format in `column_to_split`**  
- Ensure that the `column_to_split` contains values in the format `'<MonthName> <Year>'` (e.g., `'January 2023'`).
- If the format is different, adjust the splitting logic:
  ```python
  df[['month', 'year']] = df[column_to_split].str.split('-', expand=True)  # Example for 'January-2023'
  ```


**4. Month Mapping (Optional)**  
```python
month_map = {
    'January': '01', 'February': '02', ..., 'December': '12'
}
```
- Update the `month_map` dictionary if your month names are in a different language or format.
- Example:
  ```python
  month_map = {
      'Janvier': '01', 'Février': '02', ..., 'Décembre': '12'  # For French month names
  }
  ```


**5. DataFrame Input**  
```python
df = process_dataframe_with_integer_month(df, 'periodname', 'orgunitlevel5')
```
- Replace `'periodname'` and `'orgunitlevel5'` with the actual column names in your DataFrame that correspond to the column to split and the column to drop.
- Example:
  ```python
  df = process_dataframe_with_integer_month(df, 'YourPeriodColumn', 'YourColumnToDrop')
  ```


**6. Error Handling Messages (Optional)**  
```python
print(f"Error processing DataFrame: {e}")
```
- Update the error message to provide more specific feedback.
- Example:
  ```python
  print(f"Error: Unable to process the column '{column_to_split}' or '{drop_column}': {e}")
  ```

**7. New Date Format (Optional)**  
```python
df['Date'] = df['year'].astype(str) + '-' + df['month']
```
- Modify the format of the `'Date'` column if required.
- Example:
  ```python
  df['Date'] = df['month'] + '-' + df['year'].astype(str)  # Format 'MM-YYYY'
  ``` 
---

### 1.5 Rename the columns

```
def rename_columns(df):
    """
    Renames columns in a DataFrame based on predefined dictionaries.

    Args:
        df (pd.DataFrame): The DataFrame to process.

    Returns:
        pd.DataFrame: The DataFrame with renamed columns.
    """
    try:
        # Define the mappings for column renaming
        orgunit_rename = {
            'orgunitlevel1': 'adm0',
            'orgunitlevel2': 'adm1',
            'orgunitlevel3': 'adm2',
            'orgunitlevel4': 'adm3',
            'organisationunitname': 'hf'
        }

        column_rename = {
            'OPD (New and follow-up curative) 0-59m_X': 'allout_u5',
            'OPD (New and follow-up curative) 5+y_X': 'allout_ov5',


            'Admission - Child with malaria 0-59 months_X': 'maladm_u5',
            'Admission - Child with malaria 5-14 years_X': 'maladm_5_14',
            'Admission - Malaria 15+ years_X': 'maladm_ov15',



            'Child death - Malaria 1-59m_X': 'maldth_1_59m',
            'Child death - Malaria 10-14y_X': 'maldth_10_14',
            'Child death - Malaria 5-9y_X': 'maldth_5_9',


            'Death malaria 15+ years Female': 'maldth_fem_ov15',
            'Death malaria 15+ years Male': 'maldth_mal_ov15',


            'Separation - Child with malaria 0-59 months_X Death': 'maldth_u5',
            'Separation - Child with malaria 5-14 years_X Death' : 'maldth_5_14',
            'Separation - Malaria 15+ years_X Death': 'maldth_ov15',  


            'Fever case - suspected Malaria 0-59m_X': 'susp_u5_hf',
            'Fever case - suspected Malaria 5-14y_X': 'susp_5_14_hf',
            'Fever case - suspected Malaria 15+y_X': 'susp_ov15_hf',
            'Fever case in community (Suspected Malaria) 0-59m_X': 'susp_u5_com',
            'Fever case in community (Suspected Malaria) 5-14y_X': 'susp_5_14_com',
            'Fever case in community (Suspected Malaria) 15+y_X': 'susp_ov15_com',



            'Fever case in community tested for Malaria (RDT) - Negative 0-59m_X': 'tes_neg_rdt_u5_com',
            'Fever case in community tested for Malaria (RDT) - Positive 0-59m_X': 'tes_pos_rdt_u5_com',
            'Fever case in community tested for Malaria (RDT) - Negative 5-14y_X': 'tes_neg_rdt_5_14_com',
            'Fever case in community tested for Malaria (RDT) - Positive 5-14y_X': 'tes_pos_rdt_5_14_com',
            'Fever case in community tested for Malaria (RDT) - Negative 15+y_X': 'tes_neg_rdt_ov15_com',
            'Fever case in community tested for Malaria (RDT) - Positive 15+y_X': 'tes_pos_rdt_ov15_com',
            'Fever case tested for Malaria (Microscopy) - Negative 0-59m_X': 'test_neg_mic_u5_hf',
            'Fever case tested for Malaria (Microscopy) - Positive 0-59m_X': 'test_pos_mic_u5_hf',
            'Fever case tested for Malaria (Microscopy) - Negative 5-14y_X': 'test_neg_mic_5_14_hf',
            'Fever case tested for Malaria (Microscopy) - Positive 5-14y_X': 'test_pos_mic_5_14_hf',
            'Fever case tested for Malaria (Microscopy) - Negative 15+y_X': 'test_neg_mic_ov15_hf',
            'Fever case tested for Malaria (Microscopy) - Positive 15+y_X': 'test_pos_mic_ov15_hf',
            'Fever case tested for Malaria (RDT) - Negative 0-59m_X': 'tes_neg_rdt_u5_hf',
            'Fever case tested for Malaria (RDT) - Positive 0-59m_X': 'tes_pos_rdt_u5_hf',
            'Fever case tested for Malaria (RDT) - Negative 5-14y_X': 'tes_neg_rdt_5_14_hf',
            'Fever case tested for Malaria (RDT) - Positive 5-14y_X': 'tes_pos_rdt_5_14_hf',
            'Fever case tested for Malaria (RDT) - Negative 15+y_X': 'tes_neg_rdt_ov15_hf',
            'Fever case tested for Malaria (RDT) - Positive 15+y_X': 'tes_pos_rdt_ov15_hf',



            'Malaria treated in community with ACT <24 hours 0-59m_X': 'maltreat_u24_u5_com',
            'Malaria treated in community with ACT >24 hours 0-59m_X': 'maltreat_ov24_u5_com',
            'Malaria treated in community with ACT <24 hours 5-14y_X': 'maltreat_u24_5_14_com',
            'Malaria treated in community with ACT >24 hours 5-14y_X': 'maltreat_ov24_5_14_com',
            'Malaria treated in community with ACT <24 hours 15+y_X': 'maltreat_u24_ov15_com',
            'Malaria treated in community with ACT >24 hours 15+y_X': 'maltreat_ov24_ov15_com',
            'Malaria treated with ACT <24 hours 0-59m_X': 'maltreat_u24_u5_hf',
            'Malaria treated with ACT >24 hours 0-59m_X': 'maltreat_ov24_u5_hf',
            'Malaria treated with ACT <24 hours 5-14y_X': 'maltreat_u24_5_14_hf',
            'Malaria treated with ACT >24 hours 5-14y_X': 'maltreat_ov24_5_14_hf',
            'Malaria treated with ACT <24 hours 15+y_X': 'maltreat_u24_ov15_hf',
            'Malaria treated with ACT >24 hours 15+y_X': 'maltreat_ov24_ov15_hf'

        }

        # Rename columns using the mappings
        df = df.rename(columns={**orgunit_rename, **column_rename})

        return df
    except Exception as e:
        print(f"Error renaming columns: {e}")
        return None

# call the function
df = rename_columns(df)


```

**1. Define the `rename_columns` function**  
```python
def rename_columns(df):
    """
    Renames columns in a DataFrame based on predefined dictionaries.

    Args:
        df (pd.DataFrame): The DataFrame to process.

    Returns:
        pd.DataFrame: The DataFrame with renamed columns.
    """
```
- Defines a function named `rename_columns` to rename columns in a DataFrame.
- Includes a docstring to describe the purpose, input (`df`), and output (a processed DataFrame).


**2. Use a `try` block to handle errors**  
```python
    try:
```
- Starts a `try` block to rename the columns while catching any exceptions.


**3. Define the dictionary for organizational unit renaming**  
```python
        orgunit_rename = {
            'orgunitlevel1': 'adm0',
            'orgunitlevel2': 'adm1',
            'orgunitlevel3': 'adm2',
            'orgunitlevel4': 'adm3',
            'organisationunitname': 'hf'
        }
```
- Defines a dictionary (`orgunit_rename`) that maps old column names (e.g., `'orgunitlevel1'`) to new column names (e.g., `'adm0'`) for organizational unit levels.


**4. Define the dictionary for other column renaming**  
```python
        column_rename = {
            'OPD (New and follow-up curative) 0-59m_X': 'allout_u5',
            'OPD (New and follow-up curative) 5+y_X': 'allout_ov5',
            ...
        }
```
- Defines a dictionary (`column_rename`) that maps additional old column names to more concise, descriptive new names.
- The mappings cover a wide range of columns for outpatient data, malaria admissions, deaths, testing, and treatments.


**5. Combine the dictionaries and rename columns**  
```python
        df = df.rename(columns={**orgunit_rename, **column_rename})
```
- Combines the `orgunit_rename` and `column_rename` dictionaries using `{**dict1, **dict2}` syntax.
- Uses `df.rename(columns=...)` to rename the columns in the DataFrame based on the combined mappings.


**6. Return the updated DataFrame**  
```python
        return df
```
- Returns the DataFrame with the renamed columns.


**7. Handle exceptions during renaming**  
```python
    except Exception as e:
        print(f"Error renaming columns: {e}")
        return None
```
- Catches and handles any exceptions that occur during the column renaming process.
- Prints an error message and returns `None`.


**8. Call the `rename_columns` function**  
```python
df = rename_columns(df)
```
- Calls the `rename_columns` function and passes the DataFrame (`df`) to it.
- Stores the processed DataFrame with renamed columns back into the variable `df`.

---


#### User guide: what to modify

**1. Input DataFrame**  
```python
df = rename_columns(df)
```
- Replace `df` with your actual DataFrame variable name.
- Example:
  ```python
  df = rename_columns(your_dataframe)
  ```



**2. Organizational Unit Column Names**  
```python
orgunit_rename = {
    'orgunitlevel1': 'adm0',
    'orgunitlevel2': 'adm1',
    'orgunitlevel3': 'adm2',
    'orgunitlevel4': 'adm3',
    'organisationunitname': 'hf'
}
```
- Update the keys in the `orgunit_rename` dictionary if the column names in your dataset differ.

```


**3. Other Column Names**  
```python
column_rename = {
    'OPD (New and follow-up curative) 0-59m_X': 'allout_u5',
    'OPD (New and follow-up curative) 5+y_X': 'allout_ov5',
    ...
}
```
- Replace the keys in the `column_rename` dictionary to match your dataset's column names.


**4. Error Messages (Optional)**  
```python
print(f"Error renaming columns: {e}")
```
- Update the error message for clarity or to provide more specific feedback.
- Example:
  ```python
  print(f"Error: Could not rename columns. Check the column names in your dataset. {e}")
  ```


**5. Combined Dictionary (Optional)**  
```python
df = df.rename(columns={**orgunit_rename, **column_rename})
```
- If you want to split the renaming into separate steps (e.g., for debugging), modify it as follows:
  ```python
  df = df.rename(columns=orgunit_rename)
  df = df.rename(columns=column_rename)
  ```


### 1.6 Calculate new variables

```
import numpy as np
import pandas as pd

def create_variables(df):
    """
    Creates new variables in the DataFrame based on summation and subtraction of specified columns.

    Args:
        df (pd.DataFrame): The input DataFrame.

    Returns:
        pd.DataFrame: The DataFrame with new variables added.
    """
    try:
        # Create the 'allout' variable
        df['allout'] = df[['allout_u5', 'allout_ov5']].sum(axis=1, skipna=True, min_count=1)

        # Create the 'susp' variable
        df['susp'] = df[['susp_u5_hf', 'susp_5_14_hf', 'susp_ov15_hf', 'susp_u5_com', 'susp_5_14_com', 'susp_ov15_com']].sum(axis=1, skipna=True, min_count=1)

        # Create the 'test_hf' variable
        test_hf_columns = [
            'test_neg_mic_u5_hf', 'test_pos_mic_u5_hf', 'test_neg_mic_5_14_hf', 'test_pos_mic_5_14_hf',
            'test_neg_mic_ov15_hf', 'test_pos_mic_ov15_hf', 'tes_neg_rdt_u5_hf', 'tes_pos_rdt_u5_hf',
            'tes_neg_rdt_5_14_hf', 'tes_pos_rdt_5_14_hf', 'tes_neg_rdt_ov15_hf', 'tes_pos_rdt_ov15_hf'
        ]
        df['test_hf'] = df[test_hf_columns].sum(axis=1, skipna=True, min_count=1)

        # Create the 'test_com' variable
        test_com_columns = [
            'tes_neg_rdt_u5_com', 'tes_pos_rdt_u5_com', 'tes_neg_rdt_5_14_com', 'tes_pos_rdt_5_14_com',
            'tes_neg_rdt_ov15_com', 'tes_pos_rdt_ov15_com'
        ]
        df['test_com'] = df[test_com_columns].sum(axis=1, skipna=True, min_count=1)

        # Create the 'test' variable
        df['test'] = df[['test_hf', 'test_com']].sum(axis=1, skipna=True, min_count=1)

        # Create the 'conf_hf' variable
        conf_hf_columns = [
            'test_pos_mic_u5_hf', 'test_pos_mic_5_14_hf', 'test_pos_mic_ov15_hf',
            'tes_pos_rdt_u5_hf', 'tes_pos_rdt_5_14_hf', 'tes_pos_rdt_ov15_hf'
        ]
        df['conf_hf'] = df[conf_hf_columns].sum(axis=1, skipna=True, min_count=1)

        # Create the 'conf_com' variable
        conf_com_columns = [
            'tes_pos_rdt_u5_com', 'tes_pos_rdt_5_14_com', 'tes_pos_rdt_ov15_com'
        ]
        df['conf_com'] = df[conf_com_columns].sum(axis=1, skipna=True, min_count=1)

        # Create the 'conf' variable
        df['conf'] = df[['conf_hf', 'conf_com']].sum(axis=1, skipna=True, min_count=1)

        # Create the 'maltreat_com' variable
        maltreat_com_columns = [
            'maltreat_u24_u5_com', 'maltreat_ov24_u5_com', 'maltreat_u24_5_14_com',
            'maltreat_ov24_5_14_com', 'maltreat_u24_ov15_com', 'maltreat_ov24_ov15_com'
        ]
        df['maltreat_com'] = df[maltreat_com_columns].sum(axis=1, skipna=True, min_count=1)

        # Create the 'maltreat_hf' variable
        maltreat_hf_columns = [
            'maltreat_u24_u5_hf', 'maltreat_ov24_u5_hf', 'maltreat_u24_5_14_hf',
            'maltreat_ov24_5_14_hf', 'maltreat_u24_ov15_hf', 'maltreat_ov24_ov15_hf'
        ]
        df['maltreat_hf'] = df[maltreat_hf_columns].sum(axis=1, skipna=True, min_count=1)

        # Create the 'maltreat' variable
        df['maltreat'] = df[['maltreat_hf', 'maltreat_com']].sum(axis=1, skipna=True, min_count=1)

        # Create the 'pres_com' variable
        df['pres_com'] = df['maltreat_com'].sub(df['conf_com'], fill_value=0)
        df['pres_com'] = np.where(df['pres_com'] < 0, 0, df['pres_com'])

        # Create the 'pres_hf' variable
        df['pres_hf'] = df['maltreat_hf'].sub(df['conf_hf'], fill_value=0)
        df['pres_hf'] = np.where(df['pres_hf'] < 0, 0, df['pres_hf'])


        # Create the 'pres' variable
        df['pres'] = df[['pres_com', 'pres_hf']].sum(axis=1, skipna=True, min_count=1)

        # Create the 'maladm' variable
        maladm_columns = ['maladm_u5', 'maladm_5_14', 'maladm_ov15']
        df['maladm'] = df[maladm_columns].sum(axis=1, skipna=True, min_count=1)

        # Create the 'maldth' variable
        maldth_columns = [
            'maldth_u5', 'maldth_1_59m', 'maldth_10_14', 'maldth_5_9',
            'maldth_5_14', 'maldth_ov15', 'maldth_fem_ov15', 'maldth_mal_ov15'
        ]
        df['maldth'] = df[maldth_columns].sum(axis=1, skipna=True, min_count=1)

        return df
    except Exception as e:
        print(f"Error creating variables: {e}")
        return None

# Call the function
df = create_variables(df)
```

**1. Import Necessary Libraries**  
```python
import numpy as np
import pandas as pd
```
- Imports `numpy` for numerical operations and `pandas` for data manipulation.


**2. Define the `create_variables` Function**  
```python
def create_variables(df):
    """
    Creates new variables in the DataFrame based on summation and subtraction of specified columns.

    Args:
        df (pd.DataFrame): The input DataFrame.

    Returns:
        pd.DataFrame: The DataFrame with new variables added.
    """
```
- Defines a function named `create_variables` to generate new columns based on computations.
- Includes a docstring describing the function's purpose, input (`df`), and output (a modified DataFrame).



**3. Use a `try` Block to Handle Errors**  
```python
    try:
```
- Starts a `try` block to handle potential exceptions during column computations.


**4. Create the `allout` Variable**  
```python
        df['allout'] = df[['allout_u5', 'allout_ov5']].sum(axis=1, skipna=True, min_count=1)
```
- Adds a new column `allout` by summing `allout_u5` and `allout_ov5` for each row.
- `skipna=True` ensures `NaN` values are ignored, and `min_count=1` ensures at least one non-`NaN` value is required to compute the sum.



**5. Create the `susp` Variable**  
```python
        df['susp'] = df[['susp_u5_hf', 'susp_5_14_hf', 'susp_ov15_hf', 'susp_u5_com', 'susp_5_14_com', 'susp_ov15_com']].sum(axis=1, skipna=True, min_count=1)
```
- Adds a new column `susp` by summing all columns related to suspected malaria cases in both health facilities (`hf`) and communities (`com`).



**6. Create the `test_hf` Variable**  
```python
        test_hf_columns = [
            'test_neg_mic_u5_hf', 'test_pos_mic_u5_hf', 'test_neg_mic_5_14_hf', 'test_pos_mic_5_14_hf',
            'test_neg_mic_ov15_hf', 'test_pos_mic_ov15_hf', 'tes_neg_rdt_u5_hf', 'tes_pos_rdt_u5_hf',
            'tes_neg_rdt_5_14_hf', 'tes_pos_rdt_5_14_hf', 'tes_neg_rdt_ov15_hf', 'tes_pos_rdt_ov15_hf'
        ]
        df['test_hf'] = df[test_hf_columns].sum(axis=1, skipna=True, min_count=1)
```
- Sums all test-related columns in health facilities to compute the `test_hf` variable.



**7. Create the `test_com` Variable**  
```python
        test_com_columns = [
            'tes_neg_rdt_u5_com', 'tes_pos_rdt_u5_com', 'tes_neg_rdt_5_14_com', 'tes_pos_rdt_5_14_com',
            'tes_neg_rdt_ov15_com', 'tes_pos_rdt_ov15_com'
        ]
        df['test_com'] = df[test_com_columns].sum(axis=1, skipna=True, min_count=1)
```
- Sums all test-related columns in communities to compute the `test_com` variable.



**8. Create the `test` Variable**  
```python
        df['test'] = df[['test_hf', 'test_com']].sum(axis=1, skipna=True, min_count=1)
```
- Combines `test_hf` and `test_com` into a total `test` variable.



**9. Create the `conf_hf` and `conf_com` Variables**  
```python
        conf_hf_columns = [
            'test_pos_mic_u5_hf', 'test_pos_mic_5_14_hf', 'test_pos_mic_ov15_hf',
            'tes_pos_rdt_u5_hf', 'tes_pos_rdt_5_14_hf', 'tes_pos_rdt_ov15_hf'
        ]
        df['conf_hf'] = df[conf_hf_columns].sum(axis=1, skipna=True, min_count=1)
        
        conf_com_columns = [
            'tes_pos_rdt_u5_com', 'tes_pos_rdt_5_14_com', 'tes_pos_rdt_ov15_com'
        ]
        df['conf_com'] = df[conf_com_columns].sum(axis=1, skipna=True, min_count=1)
```
- Sums positive test results for health facilities (`conf_hf`) and communities (`conf_com`).



**10. Create the `conf` Variable**  
```python
        df['conf'] = df[['conf_hf', 'conf_com']].sum(axis=1, skipna=True, min_count=1)
```
- Combines `conf_hf` and `conf_com` into a total `conf` variable.



**11. Create Treatment Variables**  
```python
        maltreat_com_columns = [
            'maltreat_u24_u5_com', 'maltreat_ov24_u5_com', 'maltreat_u24_5_14_com',
            'maltreat_ov24_5_14_com', 'maltreat_u24_ov15_com', 'maltreat_ov24_ov15_com'
        ]
        df['maltreat_com'] = df[maltreat_com_columns].sum(axis=1, skipna=True, min_count=1)
        
        maltreat_hf_columns = [
            'maltreat_u24_u5_hf', 'maltreat_ov24_u5_hf', 'maltreat_u24_5_14_hf',
            'maltreat_ov24_5_14_hf', 'maltreat_u24_ov15_hf', 'maltreat_ov24_ov15_hf'
        ]
        df['maltreat_hf'] = df[maltreat_hf_columns].sum(axis=1, skipna=True, min_count=1)
        
        df['maltreat'] = df[['maltreat_hf', 'maltreat_com']].sum(axis=1, skipna=True, min_count=1)
```
- Computes treatment-related variables for communities, health facilities, and their total.



**12. Create Prescription Variables**  
```python
        df['pres_com'] = df['maltreat_com'].sub(df['conf_com'], fill_value=0)
        df['pres_com'] = np.where(df['pres_com'] < 0, 0, df['pres_com'])
        
        df['pres_hf'] = df['maltreat_hf'].sub(df['conf_hf'], fill_value=0)
        df['pres_hf'] = np.where(df['pres_hf'] < 0, 0, df['pres_hf'])
        
        df['pres'] = df[['pres_com', 'pres_hf']].sum(axis=1, skipna=True, min_count=1)
```
- Calculates prescriptions (`pres_com`, `pres_hf`) by subtracting confirmed cases from treatments and ensures no negative values.
- Combines these into the total `pres` variable.



**13. Create Malaria Admission and Death Variables**  
```python
        maladm_columns = ['maladm_u5', 'maladm_5_14', 'maladm_ov15']
        df['maladm'] = df[maladm_columns].sum(axis=1, skipna=True, min_count=1)
        
        maldth_columns = [
            'maldth_u5', 'maldth_1_59m', 'maldth_10_14', 'maldth_5_9',
            'maldth_5_14', 'maldth_ov15', 'maldth_fem_ov15', 'maldth_mal_ov15'
        ]
        df['maldth'] = df[maldth_columns].sum(axis=1, skipna=True, min_count=1)
```
- Computes total malaria admissions and deaths for various age groups.


**14. Handle Exceptions**  
```python
    except Exception as e:
        print(f"Error creating variables: {e}")
        return None
```
- Catches and handles any exceptions, printing the error message.


**15. Call the Function**  
```python
df = create_variables(df)
```
- Calls the `create_variables` function and processes the DataFrame `df`.


#### 1.6.1 User guide: what to modify


**1. Input DataFrame**  
- Replace the placeholder `df` with the name of your actual DataFrame when calling the function:
  ```python
  df = create_variables(df)
  ```

**2. Column Names in the Dataset**  
- Ensure all column names referenced in the function exist in your DataFrame. If any column names differ, update the function to match your dataset:
  - Columns for suspected cases (`susp_u5_hf`, `susp_ov15_com`, etc.).
  - Testing columns (`test_neg_mic_u5_hf`, `tes_pos_rdt_u5_com`, etc.).
  - Treatment columns (`maltreat_u24_u5_com`, `maltreat_ov24_ov15_hf`, etc.).
  - Malaria admissions and deaths columns (`maladm_u5`, `maldth_1_59m`, etc.).


**3. Missing Value Handling**  
- Verify and, if necessary, adjust the handling of missing values in `.sum()` calculations:
  ```python
  df[column_list].sum(axis=1, skipna=True, min_count=1)
  ```
  - `skipna=True`: Ignores `NaN` values.
  - `min_count=1`: Ensures at least one non-`NaN` value is required for the sum.


**4. Logical Operations for Prescription Variables**  
- Confirm the subtraction logic for prescription variables (`pres_com` and `pres_hf`):
  ```python
  df['pres_com'] = df['maltreat_com'].sub(df['conf_com'], fill_value=0)
  df['pres_hf'] = df['maltreat_hf'].sub(df['conf_hf'], fill_value=0)
  ```
  - If this logic doesn’t apply to your dataset, update or remove it.


**5. Error Handling Messages**  
- Update the error message in the `except` block to provide more context for debugging:
  ```python
  print(f"Error creating variables: {e}")
  ```

**6. New Variable Naming**  
- Review the names of the new variables (`allout`, `susp`, `test`, `conf`, `maltreat`, etc.) and ensure they align with your naming conventions. Update if necessary.


**7. Additional Variables or Columns**  
- If there are additional columns in your dataset relevant to the calculations, add them to the appropriate lists for summation.


### 1.7 Generate health facility ID

```
def create_hfid_column(df):
    """
    Create a unique HFID by grouping adm1, adm2, adm3, and hf,
    and assigning a unique ID starting from 'hf_0001'.
    """
    # Generate unique IDs for each group
    df['hf_uid'] = (
        df.groupby(['adm1', 'adm2', 'adm3', 'hf'])
        .ngroup()  # Assigns group numbers
        .apply(lambda x: f"hf_{x + 1:04}")  # Formats as hf_0001, hf_0002, etc.
    )
    return df

# Call the function
df=create_hfid_column(df)
```

**1. Define the `create_hfid_column` Function**  
```python
def create_hfid_column(df):
    """
    Create a unique HFID by grouping adm1, adm2, adm3, and hf,
    and assigning a unique ID starting from 'hf_0001'.
    """
```
- Defines a function named `create_hfid_column` to generate a unique Health Facility ID (`hf_uid`) for each group of administrative levels (`adm1`, `adm2`, `adm3`) and health facilities (`hf`).
- The docstring explains the purpose and logic of the function.



**2. Group Data by Administrative Levels and HF**  
```python
    df['hf_uid'] = (
        df.groupby(['adm1', 'adm2', 'adm3', 'hf'])
        .ngroup()  # Assigns group numbers
```
- Groups the DataFrame by the specified columns: `adm1`, `adm2`, `adm3`, and `hf`.
- Uses `.ngroup()` to assign a unique group number (`0`, `1`, `2`, etc.) for each combination of these columns.



**3. Format Group Numbers into IDs**  
```python
        .apply(lambda x: f"hf_{x + 1:04}")  # Formats as hf_0001, hf_0002, etc.
```
- Applies a `lambda` function to format the group numbers into strings with a prefix of `hf_` and zero-padded to four digits.
  - `x + 1`: Ensures the numbering starts from `1`.
  - `:04`: Pads the number to ensure it is at least four digits (e.g., `0001`).



**4. Return the Updated DataFrame**  
```python
    )
    return df
```
- Updates the DataFrame by adding the new column `hf_uid`.
- Returns the updated DataFrame.



**5. Call the `create_hfid_column` Function**  
```python
df = create_hfid_column(df)
```
- Calls the `create_hfid_column` function, passing the DataFrame (`df`) as input.
- Updates the `df` variable to include the new `hf_uid` column.

#### 1.7.1 User guide: what to modify

**1. Column Names for Grouping**  
```python
df.groupby(['adm1', 'adm2', 'adm3', 'hf'])
```
- Ensure the column names `'adm1'`, `'adm2'`, `'adm3'`, and `'hf'` exist in your DataFrame. 
- If your column names differ, replace them with the actual column names in your dataset.


**2. Unique ID Format**  
```python
.apply(lambda x: f"hf_{x + 1:04}")
```
- Update the format of the unique ID if a different prefix or padding is required:
  - To change the prefix:
    ```python
    .apply(lambda x: f"facility_{x + 1:04}")
    ```
  - To use a different padding (e.g., three digits):
    ```python
    .apply(lambda x: f"hf_{x + 1:03}")
    ```

**3. Input DataFrame**  
```python
df = create_hfid_column(df)
```
- Replace `df` with the actual name of your DataFrame when calling the function.


**4. Additional Grouping Columns (Optional)**  
- If additional columns are required for unique identification, include them in the `groupby` method:
  ```python
  df.groupby(['adm1', 'adm2', 'adm3', 'hf', 'extra_column'])
  ``` 

**5. Ensure Required Columns Exist**  
- Verify that the columns specified in the `groupby` method contain the correct data and are not missing or empty. 



### 1.8 Outlier detection and correction


```
!pip install xlsxwriter

from io import BytesIO
import pandas as pd
import numpy as np
import xlsxwriter

# Function to detect outliers using Scatterplot with Q1 and Q3 lines
def detect_outliers_scatterplot(df, col):
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    return lower_bound, upper_bound

# Function to calculate moving average
def calculate_moving_avg(series, window):
    return series.rolling(window=window, min_periods=1).mean()

# Improved function to calculate moving average excluding outliers
def calculate_moving_avg_excluding_outliers(series, window, threshold=1.5):
    # First identify outliers for the entire series
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - threshold * iqr
    upper_bound = q3 + threshold * iqr

    # Create a mask for non-outlier values
    non_outlier_mask = (series >= lower_bound) & (series <= upper_bound)

    # Replace outliers with NaN
    clean_series = series.where(non_outlier_mask)

    # Calculate moving average on clean series
    ma = clean_series.rolling(window=window, min_periods=1).mean()

    # Forward fill and then backward fill NaN values
    #ma = ma.fillna(method='ffill').fillna(method='bfill')
    ma = ma.ffill().bfill()

    return ma

# Function to process and export the results for a single column
def process_column_export(df, column, output_file):
    grouped = df.groupby(['adm1', 'adm2', 'adm3', 'hf', 'year'])

    results = []
    for (adm1, adm2, adm3, hf, year), group in grouped:
        lower_bound, upper_bound = detect_outliers_scatterplot(group, column)

        group[f'{column}_lower_bound'] = lower_bound
        group[f'{column}_upper_bound'] = upper_bound
        group[f'{column}_category'] = np.where(
            (group[column] < lower_bound) | (group[column] > upper_bound), 'Outlier', 'Non-Outlier'
        )

        mean_include_outliers = group[column].mean()
        mean_exclude_outliers = group[(group[column] >= lower_bound) & (group[column] <= upper_bound)][column].mean()
        median_include_outliers = group[column].median()
        median_exclude_outliers = group[(group[column] >= lower_bound) & (group[column] <= upper_bound)][column].median()
        moving_avg_include_outliers = calculate_moving_avg(group[column], window=3)
        moving_avg_exclude_outliers = calculate_moving_avg_excluding_outliers(group[column], window=3)
        winsorised = group[column].clip(lower=lower_bound, upper=upper_bound)

        # Replace outliers with respective correction methods
        group[f'{column}_corrected_mean_include'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', mean_include_outliers)
        group[f'{column}_corrected_mean_exclude'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', mean_exclude_outliers)
        group[f'{column}_corrected_median_include'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', median_include_outliers)
        group[f'{column}_corrected_median_exclude'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', median_exclude_outliers)
        group[f'{column}_corrected_moving_avg_include'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', moving_avg_include_outliers)
        group[f'{column}_corrected_moving_avg_exclude'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', moving_avg_exclude_outliers)
        group[f'{column}_corrected_winsorised'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', winsorised)

        results.append(group)

    final_df = pd.concat(results)

    export_columns = [
        'adm1', 'adm2', 'adm3', 'hf', 'year', 'month', column,
        f'{column}_category', f'{column}_lower_bound', f'{column}_upper_bound',
        f'{column}_corrected_mean_include', f'{column}_corrected_mean_exclude',
        f'{column}_corrected_median_include', f'{column}_corrected_median_exclude',
        f'{column}_corrected_moving_avg_include', f'{column}_corrected_moving_avg_exclude',
        f'{column}_corrected_winsorised'
    ]

    final_df[export_columns].to_excel(output_file, index=False, engine='xlsxwriter')

# Load your dataset
#input_excel_file = "/content/routine_data.xlsx"  # Replace with your actual file path
#df = pd.read_excel(input_excel_file)

# Specify the columns to process
columns_to_process = ['allout', 'susp', 'test', 'conf', 'maltreat', 'pres' ,'maladm', 'maldth']

# Process each column and export to separate Excel files
for column in columns_to_process:
    output_file = f"{column}_results.xlsx"
    process_column_export(df, column, output_file)
```
**1. Import Required Libraries**  
```python
!pip install xlsxwriter

from io import BytesIO
import pandas as pd
import numpy as np
import xlsxwriter
```
- Installs `xlsxwriter` (if not already installed).
- Imports necessary libraries:
  - `BytesIO` for in-memory file handling.
  - `pandas` for data manipulation.
  - `numpy` for numerical operations.
  - `xlsxwriter` for exporting to Excel.

---

**2. Define the Function to Detect Outliers**  
```python
def detect_outliers_scatterplot(df, col):
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    return lower_bound, upper_bound
```
- Calculates the lower and upper bounds for detecting outliers using the IQR method.
- Returns the calculated bounds for a given column.

---

**3. Define the Function to Calculate Moving Average**  
```python
def calculate_moving_avg(series, window):
    return series.rolling(window=window, min_periods=1).mean()
```
- Computes the moving average for a given series with a specified rolling window.

---

**4. Define the Function for Moving Average Excluding Outliers**  
```python
def calculate_moving_avg_excluding_outliers(series, window, threshold=1.5):
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - threshold * iqr
    upper_bound = q3 + threshold * iqr

    non_outlier_mask = (series >= lower_bound) & (series <= upper_bound)
    clean_series = series.where(non_outlier_mask)

    ma = clean_series.rolling(window=window, min_periods=1).mean()
    ma = ma.ffill().bfill()

    return ma
```
- Detects outliers and replaces them with `NaN` before calculating the moving average.
- Uses forward and backward filling to handle gaps.

---

**5. Define the Function to Process and Export Results**  
```python
def process_column_export(df, column, output_file):
```
- Processes outliers, computes corrected values, and exports the results for a specific column to an Excel file.

---

**6. Group the DataFrame by Admin and Year Levels**  
```python
    grouped = df.groupby(['adm1', 'adm2', 'adm3', 'hf', 'year'])
```
- Groups the DataFrame by administrative levels (`adm1`, `adm2`, `adm3`, `hf`) and `year`.

---

**7. Iterate Over Groups and Process Data**  
```python
    for (adm1, adm2, adm3, hf, year), group in grouped:
        lower_bound, upper_bound = detect_outliers_scatterplot(group, column)
```
- Iterates through each group.
- Calculates lower and upper bounds for outliers in the specified column.

---

**8. Add Outlier Information to the Group**  
```python
        group[f'{column}_lower_bound'] = lower_bound
        group[f'{column}_upper_bound'] = upper_bound
        group[f'{column}_category'] = np.where(
            (group[column] < lower_bound) | (group[column] > upper_bound), 'Outlier', 'Non-Outlier'
        )
```
- Adds lower and upper bounds for outliers as new columns.
- Categorizes values as "Outlier" or "Non-Outlier."

---

**9. Calculate Statistical Metrics**  
```python
        mean_include_outliers = group[column].mean()
        mean_exclude_outliers = group[(group[column] >= lower_bound) & (group[column] <= upper_bound)][column].mean()
        median_include_outliers = group[column].median()
        median_exclude_outliers = group[(group[column] >= lower_bound) & (group[column] <= upper_bound)][column].median()
```
- Computes:
  - Mean and median including outliers.
  - Mean and median excluding outliers.

---

**10. Compute Moving Averages and Winsorization**  
```python
        moving_avg_include_outliers = calculate_moving_avg(group[column], window=3)
        moving_avg_exclude_outliers = calculate_moving_avg_excluding_outliers(group[column], window=3)
        winsorised = group[column].clip(lower=lower_bound, upper=upper_bound)
```
- Calculates moving averages:
  - Including outliers.
  - Excluding outliers.
- Winsorizes data by capping outliers at the bounds.

---

**11. Create Corrected Columns**  
```python
        group[f'{column}_corrected_mean_include'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', mean_include_outliers)
        group[f'{column}_corrected_mean_exclude'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', mean_exclude_outliers)
        group[f'{column}_corrected_median_include'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', median_include_outliers)
        group[f'{column}_corrected_median_exclude'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', median_exclude_outliers)
        group[f'{column}_corrected_moving_avg_include'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', moving_avg_include_outliers)
        group[f'{column}_corrected_moving_avg_exclude'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', moving_avg_exclude_outliers)
        group[f'{column}_corrected_winsorised'] = group[column].where(group[f'{column}_category'] == 'Non-Outlier', winsorised)
```
- Creates corrected columns using different statistical and outlier correction methods.

---

**12. Append Processed Data**  
```python
        results.append(group)
```
- Appends each processed group to the results list.

---

**13. Combine and Export Results**  
```python
    final_df = pd.concat(results)
    final_df[export_columns].to_excel(output_file, index=False, engine='xlsxwriter')
```
- Combines all processed groups into a single DataFrame.
- Exports the results to an Excel file.

---

**14. Process All Specified Columns**  
```python
columns_to_process = ['allout', 'susp', 'test', 'conf', 'maltreat', 'pres', 'maladm', 'maldth']

for column in columns_to_process:
    output_file = f"{column}_results.xlsx"
    process_column_export(df, column, output_file)
```
- Loops through all specified columns, processes them, and exports results to separate Excel files.


#### 1.8.1 User guide: what to modify


**1. Input DataFrame (`df`)**
- Replace `df` with the actual name of your DataFrame containing the data.
  ```python
  df = create_variables(df)
  ```


**2. Grouping Columns**
- Ensure the following columns exist in your dataset for grouping:
  - `adm1`, `adm2`, `adm3`, `hf`, and `year`.
- If your column names differ, update them in the grouping logic:
  ```python
  grouped = df.groupby(['your_adm1_column', 'your_adm2_column', 'your_adm3_column', 'your_hf_column', 'your_year_column'])
  ```


**3. Columns to Process**
- Ensure the `columns_to_process` list includes all the columns you want to analyze:
  ```python
  columns_to_process = ['allout', 'susp', 'test', 'conf', 'maltreat', 'pres', 'maladm', 'maldth']
  ```
- If additional columns are required, add them to this list.


**4. Column Names**
- Verify that all columns referenced in the code (e.g., `allout`, `susp`, `test`) exist in your DataFrame. Update their names if they differ in your dataset:
  ```python
  'allout', 'susp', 'test', 'conf', 'maltreat', 'pres', 'maladm', 'maldth'
  ```


**5. Export File Names**
- Update the naming convention for the output Excel files if needed:
  ```python
  output_file = f"{column}_results.xlsx"
  ```
- If you want to save the files in a specific directory:
  ```python
  output_file = f"/path/to/directory/{column}_results.xlsx"
  ```


**6. Export Columns**
- Ensure the `export_columns` list includes the columns you want in the output file. Modify this list if additional columns are needed or some columns are unnecessary:
  ```python
  export_columns = [
      'adm1', 'adm2', 'adm3', 'hf', 'year', 'month', column,
      f'{column}_category', f'{column}_lower_bound', f'{column}_upper_bound',
      f'{column}_corrected_mean_include', f'{column}_corrected_mean_exclude',
      f'{column}_corrected_median_include', f'{column}_corrected_median_exclude',
      f'{column}_corrected_moving_avg_include', f'{column}_corrected_moving_avg_exclude',
      f'{column}_corrected_winsorised'
  ]
  ```


**7. Missing Values Handling**
- Ensure missing values are handled appropriately in `.sum()` operations. These parameters are already set correctly:
  ```python
  skipna=True  # Ignores NaN values
  min_count=1  # Requires at least one non-NaN value
  ```


**8. Statistical Calculations**
- Ensure the statistical calculations (mean, median, moving averages, winsorization) match your requirements. If no modifications are required, keep these unchanged:
  - Mean including/excluding outliers.
  - Median including/excluding outliers.
  - Moving averages.
  - Winsorized values.


**9. Dataset Formatting**
- Ensure your dataset is properly formatted:
  - No missing critical grouping columns (`adm1`, `adm2`, `adm3`, `hf`, `year`).
  - All data columns are numerical and free of non-numeric data if calculations are required.

### 1.9 Merge all data

```
import pandas as pd

def merge_all_results(columns_to_process):
    """
    Merge all individual result Excel files into one consolidated file
    using adm1, adm2, adm3, hf, Year, Month as merge keys
    """
    # Define the merge keys
    merge_keys = ['adm1', 'adm2', 'adm3', 'hf', 'year', 'month']

    # Initialize with the first file
    first_column = columns_to_process[0]
    merged_df = pd.read_excel(f"{first_column}_results.xlsx")

    # Merge remaining files one by one
    for column in columns_to_process[1:]:
        # Read the next file
        current_df = pd.read_excel(f"{column}_results.xlsx")

        # Merge with the accumulated result
        merged_df = pd.merge(
            merged_df,
            current_df,
            on=merge_keys,
            how='outer',
            suffixes=('', f'_{column}')  # Avoid column name conflicts
        )

    # Sort the data
    merged_df = merged_df.sort_values(by=merge_keys)

    # Export to Excel
    merged_df.to_excel('clean_routine_data.xlsx', index=False)
    print(f"Merged data saved to 'clean_routine_data.xlsx'")

    return merged_df

# List of columns/files to merge
columns_to_process = ['allout', 'susp', 'test', 'conf', 'maltreat', 'pres', 'maladm', 'maldth']

# Perform the merge
df = merge_all_results(columns_to_process)
```

**1. Import the pandas Library**  
```python
import pandas as pd
```
- Imports the `pandas` library to handle data manipulation and file reading/writing.


**2. Define the `merge_all_results` Function**  
```python
def merge_all_results(columns_to_process):
    """
    Merge all individual result Excel files into one consolidated file
    using adm1, adm2, adm3, hf, Year, Month as merge keys
    """
```
- Defines a function named `merge_all_results` to consolidate individual result files into a single Excel file.
- Includes a docstring explaining the function’s purpose and the merge keys used.


**3. Define Merge Keys**  
```python
    merge_keys = ['adm1', 'adm2', 'adm3', 'hf', 'year', 'month']
```
- Specifies the columns (`merge_keys`) to use as keys for merging the individual result files.
- These keys ensure that the data from different files aligns correctly.


**4. Initialize with the First File**  
```python
    first_column = columns_to_process[0]
    merged_df = pd.read_excel(f"{first_column}_results.xlsx")
```
- Reads the first result file (e.g., `allout_results.xlsx`) into a DataFrame named `merged_df`.
- Uses the first column from the `columns_to_process` list as the starting point.


**5. Iterate Over Remaining Columns**  
```python
    for column in columns_to_process[1:]:
        # Read the next file
        current_df = pd.read_excel(f"{column}_results.xlsx")
```
- Loops through the remaining columns in the `columns_to_process` list.
- Reads each corresponding result file (e.g., `susp_results.xlsx`, `test_results.xlsx`) into a DataFrame named `current_df`.


**6. Merge Files One by One**  
```python
        merged_df = pd.merge(
            merged_df,
            current_df,
            on=merge_keys,
            how='outer',
            suffixes=('', f'_{column}')  # Avoid column name conflicts
        )
```
- Merges `current_df` with the accumulated `merged_df` using the defined `merge_keys`.
- Performs an `outer` merge to retain all data from both files.
- Adds suffixes to avoid column name conflicts (e.g., if a column exists in multiple files).


**7. Sort the Merged Data**  
```python
    merged_df = merged_df.sort_values(by=merge_keys)
```
- Sorts the consolidated DataFrame (`merged_df`) by the `merge_keys` for better organization.


**8. Export the Merged Data**  
```python
    merged_df.to_excel('clean_routine_data.xlsx', index=False)
    print(f"Merged data saved to 'clean_routine_data.xlsx'")
```
- Exports the consolidated DataFrame to an Excel file named `clean_routine_data.xlsx`.
- Prints a confirmation message after saving the file.


**9. Return the Merged DataFrame**  
```python
    return merged_df
```
- Returns the merged DataFrame (`merged_df`) for further use or analysis.


**10. List of Columns/Files to Merge**  
```python
columns_to_process = ['allout', 'susp', 'test', 'conf', 'maltreat', 'pres', 'maladm', 'maldth']
```
- Specifies the list of columns (and corresponding result files) to be merged.
 

**11. Perform the Merge**  
```python
df = merge_all_results(columns_to_process)
```
- Calls the `merge_all_results` function, passing the `columns_to_process` list as input.
- Stores the merged DataFrame in the variable `df`.


### 1.10 User guide: what to modify

**1. Input Files**
- Ensure the files corresponding to each column in `columns_to_process` exist and are named correctly in the format `{column}_results.xlsx` (e.g., `allout_results.xlsx`, `susp_results.xlsx`).
  ```python
  columns_to_process = ['allout', 'susp', 'test', 'conf', 'maltreat', 'pres', 'maladm', 'maldth']
  ```

**2. Merge Keys**
- Verify that the merge keys (`adm1`, `adm2`, `adm3`, `hf`, `year`, `month`) exist in all the files being merged. Update the keys if your column names differ:
  ```python
  merge_keys = ['your_adm1_column', 'your_adm2_column', 'your_adm3_column', 'your_hf_column', 'your_year_column', 'your_month_column']
  ```

**3. Column Name Conflicts**
- Check for overlapping column names in the input files. The `suffixes` parameter handles conflicts, but you can customize the suffixes if needed:
  ```python
  suffixes=('', f'_{column}')  # Default behavior
  ```

**4. Output File Name**
- Update the output file name and location as required:
  ```python
  merged_df.to_excel('clean_routine_data.xlsx', index=False)
  ```
  - To save to a specific directory:
    ```python
    merged_df.to_excel('/path/to/directory/clean_routine_data.xlsx', index=False)
    ```

**5. Sorting Keys**
- Ensure the `merge_keys` used for sorting (`adm1`, `adm2`, `adm3`, `hf`, `year`, `month`) are appropriate for your dataset. Update them if the dataset requires different sorting criteria:
  ```python
  merged_df = merged_df.sort_values(by=['your_sort_key1', 'your_sort_key2', ...])
  ```

**6. Columns to Process**
- If additional or fewer columns need to be merged, modify the `columns_to_process` list:
  ```python
  columns_to_process = ['your_column1', 'your_column2', ...]
  ```

### 1.11 Output final database

```
import pandas as pd

def save_selected_columns_to_excel(input_file, selected_columns, output_file):
    """
    Save only the specified columns from the input Excel file to a new Excel file.

    Parameters:
        input_file (str): Path to the input Excel file.
        selected_columns (list): List of column names to select.
        output_file (str): Path to save the resulting Excel file.

    Returns:
        None
    """
    try:
        # Read the complete dataset
        df = pd.read_excel(input_file)

        # Select only the specified columns that exist in the dataset
        existing_columns = [col for col in selected_columns if col in df.columns]
        if not existing_columns:
            raise ValueError("None of the selected columns exist in the dataset.")

        # Create a DataFrame with only the selected columns
        selected_df = df[existing_columns]

        # Save to Excel
        selected_df.to_excel(output_file, index=False)
        print(f"Selected columns saved to '{output_file}' successfully!")
        print(f"Columns saved: {existing_columns}")
    except Exception as e:
        print(f"Error: {e}")


def rename_columns_in_excel(file_path, rename_mapping, output_file):
    """
    Rename columns in an existing Excel file and save to a new file.

    Parameters:
        file_path (str): Path to the Excel file.
        rename_mapping (dict): Dictionary mapping old column names to new column names.
        output_file (str): Path to save the renamed Excel file.

    Returns:
        None
    """
    try:
        # Read the dataset from the existing file
        df = pd.read_excel(file_path)

        # Rename the columns
        df.rename(columns=rename_mapping, inplace=True)

        # Save the renamed DataFrame to a new Excel file
        df.to_excel(output_file, index=False)
        print(f"Renamed columns saved to '{output_file}' successfully!")
        print(f"Columns renamed: {rename_mapping}")
    except Exception as e:
        print(f"Error: {e}")


# Specify input file, selected columns, and intermediate file
input_file = 'clean_routine_data.xlsx'
selected_columns = [
    'adm1', 'adm2', 'adm3', 'hf', 'year', 'month',
    'allout_corrected_winsorised', 'susp_corrected_winsorised', 'test_corrected_winsorised',
    'conf_corrected_winsorised', 'maltreat_corrected_winsorised', 'pres_corrected_winsorised',
    'maladm_corrected_winsorised', 'maldth_corrected_winsorised'
]
intermediate_file = 'intermediate_selected_data.xlsx'

# Save selected columns to an intermediate file
save_selected_columns_to_excel(input_file, selected_columns, intermediate_file)

# Specify rename mapping and final output file
rename_mapping = {
    'allout_corrected_winsorised': 'allout',
    'susp_corrected_winsorised': 'susp',
    'test_corrected_winsorised': 'test',
    'conf_corrected_winsorised': 'conf',
    'maltreat_corrected_winsorised': 'maltreat',
    'pres_corrected_winsorised': 'pres',
    'maladm_corrected_winsorised': 'maladm',
    'maldth_corrected_winsorised': 'maldth'
}
final_output_file = 'clean_malaria_routine_data.xlsx'

# Rename columns and save to the final file
rename_columns_in_excel(intermediate_file, rename_mapping, final_output_file)


```

**1. Import Necessary Libraries**  
```python
import pandas as pd
```
- Imports the `pandas` library for handling Excel files and manipulating data.



**2. Define the `save_selected_columns_to_excel` Function**  
```python
def save_selected_columns_to_excel(input_file, selected_columns, output_file):
    """
    Save only the specified columns from the input Excel file to a new Excel file.
    """
```
- Defines a function to save only specific columns from an Excel file to a new file.
- Takes the following parameters:
  - `input_file`: Path to the input Excel file.
  - `selected_columns`: List of columns to extract.
  - `output_file`: Path to save the resulting file.


**3. Read the Dataset**  
```python
        df = pd.read_excel(input_file)
```
- Reads the input Excel file into a pandas DataFrame.


**4. Filter the Selected Columns**  
```python
        existing_columns = [col for col in selected_columns if col in df.columns]
        if not existing_columns:
            raise ValueError("None of the selected columns exist in the dataset.")
```
- Filters the `selected_columns` list to include only columns that exist in the DataFrame.
- Raises an error if none of the selected columns are present.


**5. Create a New DataFrame with Selected Columns**  
```python
        selected_df = df[existing_columns]
```
- Creates a new DataFrame containing only the specified columns.


**6. Save the Selected Columns to a New Excel File**  
```python
        selected_df.to_excel(output_file, index=False)
        print(f"Selected columns saved to '{output_file}' successfully!")
        print(f"Columns saved: {existing_columns}")
```
- Saves the filtered DataFrame to a new Excel file without the index.
- Prints a confirmation message along with the saved columns.

**7. Define the `rename_columns_in_excel` Function**  
```python
def rename_columns_in_excel(file_path, rename_mapping, output_file):
    """
    Rename columns in an existing Excel file and save to a new file.
    """
```
- Defines a function to rename columns in an Excel file and save the result to a new file.
- Takes the following parameters:
  - `file_path`: Path to the existing Excel file.
  - `rename_mapping`: Dictionary mapping old column names to new column names.
  - `output_file`: Path to save the renamed file.


**8. Read the Dataset from the File**  
```python
        df = pd.read_excel(file_path)
```
- Reads the input Excel file into a pandas DataFrame.


**9. Rename Columns**  
```python
        df.rename(columns=rename_mapping, inplace=True)
```
- Renames the columns in the DataFrame using the provided `rename_mapping` dictionary.


**10. Save the Renamed Columns to a New Excel File**  
```python
        df.to_excel(output_file, index=False)
        print(f"Renamed columns saved to '{output_file}' successfully!")
        print(f"Columns renamed: {rename_mapping}")
```
- Saves the DataFrame with renamed columns to a new Excel file without the index.
- Prints a confirmation message along with the renamed columns.



**11. Specify Input File, Selected Columns, and Intermediate File**  
```python
input_file = 'clean_routine_data.xlsx'
selected_columns = [
    'adm1', 'adm2', 'adm3', 'hf', 'year', 'month',
    'allout_corrected_winsorised', 'susp_corrected_winsorised', 'test_corrected_winsorised',
    'conf_corrected_winsorised', 'maltreat_corrected_winsorised', 'pres_corrected_winsorised',
    'maladm_corrected_winsorised', 'maldth_corrected_winsorised'
]
intermediate_file = 'intermediate_selected_data.xlsx'
```
- Specifies:
  - The input Excel file (`input_file`) containing all data.
  - The list of columns to be saved to a new file (`selected_columns`).
  - The name of the intermediate Excel file (`intermediate_file`) to store the selected columns.



**12. Save Selected Columns to the Intermediate File**  
```python
save_selected_columns_to_excel(input_file, selected_columns, intermediate_file)
```
- Calls the `save_selected_columns_to_excel` function to extract the selected columns and save them to the `intermediate_file`.


**13. Specify Rename Mapping and Final Output File**  
```python
rename_mapping = {
    'allout_corrected_winsorised': 'allout',
    'susp_corrected_winsorised': 'susp',
    'test_corrected_winsorised': 'test',
    'conf_corrected_winsorised': 'conf',
    'maltreat_corrected_winsorised': 'maltreat',
    'pres_corrected_winsorised': 'pres',
    'maladm_corrected_winsorised': 'maladm',
    'maldth_corrected_winsorised': 'maldth'
}
final_output_file = 'clean_malaria_routine_data.xlsx'
```
- Specifies:
  - The mapping of old column names to new column names (`rename_mapping`).
  - The name of the final output Excel file (`final_output_file`).

---

**14. Rename Columns and Save to the Final File**  
```python
rename_columns_in_excel(intermediate_file, rename_mapping, final_output_file)
```
- Calls the `rename_columns_in_excel` function to rename the columns in the `intermediate_file` and save the result to the `final_output_file`.


#### 1.12 User guide: what to modify

**1. Input File (`input_file`)**
- Replace `'clean_routine_data.xlsx'` with the actual path to your input Excel file:
  ```python
  input_file = 'your_input_file.xlsx'
  ```


**2. Selected Columns (`selected_columns`)**
- Ensure the `selected_columns` list contains the column names you want to extract.
- Verify that these columns exist in the input file. Update the list as needed:
  ```python
  selected_columns = [
      'adm1', 'adm2', 'adm3', 'hf', 'year', 'month',
      'your_column1', 'your_column2', ...
  ]
  ```


**3. Intermediate File Name (`intermediate_file`)**
- Modify the name and path of the intermediate file if necessary:
  ```python
  intermediate_file = 'your_intermediate_file.xlsx'
  ```


**4. Rename Mapping (`rename_mapping`)**
- Update the `rename_mapping` dictionary to reflect the old column names and their desired new names:
  ```python
  rename_mapping = {
      'your_old_column1': 'your_new_column1',
      'your_old_column2': 'your_new_column2',
      ...
  }
  ```

**5. Final Output File Name (`final_output_file`)**
- Change the name and path of the final output file as needed:
  ```python
  final_output_file = 'your_final_output_file.xlsx'
  ```

**6. Verify Column Names**
- Ensure that all columns in the `rename_mapping` exist in the `intermediate_file`. Update the mapping or intermediate file if discrepancies occur.




















