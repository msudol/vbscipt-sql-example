# VBScript Database Connection and Query Execution (Written with Copilot)

This VBScript application establishes a connection to a remote SQL Server database, retrieves configuration details from a separate configuration file, and executes an SQL query. It also includes basic error handling and logs any errors to an error log file.

## Prerequisites

1. **VBScript Environment**:
   - Ensure that your system supports VBScript execution (usually available on Windows systems).

2. **Database Connection Details**:
   - Create a `ConfigFile.txt` with the following format (replace placeholders with actual values):
     ```plaintext
     ServerName=your_server
     DatabaseName=your_database
     UserName=your_username
     Password=your_password
     QueryFilePath=C:\Path\To\Your\QueryFile.sql
     ErrorLogPath=C:\Path\To\Your\error.log
     ```

3. **SQL Query File**:
   - Create an SQL query file (`QueryFile.sql`) with your desired SQL query.

## Usage

1. **Edit Configuration**:
   - Update the `ConfigFile.txt` with your database connection details and the path to your SQL query file.

2. **Execute the Script**:
   - Double-click `RunScript.bat` or run it from the command line.
   - The script will connect to the database, execute the SQL query, and log any errors to the specified error log file.

## Error Handling

- If there are errors during database connection or query execution, they will be logged to the `error.log` file.
- Customize the error logging behavior as needed.

## Additional Notes

- Keep your configuration files secure (do not commit sensitive data to version control).
- Adapt the script to your specific use case and database schema.

---

Feel free to customize this `README.md` to match your specific application. If you have any further questions or need assistance, don't hesitate to ask! 🚀
