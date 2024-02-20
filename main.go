package main

import (
    "fmt"
    "strconv"
    "bufio"
    "os"
    "strings"
    "context"
    "github.com/xuri/excelize/v2"
    "database/sql"
    _ "github.com/denisenkom/go-mssqldb"
)

func getUserInput(text string) (string, error) {
    fmt.Printf("> Enter %s : ", text)
    reader := bufio.NewReader(os.Stdin)
    // ReadString will block until the delimiter is entered
    input, err := reader.ReadString('\n')
    if err != nil {
      fmt.Println("An error occured while reading input. Please try again", err)
      return "", err
    }
    // remove the delimeter from the string
    input = strings.ReplaceAll(input, "\r", "")
    input = strings.ReplaceAll(input, "\n", "")
    return input, nil
}

func getLastIDs(db *sql.DB, tableName string) (int, int, error) {
  // Execute query to get the last Transfer_Id and Transaction_Id values. They are alwasy incremented, thus MAX
  queryGetLastIDs := fmt.Sprintf("SELECT MAX(Transfer_Id), MAX(Transaction_Id) FROM %s", tableName)

  rows, err := db.Query(queryGetLastIDs)
  if err != nil {
      fmt.Println(err)
      return 0, 0, err
  }
  defer rows.Close()

  var lastTransferID, lastTransactionID int

  // There should be only one row from MAX 
  if rows.Next() {
      err := rows.Scan(&lastTransferID, &lastTransactionID)
      if err != nil {
        fmt.Println(err)
        return 0, 0, err
      }
  }

  // Increment the retrieved values for the new Transfer_Id and Transaction_Id
  return lastTransferID+1, lastTransactionID+1, nil
}

func convertToFloat(rowValue string) (float64, error) {
    // Try parsing as integer first
    intValue, err := strconv.Atoi(rowValue)
    if err == nil {
        return float64(intValue), nil
    }

    // If parsing as integer fails, try parsing as float64
    floatValue, err := strconv.ParseFloat(rowValue, 64)
    if err != nil {
        return 0, fmt.Errorf("cannot convert value: %v", err)
    }

    return floatValue, nil
}

func main() {

    // Define the connection parameters
    server, err := getUserInput("hostname of sqlexpress server. Example localhost")
    if err != nil {
      fmt.Println(err)
      return
    }

    user, err := getUserInput("user name, example sa")
    if err != nil {
      fmt.Println(err)
      return
    }

    password, err := getUserInput("user password")
    if err != nil {
      fmt.Println(err)
      return
    }

    database, err := getUserInput("database to insert to")
    if err != nil {
      fmt.Println(err)
      return
    }

    tableName, err := getUserInput("table name to insert to. Must be Manual_Data_something")
    if err != nil {
      fmt.Println(err)
      return
    }

    connString := fmt.Sprintf("sqlserver://%s:%s@%s/SQLExpress?database=%s&connection+timeout=30&encrypt=disable",
      user, password, server, database)

    // Establish a connection to the database
    db, err := sql.Open("sqlserver", connString)
    if err != nil {
        fmt.Println(err)
        fmt.Println(connString)
        return
    }
    defer db.Close()

    // Open doesn't open a connection. Validate DSN data:
    err = db.PingContext(context.Background())
    if err != nil {
        fmt.Println(err)
        return
    }


    fileName, err := getUserInput("excel file name, with .xlsx")
    if err != nil {
      fmt.Println(err)
      return
    }
    // Open Excel file
    f, err := excelize.OpenFile(fileName)
    if err != nil {
      fmt.Println(err)
      return
    }
    defer func() {
      // Close the spreadsheet.
      if err := f.Close(); err != nil {
          fmt.Println(err)
      }
    }()

    // keep track of inserted / updated rows
    rowNumber := 0

    // get all sheets
    for _, name := range f.GetSheetMap() {
      fmt.Printf("\nAnalying sheet %s", name)
      fmt.Println()

      rows, err := f.Rows(name)
      if err != nil {
        fmt.Println(err)
        return
      }

      var collectedRows [][]string
      for rows.Next() {
        columns, _:= rows.Columns()
        // ignore empty lines
        if len(columns) > 0 {
          collectedRows = append(collectedRows, columns)
        }
      }

      fmt.Println("Here is first line of data")
      fmt.Println(collectedRows[1])

      insertOk, err := getUserInput("Continue with insert (y/n)")
      if err != nil {
        fmt.Println(err); return
      }
      if insertOk != "y" {
        fmt.Println("user chose to skip this sheet")
        continue
      } 

      // I want the IDs to be the same for the whole sheet
      fmt.Println("Getting last Transfer_Id and Transaction_Id")
      Transfer_Id, Transaction_Id, err := getLastIDs(db, tableName)
      if err != nil {
        fmt.Println(err)

        fmt.Println("... using 1, 1 as IDs")
        Transfer_Id = 1 
        Transaction_Id = 1
      }

      rowNumber += len(collectedRows)

      // skip first row
      for _, row := range collectedRows[1:] {
        // value needs to be integer, not string
        value := 0.0

        // row[2] may be empty, so the slice has size 2. Testing for row[2] == "" will panic
        if len(row) >= 3 {
          convValue, err := convertToFloat(row[2])
          if err != nil {
            fmt.Println(err)
            return
          }
          // this prevents value not used
          value = convValue
        }
        // Use this if you do not want to skip existing rows or update them
        // query := fmt.Sprintf("INSERT INTO %s (Orig_TS_UTC, Orig_TS_Local, Last_Op_TS_UTC, Last_Op_TS_Local, Tag_Name, Num_Value, Operation_Type, Status, Transfer_Id, Transaction_Id, [User]) VALUES ('%s 00:00:00.000', '%s 02:00:00:000', SYSUTCDATETIME(), SYSDATETIME(), '%s', %d, 1, 1, %d, %d, 'goepher')", 
        //   tableName, row[1], row[1], row[0], value, Transfer_Id, Transaction_Id)

        query := fmt.Sprintf(`
          MERGE INTO %s AS target
          USING (VALUES ('%s 00:00:00.000', '%s 02:00:00:000', SYSUTCDATETIME(), SYSDATETIME(), '%s', %f, 1, 1, %d, %d, 'goepher')) AS source (Orig_TS_UTC, Orig_TS_Local, Last_Op_TS_UTC, Last_Op_TS_Local, Tag_Name, Num_Value, Operation_Type, Status, Transfer_Id, Transaction_Id, [User])
          ON target.Orig_TS_Local = source.Orig_TS_Local AND target.Tag_Name = source.Tag_Name
          WHEN MATCHED THEN
            UPDATE SET target.Num_Value = source.Num_Value
          WHEN NOT MATCHED THEN
            INSERT (Orig_TS_UTC, Orig_TS_Local, Last_Op_TS_UTC, Last_Op_TS_Local, Tag_Name, Num_Value, Operation_Type, Status, Transfer_Id, Transaction_Id, [User])
            VALUES (source.Orig_TS_UTC, source.Orig_TS_Local, source.Last_Op_TS_UTC, source.Last_Op_TS_Local, source.Tag_Name, source.Num_Value, source.Operation_Type, source.Status, source.Transfer_Id, source.Transaction_Id, source.[User]);
        `, tableName, row[1], row[1], row[0], value, Transfer_Id, Transaction_Id)

        _, err = db.Exec(query)
        if err != nil {
          fmt.Println(err)
          return
        }
      }
      // basically Dream Reports "SELECT Limit 1" on trans ids
      Transaction_Id += 1
      Transfer_Id += 1
    }

    fmt.Printf("The script inserted / updated %d rows in the %s table", rowNumber, tableName)
}
