let
    // 1. Load Excel files from your folder
    Source = Folder.Files("C:\Users\Renzie\Desktop\Power BI Practice\PQ Book\Chandeep's Weekly challenge\First Challenge"),

    // 2. Extract custom Date column from file name + metadata
    #"Added Custom Column" = Table.AddColumn(Source, "Date", each Text.Combine({Text.Middle([Name], 3, 4), " 1, ", Text.Middle([Name], 6, 5)}), type text),

    // 3. Filter Excel files
    ExcelFiles = Table.SelectRows(#"Added Custom Column", each Text.EndsWith([Extension], ".xlsx") or Text.EndsWith([Extension], ".xls")),

    // 4. Extract Workbook
    WithWorkbook = Table.AddColumn(ExcelFiles, "Workbook", each try Excel.Workbook([Content], true) otherwise null),
    #"Removed Other Columns" = Table.SelectColumns(WithWorkbook,{"Workbook", "Date"}),
    ValidFiles = Table.SelectRows(#"Removed Other Columns", each [Workbook] <> null),

    // 5. Process each file and keep its Date inside its tables
    WithCleanedTables = Table.AddColumn(ValidFiles, "CleanedTables", each let
                fileDate = [Date],
                Sheets = Table.SelectRows([Workbook], each [Kind] = "Sheet" and [Hidden] <> true),
                
                WithPromoted = Table.TransformColumns(Sheets,{"Data",
                (tbl) =>
                        let
                            rows = Table.ToRows(tbl),
                            // Define your various header lists for detection
                            expectedHeaders = {"GL Code", "GL Item", "Balance for CY", "Balance for LY"},
                            expectedHeaders2 = {"GL Codes", "GL Item", "Balance for CY", "Balance for LY"},
                            expectedHeaders3 = {"GL Codes", "GL Items", "Balance for CY", "Balance for LY"},
                            expectedHeaders4 = {"GL Codes", "GL Items", "Balances for CY", "Balances for LY"},
                            
                            // Find the header row index using your chained logic
                            attempt1 = List.PositionOf(rows, expectedHeaders),
                            attempt2 = if attempt1 <> -1 then attempt1 else List.PositionOf(rows, expectedHeaders2),
                            attempt3 = if attempt2 <> -1 then attempt2 else List.PositionOf(rows, expectedHeaders3),
                            headerIndex = if attempt3 <> -1 then attempt3 else List.PositionOf(rows, expectedHeaders4),
                            
                            cleaned =
                                if headerIndex <> -1 then
                                    let
                                        skipped = Table.Skip(tbl, headerIndex),
                                        promoted = Table.PromoteHeaders(skipped, [IgnoreNulls = true])
                                        // NO DYNAMIC RENAMING HERE
                                    in
                                        promoted // Returns the promoted table as is
                                else null // If header not found, return null for this sheet
                        in cleaned}),

            Combined = Table.Combine(List.RemoveNulls(WithPromoted[Data]))
        in
            Combined
    ),
    #"Removed Other Columns1" = Table.SelectColumns(WithCleanedTables,{"Date", "CleanedTables"}),
    Custom1 = List.Distinct(List.Combine(Table.TransformColumns(#"Removed Other Columns1", {"CleanedTables", each Table.ColumnNames(_)})[CleanedTables])),
    #"Expanded CleanedTables" = Table.ExpandTableColumn(Custom1, "CleanedTables", {"GL Code", "GL Item", "Balance for CY", "Balance for LY"}, {"GL Code", "GL Item", "Balance for CY", "Balance for LY"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded CleanedTables",{{"Date", type date}, {"Balance for CY", type number}, {"Balance for LY", type number}}),
    #"Added Last Year Date" = Table.AddColumn(#"Changed Type", "Last Year Date", each Date.AddYears([Date], -1)),

    // Prepare Current Year Data
    #"Current Year Data" = Table.SelectColumns(#"Added Last Year Date", {"GL Code", "GL Item", "Balance for CY", "Date"}),
    #"Rename CY Columns" = Table.RenameColumns(#"Current Year Data", {{"Balance for CY", "Balance"}, {"Date", "Reporting Date"}}),
    #"Add CY Type" = Table.AddColumn(#"Rename CY Columns", "Year Type", each "Current Year", type text),
    
    // Prepare Last Year Data
    #"Last Year Data" = Table.SelectColumns(#"Added Last Year Date", {"GL Code", "GL Item", "Balance for LY", "Last Year Date"}),
    #"Rename LY Columns" = Table.RenameColumns(#"Last Year Data", {{"Balance for LY", "Balance"}, {"Last Year Date", "Reporting Date"}}),
    #"Add LY Type" = Table.AddColumn(#"Rename LY Columns", "Year Type", each "Last Year", type text),

    // Combine them
    #"Combined Balances" = Table.Combine({#"Add CY Type", #"Add LY Type"}),

    // MODIFIED SORTING STEP TO PUT ALL CURRENT YEAR DATA ON TOP
    #"Sorted Combined Balances" = Table.Sort(#"Combined Balances",{
        {"Year Type", Order.Ascending},
        {"Reporting Date", Order.Descending},
        {"GL Code", Order.Ascending},
        {"GL Item", Order.Ascending}
    })
in
    #"Sorted Combined Balances"
