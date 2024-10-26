let
    Source = SharePoint.Tables("https://iomint.sharepoint.com/sites/HELIOS", [ApiVersion = 15]),
    SharePointList = Source{[Id = "dbcb9786-eb84-4330-8a9f-91d4d226d9ba"]}[Items],
    SelectedColumns = Table.SelectColumns(
        SharePointList,
        {"Title", "DocumentNumber", "DocumentType", "IssuanceDate", "ExpirationDate", "Status", "PrimaryDocument"}
    ),
    Filters = Table.SelectRows(
        SelectedColumns, each ([PrimaryDocument] = true and [DocumentType] = "Residence Permit")
    ),
    // When we use flow to create an item in SP with the date column, it gets converted to the site timezone when the Date is saved.
    // Because we are using date and time format and your site is in PST timezone, in some cases, this conversion pushes the Date a day behind.
    // To fix this issue, we add 7 hours (time zone offset hours between UTC and PST) when quering the table data
    AddedHours = Table.AddColumn(Filters, "DOCUMENT ISSUANCE DATE", each [IssuanceDate] + #duration(0, 7, 0, 0)),
    AddedHours2 = Table.AddColumn(
        AddedHours, "DOCUMENT EXPIRATION DATE", each [ExpirationDate] + #duration(0, 7, 0, 0)
    ),
    RemovedColumns = Table.RemoveColumns(AddedHours2, {"IssuanceDate", "ExpirationDate", "PrimaryDocument"}),
    ChangedType = Table.TransformColumnTypes(
        RemovedColumns,
        {
            {"Title", type text},
            {"DocumentNumber", type text},
            {"DocumentType", type text},
            {"Status", type text},
            {"DOCUMENT ISSUANCE DATE", type datetime},
            {"DOCUMENT EXPIRATION DATE", type datetime}
        }
    ),
    RenamedColumns = Table.RenameColumns(
        ChangedType,
        {
            {"Title", "IndividualID"},
            {"DocumentNumber", "DOCUMENT NUMBER"},
            {"DocumentType", "DOCUMENT TYPE"},
            {"Status", "DOCUMENT STATUS"}
        }
    )
in
    RenamedColumns
