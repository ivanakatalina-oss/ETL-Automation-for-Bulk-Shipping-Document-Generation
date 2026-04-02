// Query: 04_extraction_and_enrichment.pq
// Purpose: Extract daily operational data from source workbook.
//          Normalize client names (same rules as 02_mapping_table.pq).
//          Generate traceability code (same format as 03_traceability_register.pq).
//          Enrich with ERP fiscal data via Left Outer Join on client name.
//          Derive delivery type from free-text notes field.
//
// NOTE: DeliveryType field is a candidate for dropdown replacement in OMS.
//       Current values derived from free-text: PICKUP, PROCESSING, DELIVERY.

let
    // --- STEP 1: Load daily operational table ---
    Source = Excel.CurrentWorkbook(){[Name="OperationalData"]}[Content],

    // --- STEP 2: Parse exit date with Italian locale ---
    #"Parse ExitDate" = Table.TransformColumnTypes(
        Source,
        {{"ExitDate", type date}}, "it-IT"
    ),

    // --- STEP 3: Build traceability code ---
    // Format: CLIENTNAME-SlaughterNumber-Lot-EntryDate(yyyyMMdd)
    // Null-safe for each component
    #"Add TraceabilityCode" = Table.AddColumn(
        #"Parse ExitDate",
        "TraceabilityCode",
        each Text.Upper([ClientName])
            & "-" & (if [SlaughterNumber] <> null and [SlaughterNumber] <> ""
                     then Text.From([SlaughterNumber]) else " ")
            & (if [Lot] <> null and [Lot] <> ""
               then "-" & Text.From([Lot]) else "")
            & "-" & (if [EntryDate] <> null
                     then Date.ToText(Date.From([EntryDate]), "yyyyMMdd")
                     else "NO_DATE")
    ),

    // --- STEP 4: Parse entry date ---
    #"Parse EntryDate" = Table.TransformColumnTypes(
        #"Add TraceabilityCode",
        {{"EntryDate", type date}}
    ),

    // --- STEP 5: Normalize client name ---
    // Applies same rules as 02_mapping_table.pq
    // to ensure consistent join key
    #"Remove Ampersand" = Table.ReplaceValue(#"Parse EntryDate", "&", "E", Replacer.ReplaceText, {"ClientName"}),
    #"Remove Quotes"    = Table.ReplaceValue(#"Remove Ampersand", """", "", Replacer.ReplaceText, {"ClientName"}),
    #"Norm AZ AGR"      = Table.ReplaceValue(#"Remove Quotes", "AZ.AGR.", "AZ AGR", Replacer.ReplaceText, {"ClientName"}),
    #"Norm MAC AGR"     = Table.ReplaceValue(#"Norm AZ AGR", "MAC.AGR.", "MAC AGR", Replacer.ReplaceText, {"ClientName"}),
    #"Remove Dots"      = Table.ReplaceValue(#"Norm MAC AGR", ".", "", Replacer.ReplaceText, {"ClientName"}),

    // --- STEP 6: Uppercase species and notes ---
    #"Uppercase Fields" = Table.TransformColumns(
        #"Remove Dots",
        {{"Species", Text.Upper, type text}, {"Notes", Text.Upper, type text}}
    ),

    // --- STEP 7: Join with mapping table on client name ---
    // LeftOuter: preserves all operational records
    // even if no ERP match exists
    #"Join Mapping" = Table.NestedJoin(
        #"Uppercase Fields", {"ClientName"},
        #"MAPPING_TABLE",    {"OPERATIONAL_NAME"},
        "MasterData",
        JoinKind.LeftOuter
    ),

    // --- STEP 8: Expand ERP fiscal and address fields ---
    #"Expand Master" = Table.ExpandTableColumn(
        #"Join Mapping", "MasterData",
        {"FISCAL_ID", "Client.Code", "Client.Name", "Client.City",
         "Client.Province", "Client.VAT", "Client.FiscalCode",
         "Client.Address", "Client.Mobile", "Client.Phone"},
        {"Master.FiscalID", "Master.ClientCode", "Master.ClientName",
         "Master.City", "Master.Province", "Master.VAT",
         "Master.FiscalCode", "Master.Address", "Master.Mobile", "Master.Phone"}
    ),

    // --- STEP 9: Derive delivery type from notes field ---
    // PICKUP   = client collects in person at facility
    // PROCESSING = routed to secondary processing unit (internal)
    // Default  = standard delivery, city from ERP registry
    // BACKLOG: replace with dropdown field in OMS v1.0
    #"Add DeliveryType" = Table.AddColumn(
        #"Expand Master",
        "DeliveryType",
        each if [Notes] <> null and Text.Contains(Text.From([Notes]), "PICKUP", Comparer.OrdinalIgnoreCase)
             then "PICKUP"
        else if [Notes] <> null and Text.Contains(Text.From([Notes]), "PROCESSING", Comparer.OrdinalIgnoreCase)
             then "PROCESSING"
        else [Master.City]
    ),

    // --- STEP 10: Uppercase delivery type, parse time as text ---
    #"Uppercase DeliveryType" = Table.TransformColumns(
        #"Add DeliveryType",
        {{"DeliveryType", Text.Upper, type text}}
    ),
    #"Parse Time" = Table.TransformColumnTypes(
        #"Uppercase DeliveryType",
        {{"Time", type text}}
    )
in
    #"Parse Time"
