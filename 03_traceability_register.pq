// Query: 03_traceability_register.pq
// Purpose: Load all traceability files from the regulatory data folder.
//          Combine into a single dataset.
//          Generate traceability code matching the format used in
//          04_extraction_and_enrichment.pq for downstream join.

let
    // --- STEP 1: Load all files from regulatory data folder ---
    // Folder contains periodic export files from the national 
    // traceability system (one file per batch/period)
    Source = Folder.Files("[PATH_TO_FOLDER]\TraceabilityData"),

    // --- STEP 2: Filter out hidden files ---
    #"Filter Hidden" = Table.SelectRows(
        Source, 
        each [Attributes]?[Hidden]? <> true
    ),

    // --- STEP 3: Apply custom transform function to each file ---
    #"Apply Transform" = Table.AddColumn(
        #"Filter Hidden", 
        "TransformedFile", 
        each #"Transform File"([Content])
    ),

    // --- STEP 4: Rename source name column ---
    #"Rename Source" = Table.RenameColumns(
        #"Apply Transform", 
        {"Name", "Source.Name"}
    ),

    // --- STEP 5: Select relevant columns only ---
    #"Select Columns" = Table.SelectColumns(
        #"Rename Source", 
        {"Source.Name", "TransformedFile"}
    ),

    // --- STEP 6: Expand all file data into flat table ---
    #"Expand Files" = Table.ExpandTableColumn(
        #"Select Columns", "TransformedFile",
        Table.ColumnNames(#"Transform File"(#"Sample File"))
    ),

    // --- STEP 7: Set column types ---
    #"Set Types" = Table.TransformColumnTypes(
        #"Expand Files",
        {
            {"Source.Name", type text},
            {"AnimalTag", type text},
            {"HeadCount", Int64.Type},
            {"Species", type text},
            {"Breed", type text},
            {"Sex", type text},
            {"BirthDate", type date},
            {"OriginFarmCode", type text},
            {"OriginFarmName", type text},
            {"BirthCountry", type text},
            {"OriginFarmCode2", type text},
            {"OriginFarmName2", type text},
            {"OriginCountry", type text},
            {"OriginProvince", type text},
            {"Owner", type text},
            {"DepartureCountry", type text},
            {"BatchNumber", type any},
            {"Lot", type any},
            {"SlaughterNumber", Int64.Type},
            {"CategoryCode", type text},
            {"Category", type text},
            {"WeightKg", type any},
            {"ClassificationCode", type any},
            {"Classification", type any},
            {"SlaughterType", type text},
            {"SlaughterDate", type date},
            {"AntemortemDate", type date},
            {"Form4", Int64.Type},
            {"Transporter", type any},
            {"Organic", type text},
            {"Notes", type any}
        }
    ),

    // --- STEP 8: Filter placeholder ---
    #"Filter Rows" = Table.SelectRows(#"Set Types", each true),

    // --- STEP 9: Build traceability code ---
    // Format: OWNER-SlaughterNumber-Lot-SlaughterDate(yyyyMMdd)
    // Must match format generated in 04_extraction_and_enrichment.pq
    // for the join key to work correctly
    #"Add TraceabilityCode" = Table.AddColumn(
        #"Filter Rows",
        "TRACEABILITY_CODE",
        each Text.Upper([Owner])
            & "-" & (if [SlaughterNumber] <> null and [SlaughterNumber] <> ""
                     then Text.From([SlaughterNumber]) else " ")
            & (if [Lot] <> null and [Lot] <> ""
               then "-" & Text.From([Lot]) else "")
            & "-" & (if [SlaughterDate] <> null
                     then Date.ToText(Date.From([SlaughterDate]), "yyyyMMdd")
                     else "NO_DATE")
    )
in
    #"Add TraceabilityCode"
