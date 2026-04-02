// Query: 05_ddt_line_builder.pq
// Purpose: Core DDT line construction logic.
//          Joins enriched operational data with traceability register.
//          Builds all DDT line fields required for ERP XML import:
//          - Unique DDT key
//          - Main product line description and quantity (species-specific)
//          - Offal line (conditional)
//          - Processing flag (conditional)
//          - Transport responsibility
//          - Traceability reference
//          Normalizes all fields for ERP XML compatibility.

let
    // --- STEP 1: Join operational data with traceability register ---
    // Joins on TraceabilityCode (built identically in both queries)
    // LeftOuter: preserves all operational records
    Source = Table.NestedJoin(
        Deliveries, {"TraceabilityCode"},
        TraceabilityRegister, {"TRACEABILITY_CODE"},
        "TraceData",
        JoinKind.LeftOuter
    ),

    // --- STEP 2: Expand all traceability fields ---
    #"Expand Traceability" = Table.ExpandTableColumn(
        Source, "TraceData",
        {"AnimalTag", "Breed", "Sex", "BirthDate", "BirthCountry",
         "OriginFarmCode", "OriginFarmName", "OriginCountry",
         "OriginProvince", "DepartureCountry", "BatchNumber", "Lot",
         "SlaughterNumber", "CategoryCode", "Category", "WeightKg",
         "ClassificationCode", "Classification", "SlaughterType",
         "SlaughterDate", "Form4"},
        {"Trace.AnimalTag", "Trace.Breed", "Trace.Sex", "Trace.BirthDate",
         "Trace.BirthCountry", "Trace.OriginFarmCode", "Trace.OriginFarmName",
         "Trace.OriginCountry", "Trace.OriginProvince", "Trace.DepartureCountry",
         "Trace.BatchNumber", "Trace.Lot", "Trace.SlaughterNumber",
         "Trace.CategoryCode", "Trace.Category", "Trace.WeightKg",
         "Trace.ClassificationCode", "Trace.Classification",
         "Trace.SlaughterType", "Trace.SlaughterDate", "Trace.Form4"}
    ),

    // --- STEP 3: Add row index for unique DDT key generation ---
    #"Add Index" = Table.AddIndexColumn(#"Expand Traceability", "RowIndex", 1, 1, Int64.Type),
    #"Index as Text" = Table.TransformColumnTypes(#"Add Index", {{"RowIndex", type text}}),
    #"Client Code as Text" = Table.TransformColumnTypes(#"Index as Text", {{"Master.ClientCode", type text}}),

    // --- STEP 4: Build unique DDT key ---
    // Format: ClientCode-ExitDate(yyyyMMdd)-RowIndex
    // Guarantees uniqueness per line per document per day
    #"Add DDTKey" = Table.AddColumn(
        #"Client Code as Text",
        "DDTKey",
        each Text.From([Master.ClientCode])
            & "-"
            & Date.ToText(Date.From([ExitDate]), "yyyyMMdd")
            & "-"
            & Text.From([RowIndex])
    ),

    // --- STEP 5: Filter placeholder ---
    #"Filter Rows" = Table.SelectRows(#"Add DDTKey", each true),

    // --- STEP 6: Build SlaughterNumber/Lot composite field ---
    // Used in traceability description on DDT
    #"Add SlaughterLot" = Table.AddColumn(
        #"Filter Rows",
        "SlaughterNumber_Lot",
        each if [SlaughterNumber] <> null and [SlaughterNumber] <> ""
             then Text.From([SlaughterNumber]) & "/" & Text.From([Lot])
             else Text.From([Lot])
    ),

    // --- STEP 7: Resolve animal tag for XML ---
    // Prefer traceability tag; fall back to batch number if missing
    #"Add AnimalTagXML" = Table.AddColumn(
        #"Add SlaughterLot",
        "AnimalTagXML",
        each if [Trace.AnimalTag] <> null and [Trace.AnimalTag] <> ""
             then [Trace.AnimalTag]
             else [Trace.BatchNumber]
    ),

    // --- STEP 8: Build main product line description ---
    // Species-specific cutting format:
    // Bovine/Buffalo → QUARTERS (4 per head)
    // Swine/Equine   → SIDES (2 per head)
    // Other          → WHOLE SLAUGHTERED (1 per head)
    #"Add MainDescription" = Table.AddColumn(
        #"Add AnimalTagXML",
        "MainLineDescription",
        each let s = Text.Upper([Species]) in
            if s = "BOVINE" or s = "BUFFALO" then "QUARTERS " & [Species]
            else if s = "SWINE" or s = "EQUINE" then "SIDES " & [Species]
            else [Species] & " WHOLE SLAUGHTERED"
    ),

    // --- STEP 9: Build main product quantity ---
    // Derived from species cutting factor × head count
    #"Add MainQuantity" = Table.AddColumn(
        #"Add MainDescription",
        "MainLineQuantity",
        each let s = Text.Upper([Species]) in
            if s = "BOVINE" or s = "BUFFALO" then 4 * [HeadCount]
            else if s = "SWINE" or s = "EQUINE" then 2 * [HeadCount]
            else 1 * [HeadCount]
    ),

    // --- STEP 10: Build offal line description (conditional) ---
    // Only populated when offal is flagged in the operational record
    #"Add OfalDescription" = Table.AddColumn(
        #"Add MainQuantity",
        "OfalLineDescription",
        each if [OffalFlag] <> null and Text.Length([OffalFlag]) > 0
             then "OFFAL " & [Species]
             else null
    ),

    // --- STEP 11: Build offal quantity (conditional) ---
    #"Add OfalQuantity" = Table.AddColumn(
        #"Add OfalDescription",
        "OfalLineQuantity",
        each let r = Text.Trim(Text.From([OffalFlag] & ""))
             in if r <> "" then [HeadCount] else null
    ),

    // --- STEP 12: Processing flag ---
    // Flags lines routed to internal secondary processing unit
    #"Add ProcessingFlag" = Table.AddColumn(
        #"Add OfalQuantity",
        "ProcessingFlag",
        each if [DeliveryType] <> null
                and Text.Contains(Text.Upper([DeliveryType]), "PROCESSING")
             then "PROCESSING"
             else null
    ),

    // --- STEP 13: Static DDT fields ---
    #"Add Reason"       = Table.AddColumn(#"Add ProcessingFlag", "Reason", each "GOODS RETURNED FOR PROCESSING"),
    #"Add GoodsAspect"  = Table.AddColumn(#"Add Reason", "GoodsAspect", each "FRESH MEAT"),

    // --- STEP 14: Derive transport responsibility ---
    // CONSIGNEE = client/internal unit collects (PICKUP or PROCESSING)
    // CONSIGNOR  = company delivers
    #"Add Transport" = Table.AddColumn(
        #"Add GoodsAspect",
        "TransportResponsibility",
        each let d = if [DeliveryType] = null then "" else Text.Upper([DeliveryType])
             in if Text.Contains(d, "PICKUP") or Text.Contains(d, "PROCESSING")
                then "CONSIGNEE"
                else "CONSIGNOR"
    ),

    // --- STEP 15: Add DDT reference note ---
    #"Add Note" = Table.AddColumn(
        #"Add Transport",
        "Note",
        each "REF. DDT N. " & [InboundRef]
    ),

    // --- STEP 16: Rename columns for XML field compatibility ---
    #"Rename Columns" = Table.RenameColumns(
        #"Add Note",
        {
            {"Master.City", "Master.ClientCity"},
            {"Master.Province", "Master.ClientProvince"},
            {"Master.VAT", "Master.ClientVAT"},
            {"Master.Mobile", "Master.ClientMobile"},
            {"Master.Phone", "Master.ClientPhone"},
            {"SlaughterNumber_Lot", "SlaughterNumber_Lot"},
            {"Trace.WeightKg", "Trace.WeightKG"},
            {"Master.ClientCode", "Master.ClientCode"}
        }
    ),

    // --- STEP 17: Final phone/address normalization ---
    #"Clean Address"  = Table.ReplaceValue(#"Rename Columns", "°", " ", Replacer.ReplaceText, {"Master.Address"}),
    #"Clean Mobile 1" = Table.ReplaceValue(#"Clean Address", "/", "", Replacer.ReplaceText, {"Master.ClientMobile"}),
    #"Clean Mobile 2" = Table.ReplaceValue(#"Clean Mobile 1", " ", "", Replacer.ReplaceText, {"Master.ClientMobile"}),
    #"Clean Phone 1"  = Table.ReplaceValue(#"Clean Mobile 2", "/", "", Replacer.ReplaceText, {"Master.ClientPhone"}),
    #"Clean Phone 2"  = Table.ReplaceValue(#"Clean Phone 1", " ", "", Replacer.ReplaceText, {"Master.ClientPhone"}),

    // --- STEP 18: Parse time as text for XML compatibility ---
    #"Parse Time" = Table.TransformColumnTypes(
        #"Clean Phone 2",
        {{"Time", type text}}
    )
in
    #"Parse Time"
