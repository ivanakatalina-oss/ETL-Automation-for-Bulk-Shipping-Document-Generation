// Query: 02_mapping_table.pq
// Purpose: Load operational-to-ERP name mapping table.
//          Apply same normalization rules as operational data
//          to ensure consistent join on client name.
//          Enrich with full fiscal and address data via join on UnifiedFiscalID.

let
    // --- STEP 1: Load mapping workbook ---
    // Maps operational client names (as entered by operators)
    // to ERP fiscal identifiers
    Source = Excel.Workbook(
        File.Contents("[PATH_TO_WORKBOOK]\MappingTable.xlsx"), 
        null, true
    ),
    Sheet = Source{[Item="Sheet1", Kind="Sheet"]}[Data],

    // --- STEP 2: Set column types ---
    #"Set Types" = Table.TransformColumnTypes(
        Sheet,
        {{"Column1", type text}, {"Column2", type text}}
    ),

    // --- STEP 3: Promote headers ---
    #"Promote Headers" = Table.PromoteHeaders(
        #"Set Types", [PromoteAllScalars=true]
    ),

    // --- STEP 4: Type final columns ---
    #"Type Columns" = Table.TransformColumnTypes(
        #"Promote Headers",
        {{"OPERATIONAL_NAME", type text}, {"FISCAL_ID", type text}}
    ),

    // --- STEP 5: Normalize operational name ---
    // Must match normalization applied in 04_extraction_and_enrichment.pq
    #"Remove Quotes"     = Table.ReplaceValue(#"Type Columns", """", "", Replacer.ReplaceText, {"OPERATIONAL_NAME"}),
    #"Replace Ampersand" = Table.ReplaceValue(#"Remove Quotes", "&", "E", Replacer.ReplaceText, {"OPERATIONAL_NAME"}),
    #"Normalize AZ AGR"  = Table.ReplaceValue(#"Replace Ampersand", "AZ.AGR.", "AZ AGR", Replacer.ReplaceText, {"OPERATIONAL_NAME"}),
    #"Normalize MAC AGR" = Table.ReplaceValue(#"Normalize AZ AGR", "MAC.AGR.", "MAC AGR", Replacer.ReplaceText, {"OPERATIONAL_NAME"}),
    #"Remove Dots"       = Table.ReplaceValue(#"Normalize MAC AGR", ".", "", Replacer.ReplaceText, {"OPERATIONAL_NAME"}),
    #"Fix AGR C"         = Table.ReplaceValue(#"Remove Dots", "AZ AGRC", "AZ AGR C", Replacer.ReplaceText, {"OPERATIONAL_NAME"}),
    #"Trim Name"         = Table.TransformColumns(#"Fix AGR C", {{"OPERATIONAL_NAME", Text.Trim, type text}}),

    // --- STEP 6: Clean fiscal ID ---
    #"Clean FiscalID" = Table.TransformColumns(#"Trim Name", {{"FISCAL_ID", Text.Clean, type text}}),

    // --- STEP 7: Join with client master registry on UnifiedFiscalID ---
    // LeftOuter: preserves all mapping entries even if ERP record is missing
    #"Join Clients" = Table.NestedJoin(
        #"Clean FiscalID", {"FISCAL_ID"},
        Clients,           {"UnifiedFiscalID"},
        "ClientData",
        JoinKind.LeftOuter
    ),

    // --- STEP 8: Expand client fiscal and address fields ---
    #"Expand Client Data" = Table.ExpandTableColumn(
        #"Join Clients", "ClientData",
        {"Col1", "Code", "Name", "City", "Province", "Col2",
         "VAT", "FiscalCode", "UnifiedFiscalID",
         "Address", "Mobile", "Phone", "ClientStatus", "Col14"},
        {"Client.Col1", "Client.Code", "Client.Name", "Client.City",
         "Client.Province", "Client.Col2", "Client.VAT", "Client.FiscalCode",
         "Client.UnifiedFiscalID", "Client.Address", "Client.Mobile",
         "Client.Phone", "Client.Status", "Client.Col14"}
    ),

    // --- STEP 9: Final filter placeholder ---
    #"Final Output" = Table.SelectRows(#"Expand Client Data", each true)
in
    #"Final Output"
