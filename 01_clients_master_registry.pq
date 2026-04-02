// Query: 01_clients_master_registry.pq
// Purpose: Load client master registry from external workbook.
//          Build unified fiscal ID (VAT preferred over fiscal code).
//          Clean address and phone fields for ERP compatibility.

let
    // --- STEP 1: Load client registry workbook ---
    Source = Excel.Workbook(
        File.Contents("[PATH_TO_WORKBOOK]\Clients.xlsx"), 
        null, true
    ),
    Clients_Sheet = Source{[Item="Clients", Kind="Sheet"]}[Data],

    // --- STEP 2: Promote headers ---
    #"Promote Headers" = Table.PromoteHeaders(
        Clients_Sheet, [PromoteAllScalars=true]
    ),

    // --- STEP 3: Rename fiscal code column for consistency ---
    #"Rename FiscalCode" = Table.RenameColumns(
        #"Promote Headers",
        {{"Codice fiscale", "FiscalCode"}}
    ),

    // --- STEP 4: Build unified fiscal ID ---
    // Priority: VAT number > Fiscal code > "MISSING"
    // Used as primary key for all downstream joins
    #"Add UnifiedFiscalID" = Table.AddColumn(
        #"Rename FiscalCode", 
        "UnifiedFiscalID", 
        each if [VAT] <> null and [VAT] <> "" then [VAT]
             else if [FiscalCode] <> null and [FiscalCode] <> "" then [FiscalCode]
             else "MISSING"
    ),

    // --- STEP 5: Reorder columns ---
    #"Reorder Columns" = Table.ReorderColumns(
        #"Add UnifiedFiscalID",
        {"", "Code", "Name", "City", "Province", 
         "VAT", "FiscalCode", "UnifiedFiscalID", 
         "Address", "Mobile", "Phone"}
    ),

    // --- STEP 6: Clean unified fiscal ID (remove invisible chars) ---
    #"Clean FiscalID" = Table.TransformColumns(
        #"Reorder Columns",
        {{"UnifiedFiscalID", Text.Clean, type text}}
    ),

    // --- STEP 7: Normalize address field ---
    // Remove characters incompatible with ERP XML import
    #"Clean Address 1" = Table.ReplaceValue(#"Clean FiscalID", "/", "", Replacer.ReplaceText, {"Address"}),
    #"Clean Address 2" = Table.ReplaceValue(#"Clean Address 1", "&", "", Replacer.ReplaceText, {"Address"}),
    #"Clean Address 3" = Table.ReplaceValue(#"Clean Address 2", "°", "", Replacer.ReplaceText, {"Address"}),

    // --- STEP 8: Normalize phone numbers ---
    // Remove slashes and spaces for ERP field compatibility
    #"Clean Mobile 1" = Table.ReplaceValue(#"Clean Address 3", "/", "", Replacer.ReplaceText, {"Mobile"}),
    #"Clean Mobile 2" = Table.ReplaceValue(#"Clean Mobile 1", " ", "", Replacer.ReplaceText, {"Mobile"}),
    #"Clean Phone 1"  = Table.ReplaceValue(#"Clean Mobile 2", "/", "", Replacer.ReplaceText, {"Phone"}),
    #"Clean Phone 2"  = Table.ReplaceValue(#"Clean Phone 1", " ", "", Replacer.ReplaceText, {"Phone"}),

    // --- STEP 9: Final filter placeholder (all rows) ---
    #"Final Output" = Table.SelectRows(#"Clean Phone 2", each true)
in
    #"Final Output"
