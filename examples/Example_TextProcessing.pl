use lib '../lib';
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;

# =========================================================================
# Text Processing Examples for Aspose.Cells Cloud SDK for Perl
# =========================================================================
# Demonstrates: search, replace, add/convert/extract text, translate,
# import data, data cleansing/deduplication/fill, smart markers,
# clear formats, export, and more.
# =========================================================================

my $config = AsposeCellsCloud::Configuration->new(
    client_id     => $ENV{'CellsCloudClientId'},
    client_secret => $ENV{'CellsCloudClientSecret'}
);
my $api = AsposeCellsCloud::CellsApi->new(
    AsposeCellsCloud::ApiClient->new($config)
);

my $SRC = 'EmployeeSalesSummary.xlsx';

sub save_file {
    my ($data, $path) = @_;
    open(my $fh, '>', $path) or die "Cannot open $path: $!";
    binmode($fh);
    print $fh $data;
    close($fh);
    print "  -> Saved $path\n";
}

sub print_result_files {
    my ($r) = @_;
    return unless $r && defined $r->{files} && ref($r->{files}) eq 'ARRAY';
    print "  Files: " . scalar(@{$r->{files}}) . " file(s).\n";
}

# =========================================================================
# Section 1: Search Text in Uploaded Files
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 1: Search Text in Uploaded Files\n";
print "=" x 70, "\n";

# --- 1a. Search entire workbook ---
print "\n[1a] Search for 'Sales' in uploaded workbook\n";
{
    use AsposeCellsCloud::Request::PostSearchRequest;
    my $req = AsposeCellsCloud::Request::PostSearchRequest->new();
    $req->{file}                    = $SRC;
    $req->{text}                    = 'Sales';
    $req->{check_excel_restriction} = 'true';
    my $r = $api->post_search(request => $req);
    if ($r && ref($r) eq 'ARRAY') {
        print "  Found " . scalar(@$r) . " occurrence(s).\n";
        foreach my $item (@$r) {
            print "    - Sheet: " . ($item->{worksheet} // '?')
                . ", Cell: " . ($item->{cell_name} // '?')
                . ", Text: " . ($item->{text} // '') . "\n";
        }
    }
}

# --- 1b. Search in specific worksheet only ---
print "\n[1b] Search for 'Employee' in Sheet1 only\n";
{
    use AsposeCellsCloud::Request::PostSearchRequest;
    my $req = AsposeCellsCloud::Request::PostSearchRequest->new();
    $req->{file}      = $SRC;
    $req->{text}      = 'Employee';
    $req->{sheetname} = 'Sheet1';
    my $r = $api->post_search(request => $req);
    if ($r && ref($r) eq 'ARRAY') {
        print "  Found " . scalar(@$r) . " occurrence(s).\n";
    }
}

# =========================================================================
# Section 2: Search Text in Cloud-Stored Files
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 2: Search Text in Cloud-Stored Files\n";
print "=" x 70, "\n";

# --- 2a. Search workbook in cloud ---
print "\n[2a] Search workbook for 'Sales'\n";
{
    use AsposeCellsCloud::Request::PostWorkbooksTextSearchRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbooksTextSearchRequest->new();
    $req->{name}   = $SRC;
    $req->{text}   = 'Sales';
    $req->{folder} = 'Cells';
    my $r = $api->post_workbooks_text_search(request => $req);
    if ($r && defined $r->{text_items}) {
        my $items = $r->{text_items};
        print "  Found " . scalar(@$items) . " occurrence(s).\n"
            if ref($items) eq 'ARRAY';
    }
}

# --- 2b. Search worksheet in cloud ---
print "\n[2b] Search specific worksheet for 'Employee'\n";
{
    use AsposeCellsCloud::Request::PostWorksheetTextSearchRequest;
    my $req = AsposeCellsCloud::Request::PostWorksheetTextSearchRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{text}       = 'Employee';
    $req->{folder}     = 'Cells';
    my $r = $api->post_worksheet_text_search(request => $req);
    print "  -> Done\n";
}

# --- 2c. Search content in remote spreadsheet ---
print "\n[2c] Search content in remote spreadsheet\n";
{
    use AsposeCellsCloud::Request::SearchContentInRemoteSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::SearchContentInRemoteSpreadsheetRequest->new();
    $req->{name}       = $SRC;
    $req->{text}       = 'Sales';
    $req->{worksheet}  = 'Sheet1';
    $req->{folder}     = 'Cells';
    my $r = $api->search_content_in_remote_spreadsheet(request => $req);
    print "  -> Done\n";
}

# --- 2d. Search ALL text items in remote spreadsheet ---
print "\n[2d] Search ALL text items in remote spreadsheet\n";
{
    use AsposeCellsCloud::Request::SearchAllTextItemsInRemoteSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::SearchAllTextItemsInRemoteSpreadsheetRequest->new();
    $req->{name}   = $SRC;
    $req->{folder} = 'Cells';
    my $r = $api->search_all_text_items_in_remote_spreadsheet(request => $req);
    print "  -> Done\n";
}

# --- 2e. Search ALL text items in local spreadsheet ---
print "\n[2e] Search ALL text items in local spreadsheet\n";
{
    use AsposeCellsCloud::Request::SearchSpreadsheetAllTextItemsRequest;
    my $req = AsposeCellsCloud::Request::SearchSpreadsheetAllTextItemsRequest->new();
    $req->{spreadsheet} = $SRC;
    my $r = $api->search_spreadsheet_all_text_items(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 3: Replace Text in Uploaded Files
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 3: Replace Text in Uploaded Files\n";
print "=" x 70, "\n";

# --- 3a. Replace in entire workbook ---
print "\n[3a] Replace 'Sales' -> 'Revenue' in uploaded workbook\n";
{
    use AsposeCellsCloud::Request::PostReplaceRequest;
    my $req = AsposeCellsCloud::Request::PostReplaceRequest->new();
    $req->{file}    = $SRC;
    $req->{text}    = 'Sales';
    $req->{newtext} = 'Revenue';
    my $r = $api->post_replace(request => $req);
    print_result_files($r);
}

# --- 3b. Replace in specific worksheet only ---
print "\n[3b] Replace 'Employee' -> 'Staff' in Sheet1 only\n";
{
    use AsposeCellsCloud::Request::PostReplaceRequest;
    my $req = AsposeCellsCloud::Request::PostReplaceRequest->new();
    $req->{file}      = $SRC;
    $req->{text}      = 'Employee';
    $req->{newtext}   = 'Staff';
    $req->{sheetname} = 'Sheet1';
    my $r = $api->post_replace(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 4: Replace Text in Cloud-Stored Files
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 4: Replace Text in Cloud-Stored Files\n";
print "=" x 70, "\n";

# --- 4a. Replace in cloud workbook ---
print "\n[4a] Replace 'Sales' -> 'Revenue' in cloud workbook\n";
{
    use AsposeCellsCloud::Request::PostWorkbookTextReplaceRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookTextReplaceRequest->new();
    $req->{name}      = $SRC;
    $req->{old_value} = 'Sales';
    $req->{new_value} = 'Revenue';
    $req->{folder}    = 'Cells';
    my $r = $api->post_workbook_text_replace(request => $req);
    print "  Matches: " . ($r->{matches} // 'N/A') . "\n"
        if $r && defined $r->{matches};
}

# --- 4b. Replace in cloud worksheet ---
print "\n[4b] Replace 'Employee' -> 'Staff' in cloud worksheet\n";
{
    use AsposeCellsCloud::Request::PostWorksheetTextReplaceRequest;
    my $req = AsposeCellsCloud::Request::PostWorksheetTextReplaceRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{old_value}  = 'Employee';
    $req->{new_value}  = 'Staff';
    $req->{folder}     = 'Cells';
    my $r = $api->post_worksheet_text_replace(request => $req);
    print "  -> Done\n";
}

# --- 4c. Replace content in remote spreadsheet ---
print "\n[4c] Replace content in remote spreadsheet\n";
{
    use AsposeCellsCloud::Request::ReplaceContentInRemoteSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::ReplaceContentInRemoteSpreadsheetRequest->new();
    $req->{name}       = $SRC;
    $req->{text}       = 'Sales';
    $req->{new_text}   = 'Revenue';
    $req->{worksheet}  = 'Sheet1';
    $req->{folder}     = 'Cells';
    my $r = $api->replace_content_in_remote_spreadsheet(request => $req);
    print "  -> Done\n";
}

# --- 4d. Replace content in remote worksheet ---
print "\n[4d] Replace content in remote worksheet\n";
{
    use AsposeCellsCloud::Request::ReplaceContentInRemoteWorksheetRequest;
    my $req = AsposeCellsCloud::Request::ReplaceContentInRemoteWorksheetRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{text}       = 'Sales';
    $req->{new_text}   = 'Revenue';
    $req->{folder}     = 'Cells';
    my $r = $api->replace_content_in_remote_worksheet(request => $req);
    print "  -> Done\n";
}

# --- 4e. Search content in remote worksheet ---
print "\n[4e] Search content in remote worksheet\n";
{
    use AsposeCellsCloud::Request::SearchContentInRemoteWorksheetRequest;
    my $req = AsposeCellsCloud::Request::SearchContentInRemoteWorksheetRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{text}       = 'Sales';
    $req->{folder}     = 'Cells';
    my $r = $api->search_content_in_remote_worksheet(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 5: Add Text to Cells (Local Upload)
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 5: Add Text to Cells (Local Upload)\n";
print "=" x 70, "\n";

# --- 5a. Add text after existing content ---
print "\n[5a] Add ' (Updated)' after cell content in range A1:A10\n";
{
    use AsposeCellsCloud::Request::AddTextRequest;
    my $req = AsposeCellsCloud::Request::AddTextRequest->new();
    $req->{spreadsheet}       = $SRC;
    $req->{text}              = ' (Updated)';
    $req->{position}          = 'After';
    $req->{worksheet}         = 'Sheet1';
    $req->{range}             = 'A1:A10';
    $req->{skip_empty_cells}  = bless(\1, 'boolean');
    my $r = $api->add_text(request => $req);
    if ($r) { save_file($r, 'Text_Added.xlsx'); }
}

# --- 5b. Add text before specific text match ---
print "\n[5b] Add 'FY2026-' before cells containing 'Sales' in range A1:A10\n";
{
    use AsposeCellsCloud::Request::AddTextRequest;
    my $req = AsposeCellsCloud::Request::AddTextRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{text}        = 'FY2026-';
    $req->{position}    = 'BeforeText';
    $req->{select_text} = 'Sales';
    $req->{worksheet}   = 'Sheet1';
    $req->{range}       = 'A1:A10';
    my $r = $api->add_text(request => $req);
    if ($r) { save_file($r, 'Text_Added_Before.xlsx'); }
}

# =========================================================================
# Section 6: Add Text to Cells (Cloud Remote)
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 6: Add Text to Cells (Cloud Remote)\n";
print "=" x 70, "\n";

print "\n[6a] Add text to cells in remote spreadsheet\n";
{
    use AsposeCellsCloud::Request::AddTextInRemoteSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::AddTextInRemoteSpreadsheetRequest->new();
    $req->{name}              = $SRC;
    $req->{worksheet}         = 'Sheet1';
    $req->{range}             = 'A1:A10';
    $req->{text}              = ' (Verified)';
    $req->{position}          = 'After';
    $req->{skip_empty_cells}  = bless(\1, 'boolean');
    $req->{folder}            = 'Cells';
    my $r = $api->add_text_in_remote_spreadsheet(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 7: Convert Text in Cells (Local Upload)
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 7: Convert Text in Cells (Local Upload)\n";
print "=" x 70, "\n";

# --- 7a. Replace characters ---
print "\n[7a] Replace characters 'ae' -> 'AE' in workbook\n";
{
    use AsposeCellsCloud::Request::ConvertTextRequest;
    my $req = AsposeCellsCloud::Request::ConvertTextRequest->new();
    $req->{spreadsheet}        = $SRC;
    $req->{convert_text_type}  = 'ReplaceCharacters';
    $req->{source_characters}  = 'ae';
    $req->{target_characters}  = 'AE';
    my $r = $api->convert_text(request => $req);
    if ($r) { save_file($r, 'Text_Converted.xlsx'); }
}

# =========================================================================
# Section 8: Convert Text in Cells (Cloud Remote)
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 8: Convert Text in Cells (Cloud Remote)\n";
print "=" x 70, "\n";

print "\n[8a] Convert text (replace characters) in remote spreadsheet\n";
{
    use AsposeCellsCloud::Request::ConvertTextInRemoteSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::ConvertTextInRemoteSpreadsheetRequest->new();
    $req->{name}               = $SRC;
    $req->{worksheet}          = 'Sheet1';
    $req->{range}              = 'A1:A10';
    $req->{convert_text_type}  = 'ReplaceCharacters';
    $req->{source_characters}  = 'ae';
    $req->{target_characters}  = 'AE';
    $req->{folder}             = 'Cells';
    my $r = $api->convert_text_in_remote_spreadsheet(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 9: Extract Text from Cells (Local Upload)
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 9: Extract Text from Cells (Local Upload)\n";
print "=" x 70, "\n";

# --- 9a. Extract first N characters ---
print "\n[9a] Extract first 5 characters from cells A1:A10\n";
{
    use AsposeCellsCloud::Request::ExtractTextRequest;
    my $req = AsposeCellsCloud::Request::ExtractTextRequest->new();
    $req->{spreadsheet}        = $SRC;
    $req->{extract_text_type}  = 'First';
    $req->{before_position}    = 5;
    $req->{worksheet}          = 'Sheet1';
    $req->{range}              = 'A1:A10';
    my $r = $api->extract_text(request => $req);
    if ($r) { save_file($r, 'Text_Extracted.xlsx'); }
}

# --- 9b. Extract text between delimiters ---
print "\n[9b] Extract text between '(' and ')' in cells A1:A10\n";
{
    use AsposeCellsCloud::Request::ExtractTextRequest;
    my $req = AsposeCellsCloud::Request::ExtractTextRequest->new();
    $req->{spreadsheet}        = $SRC;
    $req->{extract_text_type}  = 'BetweenDelimiter';
    $req->{before_text}        = '(';
    $req->{after_text}         = ')';
    $req->{worksheet}          = 'Sheet1';
    $req->{range}              = 'A1:A10';
    my $r = $api->extract_text(request => $req);
    if ($r) { save_file($r, 'Text_Extracted_Delim.xlsx'); }
}

# =========================================================================
# Section 10: Translate
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 10: Translate Spreadsheet Text\n";
print "=" x 70, "\n";

# --- 10a. Translate text in uploaded file ---
print "\n[10a] Translate text to German (de)\n";
{
    use AsposeCellsCloud::Request::TranslateTextFileRequest;
    my $req = AsposeCellsCloud::Request::TranslateTextFileRequest->new();
    $req->{spreadsheet}      = $SRC;
    $req->{target_language}  = 'de';
    my $r = $api->translate_text_file(request => $req);
    if ($r) { save_file($r, 'Translated_Text.xlsx'); }
}

# --- 10b. Translate spreadsheet ---
print "\n[10b] Translate entire spreadsheet to French (fr)\n";
{
    use AsposeCellsCloud::Request::TranslateSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::TranslateSpreadsheetRequest->new();
    $req->{spreadsheet}      = $SRC;
    $req->{target_language}  = 'fr';
    my $r = $api->translate_spreadsheet(request => $req);
    if ($r) { save_file($r, 'Translated_Spreadsheet.xlsx'); }
}

# =========================================================================
# Section 11: Import Data into Spreadsheet
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 11: Import Data into Spreadsheet\n";
print "=" x 70, "\n";

print "\n[11a] Import data file into spreadsheet\n";
{
    use AsposeCellsCloud::Request::ImportDataIntoSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::ImportDataIntoSpreadsheetRequest->new();
    $req->{spreadsheet}           = $SRC;
    $req->{datafile}              = 'import_data.csv';
    $req->{worksheet}             = 'Sheet1';
    $req->{startcell}             = 'B5';
    $req->{insert}                = bless(\1, 'boolean');
    $req->{convert_numeric_data}  = bless(\1, 'boolean');
    $req->{splitter}              = ',';
    my $r = $api->import_data_into_spreadsheet(request => $req);
    if ($r) { save_file($r, 'Data_Imported.xlsx'); }
}

# =========================================================================
# Section 12: Smart Marker
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 12: Smart Marker (Template Processing)\n";
print "=" x 70, "\n";

print "\n[12a] Process smart marker template with XML data\n";
{
    use AsposeCellsCloud::Request::PostWorkbookGetSmartMarkerResultRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookGetSmartMarkerResultRequest->new();
    $req->{name}             = 'SmartMarkerTemplate.xlsx';
    $req->{xml_file}         = 'SmartMarkerData.xml';
    $req->{folder}           = 'Cells';
    $req->{out_path}         = 'SmartMarkerOutput';
    my $r = $api->post_workbook_get_smart_marker_result(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 13: Data Operations (Cleansing, Deduplication, Fill)
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 13: Data Cleansing / Deduplication / Fill\n";
print "=" x 70, "\n";

# --- 13a. Data cleansing ---
print "\n[13a] Data cleansing on workbook\n";
{
    use AsposeCellsCloud::Request::PostWorkbookDataCleansingRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookDataCleansingRequest->new();
    $req->{name}                     = $SRC;
    $req->{data_cleansing}           = '';
    $req->{folder}                   = 'Cells';
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_workbook_data_cleansing(request => $req);
    print "  -> Done\n";
}

# --- 13b. Data deduplication ---
print "\n[13b] Data deduplication on workbook\n";
{
    use AsposeCellsCloud::Request::PostWorkbookDataDeduplicationRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookDataDeduplicationRequest->new();
    $req->{name}                     = $SRC;
    $req->{folder}                   = 'Cells';
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_workbook_data_deduplication(request => $req);
    print "  -> Done\n";
}

# --- 13c. Data fill ---
print "\n[13c] Data fill on workbook\n";
{
    use AsposeCellsCloud::Request::PostWorkbookDataFillRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookDataFillRequest->new();
    $req->{name}                     = $SRC;
    $req->{folder}                   = 'Cells';
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_workbook_data_fill(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 14: Clear Formats
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 14: Clear Formats\n";
print "=" x 70, "\n";

# --- 14a. Clear formats by range ---
print "\n[14a] Clear formats in range A1:C10\n";
{
    use AsposeCellsCloud::Request::PostClearFormatsRequest;
    my $req = AsposeCellsCloud::Request::PostClearFormatsRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{range}      = 'A1:C10';
    $req->{folder}     = 'Cells';
    my $r = $api->post_clear_formats(request => $req);
    print "  -> Done\n";
}

# --- 14b. Clear formats by row/column indices ---
print "\n[14b] Clear formats from rows 0-9, columns 0-2\n";
{
    use AsposeCellsCloud::Request::PostClearFormatsRequest;
    my $req = AsposeCellsCloud::Request::PostClearFormatsRequest->new();
    $req->{name}         = $SRC;
    $req->{sheet_name}   = 'Sheet1';
    $req->{start_row}    = 0;
    $req->{start_column} = 0;
    $req->{end_row}      = 9;
    $req->{end_column}   = 2;
    $req->{folder}       = 'Cells';
    my $r = $api->post_clear_formats(request => $req);
    print "  -> Done\n";
}

# =========================================================================
# Section 15: Export
# =========================================================================
print "\n", "=" x 70, "\n";
print "SECTION 15: Post Export (Uploaded File)\n";
print "=" x 70, "\n";

print "\n[15a] Export uploaded workbook worksheet to PDF\n";
{
    use AsposeCellsCloud::Request::PostExportRequest;
    my $req = AsposeCellsCloud::Request::PostExportRequest->new();
    $req->{file}          = $SRC;
    $req->{object_type}   = 'worksheet';
    $req->{format}        = 'pdf';
    my $r = $api->post_export(request => $req);
    if ($r) { save_file($r, 'Exported.pdf'); }
}

print "\nAll text processing examples completed.\n";
