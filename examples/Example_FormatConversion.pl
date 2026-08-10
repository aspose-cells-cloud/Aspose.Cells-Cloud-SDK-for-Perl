use lib '../lib';
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;

# =========================================================================
# Format Conversion Examples for Aspose.Cells Cloud SDK for Perl
# =========================================================================
# Demonstrates converting spreadsheets, worksheets, ranges, and tables
# to various formats (PDF, CSV, HTML, JSON, images, etc.).
# Covers both local file upload and cloud-storage based approaches.
# =========================================================================

my $config = AsposeCellsCloud::Configuration->new(
    client_id     => $ENV{'CellsCloudClientId'},
    client_secret => $ENV{'CellsCloudClientSecret'}
);
my $api = AsposeCellsCloud::CellsApi->new(
    AsposeCellsCloud::ApiClient->new($config)
);

my $SRC = 'EmployeeSalesSummary.xlsx';
my $FMT = 'pdf';

sub save_file {
    my ($data, $path) = @_;
    open(my $fh, '>', $path) or die "Cannot open $path: $!";
    binmode($fh);
    print $fh $data;
    close($fh);
    print "  -> Saved $path\n";
}

# =====================================================================
# Section 1: Convert entire spreadsheet (local upload)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 1: Convert Entire Spreadsheet (Local Upload)\n";
print "=" x 70, "\n";

# --- 1a. to PDF ---
print "\n[1a] Spreadsheet -> PDF\n";
{
    use AsposeCellsCloud::Request::ConvertSpreadsheetToPdfRequest;
    my $req = AsposeCellsCloud::Request::ConvertSpreadsheetToPdfRequest->new();
    $req->{spreadsheet}    = $SRC;
    $req->{fonts_location} = '';
    save_file($api->convert_spreadsheet_to_pdf(request => $req), 'Output_Spreadsheet.pdf');
}

# --- 1b. to CSV ---
print "\n[1b] Spreadsheet -> CSV\n";
{
    use AsposeCellsCloud::Request::ConvertSpreadsheetToCsvRequest;
    my $req = AsposeCellsCloud::Request::ConvertSpreadsheetToCsvRequest->new();
    $req->{spreadsheet} = $SRC;
    save_file($api->convert_spreadsheet_to_csv(request => $req), 'Output_Spreadsheet.csv');
}

# --- 1c. to JSON ---
print "\n[1c] Spreadsheet -> JSON\n";
{
    use AsposeCellsCloud::Request::ConvertSpreadsheetToJsonRequest;
    my $req = AsposeCellsCloud::Request::ConvertSpreadsheetToJsonRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{out_path}    = '';
    save_file($api->convert_spreadsheet_to_json(request => $req), 'Output_Spreadsheet.json');
}

# --- 1d. to any format (generic put_convert_workbook) ---
print "\n[1d] Spreadsheet -> XPS (generic convert)\n";
{
    use AsposeCellsCloud::Request::ConvertSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::ConvertSpreadsheetRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{format}      = 'xps';
    save_file($api->put_convert_workbook(request => $req), 'Output_Spreadsheet.xps');
}

# =====================================================================
# Section 2: Convert individual worksheet (local upload)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 2: Convert Worksheet (Local Upload)\n";
print "=" x 70, "\n";

# --- 2a. Worksheet -> PDF ---
print "\n[2a] Worksheet -> PDF\n";
{
    use AsposeCellsCloud::Request::ConvertWorksheetToPdfRequest;
    my $req = AsposeCellsCloud::Request::ConvertWorksheetToPdfRequest->new();
    $req->{spreadsheet}       = $SRC;
    $req->{worksheet}         = 'Sheet1';
    $req->{auto_rows_fit}     = bless(\1, 'boolean');
    $req->{auto_columns_fit}  = bless(\1, 'boolean');
    save_file($api->convert_worksheet_to_pdf(request => $req), 'Output_Worksheet.pdf');
}

# --- 2b. Worksheet -> CSV ---
print "\n[2b] Worksheet -> CSV\n";
{
    use AsposeCellsCloud::Request::ConvertWorksheetToCsvRequest;
    my $req = AsposeCellsCloud::Request::ConvertWorksheetToCsvRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    save_file($api->convert_worksheet_to_csv(request => $req), 'Output_Worksheet.csv');
}

# --- 2c. Worksheet -> HTML ---
print "\n[2c] Worksheet -> HTML\n";
{
    use AsposeCellsCloud::Request::ConvertWorksheetToHtmlRequest;
    my $req = AsposeCellsCloud::Request::ConvertWorksheetToHtmlRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    save_file($api->convert_worksheet_to_html(request => $req), 'Output_Worksheet.html');
}

# --- 2d. Worksheet -> HTML Table ---
print "\n[2d] Worksheet -> HTML Table\n";
{
    use AsposeCellsCloud::Request::ConvertWorksheetToHtmlTableRequest;
    my $req = AsposeCellsCloud::Request::ConvertWorksheetToHtmlTableRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    save_file($api->convert_worksheet_to_html_table(request => $req), 'Output_Worksheet_Table.html');
}

# --- 2e. Worksheet -> JSON ---
print "\n[2e] Worksheet -> JSON\n";
{
    use AsposeCellsCloud::Request::ConvertWorksheetToJsonRequest;
    my $req = AsposeCellsCloud::Request::ConvertWorksheetToJsonRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    save_file($api->convert_worksheet_to_json(request => $req), 'Output_Worksheet.json');
}

# --- 2f. Worksheet -> Image ---
print "\n[2f] Worksheet -> Image (PNG)\n";
{
    use AsposeCellsCloud::Request::ConvertWorksheetToImageRequest;
    my $req = AsposeCellsCloud::Request::ConvertWorksheetToImageRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{format}      = 'png';
    save_file($api->convert_worksheet_to_image(request => $req), 'Output_Worksheet.png');
}

# =====================================================================
# Section 3: Convert a specific range (local upload)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 3: Convert Range (Local Upload)\n";
print "=" x 70, "\n";

# --- 3a. Range -> PDF ---
print "\n[3a] Range -> PDF\n";
{
    use AsposeCellsCloud::Request::ConvertRangeToPdfRequest;
    my $req = AsposeCellsCloud::Request::ConvertRangeToPdfRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{range}       = 'A1:C10';
    save_file($api->convert_range_to_pdf(request => $req), 'Output_Range.pdf');
}

# --- 3b. Range -> CSV ---
print "\n[3b] Range -> CSV\n";
{
    use AsposeCellsCloud::Request::ConvertRangeToCsvRequest;
    my $req = AsposeCellsCloud::Request::ConvertRangeToCsvRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{range}       = 'A1:C10';
    save_file($api->convert_range_to_csv(request => $req), 'Output_Range.csv');
}

# --- 3c. Range -> HTML ---
print "\n[3c] Range -> HTML\n";
{
    use AsposeCellsCloud::Request::ConvertRangeToHtmlRequest;
    my $req = AsposeCellsCloud::Request::ConvertRangeToHtmlRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{range}       = 'A1:C10';
    save_file($api->convert_range_to_html(request => $req), 'Output_Range.html');
}

# --- 3d. Range -> JSON ---
print "\n[3d] Range -> JSON\n";
{
    use AsposeCellsCloud::Request::ConvertRangeToJsonRequest;
    my $req = AsposeCellsCloud::Request::ConvertRangeToJsonRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{range}       = 'A1:C10';
    save_file($api->convert_range_to_json(request => $req), 'Output_Range.json');
}

# --- 3e. Range -> Image ---
print "\n[3e] Range -> Image (PNG)\n";
{
    use AsposeCellsCloud::Request::ConvertRangeToImageRequest;
    my $req = AsposeCellsCloud::Request::ConvertRangeToImageRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{range}       = 'A1:C10';
    $req->{format}      = 'png';
    save_file($api->convert_range_to_image(request => $req), 'Output_Range.png');
}

# =====================================================================
# Section 4: Convert table (local upload)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 4: Convert Table (Local Upload)\n";
print "=" x 70, "\n";

# --- 4a. Table -> PDF ---
print "\n[4a] Table -> PDF\n";
{
    use AsposeCellsCloud::Request::ConvertTableToPdfRequest;
    my $req = AsposeCellsCloud::Request::ConvertTableToPdfRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{table_name}  = 'Table1';
    save_file($api->convert_table_to_pdf(request => $req), 'Output_Table.pdf');
}

# --- 4b. Table -> CSV ---
print "\n[4b] Table -> CSV\n";
{
    use AsposeCellsCloud::Request::ConvertTableToCsvRequest;
    my $req = AsposeCellsCloud::Request::ConvertTableToCsvRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{table_name}  = 'Table1';
    save_file($api->convert_table_to_csv(request => $req), 'Output_Table.csv');
}

# --- 4c. Table -> HTML ---
print "\n[4c] Table -> HTML\n";
{
    use AsposeCellsCloud::Request::ConvertTableToHtmlRequest;
    my $req = AsposeCellsCloud::Request::ConvertTableToHtmlRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{table_name}  = 'Table1';
    save_file($api->convert_table_to_html(request => $req), 'Output_Table.html');
}

# --- 4d. Table -> JSON ---
print "\n[4d] Table -> JSON\n";
{
    use AsposeCellsCloud::Request::ConvertTableToJsonRequest;
    my $req = AsposeCellsCloud::Request::ConvertTableToJsonRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{table_name}  = 'Table1';
    save_file($api->convert_table_to_json(request => $req), 'Output_Table.json');
}

# --- 4e. Table -> Image ---
print "\n[4e] Table -> Image (PNG)\n";
{
    use AsposeCellsCloud::Request::ConvertTableToImageRequest;
    my $req = AsposeCellsCloud::Request::ConvertTableToImageRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{table_name}  = 'Table1';
    $req->{format}      = 'png';
    save_file($api->convert_table_to_image(request => $req), 'Output_Table.png');
}

# =====================================================================
# Section 5: Chart conversion (local upload)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 5: Convert Chart (Local Upload)\n";
print "=" x 70, "\n";

print "\n[5a] Chart -> Image (PNG)\n";
{
    use AsposeCellsCloud::Request::ConvertChartToImageRequest;
    my $req = AsposeCellsCloud::Request::ConvertChartToImageRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{chart_index} = 0;
    $req->{format}      = 'png';
    save_file($api->convert_chart_to_image(request => $req), 'Output_Chart.png');
}

print "\n[5b] Chart -> PDF\n";
{
    use AsposeCellsCloud::Request::ConvertChartToPdfRequest;
    my $req = AsposeCellsCloud::Request::ConvertChartToPdfRequest->new();
    $req->{spreadsheet} = $SRC;
    $req->{worksheet}   = 'Sheet1';
    $req->{chart_index} = 0;
    save_file($api->convert_chart_to_pdf(request => $req), 'Output_Chart.pdf');
}

# =====================================================================
# Section 6: Cloud-storage export (file already in cloud storage)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 6: Export Cloud-Stored Objects to Format\n";
print "=" x 70, "\n";

print "\n[6a] Export spreadsheet to PDF\n";
{
    use AsposeCellsCloud::Request::ExportSpreadsheetAsFormatRequest;
    my $req = AsposeCellsCloud::Request::ExportSpreadsheetAsFormatRequest->new();
    $req->{name}    = $SRC;
    $req->{format}  = 'pdf';
    $req->{folder}  = 'Cells';
    my $r = $api->export_spreadsheet_as_format(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[6b] Export worksheet to CSV\n";
{
    use AsposeCellsCloud::Request::ExportWorksheetAsFormatRequest;
    my $req = AsposeCellsCloud::Request::ExportWorksheetAsFormatRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{format}     = 'csv';
    $req->{folder}     = 'Cells';
    my $r = $api->export_worksheet_as_format(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[6c] Export table to HTML\n";
{
    use AsposeCellsCloud::Request::ExportTableAsFormatRequest;
    my $req = AsposeCellsCloud::Request::ExportTableAsFormatRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{table_name}  = 'Table1';
    $req->{format}     = 'html';
    $req->{folder}     = 'Cells';
    my $r = $api->export_table_as_format(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[6d] Export range to JSON\n";
{
    use AsposeCellsCloud::Request::ExportRangeAsFormatRequest;
    my $req = AsposeCellsCloud::Request::ExportRangeAsFormatRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{range}      = 'A1:C10';
    $req->{format}     = 'json';
    $req->{folder}     = 'Cells';
    my $r = $api->export_range_as_format(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[6e] Export chart to PNG\n";
{
    use AsposeCellsCloud::Request::ExportChartAsFormatRequest;
    my $req = AsposeCellsCloud::Request::ExportChartAsFormatRequest->new();
    $req->{name}        = $SRC;
    $req->{sheet_name}  = 'Sheet1';
    $req->{chart_index} = 0;
    $req->{format}      = 'png';
    $req->{folder}      = 'Cells';
    my $r = $api->export_chart_as_format(request => $req);
    print "  -> Done\n" if $r;
}

# =====================================================================
# Section 7: Cloud-stored workbook conversion (post_convert_workbook_*)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 7: Cloud Workbook Format Conversion\n";
print "=" x 70, "\n";

print "\n[7a] Workbook -> PDF\n";
{
    use AsposeCellsCloud::Request::PostConvertWorkbookToPDFRequest;
    my $req = AsposeCellsCloud::Request::PostConvertWorkbookToPDFRequest->new();
    $req->{name}                     = $SRC;
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_convert_workbook_to_pdf(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[7b] Workbook -> Markdown\n";
{
    use AsposeCellsCloud::Request::PostConvertWorkbookToMarkdownRequest;
    my $req = AsposeCellsCloud::Request::PostConvertWorkbookToMarkdownRequest->new();
    $req->{name}                     = $SRC;
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_convert_workbook_to_markdown(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[7c] Workbook -> HTML\n";
{
    use AsposeCellsCloud::Request::PostConvertWorkbookToHtmlRequest;
    my $req = AsposeCellsCloud::Request::PostConvertWorkbookToHtmlRequest->new();
    $req->{name}                     = $SRC;
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_convert_workbook_to_html(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[7d] Workbook -> JSON\n";
{
    use AsposeCellsCloud::Request::PostConvertWorkbookToJsonRequest;
    my $req = AsposeCellsCloud::Request::PostConvertWorkbookToJsonRequest->new();
    $req->{name}                     = $SRC;
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_convert_workbook_to_json(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[7e] Workbook -> PPTX\n";
{
    use AsposeCellsCloud::Request::PostConvertWorkbookToPptxRequest;
    my $req = AsposeCellsCloud::Request::PostConvertWorkbookToPptxRequest->new();
    $req->{name}                     = $SRC;
    $req->{check_excel_restriction}  = 'true';
    my $r = $api->post_convert_workbook_to_pptx(request => $req);
    print "  -> Done\n" if $r;
}

# =====================================================================
# Section 8: Advanced / Batch / Export
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 8: Advanced & Batch Conversion\n";
print "=" x 70, "\n";

print "\n[8a] Convert with ConvertWorkbookOptions\n";
{
    use AsposeCellsCloud::Object::ConvertWorkbookOptions;
    use AsposeCellsCloud::Request::PostConvertWorkbookRequest;
    my $opts = AsposeCellsCloud::Object::ConvertWorkbookOptions->new();
    $opts->{name}                    = $SRC;
    $opts->{convert_format}          = 'pdf';
    $opts->{check_excel_restriction} = bless(\0, 'boolean');
    my $req = AsposeCellsCloud::Request::PostConvertWorkbookRequest->new();
    $req->{convert_workbook_options} = $opts;
    my $r = $api->post_convert_workbook(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[8b] Batch convert (files matching *.xlsx -> PDF)\n";
{
    use AsposeCellsCloud::Object::BatchConvertRequest;
    use AsposeCellsCloud::Object::SaveOptions;
    use AsposeCellsCloud::Request::PostBatchConvertRequest;
    my $batch = AsposeCellsCloud::Object::BatchConvertRequest->new();
    $batch->{source_folder}   = 'Cells';
    $batch->{source_storage}  = '';
    $batch->{match_condition} = { file_pattern => '*.xlsx', file_name_pattern => '' };
    $batch->{format}          = 'pdf';
    $batch->{out_folder}      = 'ConvertedOutput';
    $batch->{out_storage}     = '';
    $batch->{save_options}    = AsposeCellsCloud::Object::SaveOptions->new();
    my $req = AsposeCellsCloud::Request::PostBatchConvertRequest->new();
    $req->{batch_convert_request} = $batch;
    my $r = $api->post_batch_convert(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[8c] PostExport (upload file, specify object type)\n";
{
    use AsposeCellsCloud::Request::PostExportRequest;
    my $req = AsposeCellsCloud::Request::PostExportRequest->new();
    $req->{file}          = $SRC;
    $req->{object_type}   = 'worksheet';   # worksheet/table/range/chart
    $req->{format}        = 'pdf';
    my $r = $api->post_export(request => $req);
    print "  -> Done\n" if $r;
}

print "\n[8d] PostConvertWorksheetToImage (batch worksheet images)\n";
{
    use AsposeCellsCloud::Request::PostConvertWorksheetToImageRequest;
    my $req = AsposeCellsCloud::Request::PostConvertWorksheetToImageRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{format}     = 'png';
    $req->{folder}     = 'Cells';
    my $r = $api->post_convert_worksheet_to_image(request => $req);
    print "  -> Done\n" if $r;
}

print "\nAll format conversion examples completed.\n";
