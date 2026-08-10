use lib '../lib';
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;

# =========================================================================
# File Splitting Examples for Aspose.Cells Cloud SDK for Perl
# =========================================================================
# Demonstrates splitting spreadsheets by worksheets, ranges,
# and table columns. Covers both local upload and cloud-storage modes.
# =========================================================================

my $config = AsposeCellsCloud::Configuration->new(
    client_id     => $ENV{'CellsCloudClientId'},
    client_secret => $ENV{'CellsCloudClientSecret'}
);
my $api = AsposeCellsCloud::CellsApi->new(
    AsposeCellsCloud::ApiClient->new($config)
);

my $SRC = 'EmployeeSalesSummary.xlsx';

# =====================================================================
# Section 1: Split uploaded spreadsheet by worksheet
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 1: Split Uploaded Spreadsheet by Worksheet\n";
print "=" x 70, "\n";

# --- 1a. Split to PDF ---
print "\n[1a] Split uploaded file into PDF pages by worksheet\n";
{
    use AsposeCellsCloud::Request::PostSplitRequest;
    my $req = AsposeCellsCloud::Request::PostSplitRequest->new();
    $req->{file}                    = $SRC;
    $req->{out_format}              = 'pdf';
    $req->{check_excel_restriction} = 'true';
    my $r = $api->post_split(request => $req);
    if ($r && defined $r->{files}) {
        my $files = $r->{files};
        if (ref($files) eq 'ARRAY') {
            print "  Generated " . scalar(@$files) . " file(s).\n";
        }
    }
}

# --- 1b. Split to CSV ---
print "\n[1b] Split uploaded file into CSV files\n";
{
    use AsposeCellsCloud::Request::PostSplitRequest;
    my $req = AsposeCellsCloud::Request::PostSplitRequest->new();
    $req->{file}       = $SRC;
    $req->{out_format} = 'csv';
    my $r = $api->post_split(request => $req);
    print "  -> Done\n";
}

# --- 1c. Split specific worksheet range (from=0, to=1) ---
print "\n[1c] Split specific worksheets (index 0 to 1) as XLSX\n";
{
    use AsposeCellsCloud::Request::PostSplitRequest;
    my $req = AsposeCellsCloud::Request::PostSplitRequest->new();
    $req->{file}       = $SRC;
    $req->{out_format} = 'xlsx';
    $req->{from}       = 0;
    $req->{to}         = 1;
    my $r = $api->post_split(request => $req);
    print "  -> Done\n";
}

# =====================================================================
# Section 2: Split cloud-stored workbook
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 2: Split Cloud-Stored Workbook\n";
print "=" x 70, "\n";

# --- 2a. Split to PDF with sheetname rule ---
print "\n[2a] Split workbook to PDF (sheetname naming)\n";
{
    use AsposeCellsCloud::Request::PostWorkbookSplitRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookSplitRequest->new();
    $req->{name}            = $SRC;
    $req->{format}          = 'pdf';
    $req->{out_folder}      = 'SplitOutput';
    $req->{from}            = 0;
    $req->{to}              = 2;
    $req->{split_name_rule} = 'sheetname';
    $req->{folder}          = 'Cells';
    my $r = $api->post_workbook_split(request => $req);
    if ($r && defined $r->{result}) {
        print "  Generated " . scalar(@{$r->{result}}) . " file(s).\n";
    }
}

# --- 2b. Split to PNG images with custom DPI ---
print "\n[2b] Split workbook to PNG images (150 DPI)\n";
{
    use AsposeCellsCloud::Request::PostWorkbookSplitRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookSplitRequest->new();
    $req->{name}                  = $SRC;
    $req->{format}                = 'png';
    $req->{out_folder}            = 'SplitImages';
    $req->{horizontal_resolution} = 150;
    $req->{vertical_resolution}   = 150;
    $req->{split_name_rule}       = 'sheetname';
    $req->{folder}                = 'Cells';
    my $r = $api->post_workbook_split(request => $req);
    print "  -> Done\n";
}

# --- 2c. Split with GUID-based naming ---
print "\n[2c] Split workbook with GUID-based output naming\n";
{
    use AsposeCellsCloud::Request::PostWorkbookSplitRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbookSplitRequest->new();
    $req->{name}            = $SRC;
    $req->{format}          = 'xlsx';
    $req->{out_folder}      = 'SplitOutput';
    $req->{split_name_rule} = 'newguid';
    $req->{folder}          = 'Cells';
    my $r = $api->post_workbook_split(request => $req);
    print "  -> Done\n";
}

# =====================================================================
# Section 3: Split local spreadsheet to cloud storage
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 3: Split Local File to Cloud Storage\n";
print "=" x 70, "\n";

print "\n[3a] Split local file -> save to cloud storage as PDF\n";
{
    use AsposeCellsCloud::Request::SplitSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::SplitSpreadsheetRequest->new();
    $req->{spreadsheet}     = $SRC;
    $req->{format}          = 'pdf';
    $req->{out_path}        = 'SplitOutput';
    $req->{from}            = 0;
    $req->{to}              = 2;
    my $r = $api->split_spreadsheet(request => $req);
    print "  -> Done\n";
}

# =====================================================================
# Section 4: Split remote spreadsheet
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 4: Split Remote Spreadsheet\n";
print "=" x 70, "\n";

print "\n[4a] Split remote spreadsheet (by worksheet range)\n";
{
    use AsposeCellsCloud::Request::SplitRemoteSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::SplitRemoteSpreadsheetRequest->new();
    $req->{name}        = $SRC;
    $req->{folder}      = 'Cells';
    $req->{from}        = 0;
    $req->{to}          = 1;
    $req->{out_format}  = 'csv';
    $req->{out_path}    = 'RemoteSplitOutput';
    my $r = $api->split_remote_spreadsheet(request => $req);
    print "  -> Done\n";
}

# =====================================================================
# Section 5: Split table by column
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 5: Split Table by Column\n";
print "=" x 70, "\n";

# --- 5a. Split table into separate files (ZIP) ---
print "\n[5a] Split table by column -> multiple files (ZIP)\n";
{
    use AsposeCellsCloud::Request::SplitTableRequest;
    my $req = AsposeCellsCloud::Request::SplitTableRequest->new();
    $req->{spreadsheet}        = $SRC;
    $req->{worksheet}          = 'Sheet1';
    $req->{table_name}         = 'Table1';
    $req->{split_column_name}  = 'Category';
    $req->{save_split_column}  = bless(\0, 'boolean');
    $req->{to_new_workbook}    = bless(\0, 'boolean');
    $req->{to_multiple_files}  = bless(\1, 'boolean');
    my $r = $api->split_table(request => $req);
    if ($r) {
        save_file($r, 'Split_Table.zip');
    }
}

# --- 5b. Split table into one workbook (each split in separate sheet) ---
print "\n[5b] Split table by column -> single workbook with multiple sheets\n";
{
    use AsposeCellsCloud::Request::SplitTableRequest;
    my $req = AsposeCellsCloud::Request::SplitTableRequest->new();
    $req->{spreadsheet}        = $SRC;
    $req->{worksheet}          = 'Sheet1';
    $req->{table_name}         = 'Table1';
    $req->{split_column_name}  = 'Category';
    $req->{save_split_column}  = bless(\1, 'boolean');
    $req->{to_new_workbook}    = bless(\0, 'boolean');
    $req->{to_multiple_files}  = bless(\0, 'boolean');
    my $r = $api->split_table(request => $req);
    if ($r) {
        save_file($r, 'Split_Table_Sheets.xlsx');
    }
}

# =====================================================================
# Section 6: Split text in cells
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 6: Split Text in Cells\n";
print "=" x 70, "\n";

# --- 6a. Split text in cells by custom delimiter ---
print "\n[6a] Split text in cells by comma delimiter\n";
{
    use AsposeCellsCloud::Request::SplitTextRequest;
    my $req = AsposeCellsCloud::Request::SplitTextRequest->new();
    $req->{spreadsheet}                         = $SRC;
    $req->{worksheet}                           = 'Sheet1';
    $req->{range}                               = 'A1:A10';
    $req->{delimiters}                          = ',';
    $req->{keep_delimiters_in_resulting_cells}   = bless(\0, 'boolean');
    $req->{how_to_split}                        = 'Columns';
    my $r = $api->split_text(request => $req);
    if ($r) {
        save_file($r, 'Split_Text_Result.xlsx');
    }
}

# =====================================================================
# Section 7: Batch split multiple files
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 7: Batch Split\n";
print "=" x 70, "\n";

print "\n[7a] Batch split all *.xlsx files to PDF\n";
{
    use AsposeCellsCloud::Object::BatchSplitRequest;
    use AsposeCellsCloud::Request::PostBatchSplitRequest;
    my $batch = AsposeCellsCloud::Object::BatchSplitRequest->new();
    $batch->{source_folder}   = 'Cells';
    $batch->{source_storage}  = '';
    $batch->{match_condition} = {
        file_pattern      => '*.xlsx',
        file_name_pattern => ''
    };
    $batch->{format}      = 'pdf';
    $batch->{out_folder}  = 'BatchSplitOutput';
    $batch->{out_storage} = '';
    $batch->{from_index}  = 0;
    $batch->{to_index}    = 2;
    my $req = AsposeCellsCloud::Request::PostBatchSplitRequest->new();
    $req->{batch_split_request} = $batch;
    my $r = $api->post_batch_split(request => $req);
    print "  -> Done\n";
}

print "\nAll file splitting examples completed.\n";

sub save_file {
    my ($data, $path) = @_;
    open(my $fh, '>', $path) or die "Cannot open $path: $!";
    binmode($fh);
    print $fh $data;
    close($fh);
    print "  -> Saved $path\n";
}
