use lib '../lib';
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;

# =========================================================================
# File Merging Examples for Aspose.Cells Cloud SDK for Perl
# =========================================================================
# Demonstrates merging spreadsheets, worksheets, and cell ranges.
# Covers local upload, cloud-storage, and remote approaches.
# =========================================================================

my $config = AsposeCellsCloud::Configuration->new(
    client_id     => $ENV{'CellsCloudClientId'},
    client_secret => $ENV{'CellsCloudClientSecret'}
);
my $api = AsposeCellsCloud::CellsApi->new(
    AsposeCellsCloud::ApiClient->new($config)
);

my $SRC  = 'EmployeeSalesSummary.xlsx';
my $SRC2 = 'CompanySales.xlsx';

sub save_file {
    my ($data, $path) = @_;
    open(my $fh, '>', $path) or die "Cannot open $path: $!";
    binmode($fh);
    print $fh $data;
    close($fh);
    print "  -> Saved $path\n";
}

# =====================================================================
# Section 1: Merge local spreadsheets (upload & merge)
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 1: Merge Multiple Local Spreadsheets (Upload)\n";
print "=" x 70, "\n";

# --- 1a. Merge two local spreadsheets, output as XLSX ---
print "\n[1a] Merge two spreadsheets -> single file (XLSX)\n";
{
    use AsposeCellsCloud::Request::PostMergeRequest;
    my $req = AsposeCellsCloud::Request::PostMergeRequest->new();
    $req->{file}       = $SRC;
    $req->{out_format} = 'xlsx';
    $req->{merge_to}   = $SRC2;
    my $r = $api->post_merge(request => $req);
    if ($r) {
        save_file($r, 'Merged_Local.xlsx');
    }
}

# --- 1b. Merge into a single sheet ---
print "\n[1b] Merge two spreadsheets into a single sheet\n";
{
    use AsposeCellsCloud::Request::PostMergeRequest;
    my $req = AsposeCellsCloud::Request::PostMergeRequest->new();
    $req->{file}               = $SRC;
    $req->{out_format}         = 'xlsx';
    $req->{merge_to}           = $SRC2;
    $req->{merge_in_one_sheet} = bless(\1, 'boolean');
    my $r = $api->post_merge(request => $req);
    if ($r) {
        save_file($r, 'Merged_OneSheet.xlsx');
    }
}

# --- 1c. Merge and convert to PDF ---
print "\n[1c] Merge two spreadsheets -> PDF\n";
{
    use AsposeCellsCloud::Request::PostMergeRequest;
    my $req = AsposeCellsCloud::Request::PostMergeRequest->new();
    $req->{file}       = $SRC;
    $req->{out_format} = 'pdf';
    $req->{merge_to}   = $SRC2;
    my $r = $api->post_merge(request => $req);
    if ($r) {
        save_file($r, 'Merged_To_PDF.pdf');
    }
}

# =====================================================================
# Section 2: Merge local spreadsheets with merge_spreadsheets API
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 2: Merge Spreadsheets (Local Upload Variant)\n";
print "=" x 70, "\n";

print "\n[2a] Merge spreadsheets with format and merge in one sheet\n";
{
    use AsposeCellsCloud::Request::MergeSpreadsheetsRequest;
    my $req = AsposeCellsCloud::Request::MergeSpreadsheetsRequest->new();
    $req->{spreadsheet}        = $SRC;
    $req->{out_format}         = 'xlsx';
    $req->{merge_in_one_sheet} = bless(\1, 'boolean');
    $req->{out_path}           = '';
    my $r = $api->merge_spreadsheets(request => $req);
    if ($r) {
        save_file($r, 'Merged_Local_Variant.xlsx');
    }
}

# =====================================================================
# Section 3: Merge cloud-stored workbooks
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 3: Merge Cloud-Stored Workbooks\n";
print "=" x 70, "\n";

# --- 3a. Merge a workbook into another in cloud storage ---
print "\n[3a] Merge workbook into another (cloud storage)\n";
{
    use AsposeCellsCloud::Request::PostWorkbooksMergeRequest;
    my $req = AsposeCellsCloud::Request::PostWorkbooksMergeRequest->new();
    $req->{name}               = $SRC;
    $req->{merge_with}         = $SRC2;
    $req->{folder}             = 'Cells';
    $req->{merge_in_one_sheet} = bless(\0, 'boolean');
    my $r = $api->post_workbooks_merge(request => $req);
    print "  -> Done, status: " . ($r->{status} // 'OK') . "\n";
}

# --- 3b. Merge all spreadsheets from a remote folder ---
print "\n[3b] Merge all spreadsheets in a remote folder into one\n";
{
    use AsposeCellsCloud::Request::MergeSpreadsheetsInRemoteFolderRequest;
    my $req = AsposeCellsCloud::Request::MergeSpreadsheetsInRemoteFolderRequest->new();
    $req->{name}         = $SRC;
    $req->{folder}       = 'Cells';
    $req->{merge_folder} = 'Cells/MergeSource';
    my $r = $api->merge_spreadsheets_in_remote_folder(request => $req);
    print "  -> Done\n";
}

# --- 3c. Merge a remote spreadsheet into another ---
print "\n[3c] Merge remote spreadsheet into another\n";
{
    use AsposeCellsCloud::Request::MergeRemoteSpreadsheetRequest;
    my $req = AsposeCellsCloud::Request::MergeRemoteSpreadsheetRequest->new();
    $req->{name}       = $SRC;
    $req->{merge_with} = $SRC2;
    $req->{folder}     = 'Cells';
    my $r = $api->merge_remote_spreadsheet(request => $req);
    print "  -> Done\n";
}

# =====================================================================
# Section 4: Merge cells within a worksheet
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 4: Merge Cells in Worksheet\n";
print "=" x 70, "\n";

# --- 4a. Merge a range of cells ---
print "\n[4a] Merge cells range A1:B2\n";
{
    use AsposeCellsCloud::Request::PostWorksheetCellsRangeMergeRequest;
    my $req = AsposeCellsCloud::Request::PostWorksheetCellsRangeMergeRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{range}      = 'A1:B2';
    $req->{folder}     = 'Cells';
    my $r = $api->post_worksheet_cells_range_merge(request => $req);
    print "  -> Done\n";
}

# --- 4b. Merge worksheet by row/column dimensions ---
print "\n[4b] Merge worksheet (start_row=0, start_column=0, total_rows=2, total_columns=2)\n";
{
    use AsposeCellsCloud::Request::PostWorksheetMergeRequest;
    my $req = AsposeCellsCloud::Request::PostWorksheetMergeRequest->new();
    $req->{name}          = $SRC;
    $req->{sheet_name}    = 'Sheet1';
    $req->{start_row}     = 0;
    $req->{start_column}  = 0;
    $req->{total_rows}    = 2;
    $req->{total_columns} = 2;
    $req->{folder}        = 'Cells';
    my $r = $api->post_worksheet_merge(request => $req);
    print "  -> Done\n";
}

# --- 4c. Unmerge a range of cells ---
print "\n[4c] Unmerge cells range A1:B2\n";
{
    use AsposeCellsCloud::Request::PostWorksheetCellsRangeUnMergeRequest;
    my $req = AsposeCellsCloud::Request::PostWorksheetCellsRangeUnMergeRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{range}      = 'A1:B2';
    $req->{folder}     = 'Cells';
    my $r = $api->post_worksheet_cells_range_un_merge(request => $req);
    print "  -> Done\n";
}

# --- 4d. Unmerge a region by row/column dimensions ---
print "\n[4d] Unmerge worksheet region\n";
{
    use AsposeCellsCloud::Request::PostWorksheetUnmergeRequest;
    my $req = AsposeCellsCloud::Request::PostWorksheetUnmergeRequest->new();
    $req->{name}          = $SRC;
    $req->{sheet_name}    = 'Sheet1';
    $req->{start_row}     = 0;
    $req->{start_column}  = 0;
    $req->{total_rows}    = 50;
    $req->{total_columns} = 20;
    $req->{folder}        = 'Cells';
    my $r = $api->post_worksheet_unmerge(request => $req);
    print "  -> Done\n";
}

# =====================================================================
# Section 5: Get merged cells information
# =====================================================================
print "\n", "=" x 70, "\n";
print "SECTION 5: Get Merged Cells Information\n";
print "=" x 70, "\n";

# --- 5a. Get merged cells in worksheet ---
print "\n[5a] Get merged cells in worksheet\n";
{
    use AsposeCellsCloud::Request::GetMergedCellsInWorksheetRequest;
    my $req = AsposeCellsCloud::Request::GetMergedCellsInWorksheetRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{folder}     = 'Cells';
    my $r = $api->get_merged_cells_in_worksheet(request => $req);
    if ($r && defined $r->{merged_cells}) {
        my $cells = $r->{merged_cells};
        print "  Merged cells count: " . ($cells->{count} // 'N/A') . "\n";
    }
    print "  -> Done\n";
}

# --- 5b. Get merged cells in remote worksheet ---
print "\n[5b] Get merged cells in remote worksheet\n";
{
    use AsposeCellsCloud::Request::GetMergedCellsInRemotedWorksheetRequest;
    my $req = AsposeCellsCloud::Request::GetMergedCellsInRemotedWorksheetRequest->new();
    $req->{name}       = $SRC;
    $req->{sheet_name} = 'Sheet1';
    $req->{folder}     = 'Cells';
    my $r = $api->get_merged_cells_in_remoted_worksheet(request => $req);
    print "  -> Done\n";
}

print "\nAll file merging examples completed.\n";
