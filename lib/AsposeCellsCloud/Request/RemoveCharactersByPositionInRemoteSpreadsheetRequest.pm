=begin comment

Copyright (c) 2026 Aspose.Cells Cloud
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

=end comment

=cut

package AsposeCellsCloud::Request::RemoveCharactersByPositionInRemoteSpreadsheetRequest;

require 5.6.0;
use strict;
use warnings;
use utf8;
use JSON ;
use Data::Dumper;
use Module::Runtime qw(use_module);
use Log::Any qw($log);
use Date::Parse;
use DateTime;
use File::Basename;

use base ("Class::Accessor", "Class::Data::Inheritable");

__PACKAGE__->mk_classdata('attribute_map' => {});
__PACKAGE__->mk_classdata('method_documentation' => {}); 
__PACKAGE__->mk_classdata('class_documentation' => {});


# new object
sub new { 
    my ($class, %args) = @_; 

	my $self = bless {}, $class;

	foreach my $attribute (keys %{$class->attribute_map}) {
		my $args_key = $class->attribute_map->{$attribute};
		$self->$attribute( $args{ $args_key } );
	}

	return $self;
}  


# Run Operation Request
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.name : (Required) The name of the workbook file to be retrieved.  ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.worksheet : Specify the worksheet of spreadsheet.  ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.range : Specify the worksheet range of spreadsheet.  ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.theFirstNCharacters :   ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.theLastNCharacters :   ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.allCharactersBeforeText :   ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.allCharactersAfterText :   ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.caseSensitive :   ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.folder : (Optional) The folder path where the workbook is stored. The default is null.  ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.storageName : (Optional) The name of the storage if using custom cloud storage. Use default storage if omitted.  ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.region : Spreadsheet region/language setting (e.g., `en-US`, `fr-FR`). Influences number formatting, date parsing, and locale‑specific behavior.  ,
# RemoveCharactersByPositionInRemoteSpreadsheetRequest.password : The password for opening spreadsheet file.   

{
    my $params = {
       'client' =>{
            data_type => 'ApiClient',
            description => 'API Client.',
            required => '0',
       }
    };
    __PACKAGE__->method_documentation->{ 'remove_characters_by_position_in_remote_spreadsheet' } = { 
    	summary => 'Deletes characters from every cell in the target range by position (first/last N, before/after a substring, or between two delimiters) while preserving formulas, formatting and data-validation.',
        params => $params,
        returns => 'CellsCloudResponse',
    };
}

sub run_http_request {
    my ($self, %args) = @_;

    my $client = $args{'client'};

    # parse inputs
    my $_resource_path = 'v4.0/cells/{name}/worksheets/{worksheet}/range/{range}/content/remove/characters-by-position';

    my $_method = 'PUT';
    my $query_params = {};
    my $header_params = {};
    my $form_params = {};


    my $_header_accept = $client->select_header_accept('application/json');
    if ($_header_accept) {
        $header_params->{'Accept'} = $_header_accept;
    }
    $header_params->{'Content-Type'} = $client->select_header_content_type('application/json');
    if(defined $self->name){
        my $_base_variable = "{" . "name" . "}";
        my $_base_value = $client->to_path_value($self->name);
        $_resource_path =~ s/$_base_variable/$_base_value/g;        
    }

    if(defined $self->worksheet){
        my $_base_variable = "{" . "worksheet" . "}";
        my $_base_value = $client->to_path_value($self->worksheet);
        $_resource_path =~ s/$_base_variable/$_base_value/g;        
    }

    if(defined $self->range){
        my $_base_variable = "{" . "range" . "}";
        my $_base_value = $client->to_path_value($self->range);
        $_resource_path =~ s/$_base_variable/$_base_value/g;        
    } 
    if(defined $self->the_first_n_characters){
        $query_params->{'theFirstNCharacters'} = $client->to_query_value($self->the_first_n_characters);      
    }

    if(defined $self->the_last_n_characters){
        $query_params->{'theLastNCharacters'} = $client->to_query_value($self->the_last_n_characters);      
    }

    if(defined $self->all_characters_before_text){
        $query_params->{'allCharactersBeforeText'} = $client->to_query_value($self->all_characters_before_text);      
    }

    if(defined $self->all_characters_after_text){
        $query_params->{'allCharactersAfterText'} = $client->to_query_value($self->all_characters_after_text);      
    }

    if(defined $self->case_sensitive){
        $query_params->{'caseSensitive'} = $client->to_query_value($self->case_sensitive);      
    }

    if(defined $self->folder){
        $query_params->{'folder'} = $client->to_query_value($self->folder);      
    }

    if(defined $self->storage_name){
        $query_params->{'storageName'} = $client->to_query_value($self->storage_name);      
    }

    if(defined $self->region){
        $query_params->{'region'} = $client->to_query_value($self->region);      
    }

    if(defined $self->password){
        $query_params->{'password'} = $client->to_query_value($self->password);      
    } 
    my $_body_data;

 

    # authentication setting, if any
    my $auth_settings = [qw()];

    # make the API Call
    my $response = $client->call_api($_resource_path, $_method, $query_params, $form_params, $header_params, $_body_data, $auth_settings);
    return $response;
}


__PACKAGE__->method_documentation({
     'name' => {
     	datatype => 'string',
     	base_name => 'name',
     	description => '(Required) The name of the workbook file to be retrieved.',
     	format => '',
     	read_only => '',
     		},
     'worksheet' => {
     	datatype => 'string',
     	base_name => 'worksheet',
     	description => 'Specify the worksheet of spreadsheet.',
     	format => '',
     	read_only => '',
     		},
     'range' => {
     	datatype => 'string',
     	base_name => 'range',
     	description => 'Specify the worksheet range of spreadsheet.',
     	format => '',
     	read_only => '',
     		},
     'the_first_n_characters' => {
     	datatype => 'int',
     	base_name => 'theFirstNCharacters',
     	description => '',
     	format => '',
     	read_only => '',
     		},
     'the_last_n_characters' => {
     	datatype => 'int',
     	base_name => 'theLastNCharacters',
     	description => '',
     	format => '',
     	read_only => '',
     		},
     'all_characters_before_text' => {
     	datatype => 'string',
     	base_name => 'allCharactersBeforeText',
     	description => '',
     	format => '',
     	read_only => '',
     		},
     'all_characters_after_text' => {
     	datatype => 'string',
     	base_name => 'allCharactersAfterText',
     	description => '',
     	format => '',
     	read_only => '',
     		},
     'case_sensitive' => {
     	datatype => 'string',
     	base_name => 'caseSensitive',
     	description => '',
     	format => '',
     	read_only => '',
     		},
     'folder' => {
     	datatype => 'string',
     	base_name => 'folder',
     	description => '(Optional) The folder path where the workbook is stored. The default is null.',
     	format => '',
     	read_only => '',
     		},
     'storage_name' => {
     	datatype => 'string',
     	base_name => 'storageName',
     	description => '(Optional) The name of the storage if using custom cloud storage. Use default storage if omitted.',
     	format => '',
     	read_only => '',
     		},
     'region' => {
     	datatype => 'string',
     	base_name => 'region',
     	description => 'Spreadsheet region/language setting (e.g., `en-US`, `fr-FR`). Influences number formatting, date parsing, and locale‑specific behavior.',
     	format => '',
     	read_only => '',
     		},
     'password' => {
     	datatype => 'string',
     	base_name => 'password',
     	description => 'The password for opening spreadsheet file.',
     	format => '',
     	read_only => '',
     		},    
});


__PACKAGE__->attribute_map( {
    'name' => 'name',
    'worksheet' => 'worksheet',
    'range' => 'range',
    'the_first_n_characters' => 'theFirstNCharacters',
    'the_last_n_characters' => 'theLastNCharacters',
    'all_characters_before_text' => 'allCharactersBeforeText',
    'all_characters_after_text' => 'allCharactersAfterText',
    'case_sensitive' => 'caseSensitive',
    'folder' => 'folder',
    'storage_name' => 'storageName',
    'region' => 'region',
    'password' => 'password' 
} );

__PACKAGE__->mk_accessors(keys %{__PACKAGE__->attribute_map});


1;